# -*- coding: utf-8 -*-
# Spelling and Grammar Check, an NVDA add-on.
# Derived from Mute Browse Mode by Josh Kennedy (GNU GPL v2).
# This file is covered by the GNU General Public License, version 2.

"""Reads Word and Outlook's F7 spelling and grammar checker clearly.

This add-on only touches the F7 proofing checker in Microsoft Word and Microsoft
Outlook. It does not mute browse mode, does not touch page loading, and does not
change anything else about how NVDA behaves.

For every issue the checker finds, it announces "misspelled" or "Grammar error", the
word or phrase involved (spoken and then spelled out), and then, if you ask it to, moves
straight into the list of suggestions and reads the first one, spoken and spelled, with
its position in the list. If there are no suggestions at all for an issue, it says so
directly: "No suggestions."

It supports both the classic Word/Outlook spelling and grammar dialog and the newer
Microsoft Editor task pane, and it announces "Spell check is complete" once, at the end
of a check, whether or not anything was found.

Two settings, under NVDA's Preferences, Settings, Spelling and Grammar Check, control
whether the suggestions list is entered automatically and whether misspelled words and
suggestions are read at a slightly slower rate.
"""

import re
import time
from collections import deque

import addonHandler
import api
import controlTypes
import core
import globalPluginHandler
import inputCore
import scriptHandler
import speech
import speech.speech
import textInfos
import config
import wx
import winUser
from gui import guiHelper, settingsDialogs
from logHandler import log

from . import updater

try:
	addonHandler.initTranslation()
except Exception:
	log.debugWarning("Spelling and Grammar Check: translations unavailable", exc_info=True)

if "ngettext" not in globals():
	def ngettext(singular, plural, n):
		return singular if n == 1 else plural

try:
	from speech.commands import CallbackCommand as _CallbackCommand
except Exception:
	_CallbackCommand = None
	log.debugWarning("Spelling and Grammar Check: CallbackCommand unavailable", exc_info=True)

try:
	from speech.commands import EndUtteranceCommand as _EndUtteranceCommand
	from speech.commands import RateCommand as _RateCommand
except Exception:
	_EndUtteranceCommand = None
	_RateCommand = None
	log.debugWarning("Spelling and Grammar Check: speech commands unavailable", exc_info=True)


#: Section this add-on owns in nvda.ini.
CONF_SECTION = "spellGrammarCheck"

#: Running plug-in instance, used only by the input-gesture observer.
_pluginInstance = None

config.conf.spec[CONF_SECTION] = {
	"focusSpellSuggestions": "boolean(default=True)",
	"slowProofingSpeech": "boolean(default=False)",
}


def getFocusSpellSuggestions():
	return bool(config.conf[CONF_SECTION]["focusSpellSuggestions"])


def setFocusSpellSuggestions(enabled):
	config.conf[CONF_SECTION]["focusSpellSuggestions"] = bool(enabled)


def getSlowProofingSpeech():
	return bool(config.conf[CONF_SECTION]["slowProofingSpeech"])


def setSlowProofingSpeech(enabled):
	config.conf[CONF_SECTION]["slowProofingSpeech"] = bool(enabled)


### Recognising Word and Outlook

#: Executable names, lower case and without the extension, that count as Outlook.
_OUTLOOK_APP_NAMES = frozenset((
	"outlook",  # classic desktop Outlook
	"olk",  # the new Outlook for Windows
	"hxoutlook",  # Outlook / Mail from the Microsoft Store
	"hxmail",
	"msoutlook",
))

_WORD_APP_NAMES = frozenset(("winword",))

#: Window classes Outlook gives its own top level windows. These matter because modern
#: Outlook can render the message body inside an embedded Edge web view, whose objects
#: otherwise look like they belong to a different, unrelated program entirely.
_OUTLOOK_WINDOW_CLASSES = frozenset((
	"rctrl_renwnd32",  # classic Outlook: the main window and every message window
	"Outlook Host",  # the new Outlook for Windows
))


def _appNameOf(obj):
	try:
		return (obj.appModule.appName or "").lower()
	except Exception:
		return ""


def _windowClassOf(obj):
	try:
		return getattr(obj, "windowClassName", "") or ""
	except Exception:
		return ""


def _isOutlook(obj):
	return obj is not None and _appNameOf(obj) in _OUTLOOK_APP_NAMES


def _isWord(obj):
	return obj is not None and _appNameOf(obj) in _WORD_APP_NAMES


def _rootWindowOf(obj):
	"""The handle of the top level window C{obj} sits in, or 0 for none."""
	try:
		hwnd = obj.windowHandle
	except Exception:
		return 0
	if not hwnd:
		return 0
	try:
		return winUser.getAncestor(hwnd, getattr(winUser, "GA_ROOT", 2)) or 0
	except Exception:
		return 0


def _rootWindowOfObjectOrParent(obj):
	"""Find a top-level window even when a UIA text child has no window handle.

	Outlook can leave focus on the final punctuation TextInfo object while F7 opens
	Editor. That transient object reports a zero window handle, although one of its
	accessible parents still belongs to the compose window.
	"""
	current = obj
	for _step in range(12):
		if current is None:
			break
		window = _rootWindowOf(current)
		if window:
			return window
		try:
			current = current.parent
		except Exception:
			break
	return 0


def _processIDOf(obj):
	try:
		return obj.processID
	except Exception:
		return None


def _appNameOfProcess(processID):
	if not processID:
		return ""
	try:
		import appModuleHandler

		return (appModuleHandler.getAppModuleFromProcessID(processID).appName or "").lower()
	except Exception:
		return ""


def _isInOutlookWindow(obj):
	"""Whether C{obj} is anywhere inside a window belonging to Microsoft Outlook.

	Outlook renders more and more of itself, including in the newest builds the message
	body, in an embedded Edge web view, which belongs to a different process entirely.
	The top level window is the one thing that is still Outlook's own, whatever is
	embedded inside it.
	"""
	if _isOutlook(obj):
		return True
	root = _rootWindowOf(obj)
	if not root:
		return False
	try:
		if winUser.getClassName(root) in _OUTLOOK_WINDOW_CLASSES:
			return True
		processID = winUser.getWindowThreadProcessID(root)[0]
	except Exception:
		log.debugWarning("Spelling and Grammar Check: could not read the top level window", exc_info=True)
		return False
	if not processID or processID == _processIDOf(obj):
		return False
	return _appNameOfProcess(processID) in _OUTLOOK_APP_NAMES


def _isOfficeProofingWindow(obj):
	"""Whether an object is in the Word proofing UI used by Word or Outlook."""
	return _isWord(obj) or _isInOutlookWindow(obj)


def _members(enum, names):
	"""The named members of C{enum} that this NVDA actually has."""
	return frozenset(m for m in (getattr(enum, name, None) for name in names) if m is not None)


### The speech gate
#
# Two small context managers, used only around the proofing announcements below.
# _hardMute lets NVDA's own handling of a focus event run, for its property caches and
# braille, without letting it reach the synthesiser; _ownSpeech lets this add-on's own
# replacement announcement past that mute.

_muteDepth = 0
_bypassDepth = 0


class _hardMute:
	"""Context manager dropping speech for the duration of one call."""

	def __enter__(self):
		global _muteDepth
		_muteDepth += 1
		return self

	def __exit__(self, *exc):
		global _muteDepth
		_muteDepth = max(0, _muteDepth - 1)
		return False


class _ownSpeech:
	"""Context manager letting the add-on's own announcements past the mute."""

	def __enter__(self):
		global _bypassDepth
		_bypassDepth += 1
		return self

	def __exit__(self, *exc):
		global _bypassDepth
		_bypassDepth = max(0, _bypassDepth - 1)
		return False


def _isGated():
	if _bypassDepth > 0:
		return False
	return _muteDepth > 0


### Cleaning up Outlook's own proofing announcements

_OUTLOOK_EDITOR_NOISE_RE = re.compile(r"^(left\s+(?:ali(?:g)?ned|justified))$", re.IGNORECASE)

_suggestionsLabelUntil = 0.0
_suggestionsLabelKey = None
_lastSuggestionSpeechKey = None
_lastSuggestionSpeechUntil = 0.0

#: How long a just-spoken list label or suggestion counts as "already said".
_PROOFING_DUPLICATE_WINDOW = 2.5


def _suggestionContainerIdentity(container, fallback=None):
	"""Stable identity for a suggestion list whose accessible wrapper may be replaced."""
	for candidate in (container, fallback):
		if candidate is None:
			continue
		try:
			root = _rootWindowOf(candidate)
		except Exception:
			root = None
		if root:
			return ("root", root)
	try:
		window = getattr(container, "windowHandle", None)
	except Exception:
		window = None
	return ("window", window) if window else ("object", id(container or fallback))


def _filterOutlookEditorNoise(args, kwargs):
	"""Drop or clean up Word-editor labels that add no information while proofing."""
	global _suggestionsLabelUntil, _suggestionsLabelKey
	try:
		focus = api.getFocusObject()
	except Exception:
		return args, kwargs
	if not _isInOutlookWindow(focus):
		return args, kwargs
	if args:
		sequence = args[0]
	elif "speechSequence" in kwargs:
		sequence = kwargs["speechSequence"]
	else:
		return args, kwargs
	if not isinstance(sequence, list):
		return args, kwargs
	spokenText = " ".join(item for item in sequence if isinstance(item, str))
	if re.search(r"left\s+(?:ali(?:g)?ned|justified)", spokenText, re.IGNORECASE):
		pattern = re.compile(r"left\s+(?:ali(?:g)?ned|justified)", re.IGNORECASE)
		filtered = []
		for item in sequence:
			if not isinstance(item, str):
				filtered.append(item)
				continue
			cleaned = pattern.sub("", item).strip()
			if cleaned:
				filtered.append(cleaned)
	elif re.search(r"suggestions?\s*:\s*list.*?alt\s*\+?\s*n", spokenText, re.IGNORECASE):
		container = _suggestionListOf(focus)
		labelKey = _suggestionContainerIdentity(container, focus)
		alreadySaid = (
			labelKey == _suggestionsLabelKey
			and time.monotonic() < _suggestionsLabelUntil
		)
		label = "" if alreadySaid else _("Suggestions list box")
		if not alreadySaid:
			_suggestionsLabelKey = labelKey
			_suggestionsLabelUntil = time.monotonic() + _PROOFING_DUPLICATE_WINDOW
		pattern = re.compile(r"suggestions?\s*:\s*list.*?alt\s*\+?\s*n", re.IGNORECASE)
		match = pattern.search(spokenText)
		filtered = []
		cursor = 0
		replacementInserted = False
		for item in sequence:
			if not isinstance(item, str):
				filtered.append(item)
				continue
			start = cursor
			end = start + len(item)
			cursor = end + 1
			if match is None or end <= match.start() or start >= match.end():
				if item:
					filtered.append(item)
				continue
			before = item[:max(0, match.start() - start)].rstrip()
			after = item[max(0, match.end() - start):].lstrip()
			parts = []
			if before:
				parts.append(before)
			if not replacementInserted and label:
				parts.append(label)
			replacementInserted = True
			if after:
				parts.append(after)
			if parts:
				filtered.append(" ".join(parts))
	else:
		filtered = [
			item
			for item in sequence
			if not (isinstance(item, str) and _OUTLOOK_EDITOR_NOISE_RE.match(item.strip()))
		]
	if args:
		return (filtered,) + tuple(args[1:]), kwargs
	kwargs = dict(kwargs)
	kwargs["speechSequence"] = filtered
	return args, kwargs


def _caretHasGrammarError(obj):
	"""Whether Outlook marks the editing caret as a grammar error."""
	if not _isOutlookMessageBody(obj):
		return False
	for unit in (textInfos.UNIT_CHARACTER, textInfos.UNIT_WORD):
		try:
			info = obj.makeTextInfo(textInfos.POSITION_CARET)
			info.expand(unit)
			fields = info.getTextWithFields()
		except Exception:
			continue
		for item in fields:
			if not isinstance(item, textInfos.FieldCommand) or item.command != "formatChange":
				continue
			field = item.field
			try:
				if field.get("invalid-grammar") or field.get("invalidGrammar"):
					return True
			except Exception:
				continue
	return False


def _correctEditingGrammarLabel(args, kwargs):
	"""Replace NVDA's spelling label when an editor exposes a grammar annotation."""
	try:
		focus = api.getFocusObject()
	except Exception:
		return args, kwargs
	if not _caretHasGrammarError(focus):
		return args, kwargs
	if args:
		sequence = args[0]
	elif "speechSequence" in kwargs:
		sequence = kwargs["speechSequence"]
	else:
		return args, kwargs
	if not isinstance(sequence, list):
		return args, kwargs
	labels = {"misspelled", _("misspelled").strip().casefold()}
	grammarLabel = _("Grammar error")
	pattern = re.compile(
		r"(?<!\w)(?:%s)(?!\w)" % "|".join(re.escape(label) for label in labels if label),
		re.IGNORECASE,
	)
	filtered = []
	for item in sequence:
		if not isinstance(item, str):
			filtered.append(item)
		else:
			filtered.append(pattern.sub(grammarLabel, item))
	if args:
		return (filtered,) + tuple(args[1:]), kwargs
	kwargs = dict(kwargs)
	kwargs["speechSequence"] = filtered
	return args, kwargs


### Recognising the Outlook message body

_PLAIN_TEXT_BODY_CLASS = "RichEdit20W"
_PLAIN_TEXT_BODY_CONTROL_ID = 8224
_OUTLOOK_BODY_WINDOW_CLASSES = frozenset(("_WwG",))
_BODY_NAMES = frozenset(("message", "message body"))
_BODY_ROLES = _members(controlTypes.Role, ("DOCUMENT", "EDITABLETEXT"))
_STATE_READONLY = getattr(controlTypes.State, "READONLY", None)
_STATE_UNAVAILABLE = getattr(controlTypes.State, "UNAVAILABLE", None)
_STATE_MULTILINE = getattr(controlTypes.State, "MULTILINE", None)


def _hasState(states, state):
	return state is not None and state in states


def _isOutlookMessageBody(obj):
	"""Whether the focus has landed in an Outlook message body that can be typed into."""
	if not _isInOutlookWindow(obj):
		return False
	if getattr(obj, "role", None) not in _BODY_ROLES:
		return False
	try:
		states = set(obj.states or ())
	except Exception:
		states = set()
	if _hasState(states, _STATE_READONLY) or _hasState(states, _STATE_UNAVAILABLE):
		return False

	windowClass = _windowClassOf(obj)
	viewer = getattr(obj, "isReadonlyViewer", None)
	if viewer is not None:
		return not viewer and windowClass in _OUTLOOK_BODY_WINDOW_CLASSES
	if windowClass in _OUTLOOK_BODY_WINDOW_CLASSES:
		return True
	if (
		windowClass == _PLAIN_TEXT_BODY_CLASS
		and getattr(obj, "windowControlID", None) == _PLAIN_TEXT_BODY_CONTROL_ID
	):
		return True
	if not _hasState(states, _STATE_MULTILINE):
		return False
	return (getattr(obj, "name", "") or "").strip().lower() in _BODY_NAMES


### The Word and Outlook spelling and grammar checker
#
# The classic F7 window is an ordinary Win32 dialog put on screen by the Word engine
# that renders the message, and the box in it that shows the mistake is a small Word
# editing surface with a window class of _WwN or _WwO and control id 18. Current Word
# and Outlook builds instead use the Microsoft Editor task pane, a UI Automation
# interface handled separately below.

_WORD_DIALOG_WINDOW_CLASSES = frozenset(("_WwN", "_WwO"))
_SPELL_ERROR_CONTROL_ID = 18

#: How much slower than usual a misspelled word or suggestion is said and spelled. 0.8
#: is a fifth slower.
_SPELL_RATE_MULTIPLIER = 0.8

#: Longer than this and whatever we found is not one misspelled word or short phrase.
_MAX_WORD_LENGTH = 60

#: How long to wait before looking again for the word the dialog is asking about, in
#: milliseconds, and how many times.
_SPELL_RETRY_DELAYS = (120, 300, 600, 1000, 1600)

# Word/Outlook can say a check is "complete" before its own background proofing
# pass has finished analyzing text that was just typed, if F7 is pressed right
# after typing. Silently retry for a while rather than trusting that first answer.
_PREMATURE_COMPLETE_RETRY_MS = 3000
_PREMATURE_COMPLETE_MAX_ATTEMPTS = 8
_PREMATURE_COMPLETE_WINDOW = 20.0

#: A proofing dialog returning focus to the document does so immediately. This deadline
#: prevents an abandoned session being mistaken for completion during a later check.
_PROOFING_RETURN_WINDOW = 5.0

_SUGGESTION_CONTAINER_ROLES = _members(controlTypes.Role, ("LIST", "LISTBOX", "TREEVIEW", "TABLE"))
_SUGGESTION_NAMED_CONTAINER_ROLES = _members(controlTypes.Role, ("GROUPING", "PANE", "PROPERTYPAGE"))
_SUGGESTION_ITEM_ROLES = _members(controlTypes.Role, ("LISTITEM", "SUGGESTION", "BUTTON", "RADIOBUTTON"))

#: Punctuation taken off either end of what the checker has selected.
_WORD_EDGE_PUNCTUATION = " \t\r\n.,;:!?\"'`()[]{}<>…«»„“”‘’-–—/\\|*_+=@#$%^&~"

#: Roles that are, or are not, a dialog.
_DIALOG_ROLES = _members(controlTypes.Role, ("DIALOG", "ALERT", "PROPERTYPAGE", "OPTIONPANE"))
_NOT_A_DIALOG_ROLES = _members(controlTypes.Role, ("APPLICATION", "FRAME", "DESKTOP"))
_DIALOG_WINDOW_CLASS = "#32770"
_OFFICE_DIALOG_WINDOW_PREFIX = "bosa_sdm"
_DIALOG_WALK_LIMIT = 6
_DIALOG_SCAN_DEPTH = 4
_DIALOG_SCAN_LIMIT = 60

# Every phrase _classicProofingKind treats as a decisive grammar/spelling signal,
# combined so _dialogText can stop walking the dialog as soon as one shows up
# instead of always reading every descendant first. Keep this in sync with the
# marker tuples inside _classicProofingKind below.
_CLASSIC_PROOFING_STOP_MARKERS = (
	"possible word choice error", "grammatical error", "grammar error",
	"ignore rule", "next sentence",
	"not in dictionary", "spelling error", "misspelled", "unknown word",
	"word choice", "punctuation", "clarity",
	"conciseness", "refinement", "formal writing",
	"some words are similar but are used differently",
	"agreement", "subject-verb", "subject verb", "verb form", "verb tense",
	"passive voice", "fragment", "run-on", "run on", "wordiness", "wordy",
	"double negative", "article usage", "pronoun", "capitalization",
	"contraction", "cliche", "cliché", "jargon", "preposition", "modifier",
	"comma", "hyphenation", "gender-specific", "consider revising",
	"singular or plural", "singular and plural", "singular/plural",
	"plural or singular", "sticking to singular",
)

_OK_BUTTON_NAMES = frozenset(("ok", "&ok"))

#: What NVDA calls the state a disabled control is in, lower case.
try:
	_UNAVAILABLE_TEXT = (_STATE_UNAVAILABLE.displayString or "").strip().lower() or None
except Exception:
	_UNAVAILABLE_TEXT = None
	log.debugWarning("Spelling and Grammar Check: no name for the unavailable state", exc_info=True)


def _looksLikeProofingText(text, allowPhrase=False):
	"""Whether C{text} is a proofing issue, rather than a sentence or nothing."""
	word = (text or "").strip()
	limit = 120 if allowPhrase else _MAX_WORD_LENGTH
	if not word or len(word) > limit:
		return False
	if any(character in "\r\n" for character in word):
		return False
	parts = word.split()
	if len(parts) > 1 and (not allowPhrase or len(parts) > 10):
		return False
	return any(character.isalpha() for character in word)


def _isSuggestionContainer(obj):
	role = getattr(obj, "role", None)
	if role in _SUGGESTION_CONTAINER_ROLES:
		return True
	if role not in _SUGGESTION_NAMED_CONTAINER_ROLES:
		return False
	try:
		name = (getattr(obj, "name", "") or "").strip().lower()
	except Exception:
		return False
	labels = ("suggestion", _("Suggestion").strip().lower())
	return any(name.startswith(label) for label in labels if label)


def _suggestionListOf(obj):
	"""The suggestions list containing C{obj}, or C{None}."""
	current = obj
	for _step in range(8):
		if current is None:
			return None
		if _isSuggestionContainer(current):
			return current
		try:
			current = current.parent
		except Exception:
			return None
	return None


def _directSuggestionText(obj):
	"""Best short text exposed directly by one proofing-suggestion object."""
	if obj is None:
		return ""
	for attribute in ("name", "value", "description"):
		try:
			text = " ".join(str(getattr(obj, attribute, "") or "").split())
		except Exception:
			continue
		if text and len(text) <= 240:
			return text
	try:
		text = " ".join((obj.makeTextInfo(textInfos.POSITION_ALL).text or "").split())
		if text and len(text) <= 240:
			return text
	except Exception:
		pass
	# Seen live on a 2026.1.1 build: a grammar suggestion's list item exposes no
	# name, value, description, or IAccessible text at all (unlike misspelling
	# suggestions, which do). The word is still visible on screen, so read it the
	# same way this add-on already reads the boxed misspelled word: from the
	# pixels within the item's own rectangle.
	try:
		from displayModel import DisplayModelTextInfo

		text = " ".join((DisplayModelTextInfo(obj, textInfos.POSITION_ALL).text or "").split())
	except Exception:
		return ""
	return text if len(text) <= 240 else ""


_MAX_CHILDREN_SCANNED = 100


def _suggestionDescendants(container):
	"""Yield a bounded breadth-first walk below a small suggestions container."""
	try:
		queue = [(child, 1) for child in (container.children or ())]
	except Exception:
		return
	seen = set()
	visited = 0
	while queue and visited < _MAX_CHILDREN_SCANNED:
		obj, depth = queue.pop(0)
		identity = id(obj)
		if identity in seen:
			continue
		seen.add(identity)
		visited += 1
		yield obj
		if depth >= 4:
			continue
		try:
			queue.extend((child, depth + 1) for child in (obj.children or ()))
		except Exception:
			pass


def _accessibleSuggestionText(obj):
	"""Best short text exposed by a suggestion or one of its text descendants."""
	text = _directSuggestionText(obj)
	if text:
		return text
	for descendant in _suggestionDescendants(obj):
		text = _directSuggestionText(descendant)
		if text:
			return text
	return ""


def _debugDumpObject(prefix, obj, extra=""):
	"""Log a compact snapshot of obj's accessible properties, for one-time diagnosis."""
	try:
		role = getattr(obj, "role", None)
		name = getattr(obj, "name", None)
		value = getattr(obj, "value", None)
		description = getattr(obj, "description", None)
		automationId = _automationIdOf(obj)
		windowClass = _windowClassOf(obj)
		try:
			states = tuple(str(s) for s in (obj.states or ()))
		except Exception:
			states = ()
		log.debug(
			"Spelling and Grammar Check: %s role=%r name=%r value=%r description=%r "
			"automationId=%r windowClass=%r states=%r%s"
			% (prefix, role, name, value, description, automationId, windowClass, states, extra)
		)
	except Exception:
		log.debug("Spelling and Grammar Check: %s could not be inspected" % prefix, exc_info=True)


def _debugDumpSuggestionContainer(container):
	"""Log the container and up to 15 descendants, to see what the item text really is."""
	_debugDumpObject("suggestion container", container)
	try:
		count = 0
		for child in _suggestionDescendants(container):
			count += 1
			if count > 15:
				log.debug("Spelling and Grammar Check: suggestion container has more than 15 descendants, stopping dump")
				break
			_debugDumpObject("  suggestion descendant #%d" % count, child)
		if count == 0:
			log.debug("Spelling and Grammar Check: suggestion container has NO descendants at all")
	except Exception:
		log.debug("Spelling and Grammar Check: could not walk suggestion descendants", exc_info=True)


def _selectedSuggestionText(container):
	selectedStates = tuple(
		state
		for state in (
			getattr(controlTypes.State, "SELECTED", None),
			getattr(controlTypes.State, "FOCUSED", None),
		)
		if state is not None
	)
	try:
		activeChild = getattr(container, "activeChild", None)
	except Exception:
		activeChild = None
	if activeChild is not None:
		text = _accessibleSuggestionText(activeChild)
		if text:
			return text
	fallback = ""
	for child in _suggestionDescendants(container):
		if getattr(child, "role", None) not in _SUGGESTION_ITEM_ROLES:
			continue
		text = _accessibleSuggestionText(child)
		if text and not fallback:
			fallback = text
		try:
			states = child.states or ()
			isSelected = any(state in states for state in selectedStates)
		except Exception:
			isSelected = False
		if isSelected and text:
			return text
	return fallback


def _selectedItemName(obj):
	try:
		return " ".join((getattr(obj, "name", "") or "").split())
	except Exception:
		return ""


def _suggestionItemName(obj, container):
	if obj is not container and (
		getattr(obj, "role", None) in _SUGGESTION_ITEM_ROLES
		or _suggestionListOf(obj) is container
	):
		text = _accessibleSuggestionText(obj)
		if text:
			return text
	# _selectedItemName(container) is deliberately NOT used as a further fallback
	# here: it returns the suggestions LIST's own label (e.g. "Suggestions:"),
	# not a suggestion. Confirmed live on a 2026.1.1 build, where that label was
	# being read out, letter by letter, as though it were the suggested word.
	# Returning "" instead lets the caller fall back to announcing the object
	# itself rather than misreporting the list's heading as a suggestion.
	return _selectedSuggestionText(container)


def _suggestionPosition(obj, container):
	"""Return the focused suggestion's one-based position and list size."""
	try:
		info = obj.positionInfo or {}
		index = info.get("indexInGroup")
		total = info.get("similarItemsInGroup")
		if index and total:
			return int(index), int(total)
	except Exception:
		pass
	try:
		children = list(container.children or ())
		if obj in children:
			return children.index(obj) + 1, len(children)
	except Exception:
		pass
	return None


def _automationIdOf(obj):
	for attribute in ("UIAAutomationId", "automationId", "automationID"):
		try:
			value = getattr(obj, attribute, "") or ""
		except Exception:
			continue
		if value:
			return str(value)
	return ""


### The Microsoft Editor task pane

_modernEditorEventRoot = None
_modernEditorEventRootHandle = 0
_modernEditorWindowScanUntil = 0.0
_modernEditorWindowScanHandle = 0
_modernEditorSummaryChoiceKey = None
_modernEditorSummaryChoiceUntil = 0.0
_modernEditorIssueKey = None
_modernEditorSpeechUntil = 0.0
_modernEditorAnnouncedIssueKey = None
_classicProofingPunctuationLaunchUntil = 0.0
_legacySdmIssueKey = None

# An Outlook F7 pass can briefly focus an empty shell window before it shows a
# real proofing result. It is repeated for each automatic retry and contains no
# useful proofing information.
_outlookProofingShellUntil = 0.0
_suppressSuggestionsListUntil = 0.0
_outlookProofingReturnUntil = 0.0

_MODERN_EDITOR_RETRY_DELAYS = (80, 180, 350, 650, 1000, 1600)


def _rememberModernEditorRoot(root):
	global _modernEditorEventRoot, _modernEditorEventRootHandle
	_modernEditorEventRoot = root
	_modernEditorEventRootHandle = _rootWindowOf(root)


def _modernEditorRoot(obj):
	"""The modern Office Editor subtree belonging to C{obj}'s current window."""
	direct = _modernEditorAncestorRoot(obj)
	if direct is not None:
		_rememberModernEditorRoot(direct)
		return direct
	if not _isOfficeProofingWindow(obj):
		return None
	if time.monotonic() >= _modernEditorWindowScanUntil:
		return None
	objectWindow = _rootWindowOfObjectOrParent(obj)
	if not objectWindow and _modernEditorWindowScanHandle:
		objectWindow = _modernEditorWindowScanHandle
	if (
		_modernEditorEventRoot is not None
		and objectWindow
		and objectWindow == _modernEditorEventRootHandle
	):
		return _modernEditorEventRoot
	if _modernEditorWindowScanHandle:
		try:
			if objectWindow != _modernEditorWindowScanHandle:
				return None
		except Exception:
			return None
	scanRoot = _accessibleWindowRoot(obj, objectWindow)
	if scanRoot is None:
		return None
	for item in _boundedDescendants(scanRoot, limit=240, depthLimit=10):
		candidate = _modernEditorAncestorRoot(item)
		if candidate is None:
			continue
		if (
			_modernEditorIssue(candidate) is not None
			or _modernEditorCardPending(candidate)
			or _modernEditorSummaryChoice(candidate) is not None
		):
			_rememberModernEditorRoot(candidate)
			return candidate
	return None


def _modernEditorAncestorRoot(obj):
	"""The nearest modern Editor subtree containing C{obj}, without a window scan."""
	current = obj
	for _step in range(12):
		if current is None:
			return None
		identifier = _automationIdOf(current)
		try:
			name = (getattr(current, "name", "") or "").strip().lower()
		except Exception:
			name = ""
		if identifier == "DrillInPane_EditorCustomProps" or name == "editor":
			if identifier != "DrillInPane_EditorCustomProps":
				return current
		try:
			current = current.parent
		except Exception:
			return None
	return None


def _accessibleWindowRoot(obj, window=0):
	"""Highest accessible ancestor that still belongs to C{obj}'s top-level window."""
	if obj is None:
		return None
	if not window:
		window = _rootWindowOfObjectOrParent(obj)
	current = obj
	root = obj
	for _step in range(24):
		try:
			parent = current.parent
		except Exception:
			break
		if parent is None:
			break
		if window:
			try:
				parentWindow = _rootWindowOf(parent)
			except Exception:
				parentWindow = None
			if parentWindow and parentWindow != window:
				break
		root = parent
		current = parent
	return root


def _boundedDescendants(root, limit=250, depthLimit=10):
	"""Yield a breadth-first, bounded accessible subtree."""
	try:
		queue = deque((child, 1) for child in (root.children or ()))
	except Exception:
		return
	seen = set()
	while queue and len(seen) < limit:
		item, depth = queue.popleft()
		identity = id(item)
		if identity in seen:
			continue
		seen.add(identity)
		yield item
		if depth >= depthLimit:
			continue
		try:
			queue.extend((child, depth + 1) for child in (item.children or ()))
		except Exception:
			pass


def _modernEditorIssue(root):
	"""Return the issue text, kind, subtype, and concrete suggestion buttons."""
	objects = list(_boundedDescendants(root))
	issueKind = ""
	propertiesGroup = None
	suggestions = []
	for item in objects:
		identifier = _automationIdOf(item)
		if identifier == "DrillInPane_Title":
			try:
				issueKind = (getattr(item, "name", "") or "").strip()
			except Exception:
				pass
		elif identifier == "DrillInPane_EditorCustomProps":
			propertiesGroup = item
		elif re.match(r"^DrillInPane_Suggestion\d+$", identifier):
			suggestions.append(item)
	if propertiesGroup is None or not suggestions:
		return None
	rawErrorText = ""
	textRoles = _members(controlTypes.Role, ("STATICTEXT", "TEXT"))
	for item in _boundedDescendants(propertiesGroup, limit=40, depthLimit=5):
		if textRoles and getattr(item, "role", None) not in textRoles:
			continue
		try:
			candidate = " ".join((getattr(item, "name", "") or "").split())
		except Exception:
			candidate = ""
		if candidate:
			rawErrorText = candidate
			break
	if not rawErrorText:
		return None
	issueType = ""
	try:
		siblings = list(propertiesGroup.parent.children or ())
		propertiesIndex = siblings.index(propertiesGroup)
		for sibling in reversed(siblings[:propertiesIndex]):
			if _automationIdOf(sibling):
				continue
			if textRoles and getattr(sibling, "role", None) not in textRoles:
				continue
			candidate = " ".join((getattr(sibling, "name", "") or "").split())
			if candidate:
				issueType = candidate
				break
	except Exception:
		pass
	if issueType.casefold() == "possible word choice error" or rawErrorText.casefold().startswith(
		"grammatical error,"
	):
		issueKind = "Grammar"
	errorText = _modernEditorMarkedText(rawErrorText, propertiesGroup)
	return errorText, issueKind, issueType, suggestions


def _modernEditorCardPending(root):
	"""Whether Editor has an issue card whose suggestion buttons are still loading."""
	propertiesGroup = None
	hasIssueTitle = False
	for item in _boundedDescendants(root, limit=250, depthLimit=10):
		identifier = _automationIdOf(item)
		if identifier == "DrillInPane_Title":
			try:
				hasIssueTitle = bool(" ".join((getattr(item, "name", "") or "").split()))
			except Exception:
				pass
		elif identifier == "DrillInPane_EditorCustomProps":
			propertiesGroup = item
	if hasIssueTitle:
		return True
	if propertiesGroup is None:
		return False
	textRoles = _members(controlTypes.Role, ("STATICTEXT", "TEXT"))
	for item in _boundedDescendants(propertiesGroup, limit=40, depthLimit=5):
		if textRoles and getattr(item, "role", None) not in textRoles:
			continue
		try:
			if " ".join((getattr(item, "name", "") or "").split()):
				return True
		except Exception:
			pass
	return False


def _modernEditorSummaryChoice(root):
	"""The first non-empty correction category on Editor's summary screen."""
	grammarChoice = None
	spellingChoice = None
	for item in _boundedDescendants(root, limit=250, depthLimit=10):
		identifier = _automationIdOf(item).casefold()
		if identifier not in ("spelling", "grammar"):
			continue
		try:
			name = " ".join((getattr(item, "name", "") or "").split())
		except Exception:
			name = ""
		match = re.search(r"\b([1-9]\d*)\s+issues?\b", name, re.IGNORECASE)
		if not match:
			continue
		if identifier == "grammar":
			grammarChoice = item
		else:
			spellingChoice = item
	if time.monotonic() < _classicProofingPunctuationLaunchUntil:
		return grammarChoice or spellingChoice
	return spellingChoice or grammarChoice


def _activateModernEditorSummaryChoice(choice):
	try:
		choice.doAction()
	except Exception:
		log.debugWarning("Spelling and Grammar Check: could not open the Editor correction category", exc_info=True)


def _openModernEditorSummaryChoice(root, choice):
	global _modernEditorSummaryChoiceKey, _modernEditorSummaryChoiceUntil
	now = time.monotonic()
	if now >= _modernEditorWindowScanUntil:
		return False
	try:
		rootIdentity = _rootWindowOf(root) or id(root)
	except Exception:
		rootIdentity = id(root)
	try:
		choiceName = " ".join((getattr(choice, "name", "") or "").split())
	except Exception:
		choiceName = ""
	key = (rootIdentity, _automationIdOf(choice).casefold(), choiceName.casefold())
	if key == _modernEditorSummaryChoiceKey and now < _modernEditorSummaryChoiceUntil:
		return True
	_modernEditorSummaryChoiceKey = key
	_modernEditorSummaryChoiceUntil = now + 3.5
	core.callLater(0, _activateModernEditorSummaryChoice, choice)
	return True


def _modernEditorMarkedText(rawText, propertiesGroup):
	"""Extract the marked word or phrase from Editor's combined accessible label."""
	text = " ".join((rawText or "").split())
	parts = [part.strip() for part in text.split(",", 3)]
	labels = {
		"misspelled",
		_("misspelled").strip().casefold(),
		"spelling error",
		"grammar error",
		_("Grammar error").strip().casefold(),
		"grammatical error",
	}
	if len(parts) >= 2 and parts[0].casefold() in labels:
		candidate = parts[1].strip(_WORD_EDGE_PUNCTUATION)
		if _looksLikeProofingText(candidate, allowPhrase=True):
			return candidate
	try:
		selection = (_messageSelection(propertiesGroup) or "").strip().strip(_WORD_EDGE_PUNCTUATION)
	except Exception:
		selection = ""
	if (
		_looksLikeProofingText(selection, allowPhrase=True)
		and selection.casefold() in text.casefold()
	):
		return selection
	return text


def _modernEditorSuggestionText(obj):
	try:
		return " ".join((getattr(obj, "name", "") or "").split())
	except Exception:
		return ""


def _modernEditorGrammarSpeech(errorText, issueType=""):
	"""Split Editor's explanation from the incorrect text for a natural pause."""
	text = " ".join((errorText or "").split())
	if issueType.casefold() == "possible word choice error":
		parts = [part.strip() for part in text.split(",", 2)]
		if len(parts) >= 2 and parts[1]:
			incorrect = parts[1].strip(" .:")
			if incorrect:
				return [
					_("Some words are similar but are used differently:"),
					incorrect + ".",
				]
	if ":" not in text:
		return [text] if text else []
	explanation, incorrect = text.rsplit(":", 1)
	explanation = explanation.strip(" .:")
	incorrect = incorrect.strip(" .:")
	if not explanation or not incorrect:
		return [text] if text else []
	return [explanation + ".", incorrect + "."]


def _modernEditorReplacementText(suggestionText):
	text = " ".join((suggestionText or "").split())
	if not text:
		return ""
	return text.split(",", 1)[0].strip()


def _modernEditorSuggestionForObject(obj, suggestions):
	current = obj
	for _step in range(8):
		if current is None:
			break
		if current in suggestions:
			return current
		try:
			current = current.parent
		except Exception:
			break
	return None


def _focusObjectKey(obj):
	if obj is None:
		return None
	try:
		return (
			getattr(obj, "windowHandle", None),
			_windowClassOf(obj),
			_automationIdOf(obj),
			getattr(obj, "role", None),
		)
	except Exception:
		return None


def _modernEditorSuggestionForKeyboardFocus(obj, suggestions):
	direct = _modernEditorSuggestionForObject(obj, suggestions)
	if direct is not None:
		return direct
	focusKey = _focusObjectKey(obj)
	hasDistinctFocusKey = bool(focusKey and (focusKey[1] or focusKey[2]))
	for suggestion in suggestions:
		current = suggestion
		for _step in range(3):
			if current is obj or (
				hasDistinctFocusKey
				and _focusObjectKey(current) == focusKey
			):
				return suggestion
			try:
				current = current.parent
			except Exception:
				break
	return None


def _modernEditorIsSuggestionArea(obj, suggestions):
	if obj is None:
		return False
	try:
		identifier = _automationIdOf(obj).casefold()
		name = " ".join((getattr(obj, "name", "") or "").split()).casefold()
	except Exception:
		identifier = name = ""
	if "suggestion" in identifier or name.startswith("suggestion"):
		return True
	for suggestion in suggestions:
		current = suggestion
		for _step in range(8):
			if current is obj:
				return True
			try:
				current = current.parent
			except Exception:
				break
	return False


def _modernEditorIdentity(obj):
	root = _modernEditorRoot(obj)
	if root is None:
		return None
	try:
		return _rootWindowOf(root) or getattr(root, "windowHandle", None) or id(root)
	except Exception:
		return id(root)


def _sameModernEditorRoot(left, right):
	if left is None or right is None:
		return False
	if left is right:
		return True
	try:
		leftWindow = _rootWindowOf(left)
		rightWindow = _rootWindowOf(right)
		if leftWindow and rightWindow:
			return leftWindow == rightWindow
	except Exception:
		pass
	return False


def _outlookMessageBodyAncestor(obj):
	current = obj
	for _step in range(10):
		if current is None:
			break
		if _isOutlookMessageBody(current):
			return current
		try:
			current = current.parent
		except Exception:
			break
	return None


def _modernEditorFocusStillRelevant(root, focus, launchFocus):
	"""Whether a delayed automatic focus move still belongs to this issue."""
	directRoot = _modernEditorAncestorRoot(focus)
	if directRoot is not None:
		return _sameModernEditorRoot(directRoot, root)
	currentKey = _focusObjectKey(focus)
	launchKey = _focusObjectKey(launchFocus)
	if currentKey is not None and launchKey is not None and currentKey == launchKey:
		return True
	currentBody = _outlookMessageBodyAncestor(focus)
	launchBody = _outlookMessageBodyAncestor(launchFocus)
	if currentBody is None or launchBody is None:
		return False
	currentWindow = _rootWindowOfObjectOrParent(currentBody)
	launchWindow = _rootWindowOfObjectOrParent(launchBody)
	return bool(currentWindow and currentWindow == launchWindow)


def _modernEditorKeyboardFocusTarget(selected, root):
	"""Return the focusable wrapper that owns an Editor suggestion."""
	focusableState = getattr(controlTypes.State, "FOCUSABLE", None)
	current = selected
	for _step in range(8):
		if current is None:
			break
		try:
			states = current.states or ()
		except Exception:
			states = ()
		if (
			focusableState is not None
			and focusableState in states
			and callable(getattr(current, "setFocus", None))
		):
			return current
		if current is root:
			break
		try:
			current = current.parent
		except Exception:
			break
	return selected


def _focusModernEditorSuggestion(root, selected, launchFocus, attempt=0):
	"""Focus a modern Editor suggestion and verify Office did not move focus back."""
	try:
		focus = api.getFocusObject()
		if not _modernEditorFocusStillRelevant(root, focus, launchFocus):
			return
		target = _modernEditorKeyboardFocusTarget(selected, root)
		target.setFocus()
	except Exception:
		log.debugWarning("Spelling and Grammar Check: could not focus the Editor suggestion", exc_info=True)
		return
	if attempt < 3:
		core.callLater(
			(80, 180, 350)[attempt],
			_verifyModernEditorSuggestionFocus,
			root,
			selected,
			launchFocus,
			attempt + 1,
		)


def _verifyModernEditorSuggestionFocus(root, selected, launchFocus, attempt):
	try:
		focus = api.getFocusObject()
	except Exception:
		return
	if not _modernEditorFocusStillRelevant(root, focus, launchFocus):
		return
	issue = _modernEditorIssue(root)
	if issue is None:
		return
	suggestions = issue[3]
	if (
		_modernEditorSuggestionForObject(focus, suggestions) is None
		and not _modernEditorIsSuggestionArea(focus, suggestions)
	):
		_focusModernEditorSuggestion(root, selected, launchFocus, attempt)


def _announceModernEditorFocus(obj, root, issue):
	"""Announce one ready modern Editor issue after native focus speech is muted."""
	global _modernEditorIssueKey, _modernEditorSpeechUntil
	global _modernEditorAnnouncedIssueKey
	errorText, issueKind, issueType, suggestions = issue
	if not suggestions:
		return False
	selected = suggestions[0]
	selectedStates = tuple(
		state
		for state in (
			getattr(controlTypes.State, "SELECTED", None),
			getattr(controlTypes.State, "FOCUSED", None),
		)
		if state is not None
	)
	for candidate in suggestions:
		try:
			if any(state in (candidate.states or ()) for state in selectedStates):
				selected = candidate
				break
		except Exception:
			pass
	try:
		rootIdentity = _rootWindowOf(root) or id(root)
	except Exception:
		rootIdentity = id(root)
	focusedSuggestion = _modernEditorSuggestionForKeyboardFocus(obj, suggestions)
	if focusedSuggestion is None and _modernEditorIsSuggestionArea(obj, suggestions):
		focusedSuggestion = selected
	key = (
		rootIdentity,
		issueKind,
		issueType,
		errorText,
		_automationIdOf(focusedSuggestion) if focusedSuggestion is not None else "issue",
	)
	now = time.monotonic()
	previousKey = _modernEditorIssueKey
	if key == previousKey:
		return True
	issueKey = key[:4]
	if focusedSuggestion is None and issueKey == _modernEditorAnnouncedIssueKey:
		_modernEditorIssueKey = key
		return True
	_modernEditorIssueKey = key
	_modernEditorSpeechUntil = now + _PROOFING_DUPLICATE_WINDOW
	if focusedSuggestion is not None:
		firstSuggestion = previousKey is None or previousKey[:4] != key[:4] or previousKey[-1] == "issue"
		sequence = [_("Suggestions list box")] if firstSuggestion else []
		replacementText = _modernEditorReplacementText(_modernEditorSuggestionText(focusedSuggestion))
		if replacementText:
			sequence.extend(_suggestionWordSpeech(replacementText))
			sequence.append(_("{index} of {total}").format(
				index=suggestions.index(focusedSuggestion) + 1,
				total=len(suggestions),
			))
	elif issueKind.casefold() == "grammar":
		_modernEditorAnnouncedIssueKey = issueKey
		sequence = [_('Grammar error')]
		sequence.extend(_modernEditorGrammarSpeech(errorText, issueType))
	else:
		_modernEditorAnnouncedIssueKey = issueKey
		sequence = _spellingSpeech(errorText, "spelling")
	if getFocusSpellSuggestions() and focusedSuggestion is None:
		_announceProofingSequence(
			sequence,
			after=lambda: _focusModernEditorSuggestion(root, selected, obj),
		)
	else:
		with _ownSpeech():
			speech.speak(sequence)
	return True


def _retryModernEditorFocus(root, attempt):
	"""Wait briefly for a modern Editor card to finish loading its suggestions."""
	if attempt >= len(_MODERN_EDITOR_RETRY_DELAYS):
		return
	try:
		focus = api.getFocusObject()
		currentRoot = _modernEditorRoot(focus)
		if currentRoot is None or _modernEditorIdentity(currentRoot) != _modernEditorIdentity(root):
			return
		issue = _modernEditorIssue(currentRoot)
		if issue is not None:
			_announceModernEditorFocus(focus, currentRoot, issue)
			return
	except Exception:
		log.debugWarning("Spelling and Grammar Check: could not retry the Editor issue", exc_info=True)
		return
	core.callLater(
		_MODERN_EDITOR_RETRY_DELAYS[attempt],
		_retryModernEditorFocus,
		currentRoot,
		attempt + 1,
	)


def _handleModernEditorFocus(obj, nextHandler, root=None):
	"""Own one modern Editor issue: error, one list label, suggestion, and focus."""
	root = root or _modernEditorRoot(obj)
	if root is None:
		return False
	issue = _modernEditorIssue(root)
	if issue is None:
		summaryChoice = _modernEditorSummaryChoice(root)
		if summaryChoice is not None:
			if not _openModernEditorSummaryChoice(root, summaryChoice):
				return False
			with _hardMute():
				nextHandler()
			return True
		if (
			not _modernEditorCardPending(root)
			and time.monotonic() >= _modernEditorWindowScanUntil
		):
			return False
	try:
		with _hardMute():
			nextHandler()
	finally:
		if issue is None:
			core.callLater(_MODERN_EDITOR_RETRY_DELAYS[0], _retryModernEditorFocus, root, 1)
		else:
			_announceModernEditorFocus(obj, root, issue)
	return True


### The classic "Possible Word Choice Error" dialog on some older/hybrid Outlook builds

_LEGACY_SDM_WINDOW_CLASS = "bosa_sdm_mso96"
_LEGACY_WORD_CHOICE_DESCRIPTION = "possible word choice error"


def _legacySdmWordChoiceDialog(obj):
	"""Return classic Outlook's word-choice dialog containing C{obj}, if any."""
	# This handler is called for every focus event. The 1.0.11 log captured it
	# walking an unrelated Explorer list and blocking in an accessibility call.
	# Every confirmed word-choice object in that log has this dialog class, so
	# never traverse parents or query descriptions for anything else.
	if _windowClassOf(obj).casefold() != _LEGACY_SDM_WINDOW_CLASS:
		return None
	current = obj
	for _step in range(8):
		if current is None:
			return None
		if _windowClassOf(current).casefold() == _LEGACY_SDM_WINDOW_CLASS:
			if time.monotonic() < _classicProofingPunctuationLaunchUntil:
				return current
			try:
				description = " ".join((getattr(current, "description", "") or "").split()).casefold().rstrip(" :")
			except Exception:
				description = ""
			if description == _LEGACY_WORD_CHOICE_DESCRIPTION:
				return current
		try:
			current = current.parent
		except Exception:
			return None
	return None


def _legacySdmWordChoiceIssue(dialog):
	"""Return the marked word, suggestions container and selected replacement."""
	# allowSlowFallback=False: this function is reached from
	# _handleLegacySdmWordChoiceFocus, which runs on EVERY focus change, not
	# just during an active F7 check. A 1.0.68 stack trace showed it was
	# reaching this call for the MODERN in-line editor dialog too (not just
	# genuine classic Word/Outlook popups), paying for the full slow
	# accessible-tree walk -- and its ~1s freeze -- on ordinary focus changes
	# throughout the whole session, completely bypassing the careful retry
	# ladder tuned for the F7 flow in _focusSuggestionList. A real classic
	# dialog uses actual native listbox windows, so the fast scan inside
	# _findSuggestionList already finds it without ever needing the slow
	# walk -- meaning turning the slow fallback off here costs classic-dialog
	# support nothing, while stopping this path from ever paying for it
	# against a dialog it was never meant to handle.
	container = _findSuggestionList(dialog, allowSlowFallback=False)
	if container is None:
		return None
	try:
		errorText = (_messageSelection(dialog) or "").strip().strip(_WORD_EDGE_PUNCTUATION)
	except Exception:
		errorText = ""
	if not _looksLikeProofingText(errorText, allowPhrase=True):
		return None
	suggestion = _selectedSuggestionText(container)
	if not suggestion:
		try:
			suggestion = " ".join((getattr(container, "value", "") or "").split())
		except Exception:
			suggestion = ""
	if not suggestion:
		return None
	return errorText, container, suggestion


def _handleLegacySdmWordChoiceFocus(obj, nextHandler):
	"""Own the classic word-choice dialog used by some older Outlook builds."""
	global _legacySdmIssueKey
	dialog = _legacySdmWordChoiceDialog(obj)
	if dialog is None:
		return False
	role = getattr(obj, "role", None)
	if (
		obj is not dialog
		and role not in _SUGGESTION_CONTAINER_ROLES
		and role not in _SUGGESTION_ITEM_ROLES
		and _suggestionListOf(obj) is None
	):
		return False
	issue = _legacySdmWordChoiceIssue(dialog)
	if issue is None:
		return False
	errorText, container, suggestion = issue
	try:
		dialogIdentity = _rootWindowOf(dialog) or id(dialog)
	except Exception:
		dialogIdentity = id(dialog)
	inSuggestions = _suggestionListOf(obj) is container
	issueKey = (dialogIdentity, errorText.casefold())
	key = issueKey + (suggestion.casefold(), bool(inSuggestions))
	with _hardMute():
		nextHandler()
	previousKey = _legacySdmIssueKey
	if key == previousKey:
		if inSuggestions:
			_announceSuggestionEntry(obj, container, False)
		return True
	newIssue = previousKey is None or previousKey[:2] != issueKey
	_legacySdmIssueKey = key
	target = _suggestionFocusTarget(container)
	if inSuggestions and not newIssue:
		firstSuggestion = previousKey is None or previousKey[:3] != key[:3] or not previousKey[-1]
		_announceSuggestionEntry(obj, container, firstSuggestion)
		return True
	sequence = [_('Grammar error'), _("Some words are similar but are used differently:")]
	sequence.extend(_suggestionWordSpeech(errorText.rstrip(" .:")))
	if inSuggestions:
		_announceProofingSequence(
			sequence,
			after=lambda: _announceSuggestionEntry(obj, container, True),
		)
	elif getFocusSpellSuggestions():
		_announceProofingSequence(sequence, after=target.setFocus)
	else:
		with _ownSpeech():
			speech.speak(sequence)
	return True


### Speaking a suggestion

def _suggestionWordSpeech(word):
	"""The suggested word or phrase, followed by its letters."""
	# Translators: Introduces the replacement offered for the preceding error.
	sequence = [_('Suggestion'), word]
	try:
		getSpelling = getattr(speech, "getSpellingSpeech", None)
		if getSpelling is None:
			return sequence
		if _EndUtteranceCommand is not None:
			sequence.append(_EndUtteranceCommand())
		if getSlowProofingSpeech() and _RateCommand is not None:
			sequence.append(_RateCommand(multiplier=_SPELL_RATE_MULTIPLIER))
		sequence.extend(getSpelling(word))
		if getSlowProofingSpeech() and _RateCommand is not None:
			sequence.append(_RateCommand())
	except Exception:
		log.debugWarning("Spelling and Grammar Check: could not spell the suggestion", exc_info=True)
		return [word]
	return sequence


def _announceSuggestionEntry(obj, container, first):
	"""Speak a proofing suggestion without repeating its label or shortcut."""
	global _suggestionsLabelUntil, _suggestionsLabelKey
	global _lastSuggestionSpeechKey, _lastSuggestionSpeechUntil
	name = _suggestionItemName(obj, container)
	containerIdentity = _suggestionContainerIdentity(container, obj)
	key = (containerIdentity, name)
	now = time.monotonic()
	if name and key == _lastSuggestionSpeechKey and now < _lastSuggestionSpeechUntil:
		return
	_lastSuggestionSpeechKey = key
	_lastSuggestionSpeechUntil = now + _PROOFING_DUPLICATE_WINDOW
	sequence = []
	if name:
		if _pluginInstance is not None:
			_pluginInstance._cancelNoSuggestionAnnouncement()
		sequence.extend(_suggestionWordSpeech(name))
	else:
		# Outlook first fires empty list-item focus events, then exposes the real
		# suggestion a moment later. The log shows these have no accessible text;
		# say nothing here rather than announce an invented placeholder before (or
		# repeatedly after) the actual suggestion.
		return
	position = _suggestionPosition(obj, container)
	if position is not None:
		index, total = position
		# Translators: Position of an item in the spelling suggestions list.
		sequence.append(_("{index} of {total}").format(index=index, total=total))
	if sequence:
		with _ownSpeech():
			speech.speak(sequence)


### Finding and focusing the classic dialog's suggestion list

try:
	from winBindings.user32 import WNDENUMPROC as _ENUM_WINDOWS_PROC
	from winBindings.user32 import EnumChildWindows as _enumChildWindows
except Exception:
	try:
		import ctypes
		import ctypes.wintypes

		_ENUM_WINDOWS_PROC = ctypes.WINFUNCTYPE(
			ctypes.wintypes.BOOL,
			ctypes.wintypes.HWND,
			ctypes.wintypes.LPARAM,
		)
		_enumChildWindows = ctypes.windll.user32.EnumChildWindows
		_enumChildWindows.restype = ctypes.wintypes.BOOL
		_enumChildWindows.argtypes = (
			ctypes.wintypes.HWND,
			_ENUM_WINDOWS_PROC,
			ctypes.wintypes.LPARAM,
		)
	except Exception:
		_ENUM_WINDOWS_PROC = None
		_enumChildWindows = None
		log.debugWarning("Spelling and Grammar Check: cannot walk child windows", exc_info=True)

_MAX_WINDOWS_WALKED = 400


def _childWindows(window):
	"""Every window inside C{window}, however deeply nested, or an empty list."""
	found = []
	if not window or _enumChildWindows is None:
		return found

	def visit(child, _lParam):
		found.append(child)
		return len(found) < _MAX_WINDOWS_WALKED

	try:
		_enumChildWindows(window, _ENUM_WINDOWS_PROC(visit), 0)
	except Exception:
		log.debugWarning("Spelling and Grammar Check: could not walk child windows", exc_info=True)
	return found


_lastEmptySuggestionScan = (0, 0.0)
# _focusSuggestionList's retry schedule (100ms to send Alt+N and check, then
# delays of 120, 250, 450ms between the remaining attempts) adds up to about
# 920ms worst-case from the first expensive scan to the last retry. A 1.0.51
# log showed the previous 0.6s cache window expiring before that last retry,
# so a second full dialog walk (plus the window-enumeration fallback) ran
# again right before "No suggestions available" was announced -- doubling
# the cost at the worst possible moment. Cover the whole retry sequence with
# margin so only one expensive walk ever happens per issue.
_EMPTY_SUGGESTION_SCAN_TTL = 1.5

# Windows where the fast native window-class scan below has already been
# proven, this session, to never find anything -- e.g. a modern in-line
# "Editor" pane draws its suggestions inside the document window itself,
# with no separate listbox child window for the scan to ever match. Once a
# window is in here, _quickRecheckSuggestionFocus skips straight to the slow
# accessible-tree walk instead of wasting ~820ms retrying a scan that this
# same window has already shown cannot succeed. This only changes *when* the
# existing, already-correct slow walk runs -- never what it finds -- so it
# cannot introduce a wrong result, only remove pointless waiting.
_nativeScanFutileWindows = set()


def _findSuggestionList(obj, allowSlowFallback=True):
	"""Find the suggestions list in the small Office proofing dialog containing C{obj}.

	C{allowSlowFallback} controls whether the slow, COM-heavy accessible-tree
	walk below the fast native scan is allowed to run at all. A 1.0.55 log
	showed that even with the fast native scan tried first and the empty-result
	cache in place, calling this during the early retries (while Word usually
	just hasn't created the listbox window yet) still triggers that slow walk
	-- and its ~1s watchdog freeze -- on a meaningful fraction of issues,
	sometimes more than once per issue. The fast native scan is effectively
	free, so retrying *that* costs nothing; the expensive walk is deferred by
	the caller to the final retry, by which point the window has almost always
	appeared and the fast scan alone succeeds.
	"""
	global _lastEmptySuggestionScan
	root = obj
	for _step in range(_DIALOG_WALK_LIMIT):
		parent = getattr(root, "parent", None)
		if parent is None or getattr(parent, "role", None) in _NOT_A_DIALOG_ROLES:
			break
		root = parent
	rootWindow = getattr(root, "windowHandle", None) or 0
	cacheKey = (rootWindow, id(obj))
	now = time.monotonic()
	cachedKey, cachedUntil = _lastEmptySuggestionScan
	if rootWindow and cacheKey == cachedKey and now < cachedUntil:
		# _focusSuggestionList retries this exact scan up to four times, a few
		# hundred ms apart, while waiting for Word to populate a suggestions
		# list. A 1.0.49 log showed each retry repeating this same slow,
		# exhaustive dialog walk (plus the window-enumeration fallback below)
		# even when the dialog genuinely had none, adding up to several
		# seconds of silence before "No suggestions" was finally announced.
		# Nothing about the dialog's contents changes that fast, so skip
		# repeating the walk for a short window rather than redoing it from
		# scratch on every retry.
		return None
	# A native UI Automation FindFirst-based shortcut was tried here across
	# 1.0.58-1.0.62 and removed in 1.0.64: it eventually worked without
	# crashing, but it reliably grabbed the outer issue-card element instead
	# of the actual nested suggestions listbox (both satisfy ControlType=List),
	# making the addon announce the grammar rule's explanation plus the
	# ORIGINAL flagged word instead of the real suggested replacement. A wrong
	# announcement is worse than a slow correct one, so that route is
	# abandoned rather than parked half-finished.
	level = [root]
	read = 0
	# Try the cheap path first: a raw Win32 EnumChildWindows scan by class name.
	# This is a handful of native calls, not the NVDA-accessible tree walk below,
	# and in the normal case it finds the suggestions listbox directly. A 1.0.53
	# log showed the walk below taking over a second on its own -- not from
	# repeating it (1.0.52 already stopped that), but because even a single pass
	# through NVDA's generic, COM-heavy child-walk is slow on this build.
	# Checking the native window list first avoids that cost whenever it works;
	# the accessible walk stays below as a fallback for anything the window-class
	# check doesn't recognize.
	window = _rootWindowOf(obj)
	seenClasses = []
	for childWindow in _childWindows(window):
		try:
			className = (winUser.getClassName(childWindow) or "").lower()
		except Exception:
			continue
		if className not in ("listbox", "listview", "syslistview32"):
			seenClasses.append(className)
			continue
		candidate = _objectFromWindow(childWindow)
		if candidate is not None and _isSuggestionContainer(candidate):
			return candidate
		seenClasses.append("%s(rejected)" % className)
	if not allowSlowFallback:
		# Do not cache this as an empty result: the fast scan alone isn't
		# conclusive, and the caller will either retry the fast scan again
		# shortly or fall through to a final attempt that does allow the
		# slow walk. Caching here would incorrectly suppress that attempt.
		return None
	# The fast native scan above missed this dialog's suggestions list, so the
	# slow accessible-tree walk below is about to run (and, per a 1.0.56 log,
	# often costs the ~1s watchdog freeze). Record what window classes were
	# actually seen (useful if a future Word/Outlook build's dialog can be
	# recognized this way) and remember that this window's scan is futile so
	# _quickRecheckSuggestionFocus can skip straight here next time instead of
	# retrying a scan that already failed once for this exact window.
	log.debug(
		"Spelling and Grammar Check: fast suggestion-list scan missed, "
		"child window classes seen: %s" % (seenClasses or "(no child windows)"),
	)
	if window:
		# NOTE: keyed by `window` (== _rootWindowOf(obj), the real top-level
		# window handle used by the native scan above), NOT by `rootWindow`
		# (the accessible .parent-climbed root computed earlier in this
		# function). A 1.0.59 log proved that .parent walk climbs all the way
		# past this dialog to the OS Desktop object for this in-line-editor
		# style dialog, so `rootWindow` is not a usable key here -- storing it
		# in 1.0.64 meant _quickRecheckSuggestionFocus's lookup (which
		# correctly calls _rootWindowOf(obj) itself) never matched, and the
		# skip-ahead optimization never actually fired. Fixed in 1.0.65.
		_nativeScanFutileWindows.add(window)
		log.debug("Spelling and Grammar Check: marked window %r as native-scan-futile" % window)
	for _depth in range(_DIALOG_SCAN_DEPTH + 2):
		below = []
		for item in level:
			read += 1
			if read > _DIALOG_SCAN_LIMIT:
				_lastEmptySuggestionScan = (cacheKey, now + _EMPTY_SUGGESTION_SCAN_TTL)
				return None
			if _isSuggestionContainer(item):
				return item
			try:
				below.extend(item.children or ())
			except Exception:
				pass
		level = below
		if not level:
			break
	_lastEmptySuggestionScan = (cacheKey, now + _EMPTY_SUGGESTION_SCAN_TTL)
	return None


def _suggestionFocusTarget(container):
	"""The actual selected/focused suggestion, including one below wrapper objects."""
	target = None
	fallback = None
	selectedStates = tuple(
		state
		for state in (
			getattr(controlTypes.State, "SELECTED", None),
			getattr(controlTypes.State, "FOCUSED", None),
		)
		if state is not None
	)
	try:
		activeChild = getattr(container, "activeChild", None)
	except Exception:
		activeChild = None
	if activeChild is not None and callable(getattr(activeChild, "setFocus", None)):
		target = activeChild
	for child in _suggestionDescendants(container):
		if getattr(child, "role", None) not in _SUGGESTION_ITEM_ROLES:
			continue
		if not callable(getattr(child, "setFocus", None)):
			continue
		if _accessibleSuggestionText(child) and fallback is None:
			fallback = child
		try:
			states = child.states or ()
			if any(state in states for state in selectedStates):
				target = child
				break
		except Exception:
			pass
	target = target or fallback
	return target or container


def _focusSpellField(window):
	"""Move initial proofing focus out of suggestions when the user requests it."""
	field = _objectFromWindow(window)
	if field is None:
		return
	try:
		field.setFocus()
	except Exception:
		log.error("Spelling and Grammar Check: could not focus the misspelled word", exc_info=True)


def _pressSpaceWhenIgnoreOnceFocused(expectedButton, attempt=0):
	"""Press Space only after Word has actually focused Ignore Once.

	Word's focus change is asynchronous.  Sending Space immediately after
	setFocus() can therefore land on the old proofing control instead.
	"""
	try:
		focus = api.getFocusObject()
		focusName = " ".join((getattr(focus, "name", "") or "").split()).casefold()
		# The 1.0.42 log confirms that NVDA announces this exact button but
		# exposes it through a different accessibility window handle than the
		# object found in the dialog walk. Its name is the reliable evidence
		# that focus has reached the intended control.
		if focusName in ("ignore once", "&ignore once"):
			import keyboardHandler

			keyboardHandler.KeyboardInputGesture.fromName("space").send()
			return
	except Exception:
		log.debugWarning("Spelling and Grammar Check: could not verify Ignore Once focus", exc_info=True)
		return
	# The 1.0.41 log shows the button receiving focus 362 ms after setFocus().
	# Retry briefly, but never send Space until NVDA confirms that focus.
	if attempt < 4:
		delay = (75, 125, 200, 300, 450)[attempt]
		core.callLater(delay, _pressSpaceWhenIgnoreOnceFocused, expectedButton, attempt + 1)
	else:
		log.debugWarning("Spelling and Grammar Check: Ignore Once did not receive focus before activation")


def _isClassicIgnoreOnceButton(obj):
	"""Whether C{obj} is the real Ignore Once button Word has focused."""
	try:
		name = " ".join((getattr(obj, "name", "") or "").split()).casefold()
		return name in ("ignore once", "&ignore once")
	except Exception:
		return False


def _activateClassicIgnoreOnce(obj):
	"""Activate Word's real Ignore Once button in a classic proofing dialog."""
	dialog = obj if _looksLikeADialog(obj) else _dialogAbove(obj)
	if dialog is None:
		return False
	level = [dialog]
	read = 0
	for _depth in range(_DIALOG_SCAN_DEPTH + 2):
		below = []
		for item in level:
			read += 1
			if read > _DIALOG_SCAN_LIMIT:
				return False
			try:
				name = " ".join((getattr(item, "name", "") or "").split()).casefold()
			except Exception:
				name = ""
			if name in ("ignore once", "&ignore once"):
				action = getattr(item, "doAction", None)
				if callable(action):
					try:
						action()
						return True
					except Exception:
						# The 1.0.39 log showed this Office button rejecting doAction().
						# Focus it, then wait for Word to complete that focus change before
						# sending the logged manual activation. The actual focus event is
						# handled by GlobalPlugin; polling remains only as a fallback.
						try:
							if _pluginInstance is not None:
								_pluginInstance._pendingClassicIgnoreOnce = True
							item.setFocus()
							core.callLater(1200, _pressSpaceWhenIgnoreOnceFocused, item)
							return True
						except Exception:
							log.debugWarning("Spelling and Grammar Check: could not activate Ignore Once", exc_info=True)
			try:
				below.extend(item.children or ())
			except Exception:
				pass
		if not below:
			break
		level = below
	return False


def _focusSuggestionList(obj, attempt=0, expectedWord=None, quietIfMissing=False):
	"""Put focus on the selected or first spelling suggestion."""
	global _suppressSuggestionsListUntil
	try:
		expectedField = _officeSpellCheckField(obj)
		currentField = _officeSpellCheckField(api.getFocusObject())
		expectedWindow = getattr(expectedField, "windowHandle", None)
		currentWindow = getattr(currentField, "windowHandle", None)
		if not expectedWindow or currentWindow != expectedWindow:
			return
		# The checker reuses one window for every error. A callback queued for the
		# previous word must not move into (and announce) the new word's suggestions.
		if expectedWord is not None and _misspelledWord(currentField) != expectedWord:
			return
	except Exception:
		return
	if attempt == 0:
		try:
			import keyboardHandler

			_suppressSuggestionsListUntil = time.monotonic() + 3.0
			keyboardHandler.KeyboardInputGesture.fromName("alt+n").send()
			core.callLater(100, _verifySuggestionFocus, obj, 1, expectedWord, quietIfMissing)
			return
		except Exception:
			log.error("Spelling and Grammar Check: could not send the suggestions shortcut", exc_info=True)
	# Only the last retry pays for the slow accessible-tree fallback inside
	# _findSuggestionList. Every earlier attempt only tries the fast native
	# window-class scan, which is effectively free -- this keeps the common
	# case (window already exists) instant, and avoids re-triggering the
	# ~1s watchdog freeze on every retry while Word is still creating it.
	container = _findSuggestionList(obj, allowSlowFallback=attempt >= 4)
	if container is None:
		if attempt < 4:
			core.callLater((50, 120, 250, 450)[attempt], _focusSuggestionList, obj, attempt + 1, expectedWord, quietIfMissing)
		elif not quietIfMissing:
			# Translators: Said when a proofing issue has no suggestions to offer.
			with _ownSpeech():
				speech.speak([_("No suggestions.")])
		return
	target = _suggestionFocusTarget(container)
	try:
		target.setFocus()
	except Exception:
		log.error("Spelling and Grammar Check: could not focus the suggestions list", exc_info=True)
		if attempt < 4:
			core.callLater((50, 120, 250, 450)[attempt], _focusSuggestionList, obj, attempt + 1, expectedWord, quietIfMissing)
		return
	if attempt < 4:
		core.callLater((50, 120, 250, 450)[attempt], _verifySuggestionFocus, obj, attempt + 1, expectedWord, quietIfMissing)


def _quickRecheckSuggestionFocus(obj, expectedWord, quietIfMissing):
	"""One extra cheap focus-chain check before paying for the slow dialog walk.

	A 1.0.54 log showed Word sometimes not having moved focus into the
	suggestions list yet at the ~100ms mark, which forced the expensive
	accessible-tree walk in _findSuggestionList (and the ~1s watchdog
	freeze that comes with it) even on issues where focus settles into the
	list correctly a little later on its own. This recheck is the same
	near-free upward walk _verifySuggestionFocus already does, so trying it
	once more before escalating can only help, never add real cost.
	"""
	try:
		focus = api.getFocusObject()
	except Exception:
		focus = None
	field = _officeSpellCheckField(focus)
	if expectedWord is not None and (field is None or _misspelledWord(field) != expectedWord):
		log.debug(
			"Spelling and Grammar Check: _quickRecheckSuggestionFocus bailed early, "
			"expectedWord=%r field=%r" % (expectedWord, field),
		)
		return
	container = _suggestionListOf(focus)
	if container is not None:
		_announceSuggestionEntry(focus, container, True)
		return
	# If this exact window has already shown, this session, that the fast
	# native scan in _findSuggestionList never finds anything here (see
	# _nativeScanFutileWindows), skip the pointless attempts 1-3 -- which can
	# only ever repeat that same fast scan and fail the same way -- and go
	# straight to the one attempt that actually allows the slow walk. This
	# saves ~820ms of retry waiting on issues in a window we already know
	# needs it, without changing which object ends up getting found.
	rootWindow = _rootWindowOf(obj)
	if rootWindow and rootWindow in _nativeScanFutileWindows:
		_focusSuggestionList(obj, 4, expectedWord, quietIfMissing)
		return
	_focusSuggestionList(obj, 1, expectedWord, quietIfMissing)


def _verifySuggestionFocus(obj, attempt, expectedWord=None, quietIfMissing=False):
	try:
		focus = api.getFocusObject()
	except Exception:
		focus = None
	field = _officeSpellCheckField(focus)
	if expectedWord is not None and (field is None or _misspelledWord(field) != expectedWord):
		return
	container = _suggestionListOf(focus)
	if container is None:
		if attempt == 1:
			# Give Word a little longer to move focus on its own before paying
			# for the slow walk -- see _quickRecheckSuggestionFocus.
			core.callLater(150, _quickRecheckSuggestionFocus, obj, expectedWord, quietIfMissing)
			return
		_focusSuggestionList(obj, attempt, expectedWord, quietIfMissing)
		return
	_announceSuggestionEntry(focus, container, True)


### Locating the classic dialog's error field

def _anyWordDialogSurface(obj):
	"""Any Word dialog surface in the same thread as C{obj}, or 0 for none."""
	try:
		import NVDAHelper

		for windowClass in sorted(_WORD_DIALOG_WINDOW_CLASSES):
			window = NVDAHelper.localLib.findWindowWithClassInThread(
				obj.windowThreadID,
				windowClass,
				True,
			)
			if window:
				return window
	except Exception:
		log.debugWarning("Spelling and Grammar Check: could not look for a Word dialog", exc_info=True)
	return 0


def _spellBoxInside(window):
	"""The box showing the mistake, somewhere inside C{window}, or 0."""
	spare = 0
	for child in _childWindows(window):
		try:
			if winUser.getClassName(child) not in _WORD_DIALOG_WINDOW_CLASSES:
				continue
			if winUser.getControlID(child) == _SPELL_ERROR_CONTROL_ID:
				return child
		except Exception:
			continue
		spare = spare or child
	return spare


def _sameWindowFamily(window, obj):
	"""Whether C{window} and C{obj} belong to the same family of windows."""
	try:
		owner = getattr(winUser, "GA_ROOTOWNER", 3)
		return winUser.getAncestor(window, owner) == winUser.getAncestor(obj.windowHandle, owner)
	except Exception:
		return False


def _spellErrorFieldWindow(obj):
	"""The handle of the box showing the mistake, in the spelling dialog C{obj} is in."""
	if _windowClassOf(obj) in _WORD_DIALOG_WINDOW_CLASSES:
		try:
			if winUser.getControlID(obj.windowHandle) == _SPELL_ERROR_CONTROL_ID:
				return obj.windowHandle
		except Exception:
			pass
	surface = _anyWordDialogSurface(obj)
	if not surface:
		return 0
	try:
		containers = []
		for container in (_rootWindowOf(obj), winUser.getForegroundWindow()):
			if container and container not in containers:
				containers.append(container)
		for container in containers:
			window = _spellBoxInside(container)
			if window:
				return window
		found = winUser.getAncestor(surface, winUser.GA_ROOT)
		if found in containers or _sameWindowFamily(surface, obj):
			return _spellBoxInside(found) or surface
	except Exception:
		log.debugWarning("Spelling and Grammar Check: could not look for the spelling dialog", exc_info=True)
	return 0


def _objectFromWindow(window):
	"""A fresh NVDA object for C{window}, or C{None}."""
	try:
		from NVDAObjects.IAccessible import getNVDAObjectFromEvent

		return getNVDAObjectFromEvent(window, winUser.OBJID_CLIENT, 0)
	except Exception:
		log.debugWarning("Spelling and Grammar Check: could not reach the spelling box", exc_info=True)
		return None


def _officeSpellCheckField(obj):
	"""The box showing the mistake in Word or Outlook's spelling dialog, if present."""
	if not _isOfficeProofingWindow(obj):
		return None
	window = _spellErrorFieldWindow(obj)
	if not window:
		return None
	if getattr(obj, "windowHandle", None) == window:
		return obj
	return _objectFromWindow(window)


def _selectionText(obj):
	"""What is selected in C{obj}, or what the cursor is on if nothing is."""
	for position, unit in (
		(textInfos.POSITION_SELECTION, None),
		(textInfos.POSITION_CARET, textInfos.UNIT_WORD),
	):
		try:
			info = obj.makeTextInfo(position)
			if unit is not None:
				info.expand(unit)
			text = info.text or ""
		except Exception:
			continue
		if text.strip():
			return text
	return None


def _allText(obj):
	try:
		return obj.makeTextInfo(textInfos.POSITION_ALL).text or None
	except Exception:
		return None


def _boldRun(field):
	"""The bold part of the box, which is how the dialog shows which word it means."""
	from displayModel import EditableTextDisplayModelTextInfo

	info = EditableTextDisplayModelTextInfo(field, textInfos.POSITION_ALL)
	inBold = False
	bold = []
	for item in info.getTextWithFields():
		if isinstance(item, str):
			if inBold:
				bold.append(item)
		elif getattr(item, "field", None):
			inBold = item.field.get("bold", False)
		if not inBold and bold:
			break
	return "".join(bold)


def _documentBodyObject(field):
	"""The real message/document surface behind a small proofing dialog field.

	Word highlights the flagged text in the message itself while the dialog is open,
	including a format annotation saying whether it is a spelling or a grammar issue.
	The dialog's own small edit box does not carry that annotation, so callers that need
	to tell spelling and grammar apart must look at this object instead.
	"""
	import NVDAHelper

	for windowClass in sorted(_OUTLOOK_BODY_WINDOW_CLASSES):
		window = NVDAHelper.localLib.findWindowWithClassInThread(
			field.windowThreadID,
			windowClass,
			True,
		)
		if not window:
			continue
		body = _objectFromWindow(window)
		if body is not None:
			return body
	return None


def _messageSelection(field):
	"""What Word has selected in the message itself."""
	body = _documentBodyObject(field)
	if body is None:
		return None
	return _selectionText(body)


def _wordSources(field):
	"""Every way of asking what word the checker has stopped on, best answer first."""
	return (
		("errorText", lambda: getattr(field, "errorText", None)),
		("value", lambda: getattr(field, "value", None)),
		("name", lambda: getattr(field, "name", None)),
		("bold", lambda: _boldRun(field)),
		("selection", lambda: _selectionText(field)),
		("message", lambda: _messageSelection(field)),
		("description", lambda: getattr(field, "description", None)),
		("text", lambda: _allText(field)),
	)


_lastWordSourceByWindow = {}


def _misspelledWord(field):
	"""The text the Office proofing checker is asking about, or C{None}."""
	global _lastWordSourceByWindow
	try:
		field.invalidateCache()
	except Exception:
		pass
	window = getattr(field, "windowHandle", None)
	sources = dict(_wordSources(field))
	order = list(sources.keys())
	# On some Word/Outlook builds the cheap sources (errorText, value, name...)
	# always come back empty, and every single issue in a supplied 1.0.50 log
	# fell all the way through to "message" — the one source that needs
	# _documentBodyObject, which triggers NVDA's slow WinwordWindowObject/
	# generic-child-walk path. Once a source has worked for this dialog's
	# window, try it first next time instead of re-failing the same cheap
	# sources on every single issue.
	preferred = _lastWordSourceByWindow.get(window)
	if preferred in order:
		order.remove(preferred)
		order.insert(0, preferred)
	for name in order:
		try:
			word = (sources[name]() or "").strip().strip(_WORD_EDGE_PUNCTUATION)
		except Exception:
			continue
		allowPhrase = name in ("errorText", "description", "bold", "message", "text")
		if _looksLikeProofingText(word, allowPhrase=allowPhrase):
			_lastWordSourceByWindow[window] = name
			return word
	return None


def _proofingKindFromFormat(obj):
	"""Return Outlook's explicit spelling/grammar annotation at C{obj}, if exposed."""
	for position, unit in (
		(textInfos.POSITION_SELECTION, None),
		(textInfos.POSITION_CARET, textInfos.UNIT_WORD),
		(textInfos.POSITION_ALL, None),
	):
		try:
			info = obj.makeTextInfo(position)
			if unit is not None:
				info.expand(unit)
			fields = info.getTextWithFields()
		except Exception:
			continue
		for item in fields:
			if not isinstance(item, textInfos.FieldCommand) or item.command != "formatChange":
				continue
			try:
				if item.field.get("invalid-grammar") or item.field.get("invalidGrammar"):
					return "grammar"
				if item.field.get("invalid-spelling") or item.field.get("invalidSpelling"):
					return "spelling"
			except Exception:
				pass
	return None


def _windowCaption(hwnd):
	"""The real Win32 title-bar text of hwnd, read directly rather than via accessibility."""
	if not hwnd:
		return ""
	try:
		text = winUser.getWindowText(hwnd)
		if text:
			return text.strip()
	except Exception:
		pass
	try:
		import ctypes

		length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
		if length <= 0:
			return ""
		buf = ctypes.create_unicode_buffer(length + 1)
		ctypes.windll.user32.GetWindowTextW(hwnd, buf, length + 1)
		return (buf.value or "").strip()
	except Exception:
		return ""


def _classicProofingKind(obj, field, allowProvisional=True):
	"""Classify a classic Office issue from annotations and specific dialog labels."""
	formatKind = None
	for candidate in (field, obj):
		kind = _proofingKindFromFormat(candidate)
		if kind and formatKind is None:
			formatKind = kind
	texts = []
	# The window title of the dialog itself, read directly rather than through the
	# accessible tree, in case the accessible name does not match the caption.
	try:
		fieldWindow = getattr(field, "windowHandle", None)
		rootWindow = winUser.getAncestor(fieldWindow, winUser.GA_ROOT) if fieldWindow else 0
		caption = _windowCaption(rootWindow)
		if caption:
			texts.append(caption.casefold())
	except Exception:
		caption = ""
		log.debugWarning("Spelling and Grammar Check: could not read the dialog's window caption", exc_info=True)
	for attribute in ("name", "description"):
		try:
			text = " ".join((getattr(field, attribute, "") or "").split()).casefold()
		except Exception:
			text = ""
		if text:
			texts.append(text)
	current = obj
	for _step in range(7):
		if current is None:
			break
		for attribute in ("name", "description"):
			try:
				text = " ".join((getattr(current, attribute, "") or "").split()).casefold()
			except Exception:
				text = ""
			if text:
				texts.append(text)
		try:
			current = current.parent
		except Exception:
			break
	try:
		dialog = obj if _looksLikeADialog(obj) else _dialogAbove(obj)
		dialogText = _dialogText(dialog, stopMarkers=_CLASSIC_PROOFING_STOP_MARKERS) if dialog is not None else None
		if dialogText:
			texts.append(dialogText.casefold())
	except Exception:
		pass
	text = " ".join(texts)
	text = re.sub(r"spelling\s+(?:and|&)\s+grammar", "", text)
	result = None
	if any(marker in text for marker in (
		"possible word choice error", "grammatical error", "grammar error",
		"ignore rule", "next sentence",
	)):
		result = "grammar"
	elif any(marker in text for marker in (
		"not in dictionary", "spelling error", "misspelled", "unknown word",
	)):
		result = "spelling"
	elif any(marker in text for marker in (
		"word choice", "punctuation", "clarity",
		"conciseness", "refinement", "formal writing",
		"some words are similar but are used differently",
		# Common Word grammar-checker category labels, added after the "misspelled"
		# vs "grammar error" misclassification seen on 2026.1 builds: none of these
		# distinctive words appear when Word is genuinely reporting a spelling error,
		# so treating any of them as a positive grammar signal is safe.
		"agreement", "subject-verb", "subject verb", "verb form", "verb tense",
		"passive voice", "fragment", "run-on", "run on", "wordiness", "wordy",
		"double negative", "article usage", "pronoun", "capitalization",
		"contraction", "cliche", "cliché", "jargon", "preposition", "modifier",
		"comma", "hyphenation", "gender-specific", "consider revising",
		# Seen live on a 2026.1.1 build: Word's own plain-language explanation for a
		# subject/verb agreement issue ("The dogs is...") used none of the wording
		# above at all, and so fell all the way through to the "spelling" default.
		# Confirmed from NVDA's debug log, "classified as" line, 2026-08-29.
		"singular or plural", "singular and plural", "singular/plural",
		"plural or singular", "sticking to singular",
	)):
		result = "grammar"
	# The 1.0.44 timing log showed the document-body format read taking 375–458 ms
	# and coinciding with the watchdog pauses, while Word's explanation text had
	# already classified every affected issue. Read the body only if that cheap,
	# directly supplied evidence was inconclusive.
	if result is None and formatKind is None:
		try:
			body = _documentBodyObject(field)
		except Exception:
			body = None
		if body is not None:
			formatKind = _proofingKindFromFormat(body)
	if result is None:
		if formatKind and allowProvisional:
			result = formatKind
		elif allowProvisional:
			result = "spelling"
		else:
			# Outlook initially exposes neither its explanation nor a useful semantic
			# annotation for some grammar checks. In the supplied log, "doesn't" was
			# reported here as spelling and only later exposed "In formal writing...".
			# Do not name that transient state as a spelling error.
			result = None
	log.debug(
		"Spelling and Grammar Check: classified as %r (formatKind=%r, caption=%r, scanned text=%r)"
		% (result, formatKind, caption, text[:300])
	)
	return result


def _caretIsOnFinalPunctuation(obj):
	"""Whether the Office caret is on, or immediately after, final punctuation."""
	try:
		info = obj.makeTextInfo(textInfos.POSITION_CARET)
		info.expand(textInfos.UNIT_CHARACTER)
		character = info.text or ""
		if character and character[0] in ".!?;:":
			return True
		info = obj.makeTextInfo(textInfos.POSITION_CARET)
		if info.move(textInfos.UNIT_CHARACTER, -1) == 0:
			return False
		info.expand(textInfos.UNIT_CHARACTER)
		character = info.text or ""
		return bool(character) and character[0] in ".!?;:"
	except Exception:
		return False


def _announceSettledClassicIssue(fieldWindow, fallbackField, fallbackWord, after, attempt=0):
	"""Re-read a just-opened classic dialog after Outlook publishes its subtype."""
	field = _objectFromWindow(fieldWindow) if fieldWindow else None
	field = field or fallbackField
	try:
		focus = api.getFocusObject()
	except Exception:
		focus = field
	issueKind = _classicProofingKind(focus, field)
	if (
		issueKind == "spelling"
		and time.monotonic() < _classicProofingPunctuationLaunchUntil
		and attempt < 3
	):
		core.callLater((350, 550, 800)[attempt], _announceSettledClassicIssue,
			fieldWindow, fallbackField, fallbackWord, after, attempt + 1)
		return
	word = _misspelledWord(field) or fallbackWord
	try:
		if (
			_legacySdmIssueKey is not None
			and word
			and _legacySdmIssueKey[1] == word.casefold()
		):
			return
	except Exception:
		pass
	if _pluginInstance is not None:
		originalAfter = after
		after = lambda: _pluginInstance._afterErrorSpeech(originalAfter, issueKind)
	_announceMisspelledWord(word, after=after, issueKind=issueKind)


### Speaking a misspelled word or grammar issue

def _spellingSpeech(word, issueKind=None):
	""""Misspelled" or "Grammar error", then the word said and spelled slightly slower."""
	if issueKind == "grammar":
		label = _("Grammar error")
	else:
		label = _("misspelled")
	sequence = [label, word]
	spoken = [label]
	try:
		if _EndUtteranceCommand is not None:
			spoken.append(_EndUtteranceCommand())
		slow = getSlowProofingSpeech()
		if slow and _RateCommand is not None:
			spoken.append(_RateCommand(multiplier=_SPELL_RATE_MULTIPLIER))
		spoken.append(word)
		getSpelling = getattr(speech, "getSpellingSpeech", None)
		if getSpelling is not None:
			if _EndUtteranceCommand is not None:
				spoken.append(_EndUtteranceCommand())
			spoken.extend(getSpelling(word))
		if slow and _RateCommand is not None:
			spoken.append(_RateCommand())
	except Exception:
		log.debugWarning("Spelling and Grammar Check: could not spell the word out", exc_info=True)
		return sequence
	return spoken


def _announceProofingSequence(sequence, after=None):
	"""Speak a proofing sequence, then perform one guarded focus action."""
	afterCalled = [False]

	def runAfterOnce():
		if after is None or afterCalled[0]:
			return
		afterCalled[0] = True
		try:
			after()
		except Exception:
			log.debugWarning("Spelling and Grammar Check: proofing focus action failed", exc_info=True)

	spoken = list(sequence)
	if after is not None and _CallbackCommand is not None:
		spoken.append(_CallbackCommand(runAfterOnce, name="spellGrammarCheck.afterProofingSpeech"))
	with _ownSpeech():
		speech.speak(spoken)
	if after is not None:
		core.callLater(3000, runAfterOnce)


def _announceMisspelledWord(word, after=None, issueKind=None):
	try:
		sequence = _spellingSpeech(word, issueKind)
	except Exception:
		log.error("Spelling and Grammar Check: could not build the spelling announcement", exc_info=True)
		sequence = [word]
	afterCalled = [False]

	def runAfterOnce():
		if after is None or afterCalled[0]:
			return
		afterCalled[0] = True
		core.callLater(0, after)

	if after is not None and _CallbackCommand is not None:
		sequence = sequence + [
			_CallbackCommand(runAfterOnce, name="spellGrammarCheck.afterMisspelledWord"),
		]
	with _ownSpeech():
		speech.speak(sequence)
	if after is not None:
		fallbackDelay = min(6000, max(1400, len(word) * 180))
		core.callLater(fallbackDelay, runAfterOnce)


### The dialog that says the check has finished

def _wordList(english, translated):
	words = set()
	for source in (english, translated):
		for word in (source or "").split(","):
			word = word.strip().lower()
			if word:
				words.add(word)
	return frozenset(words)


_SPELLING_WORDS = _wordList("spell,grammar,proofing", _("spell,grammar,proofing"))
_COMPLETION_WORDS = _wordList("complete,finish,done", _("complete,finish,done"))


def _isOkButton(obj):
	if not _isOfficeProofingWindow(obj):
		return False
	try:
		if getattr(obj, "role", None) != controlTypes.Role.BUTTON:
			return False
		return (getattr(obj, "name", "") or "").strip().lower() in _OK_BUTTON_NAMES
	except Exception:
		return False


def _looksLikeADialog(obj):
	if getattr(obj, "role", None) in _DIALOG_ROLES:
		return True
	windowClass = _windowClassOf(obj)
	return windowClass == _DIALOG_WINDOW_CLASS or windowClass.lower().startswith(
		_OFFICE_DIALOG_WINDOW_PREFIX,
	)


def _dialogAbove(obj):
	current = obj
	for _step in range(_DIALOG_WALK_LIMIT):
		try:
			current = current.parent
		except Exception:
			return None
		if current is None:
			return None
		if _looksLikeADialog(current):
			return current
		if getattr(current, "role", None) in _NOT_A_DIALOG_ROLES:
			return None
	return None


def _dialogText(dialog, stopMarkers=None):
	"""Read text from a dialog's descendants, breadth-first.

	If stopMarkers is given, this returns as soon as any of those phrases
	appears in what has been read so far — before fetching that item's own
	children, which is the slow step (NVDA's generic IAccessible child-walk).
	A 1.0.27 log showed this walk taking roughly 1-2 seconds per classification
	because it always read the whole dialog even after a decisive phrase like
	"possible word choice error" had already turned up.
	"""
	texts = []
	read = 0
	level = [dialog]
	for _depth in range(_DIALOG_SCAN_DEPTH):
		below = []
		for item in level:
			read += 1
			if read > _DIALOG_SCAN_LIMIT:
				return " ".join(texts).lower() if stopMarkers else None
			for name in ("name", "value", "description"):
				try:
					text = getattr(item, name, None)
				except Exception:
					continue
				if isinstance(text, str) and text.strip():
					texts.append(text)
			if stopMarkers:
				joined = " ".join(texts).lower()
				if any(marker in joined for marker in stopMarkers):
					return joined
			try:
				below.extend(item.children or ())
			except Exception:
				pass
			if len(below) > _DIALOG_SCAN_LIMIT:
				return " ".join(texts).lower() if stopMarkers else None
		if not below:
			break
		level = below
	return " ".join(texts).lower()


def _saysSpellCheckIsComplete(obj):
	dialog = _dialogAbove(obj)
	if dialog is None:
		return False
	try:
		text = _dialogText(dialog)
	except Exception:
		log.debugWarning("Spelling and Grammar Check: could not read the dialog", exc_info=True)
		return False
	if not text:
		return False
	aboutSpelling = any(word in text for word in _SPELLING_WORDS)
	hasFinished = any(word in text for word in _COMPLETION_WORDS)
	return aboutSpelling and hasFinished


def _spellCheckCompleteSpeech(onTheOkButton):
	sequence = [
		# Translators: Said when the checker has finished and its final dialog, holding
		# only an OK button, comes up.
		_("Spell check is complete."),
	]
	if onTheOkButton:
		# Translators: The button on that dialog, named after the sentence above it.
		sequence.append(_("OK button"))
	return sequence


def _isBareUnavailable(args, kwargs):
	"""Whether a whole utterance is nothing but the word "unavailable"."""
	if _UNAVAILABLE_TEXT is None:
		return False
	if args:
		sequence = args[0]
	elif "speechSequence" in kwargs:
		sequence = kwargs["speechSequence"]
	else:
		return False
	if not isinstance(sequence, list) or not sequence:
		return False
	text = " ".join(item for item in sequence if isinstance(item, str))
	return text.strip().strip(_WORD_EDGE_PUNCTUATION).lower() == _UNAVAILABLE_TEXT


### Patching NVDA's speech functions

_patches = []


def _patch(owner, name, replacement):
	original = getattr(owner, name)
	wasOwn = name in getattr(owner, "__dict__", {})
	setattr(owner, name, replacement)
	_patches.append((owner, name, original, wasOwn, replacement))
	return original


def _unpatchAll():
	while _patches:
		owner, name, original, wasOwn, replacement = _patches.pop()
		try:
			if getattr(owner, name, None) is not replacement:
				continue
			if wasOwn:
				setattr(owner, name, original)
			else:
				delattr(owner, name)
		except Exception:
			log.error("Spelling and Grammar Check: could not restore %s.%s" % (owner, name), exc_info=True)


### Watching for F7 and Escape

def _onGesture(*args, **kwargs):
	"""Arm the proofing state on F7, and note cancellation on Escape."""
	global _modernEditorWindowScanUntil, _modernEditorWindowScanHandle
	global _modernEditorSummaryChoiceKey, _modernEditorSummaryChoiceUntil
	global _modernEditorIssueKey, _modernEditorSpeechUntil
	global _modernEditorAnnouncedIssueKey, _legacySdmIssueKey
	global _classicProofingPunctuationLaunchUntil, _outlookProofingShellUntil
	global _outlookProofingReturnUntil
	gesture = kwargs.get("gesture", args[0] if args else None)
	try:
		focus = api.getFocusObject()
	except Exception:
		focus = None
	if _pluginInstance is not None:
		# Other installed Outlook add-ons wrap NVDA's speech function after this
		# add-on starts. Re-install our small proofing-only filter before an F7,
		# Alt+C, or Escape gesture can cause Outlook's next native announcement.
		_pluginInstance._installSpeechHooks()
	try:
		identifiers = tuple(identifier.lower() for identifier in gesture.identifiers)
	except Exception:
		identifiers = ()
	isProofingLaunch = any(
		re.fullmatch(r"kb(?:\([^)]*\))?:f7", identifier)
		for identifier in identifiers
	)
	if isProofingLaunch:
		if _isOfficeProofingWindow(focus):
			if _isInOutlookWindow(focus):
				_outlookProofingShellUntil = time.monotonic() + _PREMATURE_COMPLETE_WINDOW + 3.0
			_modernEditorIssueKey = None
			_modernEditorSpeechUntil = 0.0
			_modernEditorAnnouncedIssueKey = None
			_legacySdmIssueKey = None
			if _pluginInstance is not None:
				_pluginInstance._proofingCancelUntil = 0.0
			_modernEditorWindowScanUntil = time.monotonic() + 4.0
			_classicProofingPunctuationLaunchUntil = (
				time.monotonic() + 4.0 if _caretIsOnFinalPunctuation(focus) else 0.0
			)
			_modernEditorSummaryChoiceKey = None
			_modernEditorSummaryChoiceUntil = 0.0
			_modernEditorWindowScanHandle = _rootWindowOfObjectOrParent(focus)
			if not _modernEditorWindowScanHandle:
				try:
					_modernEditorWindowScanHandle = winUser.getForegroundWindow() or 0
				except Exception:
					_modernEditorWindowScanHandle = 0
	isEscape = any(re.fullmatch(r"kb(?:\([^)]*\))?:escape", identifier) for identifier in identifiers)
	if isEscape and _pluginInstance is not None and _pluginInstance._inSpellDialog:
		_pluginInstance._proofingCancelUntil = time.monotonic() + 3.0
		if _isInOutlookWindow(focus):
			_outlookProofingShellUntil = time.monotonic() + 5.0
	if not isProofingLaunch:
		if _pluginInstance is not None:
			if _pluginInstance._inSpellDialog:
				# Outlook announces its message title and "unavailable" immediately
				# after Alt+C dismisses the final proofing issue. The supplied log shows
				# this happens before the add-on receives the completion focus event.
				if any(re.fullmatch(r"kb(?:\([^)]*\))?:alt\+c", identifier) for identifier in identifiers):
					_outlookProofingReturnUntil = time.monotonic() + 1.5
				_pluginInstance._noteProofingActivity()
			else:
				_pluginInstance._discardProofingSession()
	return True


def _speechSequenceText(args, kwargs):
	"""Normalized literal text from an NVDA speech sequence."""
	if args:
		sequence = args[0]
	else:
		sequence = kwargs.get("speechSequence")
	if not isinstance(sequence, (list, tuple)):
		return ""
	return " ".join(" ".join(item for item in sequence if isinstance(item, str)).split()).casefold()


def _isOutlookProofingShellSpeech(args, kwargs):
	"""Whether a speech sequence is Outlook's empty F7-opening shell label."""
	return _speechSequenceText(args, kwargs) in (
		"microsoft outlook dialog",
		"microsoft outlook window",
		"dialog",
	)


def _isSuggestionsListSpeech(args, kwargs):
	"""Whether this is NVDA's extra automatic suggestions-list announcement."""
	return _speechSequenceText(args, kwargs) == "suggestions list box"


def _isClassicProofingChromeSpeech(args, kwargs):
	"""The empty classic-dialog title/position utterances seen before a result."""
	text = _speechSequenceText(args, kwargs)
	if not text.startswith("spelling and grammar:"):
		return False
	return text.endswith(" dialog") or " row " in text or " column " in text


def _isOutlookProofingReturnSpeech(args, kwargs):
	"""The message-title/unavailable utterance emitted while Outlook closes F7."""
	text = _speechSequenceText(args, kwargs)
	return text.startswith("untitled - message (html)") and (
		"unavailable" in text or " row " in text or " column " in text
	)


def _isOutlookMessageBodyPromptSpeech(args, kwargs):
	"""Outlook's premature body-entry prompt while the completion dialog remains."""
	return _speechSequenceText(args, kwargs).startswith(
		"you are now in the message body"
	)


def _isOutlookMessageTitleSpeech(args, kwargs):
	"""Outlook's compose-window title emitted while the checker is closing."""
	return _speechSequenceText(args, kwargs).startswith("untitled - message (html)")


### Settings

def _addChoices(panel, sHelper):
	panel._spellGrammarSuggestionsCheckBox = sHelper.addItem(
		wx.CheckBox(
			panel,
			# Translators: Keep Word and Outlook's initial focus in spell suggestions.
			label=_("Move automatically into the proofing &suggestions list"),
		),
	)
	panel._spellGrammarSuggestionsCheckBox.SetValue(getFocusSpellSuggestions())
	panel._spellGrammarSlowCheckBox = sHelper.addItem(
		wx.CheckBox(
			panel,
			# Translators: Slow down words spoken by the Office proofing checker.
			label=_("&Slow down misspelled words and suggestions when read aloud"),
		),
	)
	panel._spellGrammarSlowCheckBox.SetValue(getSlowProofingSpeech())


def _saveChoices(panel):
	checkBox = getattr(panel, "_spellGrammarSuggestionsCheckBox", None)
	if checkBox is not None:
		setFocusSpellSuggestions(checkBox.IsChecked())
	checkBox = getattr(panel, "_spellGrammarSlowCheckBox", None)
	if checkBox is not None:
		setSlowProofingSpeech(checkBox.IsChecked())


def _refreshChoices(panel):
	checkBox = getattr(panel, "_spellGrammarSuggestionsCheckBox", None)
	if checkBox is not None:
		checkBox.SetValue(getFocusSpellSuggestions())
	checkBox = getattr(panel, "_spellGrammarSlowCheckBox", None)
	if checkBox is not None:
		checkBox.SetValue(getSlowProofingSpeech())


class SpellGrammarCheckPanel(settingsDialogs.SettingsPanel):
	"""The add-on's category in NVDA Settings."""

	# Translators: Title of the add-on's settings category.
	title = _("Spelling and Grammar Check")

	def makeSettings(self, settingsSizer):
		helper = guiHelper.BoxSizerHelper(self, sizer=settingsSizer)
		_addChoices(self, helper)
		self.updates = updater.SettingsControls(self, helper)

	def onSave(self):
		_saveChoices(self)
		self.updates.save()

	def onPanelActivated(self):
		_refreshChoices(self)
		super().onPanelActivated()


### The plug-in itself

class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	# Translators: Category for this add-on's commands in the Input Gestures dialog.
	scriptCategory = _("Spelling and Grammar Check")

	def __init__(self):
		global _pluginInstance
		super().__init__()
		_pluginInstance = self
		self._ownPanelAdded = False
		self._gestureHandlerRegistered = False
		#: The window and word the spelling checker was last answered for.
		self._lastSpellCheck = None
		self._lastSpellCheckUntil = 0.0
		#: Whether the focus was inside the spelling dialog last time we looked.
		self._inSpellDialog = False
		self._proofingCancelUntil = 0.0
		#: Whether the last proofing focus was inside its suggestions list.
		self._inSuggestionList = False
		self._pendingSuggestionFocusWindow = None
		self._pendingSpellFieldFocusWindow = None
		#: The window the end of a check was last announced for.
		self._saidCompleteFor = None
		self._saidCompleteUntil = 0.0
		self._suppressSpellOkUntil = 0.0
		#: Retrying a "complete" that arrived right after typing, before Word's own
		#: background proofing pass has caught up with what was just typed.
		self._prematureCompleteAttempts = 0
		self._prematureCompleteDeadline = 0.0
		#: Bumped every time a misspelled word is about to be announced, so a
		#: pending announcement can tell it has been superseded by a newer one.
		self._wordAnnounceGeneration = 0
		self._wordAnnouncementPending = False
		self._noSuggestionGeneration = 0
		self._noSuggestionActive = False
		# Set only while Alt+C or Alt+X is moving a no-suggestion classic issue
		# onto Word's Ignore Once button.
		self._pendingClassicIgnoreOnce = False
		# Modern Outlook Editor no longer implements the classic dialog's Alt+C
		# mnemonic. Preserve that established proofing command while focus is on a real
		# Editor replacement; the script passes Alt+C through everywhere else.
		self.bindGesture("kb:alt+c", "applyProofingSuggestion")
		self.bindGesture("kb:alt+x", "ignoreProofingIssue")
		updater.start()

		try:
			self._installSpeechHooks()
		except Exception:
			log.error("Spelling and Grammar Check: could not install speech hooks", exc_info=True)
			_unpatchAll()
			return

		try:
			inputCore.decide_executeGesture.register(_onGesture)
			self._gestureHandlerRegistered = True
		except Exception:
			log.error("Spelling and Grammar Check: could not hook input gestures", exc_info=True)

		try:
			if SpellGrammarCheckPanel not in settingsDialogs.NVDASettingsDialog.categoryClasses:
				settingsDialogs.NVDASettingsDialog.categoryClasses.append(SpellGrammarCheckPanel)
				self._ownPanelAdded = True
		except Exception:
			log.error("Spelling and Grammar Check: could not add its settings category", exc_info=True)

	def event_gainFocus(self, obj, nextHandler):
		"""Own each proofing issue as it arrives, and leave everything else to NVDA."""
		# Version 1.0.44's log proves that Word delivers this real focus event, but
		# the later polling callback did not activate it. Handle the event itself,
		# after Word has confirmed focus, with the same Alt+I command that succeeds
		# when the user presses it manually.
		if self._pendingClassicIgnoreOnce and _isClassicIgnoreOnceButton(obj):
			self._pendingClassicIgnoreOnce = False
			with _hardMute():
				nextHandler()
			try:
				import keyboardHandler

				keyboardHandler.KeyboardInputGesture.fromName("alt+i").send()
			except Exception:
				log.error("Spelling and Grammar Check: could not send Ignore Once after focus", exc_info=True)
			return
		if _handleModernEditorFocus(obj, nextHandler):
			return
		if _handleLegacySdmWordChoiceFocus(obj, nextHandler):
			return
		# The dialog saying the check has finished, asked about before the checker
		# itself: it comes up whether or not anything was found.
		if self._isSpellCheckComplete(obj):
			if self._retryPrematureComplete(obj):
				with _hardMute():
					nextHandler()
				return
			with _hardMute():
				nextHandler()
			self._finishSpellCheck(obj)
			return
		# The classic F7 spelling/grammar dialog.
		spellField = _officeSpellCheckField(obj)
		if spellField is not None:
			self._noteProofingActivity()
			suggestionList = _suggestionListOf(obj)
			startingProofing = not self._inSpellDialog
			firstSuggestion = suggestionList is not None and not self._inSuggestionList
			self._inSuggestionList = suggestionList is not None
			redirectToSpellField = startingProofing and firstSuggestion and not getFocusSpellSuggestions()
			if not self._inSpellDialog:
				self._inSpellDialog = True
				self._lastSpellCheck = None
				self._saidCompleteFor = None
				self._saidCompleteUntil = 0.0
			word = _misspelledWord(spellField)
			seen = (getattr(spellField, "windowHandle", None), word) if word is not None else None
			isNewWord = seen is not None and seen != self._lastSpellCheck
			# Outlook can send a second bare "to" focus event immediately after it
			# has supplied "suppose to". It is not a new issue: treating it as one
			# cancels the pending full-phrase announcement and makes its real
			# suggestion look unavailable.
			ignoreTrailingSupposeTo = bool(
				word is not None
				and word.casefold() == "to"
				and self._lastSpellCheck is not None
				and seen is not None
				and seen[0] == self._lastSpellCheck[0]
				and self._lastSpellCheck[1].casefold() == "suppose to"
			)
			onTheBox = _windowClassOf(obj) in _WORD_DIALOG_WINDOW_CLASSES
			waitingForWord = word is None and (
				onTheBox or startingProofing or suggestionList is not None
			)
			if isNewWord or onTheBox or suggestionList is not None:
				with _hardMute():
					nextHandler()
			else:
				nextHandler()
			if ignoreTrailingSupposeTo:
				return
			if isNewWord:
				self._wordAnnouncementPending = True
				self._noSuggestionActive = False
				self._pendingClassicIgnoreOnce = False
				self._cancelNoSuggestionAnnouncement()
				self._lastSpellCheck = seen
				self._prematureCompleteAttempts = 0
				self._prematureCompleteDeadline = 0.0
				log.debug("Spelling and Grammar Check: checker is asking about %r" % word)
				afterWord = None
				if redirectToSpellField:
					window = getattr(spellField, "windowHandle", None)
					afterWord = lambda: _focusSpellField(window)
				elif getFocusSpellSuggestions():
					window = getattr(spellField, "windowHandle", None)
					afterWord = lambda: _focusSuggestionList(
						_objectFromWindow(window) or spellField,
						expectedWord=word,
					)
				elif suggestionList is not None:
					afterWord = lambda: _announceSuggestionEntry(obj, suggestionList, firstSuggestion)
				self._pendingSuggestionFocusWindow = None
				self._pendingSpellFieldFocusWindow = None
				if startingProofing and time.monotonic() < _classicProofingPunctuationLaunchUntil:
					core.callLater(
						550,
						_announceSettledClassicIssue,
						getattr(spellField, "windowHandle", None),
						spellField,
						word,
						afterWord,
					)
				else:
					# A misspelled word containing punctuation (e.g. a doubled
					# apostrophe) can make Word re-fire this focus event several
					# times in quick succession with a different word each time,
					# as its own dialog settles on the final span. Debounce: only
					# the last one still standing after a brief pause gets spoken.
					self._wordAnnounceGeneration += 1
					myGeneration = self._wordAnnounceGeneration
					# Outlook can first publish a bare "to" and only expose the real
					# "suppose to" grammar issue after the dialog advances. The 1.0.32
					# log took more than seven seconds, so a timed hold still spoke the
					# wrong interim word. Keep this known interim state silent; the next
					# field event supplies the complete error and is announced normally.
					if word.casefold() == "to":
						# The normal post-error path enters Suggestions automatically when
						# that setting is on. Do the same for this silent interim state, so
						# the user never has to press Tab to make Outlook publish the full
						# "suppose to" issue. Stay quiet if its temporary list is empty.
						if getFocusSpellSuggestions():
							core.callLater(
								180,
								_focusSuggestionList,
								_objectFromWindow(getattr(spellField, "windowHandle", None)) or spellField,
								0,
								word,
								True,
							)
						return
					# The 1.0.35 log's non-stalled checks reached correct classification
					# in 0.43–0.56 seconds, including the former 180 ms settle wait.
					# Keep only a small yield for Outlook to publish the dialog, removing
					# 130 ms from every ordinary error announcement.
					announceDelay = 50
					core.callLater(
						announceDelay,
						self._announceMisspelledWordIfCurrent,
						myGeneration,
						word,
						afterWord,
						getattr(spellField, "windowHandle", None),
						spellField,
					)
			elif waitingForWord:
				if getFocusSpellSuggestions():
					self._pendingSuggestionFocusWindow = getattr(spellField, "windowHandle", None)
				elif redirectToSpellField:
					self._pendingSpellFieldFocusWindow = getattr(spellField, "windowHandle", None)
				self._lookAgainForTheWord(getattr(spellField, "windowHandle", None), 0)
			if (
				suggestionList is not None
				and not isNewWord
				and not self._wordAnnouncementPending
				and not redirectToSpellField
				and not waitingForWord
			):
				_announceSuggestionEntry(obj, suggestionList, firstSuggestion)
			return
		self._inSpellDialog = False
		self._wordAnnouncementPending = False
		self._noSuggestionActive = False
		self._pendingClassicIgnoreOnce = False
		self._cancelNoSuggestionAnnouncement()
		self._inSuggestionList = False
		self._pendingSuggestionFocusWindow = None
		self._pendingSpellFieldFocusWindow = None

		# Outlook returning focus straight to the message body, with no final dialog
		# at all: a check that found nothing to say has simply handed the message back.
		if (
			self._lastSpellCheck is not None
			and time.monotonic() < self._lastSpellCheckUntil
			and _isOutlookMessageBody(obj)
		):
			with _hardMute():
				nextHandler()
			if time.monotonic() < self._proofingCancelUntil:
				self._proofingCancelUntil = 0.0
				self._discardProofingSession()
				return
			self._finishSpellCheck(onTheOkButton=False)
			return
		nextHandler()

	def _noteProofingActivity(self):
		self._lastSpellCheckUntil = time.monotonic() + _PROOFING_RETURN_WINDOW

	def _discardProofingSession(self):
		if self._lastSpellCheck is None:
			return
		self._lastSpellCheck = None
		self._lastSpellCheckUntil = 0.0
		self._wordAnnouncementPending = False
		self._noSuggestionActive = False
		self._pendingClassicIgnoreOnce = False
		self._cancelNoSuggestionAnnouncement()
		self._pendingSuggestionFocusWindow = None
		self._pendingSpellFieldFocusWindow = None

	def _lookAgainForTheWord(self, window, attempt):
		"""Ask the spelling dialog again for the word it is asking about."""
		if not window or attempt >= len(_SPELL_RETRY_DELAYS):
			return
		core.callLater(_SPELL_RETRY_DELAYS[attempt], self._retryTheWord, window, attempt)

	def _retryTheWord(self, window, attempt):
		try:
			if not self._stillInSpellDialog(window):
				return
			self._noteProofingActivity()
			field = _objectFromWindow(window)
			word = _misspelledWord(field) if field is not None else None
			if word is None:
				if attempt + 1 < len(_SPELL_RETRY_DELAYS):
					self._lookAgainForTheWord(window, attempt + 1)
				else:
					self._sayTheBoxInstead()
				return
			seen = (window, word)
			if seen == self._lastSpellCheck:
				return
			self._lastSpellCheck = seen
			afterWord = None
			if self._pendingSuggestionFocusWindow == window:
				self._pendingSuggestionFocusWindow = None
				afterWord = lambda: _focusSuggestionList(
					_objectFromWindow(window) or field,
					expectedWord=word,
				)
			elif self._pendingSpellFieldFocusWindow == window:
				self._pendingSpellFieldFocusWindow = None
				afterWord = lambda: _focusSpellField(window)
			_announceMisspelledWord(
				word,
				after=afterWord,
				issueKind=_classicProofingKind(api.getFocusObject(), field),
			)
		except Exception:
			log.error("Spelling and Grammar Check: could not look again for the word", exc_info=True)

	def _stillInSpellDialog(self, window):
		try:
			focus = api.getFocusObject()
		except Exception:
			return False
		current = _officeSpellCheckField(focus)
		if current is None:
			return False
		return getattr(current, "windowHandle", None) == window

	def _sayTheBoxInstead(self):
		"""Say what NVDA would have said, when no word could be worked out at all."""
		try:
			focus = api.getFocusObject()
		except Exception:
			return
		if _windowClassOf(focus) not in _WORD_DIALOG_WINDOW_CLASSES:
			return
		with _ownSpeech():
			speech.speakObject(focus, reason=controlTypes.OutputReason.FOCUS)

	def _isSpellCheckOk(self, obj):
		if self._lastSpellCheck is None or time.monotonic() >= self._lastSpellCheckUntil:
			return False
		try:
			lastWindow = self._lastSpellCheck[0]
			window = getattr(obj, "windowHandle", None)
			return (
				window == lastWindow
				or winUser.getAncestor(window, winUser.GA_ROOT)
				== winUser.getAncestor(lastWindow, winUser.GA_ROOT)
			)
		except Exception:
			return False

	def _isSpellCheckComplete(self, obj):
		if not _isOkButton(obj):
			return False
		if self._isSpellCheckOk(obj):
			return True
		return _saysSpellCheckIsComplete(obj)

	def _announceMisspelledWordIfCurrent(
		self, generation, word, after, fieldWindow, fallbackField, attempt=0
	):
		if generation != self._wordAnnounceGeneration:
			return
		field = _objectFromWindow(fieldWindow) if fieldWindow else None
		field = field or fallbackField
		try:
			focus = api.getFocusObject()
		except Exception:
			focus = field
		issueKind = _classicProofingKind(focus, field, allowProvisional=False)
		if issueKind is None and attempt < 3:
			# These intervals reach the 3.8-second settling period recorded for the
			# no-suggestion grammar issue, without delaying an issue that Outlook has
			# actually identified as spelling or grammar.
			core.callLater(
				(750, 1200, 1700)[attempt],
				self._announceMisspelledWordIfCurrent,
				generation,
				word,
				after,
				fieldWindow,
				fallbackField,
				attempt + 1,
			)
			return
		if issueKind is None:
			# Retain the established spelling fallback only if Outlook never supplies
			# any classification evidence.
			issueKind = _classicProofingKind(focus, field)
		_announceMisspelledWord(
			word,
			after=lambda: self._afterErrorSpeech(after, issueKind),
			issueKind=issueKind,
		)

	def _afterErrorSpeech(self, after, issueKind):
		self._wordAnnouncementPending = False
		if after is not None:
			after()
		if getFocusSpellSuggestions():
			self._armNoSuggestionAnnouncement(issueKind)

	def _cancelNoSuggestionAnnouncement(self):
		self._noSuggestionGeneration += 1

	def _armNoSuggestionAnnouncement(self, issueKind):
		self._noSuggestionGeneration += 1
		generation = self._noSuggestionGeneration
		# A real suggestion cancels this as soon as Outlook exposes its text. The
		# 1.0.31 log showed real suggestions arriving before this point, while a
		# known-empty grammar row had already remained blank. Keep a short guard for
		# Outlook's transient empty row without adding the previous 1.8-second pause.
		core.callLater(1100, self._announceNoSuggestionIfCurrent, generation, issueKind)

	def _announceNoSuggestionIfCurrent(self, generation, issueKind):
		if generation != self._noSuggestionGeneration or not self._inSpellDialog:
			return
		# issueKind was established before the error itself was announced. Do not
		# traverse Outlook's dialog a second time here: that duplicate accessibility
		# read was immediately before this message and could cause an audible pause.
		# The timer remains, because Outlook can expose a temporary blank suggestion
		# row before a real suggestion appears.
		if issueKind == "grammar":
			message = _("No grammar suggestions available.")
		else:
			message = _("No spelling suggestions available.")
		self._noSuggestionActive = True
		with _ownSpeech():
			speech.speak([message])
		# An empty Suggestions control does not accept Outlook's Alt+C / Alt+X
		# proofing commands. Return to the error field so those commands continue
		# to advance the checker. This is done *after* speaking, not before: a
		# 1.0.52 log showed field.setFocus() here reliably taking 1-1.5 seconds
		# on this build (it triggers NVDA's own WinwordWindowObject/activePane
		# read, the same slow property access logged elsewhere as "Unable to
		# get activePane"), silently delaying the announcement itself. The
		# refocus is still needed for the next keypress, just not before this
		# message is heard.
		try:
			fieldWindow = self._lastSpellCheck[0] if self._lastSpellCheck is not None else 0
			if fieldWindow:
				_focusSpellField(fieldWindow)
		except Exception:
			pass

	def _retryPrematureComplete(self, obj):
		"""If F7 was pressed right after typing and Word already says "complete",
		that can be Word's background proofing pass not having caught up yet.
		Silently dismiss the dialog and press F7 again a few times before
		believing it. Returns True if a retry was scheduled (caller should stay
		quiet), False if this should be treated as a real completion.
		"""
		now = time.monotonic()
		# Outlook's completion dialog is a transient UIA shell. The log shows that
		# re-sending F7 here reopens it every three seconds and triggers NVDA's
		# watchdog recovery each time. Do not retry that Outlook shell.
		if _isInOutlookWindow(obj):
			self._prematureCompleteAttempts = 0
			self._prematureCompleteDeadline = 0.0
			return False
		justTyped = now < _classicProofingPunctuationLaunchUntil
		midRetry = self._prematureCompleteAttempts > 0 and now < self._prematureCompleteDeadline
		if not (justTyped or midRetry):
			self._prematureCompleteAttempts = 0
			self._prematureCompleteDeadline = 0.0
			return False
		if self._prematureCompleteAttempts == 0:
			self._prematureCompleteDeadline = now + _PREMATURE_COMPLETE_WINDOW
		if self._prematureCompleteAttempts >= _PREMATURE_COMPLETE_MAX_ATTEMPTS:
			self._prematureCompleteAttempts = 0
			self._prematureCompleteDeadline = 0.0
			return False
		self._prematureCompleteAttempts += 1
		try:
			obj.doAction()
		except Exception:
			log.debugWarning("Spelling and Grammar Check: could not dismiss the premature completion dialog", exc_info=True)
		core.callLater(_PREMATURE_COMPLETE_RETRY_MS, self._resendProofingLaunch)
		return True

	def _resendProofingLaunch(self):
		try:
			import keyboardHandler

			keyboardHandler.KeyboardInputGesture.fromName("f7").send()
		except Exception:
			log.debugWarning("Spelling and Grammar Check: could not resend F7 while retrying a premature completion", exc_info=True)

	def _finishSpellCheck(self, obj=None, onTheOkButton=True):
		"""Say the check is over, once, and stop everything that was following it."""
		now = time.monotonic()
		window = getattr(obj, "windowHandle", None) if obj is not None else None
		if window:
			try:
				window = winUser.getAncestor(window, winUser.GA_ROOT) or window
			except Exception:
				pass
		alreadySaid = (
			window is not None
			and window == self._saidCompleteFor
			and now < self._saidCompleteUntil
		)
		self._saidCompleteFor = window
		self._saidCompleteUntil = now + 2.0
		self._lastSpellCheck = None
		self._lastSpellCheckUntil = 0.0
		self._inSpellDialog = False
		self._proofingCancelUntil = 0.0
		self._inSuggestionList = False
		self._wordAnnouncementPending = False
		self._cancelNoSuggestionAnnouncement()
		self._pendingSuggestionFocusWindow = None
		self._pendingSpellFieldFocusWindow = None
		# Outlook can remain blocked for two seconds while it closes the checker.
		# Keep the narrow completion filter past that logged recovery, so its
		# premature message-body prompt cannot overtake the completion dialog.
		self._suppressSpellOkUntil = now + 6.0
		if obj is not None and _isInOutlookWindow(obj):
			global _outlookProofingShellUntil
			_outlookProofingShellUntil = max(_outlookProofingShellUntil, now + 5.0)
		self._prematureCompleteAttempts = 0
		self._prematureCompleteDeadline = 0.0
		if alreadySaid:
			return
		with _ownSpeech():
			speech.speak(_spellCheckCompleteSpeech(onTheOkButton))

	def terminate(self):
		global _pluginInstance
		_pluginInstance = None
		updater.stop()
		if self._ownPanelAdded:
			try:
				settingsDialogs.NVDASettingsDialog.categoryClasses.remove(SpellGrammarCheckPanel)
			except Exception:
				log.error("Spelling and Grammar Check: could not remove the settings panel", exc_info=True)
		if self._gestureHandlerRegistered:
			try:
				inputCore.decide_executeGesture.unregister(_onGesture)
			except Exception:
				log.error("Spelling and Grammar Check: could not unhook input gestures", exc_info=True)
		_unpatchAll()
		super().terminate()

	def _spellCheckIsRunning(self):
		return (
			self._lastSpellCheck is not None
			and time.monotonic() < self._lastSpellCheckUntil
		) or time.monotonic() < self._suppressSpellOkUntil

	def _installSpeechHooks(self):
		"""Put our proofing filter outside wrappers installed by other add-ons."""
		for owner, name, factory in (
			(speech.speech, "speak", self._makeSpeakWrapper),
			(speech, "speak", self._makeSpeakWrapper),
			(speech.speech, "speakObject", self._makeSpeakObjectWrapper),
			(speech, "speakObject", self._makeSpeakObjectWrapper),
		):
			current = getattr(owner, name)
			if getattr(current, "_spellGrammarCheckHook", False):
				continue
			wrapper = factory(current)
			wrapper._spellGrammarCheckHook = True
			_patch(owner, name, wrapper)

	def _makeSpeakWrapper(self, original):
		def speak(*args, **kwargs):
			if _muteDepth > 0 and _bypassDepth <= 0:
				return
			if (
				time.monotonic() < _outlookProofingShellUntil
				and _isClassicProofingChromeSpeech(args, kwargs)
			):
				return
			if _isOutlookProofingShellSpeech(args, kwargs) and (
				time.monotonic() < _outlookProofingShellUntil
				or self._spellCheckIsRunning()
			):
				return
			if (
				(time.monotonic() < _suppressSuggestionsListUntil
				or self._spellCheckIsRunning()
				or self._inSpellDialog)
				and _isSuggestionsListSpeech(args, kwargs)
			):
				return
			if (
				time.monotonic() < _outlookProofingReturnUntil
				and _isOutlookProofingReturnSpeech(args, kwargs)
			):
				return
			if self._spellCheckIsRunning() and _isOutlookMessageBodyPromptSpeech(args, kwargs):
				# The supplied 1.0.26 log reports this exact prompt arriving after
				# "Spell check is complete" but before its dialog is dismissed.
				return
			if self._spellCheckIsRunning() and _isOutlookMessageTitleSpeech(args, kwargs):
				# Outlook also emits its compose-window title before it has finished
				# closing the completion dialog; it adds no useful proofing information.
				return
			if self._spellCheckIsRunning() and _isBareUnavailable(args, kwargs):
				# The spelling dialog disabling itself on its way to saying it has
				# finished. See L{_isBareUnavailable}.
				return
			try:
				args, kwargs = _filterOutlookEditorNoise(args, kwargs)
				args, kwargs = _correctEditingGrammarLabel(args, kwargs)
			except Exception:
				log.debugWarning("Spelling and Grammar Check: could not filter proofing speech", exc_info=True)
			return original(*args, **kwargs)

		speak.__name__ = "speak"
		speak.__doc__ = getattr(original, "__doc__", None)
		return speak

	def _shouldDropSpellCheckOkObjectSpeech(self, args, kwargs):
		"""Drop NVDA's separate automatic OK-button speech after completion."""
		if time.monotonic() >= self._suppressSpellOkUntil:
			return False
		obj = kwargs.get("obj", args[0] if args else None)
		if obj is None or not _isOfficeProofingWindow(obj):
			return False
		if "reason" in kwargs:
			reason = kwargs["reason"]
		elif len(args) > 1:
			reason = args[1]
		else:
			return False
		automaticReasons = _members(controlTypes.OutputReason, ("FOCUS", "FOCUSENTERED", "CHANGE", "CARET"))
		if reason not in automaticReasons:
			return False
		try:
			return (
				obj.role == controlTypes.Role.BUTTON
				and (getattr(obj, "name", "") or "").strip().lower() in ("ok", "&ok")
			)
		except Exception:
			return False

	def _makeSpeakObjectWrapper(self, original):
		def speakObject(*args, **kwargs):
			if self._shouldDropSpellCheckOkObjectSpeech(args, kwargs):
				return
			return original(*args, **kwargs)

		speakObject.__name__ = "speakObject"
		speakObject.__doc__ = getattr(original, "__doc__", None)
		return speakObject

	@scriptHandler.script(
		# Translators: Description of a command, shown in the Input Gestures dialog.
		description=_("Applies the focused Microsoft Editor spelling or grammar suggestion"),
	)
	def script_applyProofingSuggestion(self, gesture):
		"""Apply a focused replacement, or ignore a classic issue with none."""
		# The latest log showed grammar items with no replacement exposing only
		# "Ignore Once, Alt+I" / "Ignore Rule, Alt+G". In that state Word has no
		# Change command, so preserve the established Alt+C and Alt+X workflow by
		# advancing with Ignore Once.
		if self._inSpellDialog and self._noSuggestionActive:
			try:
				if _activateClassicIgnoreOnce(api.getFocusObject()):
					return
				import keyboardHandler
				keyboardHandler.KeyboardInputGesture.fromName("alt+i").send()
				return
			except Exception:
				log.error("Spelling and Grammar Check: could not ignore proofing issue without suggestions", exc_info=True)
		try:
			focus = api.getFocusObject()
			root = _modernEditorAncestorRoot(focus)
			issue = _modernEditorIssue(root) if root is not None else None
			suggestion = (
				_modernEditorSuggestionForKeyboardFocus(focus, issue[3])
				if issue is not None
				else None
			)
		except Exception:
			log.debugWarning("Spelling and Grammar Check: could not inspect the Alt+C proofing focus", exc_info=True)
			suggestion = None
		if suggestion is None:
			gesture.send()
			return
		try:
			import keyboardHandler

			keyboardHandler.KeyboardInputGesture.fromName("enter").send()
		except Exception:
			log.error("Spelling and Grammar Check: could not apply the focused Editor suggestion", exc_info=True)

	@scriptHandler.script(
		# Translators: Description of a command, shown in the Input Gestures dialog.
		description=_("Ignores the current classic spelling or grammar issue"),
	)
	def script_ignoreProofingIssue(self, gesture):
		"""Make Alt+X advance a classic issue instead of re-reading its suggestion."""
		try:
			if self._inSpellDialog and _activateClassicIgnoreOnce(api.getFocusObject()):
				return
		except Exception:
			log.debugWarning("Spelling and Grammar Check: could not inspect the Alt+X proofing focus", exc_info=True)
		gesture.send()

	@scriptHandler.script(
		# Translators: Description of a command, shown in the Input Gestures dialog.
		description=_("Checks for Spelling and Grammar Check updates"),
	)
	def script_checkForUpdates(self, gesture):
		updater.checkForUpdates()
