# -*- coding: utf-8 -*-
# Update checking shared by the NVDA add-ons published at https://github.com/joshknnd1982
# Copyright (C) 2026 Josh Kennedy <joshknnd1982@gmail.com>
# This file is covered by the GNU General Public License, version 2 or later.
# It follows ClassicSpeech's update check (https://github.com/joshknnd1982/classicspeech-nvda,
# also GPL 2) and the JAWS Migration Assistant's, so all of these add-ons behave the same way.

"""Check GitHub for a newer version of this add-on, download it and install it with NVDA.

The same file ships unchanged inside each add-on and works out which add-on it
belongs to from where it is installed. Releases are published on the GitHub
repository named by the ``url`` in that add-on's manifest, such as
https://github.com/joshknnd1982/copyURL. A check asks GitHub for the latest
release and compares its tag, such as ``v1.9.3``, with the installed version.

When the release is newer, the add-on shows what is new in a dialog whose
notes can be read line by line, and offers to download the ``.nvda-addon``
file. The download must match the release's ``.sha256`` file, or failing that
the SHA-256 GitHub records for the file. NVDA's own add-on installation then
asks the user to confirm, installs the new version in place of this one,
keeping the add-on's settings, and offers to restart NVDA.

Checks run in the background. An automatic check runs at most once a day, a
little after NVDA starts, and only speaks up when there is an update. A check
can also be started from NVDA's Tools menu, under Check for add-on updates,
from the add-on's settings panel, or with a gesture assigned in Input Gestures.

The add-on's GlobalPlugin calls ``start()`` in ``__init__`` and ``stop()`` in
``terminate()``. Written for Python 3.7, so it runs in every NVDA version these
add-ons support, back to 2019.3. Keep every copy identical; the canonical one
is ``shared/updater.py`` next to the add-on repositories.
"""

import hashlib
import json
import os
import re
import shutil
import tempfile
import threading
import time

import addonHandler
import globalVars
import gui
import ui
import wx
from gui import guiHelper, settingsDialogs
from logHandler import log

try:
	addonHandler.initTranslation()
except Exception:
	# Not running from an installed add-on, such as in tests; NVDA's own _ is used.
	pass

API_URL = "https://api.github.com/repos/{repository}/releases/latest"
RELEASES_URL = "https://github.com/{repository}/releases"
CHECK_TIMEOUT_SECONDS = 20
DOWNLOAD_TIMEOUT_SECONDS = 120
MAX_DOWNLOAD_BYTES = 50 * 1024 * 1024
MAX_RESPONSE_BYTES = 5 * 1024 * 1024
AUTOMATIC_CHECK_DELAY_MS = 30 * 1000
AUTOMATIC_CHECK_INTERVAL_SECONDS = 24 * 60 * 60
ADDON_EXTENSION = ".nvda-addon"
CHECKSUM_EXTENSION = ".sha256"
#: Wide and tall enough for a paragraph of release notes without scrolling.
NOTES_SIZE = (620, 260)
#: Every add-on with this file puts its menu item in one Tools submenu, found by this label.
TOOLS_SUBMENU_LABEL = "Check for add-on &updates"

_GITHUB_REPOSITORY = re.compile(r"^https?://(?:www\.)?github\.com/([\w.-]+)/([\w.-]+?)(?:\.git)?/?$", re.IGNORECASE)
_CHECKSUM = re.compile(r"^[0-9a-fA-F]{64}$")


class UpdateError(Exception):
	"""A check or download failed; the message can be shown to the user."""


class AddonInfo(object):
	"""What the update check needs to know about the add-on this file belongs to."""

	def __init__(self, name, product, version, repository):
		#: The add-on's id from its manifest, such as ``copyURL``.
		self.name = name
		#: The add-on's name as the user knows it, its manifest summary, such as ``Copy URL``.
		self.product = product
		self.version = version
		#: ``owner/name`` on GitHub, or None when the manifest names no GitHub repository.
		self.repository = repository


class Release(object):
	def __init__(self, version, name, notes, pageUrl, addonName="", addonUrl="", addonSize=0, checksumUrl="", sha256=""):
		self.version = version
		self.name = name
		self.notes = notes
		self.pageUrl = pageUrl
		self.addonName = addonName
		self.addonUrl = addonUrl
		self.addonSize = addonSize
		#: The release's ``.sha256`` file for the add-on file, if it has one.
		self.checksumUrl = checksumUrl
		#: The SHA-256 GitHub recorded when the add-on file was uploaded, if it did.
		self.sha256 = sha256


# -- pure helpers ---------------------------------------------------------------------------


def githubRepository(url):
	"""Return ``owner/name`` for a GitHub repository URL, otherwise None."""
	match = _GITHUB_REPOSITORY.match(str(url or "").strip())
	if not match:
		return None
	return "%s/%s" % (match.group(1), match.group(2))


def parseVersion(text):
	"""Return a version such as ``1.0`` or ``v1.0.2`` as a tuple of numbers, or None."""
	text = str(text or "").strip()
	if text[:1] in ("v", "V"):
		text = text[1:]
	parts = text.split(".")
	if not text or not all(part.isdigit() for part in parts):
		return None
	return tuple(int(part) for part in parts)


def isNewer(candidate, installed):
	"""True when version text ``candidate`` is newer than ``installed``."""
	new, old = parseVersion(candidate), parseVersion(installed)
	if new is None or old is None:
		return False
	width = max(len(new), len(old))
	return new + (0,) * (width - len(new)) > old + (0,) * (width - len(old))


def releaseFromGithub(data, product):
	"""The Release described by GitHub's JSON for a release, or None."""
	if not isinstance(data, dict) or data.get("draft") or data.get("prerelease"):
		return None
	tag = str(data.get("tag_name") or "")
	version = tag[1:] if tag[:1] in ("v", "V") else tag
	if parseVersion(version) is None:
		return None
	addon = checksum = None
	assets = [asset for asset in data.get("assets") or () if isinstance(asset, dict)]
	for asset in assets:
		if str(asset.get("name") or "").lower().endswith(ADDON_EXTENSION):
			addon = asset
			break
	sha256 = ""
	if addon is not None:
		wanted = str(addon.get("name")) + CHECKSUM_EXTENSION
		checksum = next((asset for asset in assets if asset.get("name") == wanted), None)
		digest = str(addon.get("digest") or "")
		if digest[:7].lower() == "sha256:" and _CHECKSUM.match(digest[7:]):
			sha256 = digest[7:].lower()
	return Release(
		version=version,
		name=str(data.get("name") or "%s %s" % (product, version)),
		notes=str(data.get("body") or ""),
		pageUrl=str(data.get("html_url") or ""),
		addonName=str(addon.get("name") or "") if addon else "",
		addonUrl=str(addon.get("browser_download_url") or "") if addon else "",
		addonSize=int(addon.get("size") or 0) if addon else 0,
		checksumUrl=str(checksum.get("browser_download_url") or "") if checksum else "",
		sha256=sha256,
	)


def checksumFromFile(text):
	"""The SHA-256 in a ``.sha256`` file (``<hex>  <file name>``), or None."""
	words = str(text or "").split()
	if words and _CHECKSUM.match(words[0]):
		return words[0].lower()
	return None


def notesAsText(notes, limit=None):
	"""Release notes as plain text: Markdown headings, bullets, emphasis and link targets removed."""
	text = str(notes or "").replace("\r\n", "\n")
	text = re.sub(r"\[([^\]]+)\]\([^)]*\)", r"\1", text)
	text = re.sub(r"(?m)^[ \t]{0,3}#{1,6}[ \t]*", "", text)
	text = re.sub(r"(?m)^[ \t]*[-*+][ \t]+", "- ", text)
	text = text.replace("**", "").replace("__", "").replace("`", "")
	text = re.sub(r"\n{3,}", "\n\n", text).strip()
	if limit is not None and len(text) > limit:
		text = text[:limit].rsplit(" ", 1)[0].rstrip() + "..."
	return text


def isDue(lastCheck, now=None, interval=AUTOMATIC_CHECK_INTERVAL_SECONDS):
	"""True when an automatic check is due, ``lastCheck`` being seconds since the epoch."""
	now = time.time() if now is None else now
	try:
		lastCheck = float(lastCheck)
	except (TypeError, ValueError):
		return True
	return lastCheck <= 0 or now - lastCheck >= interval or lastCheck > now


# -- the running add-on ------------------------------------------------------------------------


def addonInfo():
	"""The AddonInfo for the add-on this file belongs to, or None when it is not running as an add-on."""
	try:
		manifest = addonHandler.getCodeAddon().manifest
		name = str(manifest["name"])
		info = AddonInfo(
			name=name,
			product=str(manifest["summary"] or name),
			version=str(manifest["version"]),
			repository=githubRepository(manifest.get("url")),
		)
	except Exception:
		return None
	return info


def _productName():
	info = addonInfo()
	# Translators: Stands in for the add-on's name when it can't be read from its manifest.
	return info.product if info else _("This add-on")


def _isSecure():
	return bool(getattr(globalVars.appArgs, "secure", False))


# -- settings ----------------------------------------------------------------------------------
# Kept in a small file of their own in NVDA's user configuration folder, so that
# no configuration profile, such as one NVDA switches to in Outlook, gets them.


def _statePath(name):
	return os.path.join(globalVars.appArgs.configPath, "addonUpdates", name + ".json")


def _readState(name):
	try:
		with open(_statePath(name), encoding="utf-8") as stream:
			data = json.load(stream)
	except (IOError, OSError):
		return {}
	except Exception:
		log.debugWarning("Could not read the update settings of %s" % name, exc_info=True)
		return {}
	return data if isinstance(data, dict) else {}


def _writeState(name, **changes):
	data = _readState(name)
	data.update(changes)
	path = _statePath(name)
	try:
		if not os.path.isdir(os.path.dirname(path)):
			os.makedirs(os.path.dirname(path))
		with open(path, "w", encoding="utf-8") as stream:
			json.dump(data, stream, indent="\t", sort_keys=True)
	except Exception:
		log.error("Could not save the update settings of %s" % name, exc_info=True)


def automaticChecksEnabled():
	info = addonInfo()
	if info is None:
		return False
	value = _readState(info.name).get("checkForUpdatesAutomatically", True)
	return value if isinstance(value, bool) else True


def setAutomaticChecksEnabled(enabled):
	info = addonInfo()
	if info is not None and bool(enabled) != automaticChecksEnabled():
		_writeState(info.name, checkForUpdatesAutomatically=bool(enabled))


def _lastCheck(name):
	value = _readState(name).get("lastUpdateCheck", 0)
	return value if isinstance(value, (int, float)) else 0


def _rememberCheck(name, now=None):
	_writeState(name, lastUpdateCheck=int(time.time() if now is None else now))


# -- network -----------------------------------------------------------------------------------


def _headers(info):
	return {
		"Accept": "application/vnd.github+json",
		"User-Agent": "%s/%s (NVDA add-on; +https://github.com/%s)" % (info.name, info.version, info.repository),
	}


def _get(url, headers, timeout):
	"""GET ``url``. Returns ``(status, chunks, close)``: the HTTP status code,
	an iterator over the body, and a function that closes the connection.

	Uses the requests library NVDA has shipped since 2023.2, and the standard
	library in older versions.
	"""
	try:
		import requests
	except Exception:
		requests = None
	if requests is not None:
		response = requests.get(url, headers=headers, timeout=timeout, stream=True)
		return response.status_code, response.iter_content(chunk_size=128 * 1024), response.close
	import urllib.error
	import urllib.request

	try:
		response = urllib.request.urlopen(urllib.request.Request(url, headers=headers), timeout=timeout)
	except urllib.error.HTTPError as error:
		return error.code, iter(()), error.close
	return response.getcode(), iter(lambda: response.read(128 * 1024), b""), response.close


def _read(url, headers, timeout, limit=MAX_RESPONSE_BYTES):
	"""GET ``url``; returns ``(status, body)``, the body only for a 200 answer."""
	status, chunks, close = _get(url, headers, timeout)
	try:
		if status != 200:
			return status, b""
		body = b""
		for chunk in chunks:
			body += chunk
			if len(body) > limit:
				raise UpdateError(_("GitHub's answer was larger than expected."))
		return status, body
	finally:
		close()


def fetchLatestRelease(info):
	"""The latest release of the add-on's repository on GitHub. Raises UpdateError."""
	try:
		status, body = _read(API_URL.format(repository=info.repository), _headers(info), CHECK_TIMEOUT_SECONDS)
	except UpdateError:
		raise
	except Exception as error:
		raise UpdateError(_("GitHub could not be reached. Check your internet connection.")) from error
	if status == 404:
		raise UpdateError(_("There are no {product} releases on GitHub yet.").format(product=info.product))
	if status != 200:
		raise UpdateError(_("GitHub answered with error {code}.").format(code=status))
	try:
		release = releaseFromGithub(json.loads(body.decode("utf-8")), info.product)
	except Exception as error:
		raise UpdateError(_("GitHub's answer could not be read.")) from error
	if release is None:
		raise UpdateError(
			_("The latest release on GitHub has no version number {product} understands.").format(product=info.product)
		)
	return release


def downloadRelease(release, info, folder):
	"""Download the release's add-on file into ``folder`` and check it. Returns its path.

	Raises UpdateError when the download fails, is too large, or does not match
	the release's checksum.
	"""
	if not release.addonUrl:
		raise UpdateError(_("The release has no add-on file."))
	if not release.checksumUrl and not release.sha256:
		raise UpdateError(_("The release has no checksum file, so its add-on file can't be checked."))
	if release.addonSize > MAX_DOWNLOAD_BYTES:
		raise UpdateError(_("The add-on file is larger than expected."))
	headers = _headers(info)
	name = os.path.basename(release.addonName) or info.name + ADDON_EXTENSION
	if not name.lower().endswith(ADDON_EXTENSION):
		name += ADDON_EXTENSION
	path = os.path.join(folder, name)
	expected = release.sha256
	if release.checksumUrl:
		try:
			status, body = _read(release.checksumUrl, headers, CHECK_TIMEOUT_SECONDS, limit=64 * 1024)
			expected = checksumFromFile(body.decode("utf-8", "replace")) if status == 200 else None
		except Exception as error:
			raise UpdateError(_("The checksum file could not be downloaded.")) from error
		if expected is None:
			raise UpdateError(_("The checksum file could not be read."))
	digest = hashlib.sha256()
	size = 0
	try:
		status, chunks, close = _get(release.addonUrl, headers, DOWNLOAD_TIMEOUT_SECONDS)
		try:
			if status != 200:
				raise UpdateError(_("GitHub answered with error {code}.").format(code=status))
			with open(path, "wb") as stream:
				for chunk in chunks:
					if not chunk:
						continue
					size += len(chunk)
					if size > MAX_DOWNLOAD_BYTES:
						raise UpdateError(_("The add-on file is larger than expected."))
					digest.update(chunk)
					stream.write(chunk)
		finally:
			close()
	except UpdateError:
		_removeFile(path)
		raise
	except Exception as error:
		_removeFile(path)
		raise UpdateError(_("The add-on file could not be downloaded.")) from error
	if digest.hexdigest() != expected:
		_removeFile(path)
		raise UpdateError(_("The downloaded file does not match its checksum, so it was deleted."))
	return path


def _removeFile(path):
	try:
		os.remove(path)
	except OSError:
		pass


# -- the offer ---------------------------------------------------------------------------------


class UpdateOfferDialog(wx.Dialog):
	"""What is available, what is new in it, and what happens next.

	A message box speaks its whole text once and there is nothing to move through
	afterwards, so the "What's new" part of a release would go by in one breath.
	Here the notes are a read-only multiline box: NVDA treats it as text, so it can
	be read line by line, word by word or character by character, reviewed,
	selected and copied. Focus starts in it, and Tab reaches the buttons.
	"""

	def __init__(self, parent, title, summary, notes, question="", installLabel="", closeLabel=""):
		super(UpdateOfferDialog, self).__init__(parent, title=title, style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
		sizer = wx.BoxSizer(wx.VERTICAL)
		if summary:
			sizer.Add(self._paragraph(summary), flag=wx.BOTTOM, border=6)
		# Translators: The label of the box with a release's notes in an update offer.
		sizer.Add(wx.StaticText(self, label=_("&What's new:")))
		self.notes = wx.TextCtrl(self, value=notes or "", style=wx.TE_MULTILINE | wx.TE_READONLY, size=NOTES_SIZE)
		sizer.Add(self.notes, proportion=1, flag=wx.EXPAND | wx.TOP, border=2)
		if question:
			sizer.Add(self._paragraph(question), flag=wx.TOP, border=6)
		buttons = wx.BoxSizer(wx.HORIZONTAL)
		self.installButton = None
		if installLabel:
			self.installButton = wx.Button(self, id=wx.ID_YES, label=installLabel)
			buttons.Add(self.installButton, flag=wx.RIGHT, border=8)
			self.installButton.Bind(wx.EVT_BUTTON, lambda event: self.EndModal(wx.ID_YES))
			self.installButton.SetDefault()
		# Translators: The button that closes an update offer.
		self.closeButton = wx.Button(self, id=wx.ID_CANCEL, label=closeLabel or _("&Close"))
		self.closeButton.Bind(wx.EVT_BUTTON, lambda event: self.EndModal(wx.ID_CANCEL))
		buttons.Add(self.closeButton)
		sizer.Add(buttons, flag=wx.TOP | wx.ALIGN_RIGHT, border=8)
		outer = wx.BoxSizer(wx.VERTICAL)
		outer.Add(sizer, proportion=1, flag=wx.EXPAND | wx.ALL, border=getattr(guiHelper, "BORDER_FOR_DIALOGS", 10))
		self.SetSizerAndFit(outer)
		self.CentreOnScreen()
		self.Bind(wx.EVT_CHAR_HOOK, self._onCharHook)
		self.notes.SetFocus()
		self.notes.SetInsertionPoint(0)

	def _paragraph(self, text):
		label = wx.StaticText(self, label=text)
		try:
			label.Wrap(NOTES_SIZE[0])
		except Exception:
			pass
		return label

	def _onCharHook(self, event):
		if event.GetKeyCode() == wx.WXK_ESCAPE:
			self.EndModal(wx.ID_CANCEL)
			return
		event.Skip()


def showUpdateOffer(title, summary, notes, question="", installLabel="", closeLabel=""):
	"""Show the dialog and return the button chosen.

	It waits for the user, so it runs only from wx's event loop (``wx.CallAfter``),
	never inside NVDA's core queue, which it would freeze.
	"""
	gui.mainFrame.prePopup()
	try:
		dialog = UpdateOfferDialog(
			gui.mainFrame, title, summary, notes, question=question, installLabel=installLabel, closeLabel=closeLabel
		)
		try:
			return dialog.ShowModal()
		finally:
			dialog.Destroy()
	finally:
		gui.mainFrame.postPopup()


# -- checking from NVDA ------------------------------------------------------------------------


class UpdateChecker(object):
	"""Runs checks and downloads in the background and talks to the user in NVDA."""

	def __init__(self):
		self._busy = False
		self._stopped = False
		self._timer = None

	def stop(self):
		self._stopped = True
		timer, self._timer = self._timer, None
		if timer is not None:
			try:
				timer.Stop()
			except Exception:
				pass

	def scheduleAutomaticCheck(self):
		"""Check a little after NVDA starts, if automatic checks are on and a day has passed."""
		info = addonInfo()
		if info is None or info.repository is None or _isSecure() or not automaticChecksEnabled():
			return
		if not isDue(_lastCheck(info.name)):
			return
		self._timer = wx.CallLater(AUTOMATIC_CHECK_DELAY_MS, self.check, manual=False)

	def check(self, manual=True):
		self._timer = None
		if self._stopped:
			return
		info = addonInfo()
		product = info.product if info else _productName()
		if self._busy:
			if manual:
				# Translators: Spoken when an update check is asked for while one is running.
				ui.message(_("A {product} update check is already running.").format(product=product))
			return
		if info is None or info.repository is None:
			if manual:
				# Never block the caller, which may be NVDA's core queue: see _message.
				wx.CallAfter(
					self._message,
					_(
						"This copy of {product} doesn't know where its updates are published, so it can't check for them."
					).format(product=product),
					"error",
				)
			return
		if not manual and not automaticChecksEnabled():
			return
		self._busy = True
		if manual:
			# Translators: Spoken when a manual update check starts.
			ui.message(_("Checking for {product} updates").format(product=product))
		self._run(lambda: fetchLatestRelease(info), lambda outcome: self._checked(outcome, info, manual))

	def _run(self, work, done):
		"""Run ``work`` in a background thread, then ``done(result or UpdateError)`` in NVDA's main thread."""

		def target():
			try:
				outcome = work()
			except UpdateError as error:
				outcome = error
			except Exception as error:
				log.debug("Add-on update check failed", exc_info=True)
				outcome = UpdateError(str(error) or error.__class__.__name__)
			wx.CallAfter(done, outcome)

		thread = threading.Thread(target=target, name="addonUpdates")
		thread.daemon = True
		thread.start()

	def _checked(self, outcome, info, manual):
		self._busy = False
		if self._stopped:
			return
		if isinstance(outcome, UpdateError):
			log.info("%s: update check failed: %s" % (info.product, outcome))
			if manual:
				self._message(
					_("{product} could not check for updates. {reason}").format(product=info.product, reason=outcome),
					"error",
				)
			return
		_rememberCheck(info.name)
		release = outcome
		if not isNewer(release.version, info.version):
			if manual:
				self._message(
					_("{product} is up to date. You have version {version}, the latest release.").format(
						product=info.product,
						version=info.version,
					)
				)
			return
		log.info("%s: version %s is available (installed: %s)" % (info.product, release.version, info.version))
		self._offer(release, info)

	def _offer(self, release, info):
		summary = _("{product} {new} is available. You have version {installed}.").format(
			product=info.product,
			new=release.version,
			installed=info.version,
		)
		notes = notesAsText(release.notes)
		if not notes:
			# Translators: Shown in the What's new box for a release with no notes.
			notes = _("This release has no notes.")
		# Translators: The title of the dialog offering an add-on update.
		title = _("{product} update").format(product=info.product)
		# While the offer is open, another check must not open a second one.
		self._busy = True
		try:
			if not release.addonUrl:
				showUpdateOffer(
					title,
					summary,
					notes,
					question=_("This release has no add-on file to install. Download it from {url}").format(
						url=release.pageUrl or RELEASES_URL.format(repository=info.repository)
					),
					closeLabel=_("&Close"),
				)
				return
			answer = showUpdateOffer(
				title,
				summary,
				notes,
				question=_(
					"Download and install it now? NVDA asks you to confirm the installation, "
					"then offers to restart. Your {product} settings are kept."
				).format(product=info.product),
				# Translators: The button in an update offer that downloads and installs the update.
				installLabel=_("&Download and install"),
				# Translators: The button in an update offer that closes it without updating.
				closeLabel=_("&Not now"),
			)
		finally:
			self._busy = False
		if answer == wx.ID_YES and not self._stopped:
			self._download(release, info)

	def _download(self, release, info):
		folder = tempfile.mkdtemp(prefix=info.name + "-update-")
		self._busy = True
		# Translators: Spoken when an update starts downloading.
		ui.message(_("Downloading {product} {version}").format(product=info.product, version=release.version))
		self._run(lambda: downloadRelease(release, info, folder), lambda outcome: self._downloaded(outcome, info, folder))

	def _downloaded(self, outcome, info, folder):
		self._busy = False
		try:
			if self._stopped:
				return
			if isinstance(outcome, UpdateError):
				log.info("%s: update download failed: %s" % (info.product, outcome))
				self._message(
					_("{product} could not download the update. {reason}").format(product=info.product, reason=outcome),
					"error",
				)
				return
			try:
				installWithNVDA(outcome)
			except Exception:
				self._message(_("NVDA could not install the update. Details are in the NVDA log."), "error")
		finally:
			# NVDA has copied the add-on into its add-ons folder, or the installation was cancelled.
			shutil.rmtree(folder, ignore_errors=True)

	def _message(self, message, kind="information"):
		"""Show a message box. It waits for the user, so it runs only from wx's event loop
		(``wx.CallAfter``), never inside NVDA's core queue, which it would freeze."""
		icon = wx.ICON_ERROR if kind == "error" else wx.ICON_INFORMATION
		title = _("{product} update").format(product=_productName())
		gui.mainFrame.prePopup()
		try:
			wx.MessageBox(message, title, wx.OK | icon, gui.mainFrame)
		finally:
			gui.mainFrame.postPopup()


def installWithNVDA(path):
	"""Hand a downloaded add-on file to NVDA, which confirms, installs and offers a restart."""
	try:
		from gui import addonGui

		addonGui.handleRemoteAddonInstall(path)
	except Exception:
		log.error("NVDA could not install %s" % path, exc_info=True)
		raise


# -- the Tools menu ----------------------------------------------------------------------------
# One submenu holds an item for each add-on that has this file. Whichever add-on
# starts first adds the submenu and whichever stops last removes it; they find it
# by its label, so nothing else has to be shared between them.


def _submenuLabels():
	# Translators: The NVDA Tools submenu listing the add-ons that can check GitHub for updates.
	return (_(TOOLS_SUBMENU_LABEL), TOOLS_SUBMENU_LABEL)


def _findSubmenuItem(toolsMenu):
	labels = _submenuLabels()
	for item in toolsMenu.GetMenuItems():
		if item.GetSubMenu() is not None and item.GetItemLabel() in labels:
			return item
	return None


def _addMenuItem(product):
	toolsMenu = gui.mainFrame.sysTrayIcon.toolsMenu
	submenuItem = _findSubmenuItem(toolsMenu)
	if submenuItem is None:
		submenuItem = toolsMenu.AppendSubMenu(wx.Menu(), _submenuLabels()[0])
	submenu = submenuItem.GetSubMenu()
	label = product + "..."
	# Keep the add-ons in alphabetical order, whatever order NVDA loads them in.
	position = 0
	for item in submenu.GetMenuItems():
		if item.GetItemLabelText().lower() > label.lower():
			break
		position += 1
	item = submenu.Insert(
		position,
		wx.ID_ANY,
		label,
		# Translators: The help text of an add-on's item in the Check for add-on updates menu.
		_("Checks GitHub for a newer version of {product}").format(product=product),
	)
	gui.mainFrame.sysTrayIcon.Bind(wx.EVT_MENU, _onMenuItem, item)
	return item


def _removeMenuItem(item):
	sysTrayIcon = gui.mainFrame.sysTrayIcon
	try:
		sysTrayIcon.Unbind(wx.EVT_MENU, source=item, handler=_onMenuItem)
	except Exception:
		log.debugWarning("Could not unbind the update menu item", exc_info=True)
	submenu = item.GetMenu()
	submenu.DestroyItem(item)
	if submenu.GetMenuItemCount() == 0:
		submenuItem = _findSubmenuItem(sysTrayIcon.toolsMenu)
		if submenuItem is not None and submenuItem.GetSubMenu().GetMenuItemCount() == 0:
			sysTrayIcon.toolsMenu.DestroyItem(submenuItem)


def _onMenuItem(event):
	checkForUpdates()


# -- settings panel ----------------------------------------------------------------------------


class SettingsControls(object):
	"""The update settings for an add-on's NVDA settings panel: whether to check
	automatically, and a button that checks now.

	Create it at the end of the panel's ``makeSettings`` with the panel's
	BoxSizerHelper, and call ``save()`` from its ``onSave``.
	"""

	def __init__(self, panel, helper):
		self.automatic = helper.addItem(
			wx.CheckBox(
				panel,
				# Translators: A checkbox in an add-on's settings panel.
				label=_("Check for {product} &updates automatically").format(product=_productName()),
			)
		)
		self.automatic.SetValue(automaticChecksEnabled())
		# Translators: A button in an add-on's settings panel.
		self.checkNow = helper.addItem(wx.Button(panel, label=_("Check for updates &now")))
		self.checkNow.Bind(wx.EVT_BUTTON, lambda event: checkForUpdates())

	def save(self):
		setAutomaticChecksEnabled(self.automatic.GetValue())


class UpdateSettingsPanel(settingsDialogs.SettingsPanel):
	"""A settings panel holding only the update settings, for an add-on with no settings of its own."""

	title = _productName()

	def makeSettings(self, settingsSizer):
		helper = guiHelper.BoxSizerHelper(self, sizer=settingsSizer)
		self.updates = SettingsControls(self, helper)

	def onSave(self):
		self.updates.save()


# -- what the add-on calls ---------------------------------------------------------------------

_checker = None
_menuItem = None
_ownPanel = False


def start(settingsPanel=False):
	"""Call from the GlobalPlugin's ``__init__``. Schedules the automatic check and
	adds the add-on to the Tools menu. With ``settingsPanel``, also adds
	UpdateSettingsPanel to NVDA's settings, for an add-on without a panel of its own.
	"""
	global _checker, _menuItem, _ownPanel
	if _isSecure() or _checker is not None:
		return
	_checker = UpdateChecker()
	try:
		_checker.scheduleAutomaticCheck()
	except Exception:
		log.error("Could not schedule the automatic update check", exc_info=True)
	try:
		_menuItem = _addMenuItem(_productName())
	except Exception:
		log.error("Could not add the update check to the Tools menu", exc_info=True)
	if settingsPanel:
		try:
			categories = settingsDialogs.NVDASettingsDialog.categoryClasses
			if UpdateSettingsPanel not in categories:
				categories.append(UpdateSettingsPanel)
				_ownPanel = True
		except Exception:
			log.error("Could not add the update settings panel", exc_info=True)


def stop():
	"""Call from the GlobalPlugin's ``terminate``."""
	global _checker, _menuItem, _ownPanel
	checker, _checker = _checker, None
	if checker is not None:
		checker.stop()
	item, _menuItem = _menuItem, None
	if item is not None:
		try:
			_removeMenuItem(item)
		except Exception:
			log.debugWarning("Could not remove the update check from the Tools menu", exc_info=True)
	if _ownPanel:
		_ownPanel = False
		try:
			settingsDialogs.NVDASettingsDialog.categoryClasses.remove(UpdateSettingsPanel)
		except ValueError:
			pass


def checkForUpdates():
	"""Check now and say what was found, even that the add-on is up to date."""
	if _isSecure():
		return
	checker = _checker
	if checker is None:
		# Not started, such as while NVDA is being set up; check anyway.
		checker = UpdateChecker()
	try:
		checker.check(manual=True)
	except Exception:
		log.error("Add-on update check failed", exc_info=True)
