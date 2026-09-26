"""NVDA's speech in Outlook reaches Outlook only when it says "misspelled".

python -m unittest discover -s tests

Version 1.0.71's speech hook asked Outlook about its caret (the message body's role
and states, then the formatting at the caret twice) before every utterance, while the
focus was in a message body, to see whether "misspelled" should be "Grammar error".
A tester pressed Control+C in an Outlook message. Clipspeak said "Copied selection to
clipboard", Outlook was busy copying, and NVDA's main thread waited in that hook, in
NVDA's UIA Word document role (wordDocument _get_mathMl), for 9.7 seconds. JAWS said it
at once. Now the hook looks at the speech first, and only speech that says "misspelled"
is looked at in Outlook.

The plugin is loaded outside NVDA with the stand-ins from shared/test_addons.py.
"""

import importlib
import os
import sys
import types
import unittest
from unittest import mock

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
SHARED = os.path.join(os.path.dirname(REPO), "shared")
sys.path.insert(0, SHARED)
import test_addons  # noqa: E402  (installs the NVDA stand-ins)


def loadPlugin():
	for name in list(sys.modules):
		if name.split(".")[0] in test_addons.MockModules.NAMES or name == "globalPlugins" or name.startswith("globalPlugins."):
			del sys.modules[name]
	manifest = os.path.join(REPO, "addon", "manifest.ini")
	sys.modules["addonHandler"].current = test_addons.test_updater.FakeAddon(
		test_addons.manifestValue(manifest, "name"),
		test_addons.manifestValue(manifest, "summary"),
		test_addons.manifestValue(manifest, "version"),
		test_addons.manifestValue(manifest, "url"),
	)
	globalPlugins = types.ModuleType("globalPlugins")
	globalPlugins.__path__ = [os.path.join(REPO, "addon", "globalPlugins")]
	sys.modules["globalPlugins"] = globalPlugins
	return importlib.import_module("globalPlugins.spellGrammarCheck")


class Untouchable:
	"""Outlook's message body, busy copying: any question put to it fails the test."""

	def __getattr__(self, name):
		raise AssertionError("asked Outlook for %s" % name)


class GrammarLabel(unittest.TestCase):
	@classmethod
	def setUpClass(cls):
		cls.module = loadPlugin()

	def setUp(self):
		self.focusAsked = []
		self.grammarAsked = []
		self.grammarError = False
		module = self.module

		def getFocusObject():
			self.focusAsked.append(True)
			return self.focus

		def caretHasGrammarError(obj):
			self.grammarAsked.append(obj)
			return self.grammarError

		self.focus = Untouchable()
		patches = (
			mock.patch.object(module.api, "getFocusObject", getFocusObject),
			mock.patch.object(module, "_caretHasGrammarError", caretHasGrammarError),
		)
		for patch in patches:
			patch.start()
			self.addCleanup(patch.stop)

	def test_copiedSelectionNeverAsksOutlook(self):
		sequence = ["Copied selection to clipboard"]
		args, kwargs = self.module._correctEditingGrammarLabel((sequence,), {})
		self.assertIs(args[0], sequence, "the speech goes on unchanged")
		self.assertEqual(self.focusAsked, [], "not even the focus is looked at")
		self.assertEqual(self.grammarAsked, [])

	def test_typedCharactersNeverAskOutlook(self):
		for sequence in (["a"], ["space"], ["Miller Moss out of Louisville. selected"], [object(), "misspell"]):
			with self.subTest(sequence=sequence):
				self.module._correctEditingGrammarLabel((sequence,), {"priority": 1})
		self.assertEqual(self.grammarAsked, [])
		self.assertEqual(self.focusAsked, [])

	def test_misspelledAtAGrammarErrorSaysGrammarError(self):
		self.grammarError = True
		self.focus = object()
		args, kwargs = self.module._correctEditingGrammarLabel((["their", "misspelled", "going"],), {"priority": 1})
		self.assertEqual(args[0], ["their", "Grammar error", "going"])
		self.assertEqual(kwargs, {"priority": 1})
		self.assertEqual(self.grammarAsked, [self.focus])

	def test_misspelledAsAKeywordArgument(self):
		self.grammarError = True
		self.focus = object()
		args, kwargs = self.module._correctEditingGrammarLabel((), {"speechSequence": ["recieve Misspelled"]})
		self.assertEqual(kwargs["speechSequence"], ["recieve Grammar error"])

	def test_misspelledWithoutAGrammarErrorStays(self):
		self.focus = object()
		sequence = ["recieve", "misspelled"]
		args, kwargs = self.module._correctEditingGrammarLabel((sequence,), {})
		self.assertIs(args[0], sequence)
		self.assertEqual(self.grammarAsked, [self.focus], "Outlook is asked only for this")

	def test_theSpeechHookSpeaksWithoutAskingOutlook(self):
		"""The wrapper NVDA's speech goes through, as 1.0.71's log shows it running for clipspeak's message."""
		spoken = []

		def original(*args, **kwargs):
			spoken.append((args, kwargs))

		plugin = mock.MagicMock()
		plugin._spellCheckIsRunning.return_value = False
		plugin._inSpellDialog = False
		speak = self.module.GlobalPlugin._makeSpeakWrapper(plugin, original)
		with mock.patch.object(self.module, "_filterOutlookEditorNoise", lambda args, kwargs: (args, kwargs)):
			speak(["Copied selection to clipboard"], priority=1)
		self.assertEqual(spoken, [((["Copied selection to clipboard"],), {"priority": 1})])
		self.assertEqual(self.focusAsked, [])


if __name__ == "__main__":
	unittest.main()
