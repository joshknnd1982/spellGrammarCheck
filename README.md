# Spelling and Grammar Check

An [NVDA](https://www.nvaccess.org/) screen reader add-on that reads the F7
spelling and grammar checker in Microsoft Word and Microsoft Outlook clearly,
and nothing else.

* Author: Josh Kennedy
* Version: 1.0.72
* Compatibility: NVDA 2026.1 through 2026.2
* Download: grab the `.nvda-addon` file from the
  [releases page](https://github.com/joshknnd1982/spellGrammarCheck/releases)

A focused spinoff of Mute Browse Mode, keeping only its F7 spelling and
grammar support. It does not mute browse mode, and it does not change
anything about how NVDA behaves outside the proofing checker.

## What it does

For every issue the checker finds, the add-on announces "misspelled" or
"Grammar error", then the word or phrase involved, spoken and then spelled
out. If you ask it to, it moves straight into the suggestions list and reads
the first suggestion, spoken and spelled, with its position in the list. If
an issue has no suggestions at all, it says so plainly.

It supports both the classic Word/Outlook spelling and grammar dialog and the
newer Microsoft Editor task pane, and it announces "Spell check is complete"
once, at the end of a check, whether or not anything was found.

## Settings

These settings live in the NVDA menu, under Preferences, Settings, Spelling
and Grammar Check:

* Whether the suggestions list is entered automatically.
* Whether misspelled words and suggestions are read a little more slowly than
  usual (off by default).
* Whether to check for updates automatically, with a button to check now.

## Updates

The add-on checks for updates. Once a day, a little after NVDA starts, the add-on asks its GitHub repository, [github.com/joshknnd1982/spellGrammarCheck](https://github.com/joshknnd1982/spellGrammarCheck), whether a newer version has been released, and says nothing unless there is one. When there is, a dialog shows what's new in a box you can read line by line, and offers to download and install it. The download must match the release's SHA-256 checksum. Then NVDA asks you to confirm the installation and offers to restart. Your settings are kept.

To check yourself, open the NVDA menu, choose **Tools**, then **Check for add-on updates**, and choose **Spelling and Grammar Check...**. Or press **Check for updates now** in the add-on's settings: NVDA menu, Preferences, Settings, **Spelling and Grammar Check**. You can also assign a gesture to **Checks for Spelling and Grammar Check updates** in NVDA's Input Gestures dialog, under **Spelling and Grammar Check**. To stop the daily check, clear **Check for Spelling and Grammar Check updates automatically** in the same settings panel.

## Installation

1. Download the latest `spellGrammarCheck-x.y.z.nvda-addon` file from the
   [releases page](https://github.com/joshknnd1982/spellGrammarCheck/releases).
2. Press enter on the downloaded file and confirm the installation in NVDA.
3. Restart NVDA when prompted.

## Building from source

Requires Python 3. From the repository root:

```bash
python build.py
```

This produces `spellGrammarCheck-1.0.72.nvda-addon` and its `.sha256` checksum
file in the repository root. Upload both to the GitHub release: the update check
reads the release's tag, such as `v1.0.72`, and checks the download against the
checksum.

## Repository layout

```
addon/
  manifest.ini          Add-on metadata, including the full changelog
  globalPlugins/
    spellGrammarCheck/
      __init__.py       The global plugin
      updater.py        The GitHub update check, shared by all of joshknnd1982's
                        add-ons; keep it identical
tests/                  Unit tests: python -m unittest discover -s tests
build.py                Builds the .nvda-addon package
```

## Changelog

The full version history is in the `changelog` entry of
[addon/manifest.ini](addon/manifest.ini), which NVDA shows in the Add-on
Store.

## License

GNU General Public License version 2. See [LICENSE](LICENSE).
