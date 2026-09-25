# Spelling and Grammar Check

An [NVDA](https://www.nvaccess.org/) screen reader add-on that reads the F7
spelling and grammar checker in Microsoft Word and Microsoft Outlook clearly,
and nothing else.

* Author: Josh Kennedy
* Version: 1.0.70
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

Two settings live in the NVDA menu, under Preferences, Settings, Spelling and
Grammar Check:

* Whether the suggestions list is entered automatically.
* Whether misspelled words and suggestions are read a little more slowly than
  usual (off by default).

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

This produces `spellGrammarCheck-1.0.70.nvda-addon` in the repository root.

## Repository layout

```
addon/
  manifest.ini          Add-on metadata, including the full changelog
  globalPlugins/
    spellGrammarCheck/
      __init__.py       The global plugin
build.py                Builds the .nvda-addon package
```

## Changelog

The full version history is in the `changelog` entry of
[addon/manifest.ini](addon/manifest.ini), which NVDA shows in the Add-on
Store.

## License

GNU General Public License version 2. See [LICENSE](LICENSE).
