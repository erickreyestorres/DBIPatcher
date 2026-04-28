# DBI Patcher

Automated translation system for [DBI](https://github.com/rashevskyv/dbi) - a Nintendo Switch homebrew application.

## Authors

- **DBI**: duckbill
- **Translation Script**: tg:@buinich_bohdan

## Notes from Bohdan

> Didn't have to calculate a single XOR manually - the program has 2100 real XORs and 1000 fake ones for obfuscation.

**Not translated:**
- Shadok fable strings (satirical text blocks)
- In-game language names displayed in DBI settings

## Features

- Automated translation for 22 languages using Claude Opus 4.5
- Smart formatting with automatic alignment and validation

## Supported Languages

BE (Belarusian), DE (German), EN (English US), ENGB (English UK), ES (Spanish Spain), ES419 (Spanish Latin America), ET (Estonian), FR (French), FRCA (French Canada), IT (Italian), JP (Japanese), KK (Kazakh), KR (Korean), LT (Lithuanian), LV (Latvian), NL (Dutch), PL (Polish), PT (Portuguese Portugal), PTBR (Portuguese Brazil), UA (Ukrainian), ZHCN (Chinese Simplified), ZHTW (Chinese Traditional)

## Adding a New Language

To add support for a new language:

1. **Option 1**: Open an [issue](https://github.com/rashevskyv/DBIPatcher/issues) requesting the language
2. **Option 2**: Add the language code to `data/languages.json` and submit a pull request

## Fixing Translations

To fix or improve existing translations:

1. Fork this repository
2. Edit the translation file for your language in `translations/` directory (e.g., `translations/en.csv`)
3. Commit your changes
4. Submit a pull request

## Usage

### Prerequisites
- Python 3.12+
- `openpyxl`, `requests`

### Commands

```bash
# Run full pipeline (sync, translate, align, validate, export, build, dist)
python -m src.main test

# Deploy to GitHub (creates release with assets)
python -m src.main deploy

# Run individual steps
python -m src.main sync       # Sync dictionary with source CSVs
python -m src.main translate  # Translate missing strings
python -m src.main align      # Align text formatting
python -m src.main validate   # Validate translations
python -m src.main export     # Export to CSV files
python -m src.main build      # Build binary files
python -m src.main dist       # Create distribution packages
```

## Installation on Nintendo Switch

1. Download the `translation_XX.bin` file for your language from [releases](https://github.com/rashevskyv/DBIPatcher/releases)
2. Rename the file to `translation.bin`
3. Place `translation.bin` in the same folder as your `DBI.nro`

## Official DBI Releases

Official Russian versions of DBI are available at: https://github.com/rashevskyv/dbi/releases/
