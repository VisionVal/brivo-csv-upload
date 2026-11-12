# Brivo CSV Uploader

Creates users in Brivo Access from a `;`-separated CSV. Needs only Python 3.8+ (no installs).

## Setup (per customer)
1. `cp config.example.ini config.ini` and fill in: `client_id`, `client_secret`, `api_key`,
   `admin_id` (admin username), `password`, `customfield_id`. Use `region = eu` for EU accounts.
   Use one ini file per customer, e.g. `customerA.ini`.
2. The Brivo application must be of type **Password**.

## Run
```
python3 brivo_upload.py users.csv --dry-run          # check the CSV first
python3 brivo_upload.py users.csv                    # uses config.ini
python3 brivo_upload.py users.csv -c customerA.ini
```
A `users.result.csv` report (created / failed / skipped per line) is written next to the CSV.

## CSV mapping (columns by position, first row = header is skipped)
| Col | Content | Brivo |
|---|---|---|
| 1 | Vorname | firstName |
| 2 | Nachname | lastName |
| 3 | Gesellschaft / Firma | not uploaded (Brivo has no company field) |
| 4 | E-Mail-Adresse | emails[0] |
| 7 | (number) | ignored |
| 8 | Kundennummer | custom field `customfield_id` |

## Sharing
Share this folder (zip it, without your `config.ini`). Never commit/share ini files with secrets.
