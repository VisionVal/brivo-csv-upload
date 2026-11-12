# Brivo XLSX Uploader

Creates users in Brivo Access from the first worksheet of an `.xlsx` workbook. Needs Python 3.8+; no packages are required. When available, `certifi` supplies CA certificates on Python installations without a configured system CA bundle.

## Setup (per customer)
1. Fill `config.ini` with: `client_id`, `client_secret`, `api_key`,
   `admin_id` (admin username), `password` and the Brivo custom field IDs
   `buyer_ident_field_id`, `plant_field_id`, `automat_id_field_id`, plus `shopper_group_id`.
   Set `shopper_group_id` to the ID of the existing Brivo user group named `Shopper`.
   Use `region = eu` for EU accounts.
   The script reads `config.ini` from the current directory.
2. The Brivo application must be of type **Password**.

## Run
```
python3 brivo_upload.py
```
A `<dateiname>.result.xlsx` report (created / failed / skipped per line) is written next to the workbook.

The start menu offers:
- `(1) Liste die Felder` to show Brivo custom-field names and IDs, plus the configured Shopper group ID. It then asks whether to write the IDs to `config.ini`; after saving, it displays the file with secret values masked.
- `(2) Test XLSX` to validate and preview workbook rows without API changes. It also checks existing Brivo users by first and last name.
- `(3) Upload XLSX` to upload the workbook at the path you enter.
- `(q) Quit` to exit.

## Workbook mapping (first worksheet; columns are matched by header)
| Header | Brivo |
|---|---|
| Vorname | firstName |
| Nachname | lastName |
| Gesellschaft / Firma | not uploaded (Brivo has no company field) |
| E-Mail-Adresse | emails[0] |
| Buyer_Ident | custom field `buyer_ident_field_id` |
| plant | custom field `plant_field_id` |
| automatisch erzeugt | custom field `automat_id_field_id`, Format `Vorname_Nachname` |

Custom field values may be empty; non-empty values are uploaded. The custom field IDs
must match the corresponding fields configured in Brivo. Every created user is also
assigned to the `Shopper` group using `shopper_group_id`.
`Automat_Id` is generated from the row's first and last name; an `automat_id` column in the workbook is ignored and not required.
Before uploading, the script checks existing Brivo users by first and last name (case-insensitive, ignoring repeated whitespace) and skips matches. The XLSX test reports the same matches; it also detects duplicate names within the workbook.

## Sharing
Share this folder (zip it, without your `config.ini`). Never commit/share ini files with secrets.
