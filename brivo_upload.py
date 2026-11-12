#!/usr/bin/env python3
"""Upload users from an XLSX workbook to Brivo Access. Python 3.8+, no dependencies.

"""
import base64, configparser, json, re, ssl, sys, time, urllib.error, urllib.parse, urllib.request
from zipfile import BadZipFile, ZipFile
from xml.etree import ElementTree
from pathlib import Path

try:
    import certifi
except ImportError:
    certifi = None

SSL_CONTEXT = ssl.create_default_context(cafile=certifi.where() if certifi else None)

# Workbook header labels and their keys in the Brivo user payload.
HEADERS = {
    "vorname": "first",
    "nachname": "last",
    "gesellschaftfirma": "company",
    "emailadresse": "email",
    "buyerident": "Buyer_Ident",
    "plant": "Plant",
}
CUSTOM_FIELDS = ("Buyer_Ident", "Plant", "Automat_Id")
OPTIONAL_FIELDS = ("company", "email")
NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

HOSTS = {
    "us": ("https://auth.brivo.com", "https://api.brivo.com/v1/api"),
    "eu": ("https://auth.eu.brivo.com", "https://api.eu.brivo.com/v1/api"),
}
AUTH_REQUIRED = ["client_id", "client_secret", "api_key", "admin_id", "password"]
UPLOAD_REQUIRED = ["buyer_ident_field_id", "plant_field_id", "automat_id_field_id", "shopper_group_id"]


class ApiError(Exception):
    def __init__(self, status, body):
        super().__init__(f"HTTP {status}: {body}")
        self.status = status


class ProvisioningError(Exception):
    def __init__(self, operation, user_id, cause):
        super().__init__(f"{operation}: {cause}")
        self.operation = operation
        self.user_id = user_id


def load_config(path, for_upload=False):
    p = Path(path)
    if not p.exists():
        sys.exit(f"Config file '{path}' not found. Copy config.example.ini to config.ini and fill it in.")
    cp = configparser.ConfigParser(interpolation=None)
    cp.read(p, encoding="utf-8")
    cfg = dict(cp["brivo"]) if "brivo" in cp else {}
    required = AUTH_REQUIRED + (UPLOAD_REQUIRED if for_upload else [])
    missing = [k for k in required if not cfg.get(k, "").strip()]
    if missing:
        sys.exit(f"Missing values in {path}: {', '.join(missing)}")
    cfg = {k: v.strip() for k, v in cfg.items()}
    cfg.setdefault("region", "us")
    return cfg


class Brivo:
    def __init__(self, cfg):
        self.cfg = cfg
        self.auth_host, self.api = HOSTS.get(cfg["region"].lower(), HOSTS["us"])
        self.token = None

    def login(self):
        sig = base64.b64encode(f'{self.cfg["client_id"]}:{self.cfg["client_secret"]}'.encode()).decode()
        data = urllib.parse.urlencode({
            "grant_type": "password",
            "username": self.cfg["admin_id"],
            "password": self.cfg["password"],
        }).encode()
        req = urllib.request.Request(self.auth_host + "/oauth/token", data=data, method="POST", headers={
            "Authorization": "Basic " + sig,
            "api-key": self.cfg["api_key"],
            "Content-Type": "application/x-www-form-urlencoded",
        })
        self.token = self._send(req)["access_token"]

    def call(self, method, path, body=None):
        if not self.token:
            self.login()
        for attempt in (1, 2):
            req = urllib.request.Request(
                self.api + path, method=method,
                data=json.dumps(body).encode() if body is not None else None,
                headers={"Authorization": "bearer " + self.token, "api-key": self.cfg["api_key"],
                         "Content-Type": "application/json"})
            try:
                return self._send(req)
            except ApiError as e:
                if e.status == 401 and attempt == 1:  # token lives only 300s -> re-login
                    self.login()
                    continue
                if e.status == 429 and attempt == 1:
                    time.sleep(2)
                    continue
                raise

    @staticmethod
    def _send(req):
        try:
            with urllib.request.urlopen(req, timeout=30, context=SSL_CONTEXT) as r:
                raw = r.read().decode("utf-8")
                return json.loads(raw) if raw else {}
        except urllib.error.HTTPError as e:
            raise ApiError(e.code, e.read().decode("utf-8", "replace")[:500])
        except urllib.error.URLError as e:
            raise ApiError(0, str(e.reason))


def _column_index(cell_ref):
    letters = re.match(r"[A-Z]+", cell_ref.upper())
    if not letters:
        return None
    index = 0
    for char in letters.group():
        index = index * 26 + ord(char) - ord("A") + 1
    return index - 1


def read_rows(path):
    """Read the first worksheet of an XLSX file using only the standard library."""
    path = Path(path)
    if path.suffix.lower() != ".xlsx":
        raise ValueError("Input must be an .xlsx workbook")
    try:
        with ZipFile(path) as workbook:
            workbook_xml = ElementTree.fromstring(workbook.read("xl/workbook.xml"))
            relations = ElementTree.fromstring(workbook.read("xl/_rels/workbook.xml.rels"))
            relation_targets = {
                rel.attrib["Id"]: rel.attrib["Target"]
                for rel in relations.findall("pkgrel:Relationship", NS)
            }
            sheet = workbook_xml.find("main:sheets/main:sheet", NS)
            if sheet is None:
                raise ValueError("Workbook contains no worksheets")
            rel_id = sheet.attrib[f"{{{NS['rel']}}}id"]
            target = relation_targets[rel_id].lstrip("/")
            sheet_path = target if target.startswith("xl/") else f"xl/{target}"

            shared_strings = []
            if "xl/sharedStrings.xml" in workbook.namelist():
                strings_xml = ElementTree.fromstring(workbook.read("xl/sharedStrings.xml"))
                shared_strings = [
                    "".join(t.text or "" for t in item.findall(".//main:t", NS))
                    for item in strings_xml.findall("main:si", NS)
                ]

            sheet_xml = ElementTree.fromstring(workbook.read(sheet_path))
            rows = []
            for row in sheet_xml.findall(".//main:sheetData/main:row", NS):
                cells = {}
                for cell in row.findall("main:c", NS):
                    index = _column_index(cell.attrib.get("r", ""))
                    if index is None:
                        continue
                    value = cell.find("main:v", NS)
                    inline = cell.find("main:is", NS)
                    if cell.attrib.get("t") == "s" and value is not None:
                        text = shared_strings[int(value.text)]
                    elif inline is not None:
                        text = "".join(t.text or "" for t in inline.findall(".//main:t", NS))
                    else:
                        text = value.text if value is not None else ""
                    cells[index] = text.strip()
                rows.append([cells.get(i, "") for i in range(max(cells, default=-1) + 1)])
            return rows
    except (BadZipFile, KeyError, ElementTree.ParseError) as e:
        raise ValueError(f"Could not read XLSX workbook '{path}': {e}") from e


def _normalized_header(value):
    return "".join(char for char in value.casefold() if char.isalnum())


def _normalized_name(value):
    return " ".join(str(value or "").split()).casefold()


def _response_items(result, keys):
    """Extract list items from common Brivo response envelopes."""
    if isinstance(result, list):
        return result
    if isinstance(result, dict):
        for key in keys:
            value = result.get(key)
            if isinstance(value, list):
                return value
        for key in ("data", "_embedded", "response", "result"):
            nested = result.get(key)
            if isinstance(nested, (dict, list)):
                items = _response_items(nested, keys)
                if items or isinstance(nested, list):
                    return items
    return []


def fetch_existing_users(api):
    """Read all Brivo users, following pagination when present."""
    users = []
    page_size = 100
    offset = 0
    path = f"/users?offset={offset}&pageSize={page_size}"
    seen_paths = set()
    while path:
        if path in seen_paths:
            raise ValueError("Brivo returned a pagination loop while listing users")
        seen_paths.add(path)
        result = api.call("GET", path)
        item_keys = ("data", "items", "users", "results", "records", "value")
        page = _response_items(result, item_keys)
        if isinstance(result, dict) and not any(isinstance(result.get(key), list) for key in item_keys):
            raise ValueError("Brivo /users response did not contain a recognizable user list")
        if not isinstance(result, (list, dict)):
            raise ValueError("Brivo /users response had an unexpected format")
        if any(not isinstance(user, dict) for user in page):
            raise ValueError("Brivo /users response contained an invalid user record")
        users.extend(page)

        next_path = None
        count = None
        response_offset = offset
        response_page_size = page_size
        if isinstance(result, dict):
            try:
                count = int(result["count"]) if result.get("count") is not None else None
                response_offset = int(result.get("offset", offset))
                response_page_size = int(result.get("pageSize", page_size))
            except (TypeError, ValueError):
                count = None
            for key in ("next", "nextPage", "next_page", "nextLink", "nextLinkUrl"):
                candidate = result.get(key)
                if candidate:
                    next_path = candidate
                    break
            if not next_path and isinstance(result.get("pagination"), dict):
                pagination = result["pagination"]
                next_path = pagination.get("next", pagination.get("nextPage", pagination.get("next_page")))
            if not next_path and isinstance(result.get("_links"), dict):
                next_link = result["_links"].get("next")
                next_path = next_link.get("href") if isinstance(next_link, dict) else next_link
        if next_path:
            parsed = urllib.parse.urlparse(str(next_path))
            if parsed.scheme:
                route = parsed.path
                api_prefix = urllib.parse.urlparse(api.api).path.rstrip("/")
                if api_prefix and route.startswith(api_prefix):
                    route = route[len(api_prefix):]
                path = route + ("?" + parsed.query if parsed.query else "")
            else:
                path = str(next_path)
        elif count is not None and response_page_size > 0 and response_offset + len(page) < count:
            offset = response_offset + len(page)
            path = f"/users?offset={offset}&pageSize={page_size}"
        else:
            path = None
    return users


def index_existing_names(existing_users):
    names = {}
    for user in existing_users:
        if isinstance(user.get("user"), dict):
            user = {**user, **user["user"]}
        first = next((user.get(key) for key in ("firstName", "firstname", "first_name", "first", "givenName", "given_name") if user.get(key)), "")
        last = next((user.get(key) for key in ("lastName", "lastname", "last_name", "last", "surname", "familyName", "family_name") if user.get(key)), "")
        key = (_normalized_name(first), _normalized_name(last))
        if key != ("", ""):
            names.setdefault(key, []).append(user)
    return names


def matching_existing_user(user, existing_names):
    matches = existing_names.get((_normalized_name(user["first"]), _normalized_name(user["last"])), [])
    return matches[0] if matches else None


def parse(rows):
    users, problems = [], []
    if not rows:
        raise ValueError("Workbook is empty")
    indexes = {}
    for index, label in enumerate(rows[0]):
        key = HEADERS.get(_normalized_header(label))
        if key:
            indexes[key] = index
    missing = [key for key in ("first", "last", "Buyer_Ident", "Plant") if key not in indexes]
    if missing:
        raise ValueError("Missing required worksheet columns: " + ", ".join(missing))

    for n, r in enumerate(rows[1:], start=2):
        if not any(c.strip() for c in r):
            continue
        u = {key: "" for key in (*OPTIONAL_FIELDS, *CUSTOM_FIELDS)}
        u["line"] = n
        u.update({key: r[index].strip() if index < len(r) else "" for key, index in indexes.items()})
        if not u["first"] or not u["last"]:
            problems.append((n, "first or last name missing"))
        elif u["email"] and "@" not in u["email"]:
            problems.append((n, f"invalid e-mail '{u['email']}'"))
        else:
            u["Automat_Id"] = f"{u['first']}_{u['last']}"
            users.append(u)
    return users, problems


def write_xlsx_report(path, rows):
    """Write a minimal XLSX report using inline strings and the standard library."""
    main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    pkgrel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    content_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    ElementTree.register_namespace("", main_ns)
    ElementTree.register_namespace("r", rel_ns)

    def tag(namespace, name):
        return f"{{{namespace}}}{name}"

    def column_name(index):
        name = ""
        while index:
            index, remainder = divmod(index - 1, 26)
            name = chr(65 + remainder) + name
        return name

    workbook = ElementTree.Element(tag(main_ns, "workbook"))
    sheets = ElementTree.SubElement(workbook, tag(main_ns, "sheets"))
    ElementTree.SubElement(sheets, tag(main_ns, "sheet"), {
        "name": "Bericht", "sheetId": "1", tag(rel_ns, "id"): "rId1"
    })

    sheet = ElementTree.Element(tag(main_ns, "worksheet"))
    sheet_data = ElementTree.SubElement(sheet, tag(main_ns, "sheetData"))
    for row_number, values in enumerate(rows, start=1):
        row = ElementTree.SubElement(sheet_data, tag(main_ns, "row"), {"r": str(row_number)})
        for column_number, value in enumerate(values, start=1):
            cell = ElementTree.SubElement(row, tag(main_ns, "c"), {
                "r": f"{column_name(column_number)}{row_number}", "t": "inlineStr"
            })
            inline = ElementTree.SubElement(cell, tag(main_ns, "is"))
            text = ElementTree.SubElement(inline, tag(main_ns, "t"))
            text.text = str(value)
            if text.text != text.text.strip():
                text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    package_rels = ElementTree.Element(tag(pkgrel_ns, "Relationships"))
    ElementTree.SubElement(package_rels, tag(pkgrel_ns, "Relationship"), {
        "Id": "rId1", "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
        "Target": "xl/workbook.xml"
    })
    workbook_rels = ElementTree.Element(tag(pkgrel_ns, "Relationships"))
    ElementTree.SubElement(workbook_rels, tag(pkgrel_ns, "Relationship"), {
        "Id": "rId1", "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
        "Target": "worksheets/sheet1.xml"
    })
    content_types = ElementTree.Element(tag(content_ns, "Types"))
    ElementTree.SubElement(content_types, tag(content_ns, "Default"), {
        "Extension": "rels", "ContentType": "application/vnd.openxmlformats-package.relationships+xml"
    })
    ElementTree.SubElement(content_types, tag(content_ns, "Default"), {
        "Extension": "xml", "ContentType": "application/xml"
    })
    ElementTree.SubElement(content_types, tag(content_ns, "Override"), {
        "PartName": "/xl/workbook.xml",
        "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
    })
    ElementTree.SubElement(content_types, tag(content_ns, "Override"), {
        "PartName": "/xl/worksheets/sheet1.xml",
        "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
    })

    with ZipFile(path, "w") as report:
        report.writestr("[Content_Types].xml", ElementTree.tostring(content_types, encoding="utf-8", xml_declaration=True))
        report.writestr("_rels/.rels", ElementTree.tostring(package_rels, encoding="utf-8", xml_declaration=True))
        report.writestr("xl/workbook.xml", ElementTree.tostring(workbook, encoding="utf-8", xml_declaration=True))
        report.writestr("xl/_rels/workbook.xml.rels", ElementTree.tostring(workbook_rels, encoding="utf-8", xml_declaration=True))
        report.writestr("xl/worksheets/sheet1.xml", ElementTree.tostring(sheet, encoding="utf-8", xml_declaration=True))


def create_user(api, cfg, u, existing_names):
    existing = matching_existing_user(u, existing_names)
    if existing:
        return None, existing
    body = {"firstName": u["first"], "lastName": u["last"]}
    if u["email"]:
        body["emails"] = [{"address": u["email"], "type": "work"}]
    field_ids = {
        "Buyer_Ident": int(cfg["buyer_ident_field_id"]),
        "Plant": int(cfg["plant_field_id"]),
        "Automat_Id": int(cfg["automat_id_field_id"]),
    }
    populated_fields = [
        {"id": field_ids[name], "value": u[name]}
        for name in CUSTOM_FIELDS if u[name]
    ]
    try:
        res = api.call("POST", "/users", body)
    except Exception as e:
        raise ProvisioningError("Nutzer anlegen (POST /users)", None, e) from e
    uid = res["id"]
    for field in populated_fields:
        try:
            api.call("PUT", f"/users/{uid}/custom-fields/{field['id']}", {"value": field["value"]})
        except Exception as e:
            raise ProvisioningError(f"Custom Field {field['id']} setzen", uid, e) from e
    group_id = int(cfg["shopper_group_id"])
    try:
        # Brivo exposes group membership changes from the group resource.
        api.call("PUT", f"/groups/{group_id}/users/{uid}")
    except Exception as e:
        raise ProvisioningError(f"Shopper-Gruppe {group_id} zuweisen", uid, e) from e

    try:
        memberships = api.call("GET", f"/users/{uid}/groups")
    except Exception as e:
        raise ProvisioningError(f"Shopper-Zuweisung für Gruppe {group_id} prüfen", uid, e) from e
    groups = memberships if isinstance(memberships, list) else memberships.get("data", memberships.get("items", []))
    if not any(str(group.get("id", group.get("groupId", ""))) == str(group_id) for group in groups):
        raise ProvisioningError(
            f"Shopper-Gruppe {group_id} verifizieren",
            uid,
            RuntimeError("Brivo meldet den User nach der Zuweisung nicht als Gruppenmitglied"),
        )
    existing_names.setdefault((_normalized_name(u["first"]), _normalized_name(u["last"])), []).append(
        {"id": uid, "firstName": u["first"], "lastName": u["last"]}
    )
    return uid, None


def list_custom_fields(api):
    result = api.call("GET", "/custom-fields")
    fields = result if isinstance(result, list) else result.get("data", result.get("items", []))
    if fields:
        print("Brivo Custom Fields:")
        for field in fields:
            print(f"  ID {field.get('id', '?')}: {field.get('name', field.get('fieldName', '(ohne Namen)'))}")
    else:
        print("Keine Custom Fields gefunden.")
    group_id = api.cfg.get("shopper_group_id", "").strip()
    print(f"Shopper Gruppe: ID {group_id if group_id else '(nicht konfiguriert)'}")
    return fields or []


def update_ini_ids(path, values):
    """Update only requested [brivo] values while preserving other ini content."""
    path = Path(path)
    lines = path.read_text(encoding="utf-8").splitlines(keepends=True)
    in_brivo = False
    found = set()
    section_seen = False
    for index, line in enumerate(lines):
        stripped = line.strip()
        if stripped.startswith("[") and stripped.endswith("]"):
            in_brivo = stripped[1:-1].strip().casefold() == "brivo"
            section_seen = section_seen or in_brivo
            continue
        if not in_brivo or not stripped or stripped.startswith((";", "#")):
            continue
        match = re.match(r"^([ \t]*)([A-Za-z0-9_]+)([ \t]*=[ \t]*)([^\r\n]*)(\r?\n)?$", line)
        if match and match.group(2).casefold() in values:
            key = match.group(2).casefold()
            lines[index] = f"{match.group(1)}{match.group(2)}{match.group(3)}{values[key]}{match.group(5) or ''}"
            found.add(key)
    if not section_seen:
        raise ValueError("Abschnitt [brivo] fehlt in der INI-Datei")
    missing = [key for key in values if key not in found]
    if missing:
        insert_at = next((i for i, line in enumerate(lines) if i > 0 and line.strip().startswith("[") and line.strip().endswith("]")), len(lines))
        additions = [f"{key} = {values[key]}\n" for key in missing]
        if insert_at and lines[insert_at - 1].strip() and not lines[insert_at - 1].endswith(("\n", "\r")):
            lines[insert_at - 1] += "\n"
        lines[insert_at:insert_at] = additions
    path.write_text("".join(lines), encoding="utf-8")


def show_ini_config(path):
    """Display ini contents while masking credentials and secrets."""
    secret_keys = {"client_secret", "api_key", "password"}
    print(f"\n--- {path} ---")
    for line in Path(path).read_text(encoding="utf-8").splitlines():
        match = re.match(r"^([ \t]*)([A-Za-z0-9_]+)([ \t]*=[ \t]*).*$", line)
        if match and match.group(2).casefold() in secret_keys:
            line = f"{match.group(1)}{match.group(2)}{match.group(3)}[MASKIERT]"
        print(line)
    print("--- Ende ---\n")


def upload_workbook(path, config_path):
    try:
        users, problems = parse(read_rows(path))
    except ValueError as e:
        print(f"Fehler beim Lesen der XLSX-Datei: {e}")
        return 1
    print(f"{len(users)} gültige Zeilen, {len(problems)} ungültige Zeilen")
    for n, msg in problems:
        print(f"  Zeile {n}: ÜBERSPRUNGEN - {msg}")
    cfg = load_config(config_path, for_upload=True)
    api = Brivo(cfg)
    api.login()
    existing_names = index_existing_names(fetch_existing_users(api))
    ok = partial = failed = skipped = 0
    report = Path(path).with_suffix(".result.xlsx")
    report_rows = [["Zeile", "Vorname", "Nachname", "E-Mail", "Status", "Brivo-ID / Details"]]
    for n, msg in problems:
        report_rows.append([n, "", "", "", "skipped", msg])
    for u in users:
        try:
            uid, existing = create_user(api, cfg, u, existing_names)
            if existing:
                existing_id = existing.get("id", "?")
                skipped += 1
                report_rows.append([u["line"], u["first"], u["last"], u["email"], "exists", existing_id])
                print(f"  Zeile {u['line']}: übersprungen, User {u['first']} {u['last']} existiert bereits (ID {existing_id})")
                continue
            ok += 1
            report_rows.append([u["line"], u["first"], u["last"], u["email"], "created", uid])
            print(f"  Zeile {u['line']}: angelegt {u['first']} {u['last']} (ID {uid})")
        except ProvisioningError as e:
            if e.user_id is not None:
                partial += 1
                status = "partial"
                detail = f"User-ID {e.user_id}; {e}"
                print(f"  Zeile {u['line']}: TEILWEISE angelegt (ID {e.user_id}); {e}")
            else:
                failed += 1
                status = "failed"
                detail = str(e)
                print(f"  Zeile {u['line']}: FEHLER {u['first']} {u['last']} - {e}")
            report_rows.append([u["line"], u["first"], u["last"], u["email"], status, detail])
        except Exception as e:
            failed += 1
            report_rows.append([u["line"], u["first"], u["last"], u["email"], "failed", e])
            print(f"  Zeile {u['line']}: FEHLER {u['first']} {u['last']} - {e}")
    write_xlsx_report(report, report_rows)
    skipped += len(problems)
    print(f"Fertig: {ok} angelegt, {partial} teilweise angelegt, {failed} fehlgeschlagen, {skipped} übersprungen. Bericht: {report}")
    return 1 if failed or partial else 0


def main():
    while True:
        print("\nVisionVal Brivo XLSX Uploader")
        print("Version 1.1")
        print("(1) Liste die Felder")
        print("(2) Test XLSX")
        print("(3) Upload XLSX")
        print("(q) Quit")
        choice = input("Auswahl: ").strip().lower()

        if choice == "1":
            cfg = load_config("config.ini")
            try:
                api = Brivo(cfg)
                api.login()
                fields = list_custom_fields(api)
                answer = input("Soll config.ini mit den gefundenen IDs aktualisiert werden? (j/N): ").strip().casefold()
                if answer in ("j", "ja", "y", "yes"):
                    wanted = {
                        "buyerident": "buyer_ident_field_id",
                        "plant": "plant_field_id",
                        "automatid": "automat_id_field_id",
                    }
                    values = {}
                    for field in fields:
                        name = field.get("name", field.get("fieldName", field.get("label", "")))
                        key = wanted.get(_normalized_header(str(name)))
                        field_id = field.get("id")
                        if key and field_id is not None:
                            values[key] = str(field_id)
                    if cfg.get("shopper_group_id", "").strip():
                        values["shopper_group_id"] = cfg["shopper_group_id"].strip()
                    update_ini_ids("config.ini", values)
                    print("config.ini wurde aktualisiert: " + (", ".join(sorted(values)) if values else "keine Feld-IDs gefunden"))
                    show_ini_config("config.ini")
                else:
                    print("config.ini wurde nicht geändert.")
            except Exception as e:
                print(f"Fehler beim Abrufen der Felder: {e}")
        elif choice == "2":
            path = input("Pfad zur XLSX-Datei: ").strip()
            if not path:
                print("Bitte einen XLSX-Dateipfad eingeben.")
                continue
            try:
                users, problems = parse(read_rows(path))
                cfg = load_config("config.ini")
                api = Brivo(cfg)
                api.login()
                existing_names = index_existing_names(fetch_existing_users(api))
                existing_count = 0
                skipped_count = len(problems)
                print(f"{len(users)} gültige Zeilen, {len(problems)} ungültige Zeilen")
                for u in users:
                    fields = ", ".join(f"{name}={u[name]}" for name in CUSTOM_FIELDS)
                    existing = matching_existing_user(u, existing_names)
                    if existing:
                        existing_count += 1
                        skipped_count += 1
                        print(f"  Zeile {u['line']}: ÜBERSPRUNGEN, User existiert bereits (ID {existing.get('id', '?')}) - {u['first']} {u['last']}")
                    else:
                        print(f"  Zeile {u['line']}: NEU - {u['first']} {u['last']} {u['email']} {fields}")
                        existing_names.setdefault((_normalized_name(u["first"]), _normalized_name(u["last"])), []).append(
                            {"firstName": u["first"], "lastName": u["last"]}
                        )
                print(f"{existing_count} vorhandene User würden übersprungen; insgesamt würden {skipped_count} Zeilen übersprungen.")
                if problems:
                    print("\nUngültige Zeilen:")
                    for n, msg in problems:
                        print(f"  Zeile {n}: {msg}")
            except ValueError as e:
                print(f"Fehler beim Lesen der XLSX-Datei: {e}")
            except Exception as e:
                print(f"Fehler beim Prüfen bestehender Brivo-User: {e}")
        elif choice == "3":
            path = input("Pfad zur XLSX-Datei: ").strip()
            if not path:
                print("Bitte einen XLSX-Dateipfad eingeben.")
                continue
            result = upload_workbook(path, "config.ini")
            if result:
                print("Upload mit Fehlern beendet.")
        elif choice == "q":
            print("Programm beendet.")
            return
        else:
            print("Ungültige Auswahl. Bitte 1, 2, 3 oder q eingeben.")


if __name__ == "__main__":
    main()
