#!/usr/bin/env python3
"""Upload users from a CSV file to Brivo Access. Python 3.8+, no dependencies.

Usage:
    python3 brivo_upload.py users.csv                    # uses config.ini
    python3 brivo_upload.py users.csv -c customerA.ini   # per-customer config
    python3 brivo_upload.py users.csv --dry-run          # parse only, no API calls
"""
import argparse, base64, configparser, csv, json, sys, time, urllib.error, urllib.parse, urllib.request
from pathlib import Path

# Column positions (0-based) in the CSV
COL_FIRST, COL_LAST, COL_COMPANY, COL_EMAIL, COL_CUSTOMER_NO = 0, 1, 2, 3, 7

HOSTS = {
    "us": ("https://auth.brivo.com", "https://api.brivo.com/v1/api"),
    "eu": ("https://auth.eu.brivo.com", "https://api.eu.brivo.com/v1/api"),
}
REQUIRED = ["client_id", "client_secret", "api_key", "admin_id", "password", "customfield_id"]


class ApiError(Exception):
    def __init__(self, status, body):
        super().__init__(f"HTTP {status}: {body}")
        self.status = status


def load_config(path):
    p = Path(path)
    if not p.exists():
        sys.exit(f"Config file '{path}' not found. Copy config.example.ini to config.ini and fill it in.")
    cp = configparser.ConfigParser(interpolation=None)
    cp.read(p, encoding="utf-8")
    cfg = dict(cp["brivo"]) if "brivo" in cp else {}
    missing = [k for k in REQUIRED if not cfg.get(k, "").strip()]
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
            with urllib.request.urlopen(req, timeout=30) as r:
                raw = r.read().decode("utf-8")
                return json.loads(raw) if raw else {}
        except urllib.error.HTTPError as e:
            raise ApiError(e.code, e.read().decode("utf-8", "replace")[:500])
        except urllib.error.URLError as e:
            raise ApiError(0, str(e.reason))


def read_rows(path):
    raw = Path(path).read_bytes()
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = raw.decode("cp1252")  # typical Excel export
    rows = list(csv.reader(text.splitlines(), delimiter=";"))
    return rows[1:]  # skip header


def parse(rows):
    users, problems = [], []
    for n, r in enumerate(rows, start=2):
        if not any(c.strip() for c in r):
            continue
        r = [c.strip() for c in r] + [""] * 8
        u = {"line": n, "first": r[COL_FIRST], "last": r[COL_LAST], "company": r[COL_COMPANY],
             "email": r[COL_EMAIL], "customer_no": r[COL_CUSTOMER_NO]}
        if not u["first"] or not u["last"]:
            problems.append((n, "first or last name missing"))
        elif u["email"] and "@" not in u["email"]:
            problems.append((n, f"invalid e-mail '{u['email']}'"))
        else:
            users.append(u)
    return users, problems


def create_user(api, cfg, u):
    cf_id = int(cfg["customfield_id"])
    body = {"firstName": u["first"], "lastName": u["last"]}
    if u["email"]:
        body["emails"] = [{"address": u["email"], "type": "work"}]
    if u["customer_no"]:
        body["customFields"] = [{"id": cf_id, "value": u["customer_no"]}]
    res = api.call("POST", "/users", body)
    uid = res["id"]
    # Fallback: if the custom field wasn't stored with the create call, set it explicitly.
    if u["customer_no"]:
        got = {c.get("id"): c.get("value") for c in res.get("customFields", [])}
        if got.get(cf_id) != u["customer_no"]:
            api.call("PUT", f"/users/{uid}/custom-fields/{cf_id}", {"value": u["customer_no"]})
    return uid


def main():
    ap = argparse.ArgumentParser(description="Upload a CSV of users to Brivo Access")
    ap.add_argument("csv")
    ap.add_argument("-c", "--config", default="config.ini")
    ap.add_argument("--dry-run", action="store_true", help="validate the CSV only, no API calls")
    a = ap.parse_args()

    users, problems = parse(read_rows(a.csv))
    print(f"{len(users)} valid rows, {len(problems)} invalid rows")
    for n, msg in problems:
        print(f"  line {n}: SKIPPED - {msg}")
    if a.dry_run:
        for u in users:
            print(f"  line {u['line']}: {u['first']} {u['last']} <{u['email']}> customer_no={u['customer_no']}")
        return

    cfg = load_config(a.config)
    api = Brivo(cfg)
    api.login()
    ok = failed = 0
    report = Path(a.csv).with_suffix(".result.csv")
    with open(report, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["line", "first", "last", "email", "status", "brivo_user_id_or_error"])
        for n, msg in problems:
            w.writerow([n, "", "", "", "skipped", msg])
        for u in users:
            try:
                uid = create_user(api, cfg, u)
                ok += 1
                w.writerow([u["line"], u["first"], u["last"], u["email"], "created", uid])
                print(f"  line {u['line']}: created {u['first']} {u['last']} (id {uid})")
            except Exception as e:
                failed += 1
                w.writerow([u["line"], u["first"], u["last"], u["email"], "failed", e])
                print(f"  line {u['line']}: FAILED {u['first']} {u['last']} - {e}")
    print(f"Done: {ok} created, {failed} failed, {len(problems)} skipped. Report: {report}")
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
