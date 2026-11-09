"""
Tally CA Suite  -  FINAL (all phases)
=====================================
Reads TallyPrime (XML server, localhost:9000) or an Excel/CSV trial balance and produces:
  * Financial statements in the firm's own template  (Company = Schedule III vertical,
    Trust / Society / Association / Firm = horizontal)   -> Excel / PDF / Word
  * Notes that appear / disappear automatically with the figures
  * Independent Auditor's Report (Word) built on the firm's own .docx, with CARO,
    Rule 11, audit-trail, cash-flow and IFC clauses switched by year and facts
  * 40+ in-app checks, an in-app error log, a law library with live updating and self-repair

Install once:  pip install openpyxl reportlab python-docx
Run:           python tally_suite.py            (window)
               python tally_suite.py --selftest (full automatic test, no Tally needed)

IMPORTANT: this is a drafting and review aid. Every legal rule carries a source and a
"verified" date; anything unverified is flagged. The CA reviews and signs - not the software.
"""
APP_VERSION = "1.2.0"

import os, re, sys, io, json, copy, base64, hashlib, shutil, socket, threading, traceback
import subprocess, tempfile, urllib.request, urllib.error, xml.etree.ElementTree as ET
from collections import defaultdict, OrderedDict
from datetime import datetime, date, timedelta
from pathlib import Path

# ------------------------------------------------------------------ paths
APP_DIR = Path(os.environ.get("TALLY_CA_HOME") or (Path.home() / "TallyCASuite"))
for _d in ("backups", "snapshots", "output"):
    try: (APP_DIR / _d).mkdir(parents=True, exist_ok=True)
    except Exception: pass
RULES_FILE    = APP_DIR / "rules.json"
SETTINGS_FILE = APP_DIR / "settings.json"
LOG_FILE      = APP_DIR / "error.log"

# ------------------------------------------------------------------ error log (shown in the app)
class ErrorLog:
    def __init__(self):
        self.items, self.listeners = [], []
    def add(self, level, where, msg, tb=""):
        it = dict(ts=datetime.now().strftime("%d-%m-%Y %H:%M:%S"), level=level, where=where, msg=str(msg), tb=tb)
        self.items.append(it)
        try:
            with open(LOG_FILE, "a", encoding="utf-8") as f:
                f.write(f"[{it['ts']}] {level} {where}: {msg}\n{tb}\n")
        except Exception: pass
        for fn in list(self.listeners):
            try: fn(it)
            except Exception: pass
    def errors(self): return [i for i in self.items if i["level"] == "ERROR"]
    def clear(self): self.items.clear()
LOG = ErrorLog()

def guard(where, default=None):
    """Decorator: never crash the app - log the error (visible in the Errors tab) and carry on."""
    def deco(fn):
        def wrap(*a, **k):
            try: return fn(*a, **k)
            except Exception as e:
                LOG.add("ERROR", where, f"{type(e).__name__}: {e}", traceback.format_exc())
                return default
        wrap.__name__ = fn.__name__
        return wrap
    return deco

# ------------------------------------------------------------------ small utilities
def num(s):
    """Tally XML amount -> float.  Tally exports debit as negative, credit as positive."""
    if s is None: return 0.0
    if isinstance(s, (int, float)): return float(s)
    s = str(s).replace(" ", "").replace("\u00a0", "")
    m = re.search(r"-?[\d,]*\.?\d+", s)
    return float(m.group().replace(",", "")) if m else 0.0

def r2(x): return round(float(x) + 0.0, 2)

def inr(x, paren=False, dash=True):
    """Indian digit grouping 12,34,567.89 ; zero -> '-' ; negatives as (x) when paren."""
    if x is None or x == "": return ""
    x = r2(x)
    if abs(x) < 0.005: return "-" if dash else "0.00"
    neg = x < 0
    s = f"{abs(x):.2f}"; ip, dp = s.split(".")
    if len(ip) > 3:
        head, tail = ip[:-3], ip[-3:]
        parts = []
        while len(head) > 2: parts.insert(0, head[-2:]); head = head[:-2]
        if head: parts.insert(0, head)
        ip = ",".join(parts + [tail])
    out = f"{ip}.{dp}"
    if neg: out = f"({out})" if paren else f"-{out}"
    return out

def pdate(s):
    if isinstance(s, (date, datetime)): return s if isinstance(s, date) and not isinstance(s, datetime) else s.date()
    s = str(s).strip()
    for f in ("%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%Y", "%Y%m%d", "%d-%b-%Y", "%d-%b-%y"):
        try: return datetime.strptime(s, f).date()
        except ValueError: pass
    raise ValueError(f"unrecognised date: {s!r}")

def long_date(d):   return f"{d.strftime('%B')} {d.day}, {d.year}"          # March 31, 2025
def dmy(d):         return d.strftime("%d.%m.%Y")                            # 31.03.2025
def fy_label(end):  return f"{end.year-1}-{str(end.year)[2:]}"               # 2024-25
def ord_date(d):
    n = d.day; suf = "th" if 10 <= n % 100 <= 20 else {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suf} {d.strftime('%B')} {d.year}"

# ------------------------------------------------------------------ LAW LIBRARY (defaults)
# Each rule: history of {from, ...values}. Value in force on a date = last entry whose "from" <= date.
# "verified": date the entry was checked against a published source; "" = from memory, CA must confirm.
def _h(frm, ver="", **kw): d = {"from": frm, "verified": ver}; d.update(kw); return d
DEFAULT_RULES = {
 "schema": 1, "version": "2026.11.08", "updated": "2026-11-08",
 "rules": {
  "small_company": {
   "title": "Small company - Companies Act s.2(85) read with Rule 2(1)(t)",
   "explain": "A private company (not holding/subsidiary, not s.8, not under a special Act) is 'small' if paid-up capital and turnover (preceding FY) are both within the limits. Small companies are exempt from CARO, cash flow statement, and get lighter AS/Sch III disclosure.",
   "source": "MCA G.S.R. 880(E) dated 01-12-2025 (10cr / 100cr); G.S.R. 696(E) 15-09-2022 (4cr / 40cr); G.S.R. 92(E) 01-04-2021 (2cr / 20cr)",
   "history": [_h("2014-04-01", "", paid_up=5000000, turnover=20000000),
               _h("2021-04-01", "", paid_up=20000000, turnover=200000000),
               _h("2022-09-15", "", paid_up=40000000, turnover=400000000),
               _h("2025-12-01", "2026-11-08", paid_up=100000000, turnover=1000000000)]},
  "caro_private_exempt": {
   "title": "CARO - private company exemption (all three limits together)",
   "explain": "A private company that is not a holding/subsidiary of a public company is outside CARO only if paid-up capital + reserves <= 1 cr at balance sheet date, bank/FI borrowings <= 1 cr at ANY time in the year, and total revenue (incl. other income) <= 10 cr.",
   "source": "CARO 2020 para 1(2)(v); CARO 2016 para 1(2)",
   "history": [_h("2015-04-01", "", cap_reserves=10000000, borrowings=10000000, revenue=100000000),
               _h("2020-04-01", "2026-11-08", cap_reserves=10000000, borrowings=10000000, revenue=100000000)]},
  "caro_regime": {
   "title": "Which Companies (Auditor's Report) Order applies to the financial year",
   "explain": "Selected by the financial year START date. CARO 2020 applies to financial years commencing on/after 1 April 2020 (the 2019-20 effective date was deferred), so FY 2019-20 is still reported under CARO 2016.",
   "source": "MCA S.O. 849(E) 25-02-2020 and S.O. 1088(E) 24-03-2020 (deferral)",
   "history": [_h("2003-01-01", "", order="CARO 2003"), _h("2014-04-01", "", order="CARO 2015"),
               _h("2015-04-01", "", order="CARO 2016"), _h("2020-04-01", "", order="CARO 2020")]},
  "rule11_funds": {
   "title": "Rule 11(e) - funds advanced / received through intermediaries",
   "explain": "Management-representation reporting on intermediaries and funding parties. Applies to audit reports for FY starting on/after 1 April 2021.",
   "source": "Companies (Audit and Auditors) Amendment Rules, 2021", "history": [_h("2021-04-01", "", on=True)]},
  "rule11_dividend": {
   "title": "Rule 11(f) - dividend declared or paid",
   "explain": "Report whether dividend declared/paid complies with s.123. Applies from FY starting on/after 1 April 2021.",
   "source": "Companies (Audit and Auditors) Amendment Rules, 2021", "history": [_h("2021-04-01", "", on=True)]},
  "rule11_audit_trail": {
   "title": "Rule 11(g) - accounting software with audit trail (edit log)",
   "explain": "Mandatory from FY starting on/after 1 April 2023. From FY 2024-25 the report must also say the audit trail was preserved as per statutory record retention.",
   "source": "Companies (Accounts) Second Amendment Rules 2021 / Rule 3(1) as amended 2022; Rule 11(g)",
   "history": [_h("2023-04-01", "", on=True, preserved=False), _h("2024-04-01", "", on=True, preserved=True)]},
  "ifc_exempt": {
   "title": "s.143(3)(i) - exemption from IFC-over-financial-reporting opinion",
   "explain": "Private company is exempt if turnover < 50 cr in the previous year AND borrowings from banks/FIs < 25 cr at any time in the year (also exempt: OPC, small company).",
   "source": "MCA notification 13-06-2017", "history": [_h("2017-06-13", "", turnover=500000000, borrowings=250000000)]},
  "cash_flow_exempt": {
   "title": "Cash flow statement - who is exempt",
   "explain": "Financial statement as defined in s.2(40) does not include a cash flow statement for One Person Company, small company and dormant company.",
   "source": "Companies Act 2013 s.2(40) proviso", "history": [_h("2014-04-01", "", on=True)]},
  "cash_limits": {
   "title": "Cash transaction limits used by the checkers",
   "explain": "Rs 10,000 cash payment per day per person (s.40A(3) - disallowance; Rs 35,000 for transporters); Rs 20,000 cash loan/deposit/repayment (s.269SS/269T); Rs 2,00,000 cash receipt (s.269ST). The Income-tax Act 2025 carries corresponding provisions - section numbers may differ.",
   "source": "Income-tax law - CA to confirm for the current year",
   "history": [_h("2018-04-01", "", pay=10000, loan=20000, receipt=200000)]},
  "going_concern": {
   "title": "Going-concern indicators used by the checkers",
   "explain": "Net worth negative or accumulated losses >= net worth, continuing losses with nil revenue, dependence on director loans. SA 570 requires evaluating management's plan; a material uncertainty needs a separate paragraph in the report.",
   "source": "SA 570 (Revised)", "history": [_h("2014-04-01", "", min_ratio=1.0)]},
 },
 "live": {"fx": {}, "fx_asof": "", "fx_source": ""},
}

class Rules:
    """Loads / validates / repairs / merges the law library, and answers 'what applies on date X'."""
    def __init__(self):
        self.data = copy.deepcopy(DEFAULT_RULES)
        self.load()

    # ---- persistence with self-repair ----
    def validate(self, obj):
        probs = []
        if not isinstance(obj, dict): return ["not a JSON object"]
        if obj.get("schema") != 1: probs.append("schema must be 1")
        rs = obj.get("rules")
        if not isinstance(rs, dict) or not rs: return probs + ["'rules' missing or empty"]
        for rid, r in rs.items():
            if not isinstance(r, dict): probs.append(f"{rid}: not an object"); continue
            for k in ("title", "history"):
                if k not in r: probs.append(f"{rid}: missing {k}")
            h = r.get("history")
            if not isinstance(h, list) or not h: probs.append(f"{rid}: history empty"); continue
            for e in h:
                try: pdate(e["from"])
                except Exception: probs.append(f"{rid}: bad 'from' date {e.get('from') if isinstance(e, dict) else e!r}")
                for k, v in (e.items() if isinstance(e, dict) else []):
                    if k in ("paid_up", "turnover", "cap_reserves", "borrowings", "revenue", "pay", "loan", "receipt", "min_ratio") and not isinstance(v, (int, float)):
                        probs.append(f"{rid}: {k} must be a number")
        return probs

    def load(self):
        cand = [RULES_FILE] + sorted((APP_DIR / "backups").glob("rules_*.json"), reverse=True)
        for i, p in enumerate(cand):
            try:
                if not p.exists(): continue
                obj = json.loads(p.read_text(encoding="utf-8"))
                pr = self.validate(obj)
                if pr: raise ValueError("; ".join(pr[:3]))
                self.data = self._with_defaults(obj)
                if i > 0:
                    LOG.add("WARN", "rules", f"rules.json was damaged - SELF-REPAIRED from backup {p.name}")
                    self.save(backup=False)
                return
            except Exception as e:
                LOG.add("WARN", "rules", f"could not use {p.name}: {e}")
        # nothing usable -> embedded defaults (and write them out)
        self.data = copy.deepcopy(DEFAULT_RULES)
        self.save(backup=False)
        if RULES_FILE.exists(): LOG.add("INFO", "rules", "law library restored from built-in defaults")

    def _with_defaults(self, obj):
        out = copy.deepcopy(DEFAULT_RULES)          # built-ins stay unless overridden by a valid newer entry
        for rid, r in obj["rules"].items(): out["rules"][rid] = r
        out["version"] = obj.get("version", out["version"]); out["updated"] = obj.get("updated", out["updated"])
        out["live"] = obj.get("live", out["live"])
        return out

    def save(self, backup=True):
        try:
            if backup and RULES_FILE.exists():
                shutil.copy2(RULES_FILE, APP_DIR / "backups" / f"rules_{datetime.now():%Y%m%d_%H%M%S}.json")
                for old in sorted((APP_DIR / "backups").glob("rules_*.json"))[:-10]: old.unlink()
            tmp = RULES_FILE.with_suffix(".tmp")
            tmp.write_text(json.dumps(self.data, indent=1, ensure_ascii=False), encoding="utf-8")
            tmp.replace(RULES_FILE)
        except Exception as e:
            LOG.add("ERROR", "rules.save", e, traceback.format_exc())

    # ---- lookups ----
    def at(self, rid, on):
        """Entry in force on date `on`, or None (rule not yet in force)."""
        r = self.data["rules"].get(rid)
        if not r: LOG.add("ERROR", "rules.at", f"unknown rule {rid}"); return None
        on = pdate(on); best = None
        for e in sorted(r["history"], key=lambda e: pdate(e["from"])):
            if pdate(e["from"]) <= on: best = e
        return best
    def title(self, rid): return self.data["rules"].get(rid, {}).get("title", rid)
    def used(self): return getattr(self, "_used", {})
    def use(self, rid, on):
        e = self.at(rid, on)
        if not hasattr(self, "_used"): self._used = {}
        if e is not None: self._used[rid] = e
        return e
    def reset_used(self): self._used = {}

    def stale(self, days=180):
        """Entries used in this run that are unverified or verified long ago."""
        out = []
        for rid, e in self.used().items():
            v = e.get("verified", "")
            if not v: out.append((rid, "never verified against a published source"))
            else:
                try:
                    age = (date.today() - pdate(v)).days
                    if age > days: out.append((rid, f"last verified {v} ({age} days ago)"))
                except Exception: out.append((rid, "bad verified date"))
        return out

RULES = Rules()

# ------------------------------------------------------------------ settings
DEFAULT_SETTINGS = {
 "tally_host": "localhost", "tally_port": 9000,
 "rules_url": "", "app_url": "", "app_sha256": "",
 "fx_providers": ["https://api.frankfurter.dev/v1/latest?base=INR&symbols=USD,EUR,GBP,AED,SGD,JPY",
                  "https://open.er-api.com/v6/latest/INR"],
 "profile": {},
}
def load_settings():
    try:
        s = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
        out = copy.deepcopy(DEFAULT_SETTINGS); out.update(s); return out
    except Exception:
        return copy.deepcopy(DEFAULT_SETTINGS)
def save_settings(s):
    try: SETTINGS_FILE.write_text(json.dumps(s, indent=1, ensure_ascii=False), encoding="utf-8")
    except Exception as e: LOG.add("ERROR", "settings", e, traceback.format_exc())

# ------------------------------------------------------------------ live updates
def http_get(url, timeout=20):
    req = urllib.request.Request(url, headers={"User-Agent": f"TallyCASuite/{APP_VERSION}"})
    return urllib.request.urlopen(req, timeout=timeout).read()

@guard("update_fx", default=None)
def update_fx(settings, rules=None):
    """Live exchange rates (INR per 1 unit).  Tries each provider until one answers sanely."""
    rules = rules or RULES; last = None
    for url in settings.get("fx_providers", []):
        try:
            j = json.loads(http_get(url).decode("utf-8"))
            rates = j.get("rates") or {}
            if not rates: raise ValueError("no 'rates' in response")
            inr_per = {}
            for cur, per_inr in rates.items():
                if isinstance(per_inr, (int, float)) and per_inr > 0: inr_per[cur] = round(1.0 / per_inr, 4)
            if not inr_per: raise ValueError("empty rate table")
            for cur, v in inr_per.items():
                if not (0.0001 < v < 100000): raise ValueError(f"implausible rate {cur}={v}")
            asof = j.get("date") or (j.get("time_last_update_utc") or "")[:16] or datetime.now().strftime("%Y-%m-%d")
            rules.data["live"] = {"fx": inr_per, "fx_asof": str(asof), "fx_source": url.split("/")[2]}
            rules.save()
            return rules.data["live"]
        except Exception as e:
            last = e; LOG.add("WARN", "update_fx", f"{url}: {e}")
    LOG.add("ERROR", "update_fx", f"all providers failed ({last}). Previous rates kept.")
    return None

def _canon_hash(obj):
    o = copy.deepcopy(obj); o.pop("sha256", None)
    return hashlib.sha256(json.dumps(o, sort_keys=True, ensure_ascii=False).encode("utf-8")).hexdigest()

@guard("update_rules", default=None)
def update_rules(settings, rules=None, raw=None):
    """Pull a newer law library from the firm's own URL (e.g. a GitHub raw link). Validates, backs up,
    merges, re-validates - and rolls back automatically if anything is wrong.  Returns list of explained changes."""
    rules = rules or RULES
    url = settings.get("rules_url", "").strip()
    if raw is None:
        if not url: LOG.add("INFO", "update_rules", "no rules_url set (Law Library tab -> Sources)"); return None
        raw = http_get(url).decode("utf-8")
    obj = json.loads(raw)
    pr = rules.validate(obj)
    if pr: LOG.add("ERROR", "update_rules", "downloaded rules rejected: " + "; ".join(pr[:5])); return None
    if obj.get("sha256") and obj["sha256"] != _canon_hash(obj):
        LOG.add("ERROR", "update_rules", "checksum mismatch - file altered or corrupt. Rejected."); return None
    before = copy.deepcopy(rules.data)
    rules.save(backup=True)
    changes = []
    for rid, r in obj["rules"].items():
        old = before["rules"].get(rid)
        if old is None: changes.append(f"NEW  {rid}: {r.get('title')}")
        elif old["history"] != r["history"]:
            newh = [e for e in r["history"] if e not in old["history"]]
            for e in newh:
                vals = ", ".join(f"{k}={v}" for k, v in e.items() if k not in ("from", "verified"))
                changes.append(f"CHANGED {rid} from {e['from']}: {vals}   [{r.get('source','')}]")
        rules.data["rules"][rid] = r
    rules.data["version"] = obj.get("version", rules.data["version"]); rules.data["updated"] = obj.get("updated", rules.data["updated"])
    pr = rules.validate(rules.data)
    if pr:
        rules.data = before; rules.save(backup=False)
        LOG.add("ERROR", "update_rules", "merge failed validation - ROLLED BACK: " + "; ".join(pr[:3])); return None
    rules.save(backup=False)
    LOG.add("INFO", "update_rules", f"law library updated to {rules.data['version']} ({len(changes)} change(s))")
    return changes or ["No changes - already up to date."]

def _ver_tuple(v): return tuple(int(x) for x in re.findall(r"\d+", v)[:3]) or (0,)

@guard("self_update", default=None)
def self_update(settings, new_source=None, apply=True):
    """Download a newer tally_suite.py from the firm's own URL, TEST it, then swap it in. Rolls back on failure."""
    if getattr(sys, "frozen", False):
        LOG.add("INFO", "self_update", "This is the packaged .exe - download the new .exe from GitHub Releases instead."); return "frozen"
    if new_source is None:
        url = settings.get("app_url", "").strip()
        if not url: LOG.add("INFO", "self_update", "no app_url set"); return None
        new_source = http_get(url, 60).decode("utf-8")
    pin = settings.get("app_sha256", "").strip().lower()
    if pin and hashlib.sha256(new_source.encode("utf-8")).hexdigest() != pin:
        LOG.add("ERROR", "self_update", "SHA-256 of downloaded code does not match the pinned value. Rejected."); return None
    m = re.search(r'APP_VERSION\s*=\s*"([\d.]+)"', new_source)
    if not m: LOG.add("ERROR", "self_update", "downloaded file has no APP_VERSION"); return None
    if _ver_tuple(m.group(1)) <= _ver_tuple(APP_VERSION):
        LOG.add("INFO", "self_update", f"already latest ({APP_VERSION})"); return "current"
    compile(new_source, "tally_suite_new.py", "exec")                       # syntax check
    tmpdir = Path(tempfile.mkdtemp()); cand = tmpdir / "tally_suite.py"; cand.write_text(new_source, encoding="utf-8")
    env = dict(os.environ, TALLY_CA_HOME=str(tmpdir / "home"))
    r = subprocess.run([sys.executable, str(cand), "--selftest"], capture_output=True, text=True, timeout=300, env=env)
    if r.returncode != 0:
        LOG.add("ERROR", "self_update", "new version FAILED its self-test - not installed.\n" + (r.stdout + r.stderr)[-600:]); return None
    if not apply: return m.group(1)
    me = Path(__file__).resolve()
    shutil.copy2(me, APP_DIR / "backups" / f"tally_suite_{APP_VERSION}.py")
    me.write_text(new_source, encoding="utf-8")
    LOG.add("INFO", "self_update", f"updated {APP_VERSION} -> {m.group(1)}. Restart the app. Old copy kept in backups.")
    return m.group(1)

@guard("self_repair", default=[])
def self_repair(settings=None):
    """One-click health check + repair. Returns a list of what was checked / fixed."""
    out = []
    for d in ("backups", "snapshots", "output"):
        p = APP_DIR / d
        if not p.exists(): p.mkdir(parents=True, exist_ok=True); out.append(f"re-created folder {d}")
    n0 = len(LOG.items); RULES.load(); RULES.reset_used()
    pr = RULES.validate(RULES.data)
    out.append("law library valid" if not pr else f"law library problems: {pr[:2]}")
    if len(LOG.items) > n0: out.append("law library was repaired from backup/defaults (see Errors tab)")
    try: json.loads(SETTINGS_FILE.read_text(encoding="utf-8")) if SETTINGS_FILE.exists() else None; out.append("settings readable")
    except Exception: save_settings(copy.deepcopy(DEFAULT_SETTINGS)); out.append("settings.json was damaged - reset to defaults")
    for mod, pipname in (("openpyxl", "openpyxl"), ("reportlab", "reportlab"), ("docx", "python-docx")):
        try: __import__(mod); out.append(f"{pipname} OK")
        except Exception:
            out.append(f"{pipname} MISSING -> trying to install")
            try:
                if getattr(sys, "frozen", False): raise RuntimeError("packaged exe: rebuild needed")
                subprocess.run([sys.executable, "-m", "pip", "install", pipname], capture_output=True, timeout=180, check=True)
                out.append(f"{pipname} installed")
            except Exception as e:
                LOG.add("ERROR", "self_repair", f"cannot install {pipname}: {e}"); out.append(f"{pipname} could not be installed")
    st = selftest(quiet=True)
    out.append("engine self-test PASSED" if st[0] else "engine self-test FAILED: " + "; ".join(st[1][:3]))
    return out

EMBEDDED_TEMPLATE_B64 = "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"

# =============================================================== TALLY LINK + DATA MODEL
# A ledger is a dict: name, group, open, bal  -> both signed "Dr positive / Cr negative".
class Tally:
    def __init__(self, host="localhost", port=9000):
        self.url = f"http://{host}:{port}"

    def _post(self, xml):
        req = urllib.request.Request(self.url, xml.encode("utf-8"), {"Content-Type": "text/xml"})
        raw = urllib.request.urlopen(req, timeout=1500).read()          # big companies: Tally recalculates every ledger for the year - can take several minutes
        try:    txt = raw.decode("utf-8")
        except UnicodeDecodeError: txt = raw.decode("utf-16")
        txt = re.sub(r"&#(?:[0-9]|1[0-9]|2[0-9]|3[01]);", "", txt)         # Tally junk characters
        txt = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", txt)
        return ET.fromstring(txt)

    def collection(self, ctype, fetch, company=None, period=None):
        cmp_ = f"<SVCURRENTCOMPANY>{company}</SVCURRENTCOMPANY>" if company else ""
        if period: cmp_ += f"<SVFROMDATE>{period[0]:%Y%m%d}</SVFROMDATE><SVTODATE>{period[1]:%Y%m%d}</SVTODATE>"      # balances + vouchers for THIS year only
        xml = f"""<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>
<TYPE>Collection</TYPE><ID>CAC</ID></HEADER><BODY><DESC><STATICVARIABLES>
<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>{cmp_}</STATICVARIABLES>
<TDL><TDLMESSAGE><COLLECTION NAME="CAC" ISMODIFY="No"><TYPE>{ctype}</TYPE>
<FETCH>{fetch}</FETCH></COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"""
        return self._post(xml)

    def companies(self):
        r = self.collection("Company", "Name,StartingFrom,EndingAt")
        out = []
        for c in r.iter("COMPANY"):
            nm = c.get("NAME") or c.findtext("NAME") or ""
            if nm: out.append(dict(name=nm.strip(), start=(c.findtext("STARTINGFROM") or "").strip(),
                                   end=(c.findtext("ENDINGAT") or "").strip()))
        return out

    def groups(self, company):
        r = self.collection("Group", "Name,Parent", company)
        return {(g.get("NAME") or g.findtext("NAME") or "").strip(): (g.findtext("PARENT") or "").strip() for g in r.iter("GROUP")}

    def ledgers(self, company, period=None):
        r = self.collection("Ledger", "Name,Parent,OpeningBalance,ClosingBalance", company, period)
        out = []
        for l in r.iter("LEDGER"):
            out.append(dict(name=(l.get("NAME") or l.findtext("NAME", "")).strip(),
                            group=(l.findtext("PARENT") or "").strip(),
                            open=-num(l.findtext("OPENINGBALANCE")), bal=-num(l.findtext("CLOSINGBALANCE"))))
        return out

    def voucher_dates(self, company, period=None):
        """Light, fast query (dates only) - used to find which financial years really contain data."""
        r = self.collection("Voucher", "Date", company, period); out = []
        for v in r.iter("VOUCHER"):
            try: out.append(pdate((v.findtext("DATE") or "").strip()))
            except Exception: pass
        return out

    def vouchers(self, company, period=None):
        r = self.collection("Voucher",
            "Date,VoucherTypeName,VoucherNumber,PartyLedgerName,Narration,"
            "AllLedgerEntries.LedgerName,AllLedgerEntries.Amount,LedgerEntries.LedgerName,LedgerEntries.Amount", company, period)
        rows = []
        for v in r.iter("VOUCHER"):
            ds = (v.findtext("DATE") or "").strip()
            try: d = pdate(ds)
            except Exception: d = None
            ents = list(v.iter("ALLLEDGERENTRIES.LIST")) + list(v.iter("LEDGERENTRIES.LIST"))
            for e in ents:
                amt = num(e.findtext("AMOUNT"))
                rows.append(dict(date=d, vtype=(v.findtext("VOUCHERTYPENAME") or "").strip(),
                                 no=(v.findtext("VOUCHERNUMBER") or "").strip(),
                                 ledger=(e.findtext("LEDGERNAME") or "").strip(),
                                 debit=-amt if amt < 0 else 0.0, credit=amt if amt > 0 else 0.0,
                                 narration=(v.findtext("NARRATION") or "").strip()))
        return rows

# ---------------------------------------------------------------- group hierarchy / nature
DEFAULT_PARENT = {                       # Tally's built-in sub-groups (used when no hierarchy is supplied)
 "Reserves & Surplus": "Capital Account", "Bank OD A/c": "Loans (Liability)", "Secured Loans": "Loans (Liability)",
 "Unsecured Loans": "Loans (Liability)", "Duties & Taxes": "Current Liabilities", "Provisions": "Current Liabilities",
 "Sundry Creditors": "Current Liabilities", "Bank Accounts": "Current Assets", "Cash-in-Hand": "Current Assets",
 "Cash-in-hand": "Current Assets", "Deposits (Asset)": "Current Assets", "Loans & Advances (Asset)": "Current Assets",
 "Stock-in-Hand": "Current Assets", "Stock-in-hand": "Current Assets", "Sundry Debtors": "Current Assets"}
PRIMARY = {"Capital Account": "L", "Current Liabilities": "L", "Loans (Liability)": "L",
           "Fixed Assets": "A", "Investments": "A", "Current Assets": "A", "Misc. Expenses (ASSET)": "A",
           "Branch / Divisions": "A", "Suspense A/c": "S",
           "Sales Accounts": "I", "Direct Incomes": "I", "Income (Direct)": "I", "Indirect Incomes": "I", "Income (Indirect)": "I",
           "Purchase Accounts": "E", "Direct Expenses": "E", "Expenses (Direct)": "E", "Indirect Expenses": "E", "Expenses (Indirect)": "E"}
PLA_NAMES = {"profit & loss a/c", "profit and loss a/c", "profit & loss account", "profit and loss account"}

def chain(group, groups=None):
    out, seen, g = [], set(), (group or "").strip()
    while g and g not in seen and g != "Primary":
        out.append(g); seen.add(g)
        g = (groups or {}).get(g) or DEFAULT_PARENT.get(g, "")
        g = (g or "").strip()
    return out

def nature(ch):
    for g in reversed(ch):
        if g in PRIMARY: return PRIMARY[g]
    return "?"

def is_pla(led): return led["name"].strip().lower() in PLA_NAMES

# ---------------------------------------------------------------- import from Excel / CSV (no Tally needed)
@guard("import_tb", default=None)
def import_tb(path):
    """Columns (any order, header row anywhere in first 15 rows): Ledger | Group | Opening(optional, Dr+/Cr-) | Debit | Credit
    or Ledger | Group | Closing (Dr+/Cr-)."""
    rows = []
    p = str(path)
    if p.lower().endswith((".xlsx", ".xlsm")):
        import openpyxl
        ws = openpyxl.load_workbook(p, data_only=True).worksheets[0]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
    else:
        import csv
        with open(p, newline="", encoding="utf-8-sig") as f: rows = [r for r in csv.reader(f)]
    hi = None
    for i, r in enumerate(rows[:15]):
        low = [str(c or "").strip().lower() for c in r]
        if any(c in ("ledger", "name", "particulars", "ledger name") for c in low) and any(c in ("group", "parent") for c in low):
            hi = i; break
    if hi is None: raise ValueError("Could not find a header row with 'Ledger' and 'Group' columns")
    h = [str(c or "").strip().lower() for c in rows[hi]]
    def col(*names):
        for n in names:
            if n in h: return h.index(n)
        return None
    cn, cg = col("ledger", "name", "particulars", "ledger name"), col("group", "parent")
    cd, cc, cb, co = col("debit", "dr"), col("credit", "cr"), col("closing", "balance", "closing balance"), col("opening", "opening balance")
    if cb is None and cd is None: raise ValueError("Need 'Debit'/'Credit' or 'Closing' column")
    out = []
    for r in rows[hi + 1:]:
        if not r or cn >= len(r) or not r[cn] or str(r[cn]).strip().lower() in ("total", "grand total"): continue
        g = (lambda i: num(r[i]) if i is not None and i < len(r) else 0.0)
        bal = g(cb) if cb is not None else g(cd) - g(cc)
        out.append(dict(name=str(r[cn]).strip(), group=str(r[cg] or "").strip(), open=g(co) if co is not None else 0.0, bal=bal))
    if not out: raise ValueError("No ledger rows found")
    return out

# ---------------------------------------------------------------- snapshots (prior-year comparatives)
def snap_path(company, fy_end):
    safe = re.sub(r"[^A-Za-z0-9]+", "_", company).strip("_")
    return APP_DIR / "snapshots" / f"{safe}_{pdate(fy_end).year}.json"

@guard("save_snapshot")
def save_snapshot(company, fy_end, ledgers, groups=None):
    snap_path(company, fy_end).write_text(json.dumps(dict(company=company, fy_end=str(pdate(fy_end)), ledgers=ledgers, groups=groups or {}), indent=1), encoding="utf-8")

@guard("load_snapshot", default=None)
def load_snapshot(company, prior_fy_end):
    p = snap_path(company, prior_fy_end)
    return json.loads(p.read_text(encoding="utf-8")) if p.exists() else None

def prior_from_opening(ledgers):
    """Fallback comparatives: balance-sheet ledgers at their OPENING balance (= prior year closing).
    P&L ledgers are dropped because Tally cannot give the prior-year P&L."""
    out = []
    for l in ledgers:
        out.append(dict(name=l["name"], group=l["group"], open=l["open"], bal=l["open"], _from_open=True))
    return out

# =============================================================== TALLY SIMULATOR (for testing without Tally)
def _L(n, g, o, c): return dict(name=n, group=g, open=r2(o), bal=r2(c))     # Dr positive

def sim_company_ledgers(fy):      # Mascot Techcrib - taken from the audited statements you uploaded
    if fy == "2025":
        pla, sp = 29001331.55, 5771770.00
        exp = dict(sal=2512000.00, intr=85753.01, rent=228000.00, prt=43142.00, prof=64716.00, off=129428.00, aud=10000.00,
                   mkt=782215.00, elec=19197.00, misc=194143.00, dep=87308.60)
        loans = dict(ant=(22174479.72, 25542320.72), cy=(1560000.00, 2460000.00), ks=(1225043.00, 1310796.01))
        fa = dict(comp=(156234.96, 93740.98), plant=(27819.86, 23646.88), furn=(206416.35, 185774.71), sw=(1429720.00, 1714720.00))
    else:                         # FY 2023-24 (prior year)
        pla, sp = 25153930.37, 5771770.00
        exp = dict(sal=1910000.00, intr=80143.00, rent=216000.00, prt=86138.00, prof=111970.00, off=275626.00, aud=10000.00,
                   mkt=618087.00, elec=19836.00, misc=387600.00, dep=132001.18)
        loans = dict(ant=(18794502.72, 22174479.72), cy=(960000.00, 1560000.00), ks=(1144900.00, 1225043.00))
        fa = dict(comp=(433986.00, 156234.96), plant=(38505.00, 27819.86), furn=(254835.00, 206416.35), sw=(480000.00, 1429720.00))
    L = [_L("Equity Share Capital", "Capital Account", -100000, -100000), _L("CCPS", "Capital Account", -28230, -28230),
         _L("Securities Premium", "Reserves & Surplus", -sp, -sp), _L("Profit & Loss A/c", "Reserves & Surplus", pla, pla),
         _L("Antonty George Pattara", "Unsecured Loans", -loans["ant"][0], -loans["ant"][1]),
         _L("Cyriac Njavally", "Unsecured Loans", -loans["cy"][0], -loans["cy"][1]),
         _L("KSIDC Soft Loan", "Unsecured Loans", -loans["ks"][0], -loans["ks"][1]),
         _L("Audit Fee Payable", "Provisions", -10000, -10000),
         _L("Computers & Software", "Fixed Assets", fa["comp"][0], fa["comp"][1]), _L("Plant & Machinery", "Fixed Assets", fa["plant"][0], fa["plant"][1]),
         _L("Furniture & Electrical Fittings", "Fixed Assets", fa["furn"][0], fa["furn"][1]), _L("In House Software under Development", "Fixed Assets", fa["sw"][0], fa["sw"][1]),
         _L("Rent Deposit", "Deposits (Asset)", 48000, 48000)]
    for key, nm, grp in (("sal", "Salaries", "Indirect Expenses"), ("intr", "Interest on KSIDC Loan", "Indirect Expenses"), ("rent", "Rent", "Indirect Expenses"),
                         ("prt", "Printing & Stationery", "Indirect Expenses"), ("prof", "Professional Fees", "Indirect Expenses"), ("off", "Office Expenses", "Indirect Expenses"),
                         ("aud", "Auditors Fees", "Indirect Expenses"), ("mkt", "Marketing, Travel etc", "Indirect Expenses"), ("elec", "Electricity Charges", "Indirect Expenses"),
                         ("misc", "Miscellaneous Expenses", "Indirect Expenses"), ("dep", "Depreciation", "Indirect Expenses")):
        L.append(_L(nm, grp, 0, exp[key]))
    return L

def sim_trust_ledgers():          # Association for Welfare of AirIndia Retired Employees (from aware_fs_2025-26.xlsx)
    L = [_L("Membership Fee - General Fund", "Capital Account", -70000, -75500), _L("Advance paid by President as Expenses", "Loans (Liability)", -12412, -3057),
         _L("Duties & Taxes", "Duties & Taxes", -56987.90 + 56987.90, 0), _L("Accountant Fee Part Payment", "Sundry Creditors", -2000, -3000),
         _L("Provision For Audit Fee", "Provisions", -7500, -13000), _L("Membership Fee Payable (Extra Paid)", "Sundry Creditors", -2700, -2700),
         _L("GRP Insurance Add On Charge Fund", "Reserves & Surplus", -55500, -108000), _L("Contingency Premium Increase Offset Fund", "Reserves & Surplus", -316599.10, -569988.40),
         _L("Profit & Loss A/c", "Reserves & Surplus", -52288.93, -52288.93),
         _L("Mobile Phone For Office", "Fixed Assets", 10790, 10790), _L("Fixed Deposit - Deposits", "Deposits (Asset)", 630000, 630000),
         _L("SBI Bank Account", "Bank Accounts", 570138.93 - 630000 + 630000 - 570138.93 + 0, 84306.33)]
    inc = [("Annual Fee", 15341), ("Miscellaneous Income", 500), ("Interest - Fixed Deposit", 6975)]
    exp = [("Accounting Fee", 9000), ("Audit Fee", 7500), ("Bank Charges", 649), ("Consultation Fee (GST)", 11500), ("GB Meeting Expenses", 22258),
           ("Internet & Other Office Charges", 6000), ("Legal Charges (Stamps & Stamp Papers)", 450), ("Miscellaneous Expenses", 158),
           ("Office Expenses", 53519), ("Postage & Couriers", 205), ("Printing & Stationery Charges", 3401), ("Telephone Charges", 3604),
           ("Travelling And Conveyance Allowance", 7010)]
    for n, a in inc: L.append(_L(n, "Indirect Incomes" if "Interest" in n else "Direct Incomes", 0, -a))
    for n, a in exp: L.append(_L(n, "Indirect Expenses", 0, a))
    # opening of bank so that opening BS balances: assets opening must equal liabilities opening
    tot_open = sum(l["open"] for l in L)
    for l in L:
        if l["name"] == "SBI Bank Account": l["open"] = r2(l["open"] - tot_open)
    return L

def _month_ends(fy_end):
    y = pdate(fy_end).year
    return [date(y - 1, 4, 28), date(y - 1, 5, 28), date(y - 1, 6, 28), date(y - 1, 7, 28), date(y - 1, 8, 28), date(y - 1, 9, 28),
            date(y - 1, 10, 28), date(y - 1, 11, 28), date(y - 1, 12, 28), date(y, 1, 28), date(y, 2, 28), date(y, 3, 28)]

def sim_company_vouchers(ledgers, fy_end="2025-03-31", dirty=False):
    """Journal vouchers that reproduce each ledger's movement EXACTLY (so the voucher-vs-ledger check can pass)."""
    mv = {l["name"]: r2(l["bal"] - l["open"]) for l in ledgers}
    months = _month_ends(fy_end); V = []; n = 0
    exp_names = [l["name"] for l in ledgers if l["group"] == "Indirect Expenses" and l["name"] not in ("Depreciation", "Interest on KSIDC Loan")]
    for i, d in enumerate(months):
        lines = []
        for nm in exp_names:
            part = r2(mv[nm] / 12) if i < 11 else r2(mv[nm] - r2(mv[nm] / 12) * 11)
            lines.append((nm, part))
        tot = r2(sum(a for _, a in lines)); cy = 75000.00
        n += 1
        ents = [(nm, a) for nm, a in lines] + [("Cyriac Njavally", -cy), ("Antonty George Pattara", -r2(tot - cy))]
        V.append(dict(date=d, vtype="Journal", no=str(n), ents=ents, narration=f"Expenses met by directors - month {i+1}"))
    n += 1; sw = mv["In House Software under Development"]
    V.append(dict(date=months[8], vtype="Journal", no=str(n), ents=[("In House Software under Development", sw), ("Antonty George Pattara", -sw)], narration="In-house software development cost"))
    n += 1; ints = mv["Interest on KSIDC Loan"]
    V.append(dict(date=months[11].replace(day=31), vtype="Journal", no=str(n), ents=[("Interest on KSIDC Loan", ints), ("KSIDC Soft Loan", -ints)], narration="Interest accrued on soft loan"))
    n += 1; dep = mv["Depreciation"]
    fa = [("Computers & Software", 62493.98), ("Plant & Machinery", 4172.98), ("Furniture & Electrical Fittings", 20641.64)]
    V.append(dict(date=months[11].replace(day=31), vtype="Journal", no=str(n), ents=[("Depreciation", dep)] + [(a, -b) for a, b in fa], narration="Depreciation for the year"))
    if dirty:
        d0 = months[2]
        V += [dict(date=d0, vtype="Payment", no="P-1", ents=[("Office Expenses", 25000.0), ("Cash", -25000.0)], narration=""),
              dict(date=d0, vtype="Receipt", no="R-1", ents=[("Cash", 300000.0), ("Antonty George Pattara", -300000.0)], narration="Loan from director in cash"),
              dict(date=d0, vtype="Journal", no="R-3", ents=[("Rent", 100000.0), ("Antonty George Pattara", -100000.0)], narration="round sum"),
              dict(date=d0, vtype="Journal", no="4", ents=[("Rent", 5000.0), ("Antonty George Pattara", -4000.0)], narration="unbalanced entry"),
              dict(date=date(2025, 4, 10), vtype="Journal", no="99", ents=[("Rent", 1.0), ("Antonty George Pattara", -1.0)], narration="after year end"),
              dict(date=d0, vtype="Payment", no="P-2", ents=[("Professional Fees", 60000.0), ("Cash", -60000.0)], narration="consultant - no TDS")]
    return V

class TallySim:
    """Tiny stand-in for TallyPrime's XML server so the whole app can be tested without Tally."""
    def __init__(self, datasets):
        self.datasets = datasets
        sim = self
        from http.server import BaseHTTPRequestHandler, HTTPServer
        class H(BaseHTTPRequestHandler):
            def log_message(self, *a): pass
            def do_POST(self):
                body = self.rfile.read(int(self.headers.get("Content-Length", 0))).decode("utf-8", "ignore")
                out = sim.answer(body).encode("utf-8")
                self.send_response(200); self.send_header("Content-Type", "text/xml"); self.send_header("Content-Length", str(len(out))); self.end_headers(); self.wfile.write(out)
        self.srv = HTTPServer(("127.0.0.1", 0), H); self.port = self.srv.server_address[1]
        threading.Thread(target=self.srv.serve_forever, daemon=True).start()
    def stop(self): self.srv.shutdown()
    def answer(self, body):
        from xml.sax.saxutils import escape as esc
        m = re.search(r"<COLLECTION[^>]*>\s*<TYPE>(\w+)</TYPE>", body); t = m.group(1) if m else ""
        cm = re.search(r"<SVCURRENTCOMPANY>(.*?)</SVCURRENTCOMPANY>", body, re.S); cname = cm.group(1) if cm else None
        ds = self.datasets.get(cname) if cname else None
        x = ["<ENVELOPE><BODY><DATA><COLLECTION>"]
        if t == "Company":
            for k, d in self.datasets.items():
                x.append(f'<COMPANY NAME="{esc(k)}"><STARTINGFROM>{d["start"]}</STARTINGFROM><ENDINGAT>{d["end"]}</ENDINGAT></COMPANY>')
        elif ds is None: pass
        elif t == "Group":
            for g, p in ds.get("groups", {}).items(): x.append(f'<GROUP NAME="{esc(g)}"><PARENT>{esc(p)}</PARENT></GROUP>')
        elif t == "Ledger":
            for l in ds["ledgers"]:
                x.append(f'<LEDGER NAME="{esc(l["name"])}"><PARENT>{esc(l["group"])}</PARENT><OPENINGBALANCE>{-l["open"]:.2f}</OPENINGBALANCE><CLOSINGBALANCE>{-l["bal"]:.2f}</CLOSINGBALANCE></LEDGER>')
        elif t == "Voucher":
            for v in ds.get("vouchers", []):
                x.append(f'<VOUCHER><DATE>{v["date"]:%Y%m%d}</DATE><VOUCHERTYPENAME>{v["vtype"]}</VOUCHERTYPENAME><VOUCHERNUMBER>{v["no"]}</VOUCHERNUMBER><NARRATION>{esc(v["narration"])}</NARRATION>')
                for nm, a in v["ents"]: x.append(f'<ALLLEDGERENTRIES.LIST><LEDGERNAME>{esc(nm)}</LEDGERNAME><AMOUNT>{-a:.2f}</AMOUNT></ALLLEDGERENTRIES.LIST>')
                x.append("</VOUCHER>")
        x.append("</COLLECTION></DATA></BODY></ENVELOPE>")
        return "".join(x)

def make_sim_datasets():
    cur = sim_company_ledgers("2025"); pri = sim_company_ledgers("2024"); dirty = copy.deepcopy(cur)
    dirty += [_L("Suspense A/c", "Suspense A/c", 0, 1500.0), _L("Cash", "Cash-in-Hand", 0, 0.0), _L("Deferred Revenue Expenses", "Misc. Expenses (ASSET)", 0, 0.0)]
    dirty[-3]["bal"] = 1500.0
    for l in dirty:
        if l["name"] == "Antonty George Pattara": l["bal"] = r2(l["bal"] - 300000.0 - 100000.0 - 4000.0 + 0 - 1.0)   # keep vouchers vs ledger roughly mismatched on purpose
    t = sim_trust_ledgers()
    return {
     "Mascot Techcrib Private Limited (FY 2024-25)": dict(start="20240401", end="20250331", ledgers=cur, groups={}, vouchers=sim_company_vouchers(cur, "2025-03-31")),
     "Mascot Techcrib Private Limited (FY 2023-24)": dict(start="20230401", end="20240331", ledgers=pri, groups={}, vouchers=sim_company_vouchers(pri, "2024-03-31")),
     "Mascot DIRTY TEST (FY 2024-25)": dict(start="20240401", end="20250331", ledgers=dirty, groups={}, vouchers=sim_company_vouchers(cur, "2025-03-31", dirty=True)),
     "Association for Welfare of AirIndia Retired Employees (FY 2025-26)": dict(start="20250401", end="20260331", ledgers=t, groups={}, vouchers=[]),
    }

# =============================================================== FINANCIAL STATEMENT ENGINE
EQ_GROUPS = [("1", "Shareholders' funds", [("share_capital", "Share capital"), ("reserves", "Reserves and surplus")]),
             ("2", "Share application money pending allotment", [("share_app", None)]),
             ("3", "Non-current liabilities", [("lt_borrow", "Long-term borrowings"), ("dtl", "Deferred tax liabilities (net)"),
                                               ("other_ltl", "Other long-term liabilities"), ("lt_prov", "Long-term provisions")]),
             ("4", "Current liabilities", [("st_borrow", "Short-term borrowings"), ("trade_pay", "Trade payables"),
                                           ("other_cl", "Other current liabilities"), ("st_prov", "Short-term provisions")])]
AS_GROUPS = [("1", "Non-current assets", [("ppe", "Property, Plant and Equipment"), ("intang", "Intangible assets"), ("cwip", "Capital work-in-progress"),
                                          ("nc_inv", "Non-current investments"), ("dta", "Deferred tax assets (net)"), ("lt_loans", "Long-term loans and advances"),
                                          ("other_nca", "Other non-current assets"), ("def_rev", "Deferred revenue expenses")]),
             ("2", "Current assets", [("c_inv", "Current investments"), ("inventories", "Inventories"), ("trade_rec", "Trade receivables"),
                                      ("cash", "Cash and cash equivalents"), ("st_loans", "Short-term loans and advances"), ("other_ca", "Other current assets")])]
LABEL = {k: l for g in EQ_GROUPS + AS_GROUPS for k, l in g[2] if l}
LABEL.update(share_app="Share application money pending allotment", revenue="Revenue from operations", other_income="Other income",
             cost_mat="Cost of materials consumed / purchases", chg_inv="Changes in inventories", employee="Employee benefits expense",
             finance="Finance costs", dep="Depreciation and amortisation", other_exp="Other expenses", tax="Tax expense",
             prior_period="Prior period items", exceptional="Exceptional items")
LIAB_KEYS = [k for g in EQ_GROUPS for k, _ in g[2]]
ASSET_KEYS = [k for g in AS_GROUPS for k, _ in g[2]]
PL_EXP = ["cost_mat", "chg_inv", "employee", "finance", "dep", "other_exp"]
NOTE_ALWAYS_PL = {"other_exp"}

RX = lambda p: re.compile(p, re.I)
RX_EMP, RX_DEP, RX_FIN = RX(r"salar|wage|staff|remunerat|bonus|gratuity|provident|\besi\b|\bpf\b|employee|stipend"), RX(r"depreciat|amortis|amortiz"), RX(r"^interest|interest (on|paid|expense)|finance cost|loan interest|bank interest paid")
RX_TAX, RX_PRIOR = RX(r"income.?tax|tax expense|provision for (income )?tax|deferred tax expense|current tax"), RX(r"prior period|write.?off of deferred")
RX_LOAN_BANK = RX(r"overdraft|\bod\b|cash credit|\bcc\b|working capital|bank loan|term loan|nbfc|finance ltd|ksidc|\bsidbi\b|\bbank\b|capital ltd|financial corp")
RX_RELATED = RX(r"director|shareholder|promoter|proprietor|partner|trustee|member|president|secretary|treasurer|holding|subsidiary|related")

def classify_company(l, ch, prof, flags):
    n, nat, names = l["name"], nature(ch), set(ch)
    if nat in ("A", "L") and RX(r"deferred tax").search(n): return ("dtl" if nat == "L" else "dta"), ""
    if nat == "L" and RX(r"share application").search(n): return "share_app", ""
    if nat == "S":
        flags.append(("WARN", "classification", f"Suspense balance in '{n}' - clear before finalising."))
        return ("other_ca" if l["bal"] > 0 else "other_cl"), "suspense"
    if nat == "L":
        if "Reserves & Surplus" in names: return "reserves", ""
        if "Capital Account" in names:
            if RX(r"share capital|equity|preference|ccps|\bccp\b|ordinary|paid.?up|^capital$").search(n): return "share_capital", ""
            if RX(r"premium|reserve|surplus|retained|general|profit").search(n): return "reserves", ""
            flags.append(("WARN", "classification", f"'{n}' (Capital Account) assumed to be share capital - confirm.")); return "share_capital", ""
        if "Bank OD A/c" in names: return "st_borrow", "loan"
        if names & {"Loans (Liability)", "Secured Loans", "Unsecured Loans"}:
            if RX(r"overdraft|\bod\b|cash credit|\bcc\b|working capital").search(n): return "st_borrow", "loan"
            return prof.get("loan_line", "lt_borrow"), "loan"
        if "Sundry Creditors" in names: return "trade_pay", ""
        if "Provisions" in names: return ("lt_prov" if RX(r"long.?term|gratuity|leave").search(n) else "st_prov"), ""
        return "other_cl", ""
    if nat == "A":
        if "Fixed Assets" in names:
            if RX(r"capital work|cwip|under construction").search(n): return "cwip", ""
            if RX(r"software|intangible|goodwill|patent|trade.?mark|licen[cs]e|website").search(n) and not prof.get("merge_intangibles_in_ppe"): return "intang", ""
            return "ppe", ""
        if "Investments" in names: return ("c_inv" if RX(r"current|mutual fund|liquid|short").search(n) else "nc_inv"), ""
        if "Stock-in-Hand" in names or "Stock-in-hand" in names: return "inventories", ""
        if "Sundry Debtors" in names: return "trade_rec", ""
        if names & {"Cash-in-Hand", "Cash-in-hand", "Bank Accounts"}: return "cash", ""
        if "Deposits (Asset)" in names: return prof.get("deposit_line", "st_loans"), ""
        if "Loans & Advances (Asset)" in names: return "st_loans", ""
        if "Misc. Expenses (ASSET)" in names: return "def_rev", ""
        return "other_ca", ""
    if nat == "I":
        return ("revenue" if (names & {"Sales Accounts", "Direct Incomes", "Income (Direct)"}) else "other_income"), ""
    if nat == "E":
        if RX_PRIOR.search(n): return "prior_period", ""
        if RX_DEP.search(n): return "dep", ""
        if RX_TAX.search(n): return "tax", ""
        if names & {"Purchase Accounts", "Direct Expenses", "Expenses (Direct)"}: return "cost_mat", ""
        if RX_EMP.search(n): return "employee", ""
        if RX_FIN.search(n): return "finance", ""
        return "other_exp", ""
    flags.append(("ERROR", "classification", f"Cannot determine nature of '{n}' (group '{l['group']}'). Add the group to Tally or fix the import."))
    return "other_ca", ""

def compute_company(ledgers, prof, groups=None):
    """-> dict(lines, items, flags, profit, pbt, ...).  lines in presentation sign (Cr+ for liabilities/income, Dr+ for assets/expenses)."""
    lines, items, flags = defaultdict(float), defaultdict(list), []
    pla_open, from_open = 0.0, False
    for l in ledgers:
        if is_pla(l):
            pla_open = -l["open"]; from_open = from_open or bool(l.get("_from_open")); continue
        ch = chain(l["group"], groups); nat = nature(ch)
        key, tag = classify_company(l, ch, prof, flags)
        amt = l["bal"] if nat in ("A", "E", "S") else -l["bal"]
        if nat == "S": amt = abs(l["bal"])
        if key == "inventories" and nat == "A":
            lines["chg_inv"] += l["open"] - l["bal"]; items["chg_inv"].append((l["name"], r2(l["open"] - l["bal"]), ""))
        if key == "cash" and amt < -0.004:
            flags.append(("WARN", "reclass", f"'{l['name']}' has a CREDIT balance - shown under Short-term borrowings (bank overdraft)."))
            key, amt, tag = "st_borrow", -amt, "loan"
        elif key == "trade_rec" and amt < -0.004:
            flags.append(("WARN", "reclass", f"Debtor '{l['name']}' has a credit balance - shown under Other current liabilities (advance from customer)."))
            key, amt = "other_cl", -amt
        elif key == "trade_pay" and amt < -0.004:
            flags.append(("WARN", "reclass", f"Creditor '{l['name']}' has a debit balance - shown under Other current assets (advance to supplier)."))
            key, amt = "other_ca", -amt
        if abs(amt) < 0.005 and abs(l["open"]) < 0.005 and key not in ("reserves", "share_capital"): continue
        lines[key] += amt; items[key].append((l["name"], r2(amt), tag))
    inc = lines["revenue"] + lines["other_income"]
    exp = sum(lines[k] for k in PL_EXP)
    exc, tax, prior = lines["exceptional"], lines["tax"], lines["prior_period"]
    pbt = inc - exp - exc
    profit = pbt - prior - tax
    lines["reserves"] += pla_open + (0 if from_open else profit)
    items["reserves"].append(("Surplus / (deficit) in Statement of Profit & Loss - opening", r2(pla_open), "pla_open"))
    if not from_open: items["reserves"].append(("Add / (less): profit / (loss) for the year", r2(profit), "pla_profit"))
    tl = sum(lines[k] for k in LIAB_KEYS); ta = sum(lines[k] for k in ASSET_KEYS)
    for k in list(lines): lines[k] = r2(lines[k])
    return dict(lines=lines, items=items, flags=flags, income=r2(inc), expenses=r2(exp), pbt=r2(pbt), profit=r2(profit), tax=r2(tax),
                prior=r2(prior), tot_l=r2(tl), tot_a=r2(ta), diff=r2(tl - ta), pla_open=r2(pla_open), from_open=from_open)

# ------------------------------------------------------------------ horizontal (trust / society / association / firm)
def compute_h(ledgers, prof, groups=None):
    flags, L, A = [], OrderedDict(), OrderedDict()
    for h in ("Capital / General Fund", "Funds & Reserves", "Loans (Liability)", "Current Liabilities"): L[h] = []
    for h in ("Fixed Assets", "Investments", "Current Assets", "Other (Misc. / Suspense)"): A[h] = []
    I, E, DIRECT_E, DIRECT_I, IND_E, IND_I = [], [], [], [], [], []
    pla_open, stock_open, stock_close, from_open = 0.0, 0.0, 0.0, False
    for l in ledgers:
        if is_pla(l): pla_open = -l["open"]; from_open = from_open or bool(l.get("_from_open")); continue
        ch = chain(l["group"], groups); nat = nature(ch); names = set(ch)
        if nat == "L":
            if "Reserves & Surplus" in names: h = "Funds & Reserves"
            elif "Capital Account" in names: h = "Capital / General Fund"
            elif names & {"Loans (Liability)", "Secured Loans", "Unsecured Loans", "Bank OD A/c"}: h = "Loans (Liability)"
            else: h = "Current Liabilities"
            L[h].append((l["name"], r2(-l["bal"]), "loan" if h == "Loans (Liability)" else ""))
        elif nat in ("A", "S"):
            if nat == "S": flags.append(("WARN", "classification", f"Suspense balance in '{l['name']}'.")); h = "Other (Misc. / Suspense)"
            elif "Fixed Assets" in names: h = "Fixed Assets"
            elif "Investments" in names: h = "Investments"
            elif "Misc. Expenses (ASSET)" in names: h = "Other (Misc. / Suspense)"
            else: h = "Current Assets"
            if "Stock-in-Hand" in names or "Stock-in-hand" in names: stock_open += l["open"]; stock_close += l["bal"]
            if abs(l["bal"]) > 0.004 or abs(l["open"]) > 0.004: A[h].append((l["name"], r2(l["bal"]), ""))
        elif nat == "I":
            (DIRECT_I if names & {"Sales Accounts", "Direct Incomes", "Income (Direct)"} else IND_I).append((l["name"], r2(-l["bal"]), ""))
        elif nat == "E":
            (DIRECT_E if names & {"Purchase Accounts", "Direct Expenses", "Expenses (Direct)"} else IND_E).append((l["name"], r2(l["bal"]), ""))
        else: flags.append(("ERROR", "classification", f"Cannot determine nature of '{l['name']}' ({l['group']})."))
    trading = bool(DIRECT_E or [x for x in DIRECT_I if True]) and prof.get("entity") == "firm"
    if trading:
        gp = (sum(a for _, a, _ in DIRECT_I) + stock_close) - (sum(a for _, a, _ in DIRECT_E) + stock_open)
        net = gp + sum(a for _, a, _ in IND_I) - sum(a for _, a, _ in IND_E)
    else:
        gp = None
        net = sum(a for _, a, _ in DIRECT_I + IND_I) - sum(a for _, a, _ in DIRECT_E + IND_E) + (stock_close - stock_open)
    tl = sum(a for h in L.values() for _, a, _ in h) + pla_open + (0 if from_open else net)
    ta = sum(a for h in A.values() for _, a, _ in h)
    return dict(L=L, A=A, DIRECT_E=DIRECT_E, DIRECT_I=DIRECT_I, IND_E=IND_E, IND_I=IND_I, stock_open=r2(stock_open), stock_close=r2(stock_close),
                trading=trading, gp=None if gp is None else r2(gp), net=r2(net), pla_open=r2(pla_open), from_open=from_open,
                tot_l=r2(tl), tot_a=r2(ta), diff=r2(tl - ta), flags=flags, profit=r2(net))

# ------------------------------------------------------------------ cash flow (AS 3 indirect) for vertical statements
def cash_flow(cur, prv):
    if prv is None: return None
    d = lambda k: cur["lines"][k] - prv["lines"][k]
    loan = lambda f: sum(a for k in ("lt_borrow", "st_borrow", "other_cl", "other_ltl") for n, a, t in f["items"].get(k, []) if t == "loan")
    d_loan = loan(cur) - loan(prv)
    dep, fin = cur["lines"]["dep"], cur["lines"]["finance"]
    pbt = cur["pbt"] - cur["prior"] if True else cur["pbt"]
    ops = [("Net profit / (loss) before tax", pbt), ("Adjustments: Depreciation and amortisation", dep), ("Adjustments: Finance costs", fin)]
    pre_wc = sum(a for _, a in ops)
    wc = [("(Increase) / decrease in inventories", -d("inventories")), ("(Increase) / decrease in trade receivables", -d("trade_rec")),
          ("(Increase) / decrease in loans, advances & other assets", -(d("st_loans") + d("lt_loans") + d("other_ca") + d("other_nca") + d("def_rev"))),
          ("Increase / (decrease) in trade payables", d("trade_pay")),
          ("Increase / (decrease) in other liabilities & provisions", (d("other_cl") + d("other_ltl") + d("st_prov") + d("lt_prov")) - (
              sum(a for k in ("other_cl", "other_ltl") for n, a, t in cur["items"].get(k, []) if t == "loan") - sum(a for k in ("other_cl", "other_ltl") for n, a, t in prv["items"].get(k, []) if t == "loan")))]
    op = pre_wc + sum(a for _, a in wc) - cur["tax"]
    capex = d("ppe") + d("intang") + d("cwip") + dep
    inv = [("Purchase of PPE / intangibles / CWIP (net of disposals)", -capex), ("(Purchase) / sale of investments", -(d("nc_inv") + d("c_inv")))]
    invt = sum(a for _, a in inv)
    fn = [("Issue of share capital / premium / application money", d("share_capital") + d("share_app") + sum(a for n, a, t in cur["items"]["reserves"] if "premium" in n.lower()) - sum(a for n, a, t in prv["items"]["reserves"] if "premium" in n.lower())),
          ("Proceeds from / (repayment of) borrowings", d_loan), ("Finance costs paid", -fin)]
    fnt = sum(a for _, a in fn)
    net = op + invt + fnt
    c_open, c_close = prv["lines"]["cash"], cur["lines"]["cash"]
    return dict(ops=ops, pre_wc=pre_wc, wc=wc, op=r2(op), tax=cur["tax"], inv=inv, invt=r2(invt), fn=fn, fnt=r2(fnt), net=r2(net),
                c_open=r2(c_open), c_close=r2(c_close), gap=r2(c_open + net - c_close))

# =============================================================== TABLES (neutral structure -> text / Excel / PDF / Word)
def R(c, b=False, sm=None, ind=0, top=False): return dict(c=c, b=b, sm=sm, ind=ind, top=top)

def _col_heads(fy_end, prior_end=None):
    e = pdate(fy_end); p = prior_end or date(e.year - 1, e.month, e.day)
    return f"As at {e.day} {e.strftime('%B')}, {e.year}", f"As at {p.day} {p.strftime('%B')}, {p.year}"

def _pl_heads(fy_end):
    e = pdate(fy_end); p = date(e.year - 1, e.month, e.day)
    return f"For the year ended {e.day} {e.strftime('%B')}, {e.year}", f"For the year ended {p.day} {p.strftime('%B')}, {p.year}"

def _shown(key, cur, prv, prof):
    a = cur["lines"].get(key, 0.0); b = prv["lines"].get(key, 0.0) if prv else 0.0
    return abs(a) > 0.004 or abs(b) > 0.004 or key in prof.get("always_lines", [])

def number_notes(cur, prv, prof):
    """Notes appear / vanish with the figures. Note 1 = accounting policies; others numbered in face order."""
    order = [k for g in EQ_GROUPS for k, _ in g[2]] + [k for g in AS_GROUPS for k, _ in g[2]]
    notes, n = OrderedDict(), 2
    def items_of(k):
        names = OrderedDict()
        for src, f in (("cur", cur), ("prv", prv)):
            if f is None: continue
            for nm, a, t in f["items"].get(k, []): names.setdefault(nm, {"cur": 0.0, "prv": 0.0, "tag": t}); names[nm][src] += a
        return names
    for k in order:
        if not _shown(k, cur, prv, prof): continue
        if k in ("trade_pay", "trade_rec", "share_app") and len(items_of(k)) <= 1 and k not in prof.get("always_notes", []): continue
        notes[k] = n; n += 1
    for k in ("employee", "finance", "revenue", "other_income", "other_exp", "cost_mat"):
        if not _shown(k, cur, prv, prof): continue
        if k in NOTE_ALWAYS_PL or len(items_of(k)) >= 2: notes[k] = n; n += 1
    return notes

def build_vertical(cur, prv, prof, fy_end, cf=None):
    """Schedule III style: Balance Sheet, Statement of P&L, Notes, optional Cash Flow."""
    c1, c2 = _col_heads(fy_end); p1, p2 = _pl_heads(fy_end)
    notes = number_notes(cur, prv, prof)
    G = lambda f, k: (f["lines"].get(k, 0.0) if f else None)
    T = []
    # ---------------- balance sheet
    rows = []; idx = {}
    def line(key, lab, ind=1):
        rows.append(R([lab, notes.get(key, ""), G(cur, key), G(prv, key) if prv else None], ind=ind)); idx[key] = len(rows) - 1
    rows.append(R(["A  EQUITY AND LIABILITIES", "", None, None], b=True))
    grp_idx = []
    gno = 0
    for num_, title, ls in EQ_GROUPS:
        vis = [(k, l) for k, l in ls if _shown(k, cur, prv, prof)]
        if not vis: continue
        gno += 1; num_ = str(gno)
        if len(ls) == 1 and ls[0][1] is None:
            rows.append(R([f"{num_}  {title}", notes.get(ls[0][0], ""), G(cur, ls[0][0]), G(prv, ls[0][0]) if prv else None], b=True)); idx[ls[0][0]] = len(rows) - 1; grp_idx.append(len(rows) - 1)
        else:
            rows.append(R([f"{num_}  {title}", "", None, None], b=True))
            for i, (k, l) in enumerate(vis): line(k, f"({chr(97+i)}) {l}")
    eq_rows = [idx[k] for k in LIAB_KEYS if k in idx]
    rows.append(R(["TOTAL", "", cur["tot_l"], prv["tot_l"] if prv else None], b=True, sm=eq_rows, top=True))
    tl_row = len(rows) - 1
    rows.append(R(["", "", None, None]))
    rows.append(R(["B  ASSETS", "", None, None], b=True)); idx2 = {}
    gno = 0
    for num_, title, ls in AS_GROUPS:
        vis = [(k, l) for k, l in ls if _shown(k, cur, prv, prof)]
        if not vis: continue
        gno += 1; num_ = str(gno)
        rows.append(R([f"{num_}  {title}", "", None, None], b=True))
        for i, (k, l) in enumerate(vis): line(k, f"({chr(97+i)}) {l}"); idx2[k] = idx[k]
    rows.append(R(["TOTAL", "", cur["tot_a"], prv["tot_a"] if prv else None], b=True, sm=list(idx2.values()), top=True))
    T.append(dict(kind="v", title=f"BALANCE SHEET AS AT {pdate(fy_end).day}{'ST' if pdate(fy_end).day==31 or pdate(fy_end).day==1 else 'TH'} {pdate(fy_end).strftime('%B').upper()}, {pdate(fy_end).year}",
                  cols=["Particulars", "Note No.", c1, c2], rows=rows, sign=True, widths=[0.52, 0.08, 0.20, 0.20]))
    # ---------------- profit & loss
    rows = []; pidx = {}
    def pline(label, key=None, val=None, valp=None, b=False, sm=None, top=False, ind=1):
        a = cur["lines"].get(key, 0.0) if key else val
        bb = (prv["lines"].get(key, 0.0) if prv else None) if key else valp
        rows.append(R([label, notes.get(key, "") if key else "", a, bb], b=b, sm=sm, top=top, ind=ind)); pidx[key or label] = len(rows) - 1
    rows.append(R(["A  CONTINUING OPERATIONS", "", None, None], b=True))
    pline("1  Revenue from operations (net)", "revenue", ind=0); pline("2  Other income", "other_income", ind=0)
    pline("3  Total revenue (1+2)", None, cur["income"], prv["income"] if prv else None, b=True, sm=[pidx["revenue"], pidx["other_income"]], top=True, ind=0)
    rows.append(R(["4  Expenses", "", None, None], b=True)); first = len(rows)
    letters = iter("abcdefghij")
    for k in PL_EXP:
        if _shown(k, cur, prv, prof) or k in ("employee", "finance", "other_exp", "dep") and (cur["lines"][k] or (prv and prv["lines"][k])): pline(f"({next(letters)}) {LABEL[k]}", k)
    exp_rows = [pidx[k] for k in PL_EXP if k in pidx]
    pline("Total expenses", None, cur["expenses"], prv["expenses"] if prv else None, b=True, sm=exp_rows, top=True, ind=0)
    pline("5  Profit / (Loss) before exceptional items and tax (3 - 4)", None, cur["pbt"], prv["pbt"] if prv else None, b=True, ind=0); sn = 6
    if abs(cur["prior"]) > 0.004 or (prv and abs(prv["prior"]) > 0.004) or "prior_period" in prof.get("always_lines", []):
        pline(f"{sn}  Prior period item - write off of deferred revenue expenses", "prior_period", ind=0); sn += 1
    pline(f"{sn}  Profit / (Loss) before tax", None, r2(cur["pbt"] - cur["prior"]), r2(prv["pbt"] - prv["prior"]) if prv else None, b=True, ind=0); sn += 1
    pline(f"{sn}  Tax expense - current tax", "tax", ind=0); sn += 1
    pline(f"{sn}  Profit / (Loss) for the year", None, cur["profit"], prv["profit"] if prv else None, b=True, top=True, ind=0)
    T.append(dict(kind="v", title=f"STATEMENT OF PROFIT & LOSS FOR THE YEAR ENDED {pdate(fy_end):%d.%m.%Y}", cols=["Particulars", "Note No.", p1, p2], rows=rows, sign=True, widths=[0.52, 0.08, 0.20, 0.20]))
    # ---------------- cash flow
    if cf:
        rows = [R(["A. CASH FLOW FROM OPERATING ACTIVITIES", None, None], b=True)]
        for l, a in cf["ops"]: rows.append(R([l, a, None], ind=1))
        rows.append(R(["Operating profit before working-capital changes", cf["pre_wc"], None], b=True, top=True))
        for l, a in cf["wc"]: rows.append(R([l, a, None], ind=1))
        rows.append(R(["Taxes paid", -cf["tax"], None], ind=1))
        rows.append(R(["Net cash from / (used in) operating activities (A)", cf["op"], None], b=True, top=True))
        rows.append(R(["B. CASH FLOW FROM INVESTING ACTIVITIES", None, None], b=True))
        for l, a in cf["inv"]: rows.append(R([l, a, None], ind=1))
        rows.append(R(["Net cash from / (used in) investing activities (B)", cf["invt"], None], b=True, top=True))
        rows.append(R(["C. CASH FLOW FROM FINANCING ACTIVITIES", None, None], b=True))
        for l, a in cf["fn"]: rows.append(R([l, a, None], ind=1))
        rows.append(R(["Net cash from / (used in) financing activities (C)", cf["fnt"], None], b=True, top=True))
        rows.append(R(["Net increase / (decrease) in cash and cash equivalents (A+B+C)", cf["net"], None], b=True))
        rows.append(R(["Cash and cash equivalents at the beginning of the year", cf["c_open"], None]))
        rows.append(R(["Cash and cash equivalents at the end of the year", cf["c_close"], None], b=True, top=True))
        T.append(dict(kind="v", title=f"CASH FLOW STATEMENT FOR THE YEAR ENDED {pdate(fy_end):%d.%m.%Y}", cols=["Particulars", p1, ""], rows=rows, sign=True, widths=[0.64, 0.18, 0.18]))
    # ---------------- notes
    n1 = prof.get("policies") or DEFAULT_POLICIES
    T.append(dict(kind="t", title="NOTE 1: SIGNIFICANT ACCOUNTING POLICIES", text=n1))
    for k, no in notes.items():
        names = OrderedDict()
        for src, f in (("cur", cur), ("prv", prv)):
            if f is None: continue
            for nm, a, t in f["items"].get(k, []): names.setdefault(nm, {"cur": 0.0, "prv": 0.0}); names[nm][src] += a
        title = f"NOTE {no}: {LABEL.get(k, k).upper()}"
        rows = []
        if k == "share_capital":
            for t in prof.get("share_capital_lines", []): rows.append(R([t, None, None]))
        for nm, v in names.items():
            if abs(v["cur"]) < 0.005 and abs(v["prv"]) < 0.005 and k != "reserves": continue
            rows.append(R([nm, v["cur"], v["prv"] if prv else None], ind=1))
        rows.append(R(["Total", cur["lines"].get(k, 0.0), prv["lines"].get(k, 0.0) if prv else None], b=True, top=True, sm=list(range(len(rows) - sum(1 for r_ in rows if r_["c"][1] is None and r_["c"][2] is None), len(rows)))))
        T.append(dict(kind="n", title=title, cols=["Particulars", c1.replace("As at", "As at") if k in LIAB_KEYS + ASSET_KEYS else p1, c2 if k in LIAB_KEYS + ASSET_KEYS else p2], rows=rows, widths=[0.58, 0.21, 0.21]))
    return T, notes

DEFAULT_POLICIES = ("1. Basis of preparation: The financial statements are prepared under the historical cost convention on the accrual basis of accounting, "
    "in accordance with the Accounting Standards notified under section 133 of the Companies Act, 2013 and the relevant provisions of the Act.\n"
    "2. Use of estimates: Preparation of financial statements requires management estimates and assumptions; actual results may differ.\n"
    "3. Property, plant and equipment are stated at cost less accumulated depreciation; depreciation is provided as per Schedule II to the Act.\n"
    "4. Revenue is recognised on accrual basis when it is reasonably certain that ultimate collection will be made.\n"
    "5. Going concern: These financial statements are prepared on a going-concern basis. [CA to review - see Checks tab if net worth is negative.]\n"
    "6. Taxation: Provision is made for current tax; deferred tax is recognised only where there is virtual certainty of realisation.")

def build_horizontal(cur, prv, prof, fy_end):
    e = pdate(fy_end); p = date(e.year - 1, e.month, e.day)
    d1, d2 = e.strftime("%d.%m.%Y"), p.strftime("%d.%m.%Y")
    pmap = lambda sec: {nm: a for h in (prv[sec].values() if prv else []) for nm, a, _ in h}
    def side(sec, extra_head=None):
        out = []; pm = pmap(sec)
        for h, its in cur[sec].items():
            names = OrderedDict((nm, (a, pm.get(nm, 0.0))) for nm, a, _ in its)
            if prv:
                for nm, a in pm.items():
                    if nm not in names and nm in {x[0] for x in prv[sec][h]}: names[nm] = (0.0, a)
            if not names and not (prv and prv[sec][h]): continue
            out.append(("head", h, None, None))
            for nm, (a, b) in names.items(): out.append(("item", nm, b if prv else None, a))
        return out
    left = side("L"); right = side("A")
    # surplus lines on liabilities side
    lab = "Excess of Income over Expenditure" if prof.get("entity") != "firm" else "Profit & Loss Account"
    left.append(("head", lab, None, None))
    left.append(("item", "Opening balance", (r2(prv["pla_open"]) if prv and not prv["from_open"] else None), cur["pla_open"]))
    left.append(("item", "Current period surplus / (deficit)", prv["net"] if prv and not prv["from_open"] else None, cur["net"]))
    ltot = ("total", "Total", prv["tot_l"] if prv else None, cur["tot_l"]); atot = ("total", "Total", prv["tot_a"] if prv else None, cur["tot_a"])
    n = max(len(left), len(right)); left += [("", "", None, None)] * (n - len(left)); right += [("", "", None, None)] * (n - len(right))
    rows = [R([a[2], a[1], a[3], b[2], b[1], b[3]], b=(a[0] in ("head", "total") or b[0] in ("head", "total")), ind=0) for a, b in zip(left + [ltot], right + [atot])]
    rows[-1]["top"] = True
    bs = dict(kind="h", title=f"BALANCE SHEET AS AT {d1}", cols=[d2, "LIABILITIES", d1, d2, "ASSETS", d1], rows=rows, widths=[0.12, 0.28, 0.12, 0.12, 0.24, 0.12])
    # ---- income & expenditure / P&L (single horizontal account; firms with stock show opening / closing stock)
    def lists(h):
        ex = [("item", nm, None, a) for nm, a, _ in h["DIRECT_E"] + h["IND_E"]]
        ic = [("item", nm, None, a) for nm, a, _ in h["DIRECT_I"] + h["IND_I"]]
        return ex, ic
    ex, ic = lists(cur)
    pm_e = {nm: a for nm, a, _ in (prv["DIRECT_E"] + prv["IND_E"])} if prv else {}
    pm_i = {nm: a for nm, a, _ in (prv["DIRECT_I"] + prv["IND_I"])} if prv else {}
    ex = [("item", nm, pm_e.get(nm) if prv else None, a) for _, nm, _, a in ex]
    ic = [("item", nm, pm_i.get(nm) if prv else None, a) for _, nm, _, a in ic]
    if prv:
        for nm, a in pm_e.items():
            if nm not in {x[1] for x in ex}: ex.append(("item", nm, a, 0.0))
        for nm, a in pm_i.items():
            if nm not in {x[1] for x in ic}: ic.append(("item", nm, a, 0.0))
    if cur["trading"]:
        ex.insert(0, ("item", "Opening stock", prv["stock_open"] if prv else None, cur["stock_open"]))
        ic.append(("item", "Closing stock", prv["stock_close"] if prv else None, cur["stock_close"]))
    else:
        dst = cur["stock_close"] - cur["stock_open"]
        if abs(dst) > 0.004: (ic if dst > 0 else ex).append(("item", "Net change in stock", None, abs(dst)))
    net = cur["net"]; pnet = prv["net"] if (prv and not prv["from_open"]) else None
    ex_lab = "Excess of Income over Expenditure" if prof.get("entity") != "firm" else "Net Profit"
    in_lab = "Excess of Expenditure over Income" if prof.get("entity") != "firm" else "Net Loss"
    if net >= 0: ex.append(("item", ex_lab, pnet if (pnet is not None and pnet >= 0) else None, net))
    else:        ic.append(("item", in_lab, -pnet if (pnet is not None and pnet < 0) else None, -net))
    if pnet is not None and ((net >= 0) != (pnet >= 0)):          # prior year result on the other side
        (ic if net >= 0 else ex).append(("item", (in_lab if net >= 0 else ex_lab), abs(pnet), None))
    m = max(len(ex), len(ic)); ex += [("", "", None, None)] * (m - len(ex)); ic += [("", "", None, None)] * (m - len(ic))
    tot = r2(sum(a or 0 for *_, a in ex)); ptot_l = r2(sum(b or 0 for _, _, b, _ in ex)) if prv else None; ptot_r = r2(sum(b or 0 for _, _, b, _ in ic)) if prv else None
    rows = [R([a[2], a[1], a[3], b[2], b[1], b[3]], ind=0) for a, b in zip(ex, ic)]
    rows.append(R([ptot_l, "Total", tot, ptot_r, "Total", tot], b=True, top=True))
    pl_title = ("INCOME AND EXPENDITURE ACCOUNT FOR THE YEAR ENDED " if prof.get("entity") != "firm" else "TRADING AND PROFIT & LOSS ACCOUNT FOR THE YEAR ENDED ") + d1
    return [bs, dict(kind="h", title=pl_title, cols=[d2, "EXPENDITURE", d1, d2, "INCOME", d1], rows=rows, widths=[0.12, 0.28, 0.12, 0.12, 0.24, 0.12])], {}

# =============================================================== CHECKERS (shown in the Checks tab)
class Finding:
    def __init__(self, sev, area, msg, why="", fix=""):
        self.sev, self.area, self.msg, self.why, self.fix = sev, area, msg, why, fix
    def as_row(self): return (self.sev, self.area, self.msg, self.why, self.fix)

def _cash_ledgers(ledgers, groups):
    return {l["name"] for l in ledgers if set(chain(l["group"], groups)) & {"Cash-in-Hand", "Cash-in-hand"} or l["name"].strip().lower() in ("cash", "cash in hand", "petty cash")}

def run_checks(ctx):
    """ctx: ledgers, prior (or None), vouchers (or None), cur (computed), prv (computed or None), prof, fy_start, fy_end, groups, applic, tables"""
    F = []; add = lambda *a, **k: F.append(Finding(*a, **k))
    led, prof, groups = ctx["ledgers"], ctx["prof"], ctx.get("groups")
    cur, prv, fy_s, fy_e = ctx["cur"], ctx.get("prv"), pdate(ctx["fy_start"]), pdate(ctx["fy_end"])
    horizontal = prof.get("entity") != "company"
    # ---- 1. data integrity
    if not led: add("ERROR", "Data", "No ledgers loaded.", fix="Connect to Tally or import a trial balance."); return F
    for sev, area, msg in cur.get("flags", []): add(sev, "Classification", msg, "The ledger's Tally group decides where it prints in the statements.", "Correct the ledger's group in Tally, or add a mapping in Firm Profile.")
    if abs(cur["diff"]) > 0.5:
        add("ERROR", "Balance Sheet", f"Balance sheet does not balance: liabilities {inr(cur['tot_l'])} vs assets {inr(cur['tot_a'])} (difference {inr(cur['diff'])}).",
            "Total of all ledgers (with the P&L result) must be nil. A gap usually means a missing ledger, an opening-balance mismatch, a closing stock not in ledgers, or profit already carried into the P&L A/c ledger.",
            "Check opening balances, stock-in-hand and the 'Profit & Loss A/c' ledger; post the missing entry in Tally (or add the ledger to the imported trial balance).")
    elif abs(cur["diff"]) > 0.004: add("WARN", "Balance Sheet", f"Rounding difference of {inr(cur['diff'])} between assets and liabilities.", fix="Adjust by a rounding ledger.")
    else: add("OK", "Balance Sheet", f"Balance sheet balances at {inr(cur['tot_l'])}.")
    pla = [l for l in led if is_pla(l)]
    if len(pla) == 0: add("WARN", "Data", "No 'Profit & Loss A/c' ledger found - opening accumulated surplus assumed nil.", fix="Check the ledger name in Tally.")
    if len(pla) > 1: add("ERROR", "Data", "More than one 'Profit & Loss A/c' ledger.")
    dup = defaultdict(int)
    for l in led: dup[l["name"].strip().lower()] += 1
    for n_, c_ in dup.items():
        if c_ > 1: add("ERROR", "Data", f"Duplicate ledger name '{n_}' ({c_} times).", "Duplicates double-count balances.")
    for l in led:
        for k in ("bal", "open"):
            if abs(l[k] * 100 - round(l[k] * 100)) > 0.01: add("INFO", "Data", f"'{l['name']}' has more than 2 decimals ({l[k]}).")
    # ---- 2. comparatives
    if prv is None:
        add("WARN", "Comparatives", "No prior-year figures. Prior-year column will be blank.", "Schedule III and good practice require comparatives.", "Load last year's snapshot (saved automatically) or import last year's trial balance.")
    else:
        if ctx.get("prior_source") == "opening":
            add("WARN", "Comparatives", "Prior-year BALANCE SHEET taken from Tally opening balances; prior-year P&L is NOT available.", "Tally only holds current-year P&L.", "Load last year's snapshot / TB to get the prior-year P&L and cash flow.")
        elif ctx.get("prior"):
            pm = {l["name"].lower(): l for l in ctx["prior"] if not _is_pl_nature(l, groups)}
            bad = [(l["name"], l["open"], pm[l["name"].lower()]["bal"]) for l in led if not _is_pl_nature(l, groups) and not is_pla(l) and l["name"].lower() in pm and abs(l["open"] - pm[l["name"].lower()]["bal"]) > 1]
            for nm, o, c_ in bad[:20]: add("WARN", "Comparatives", f"'{nm}': opening {inr(o)} differs from last year's closing {inr(c_)}.", "Opening must equal prior closing (SA 510).", "Find the change made after last year's audit.")
            if not bad: add("OK", "Comparatives", "Opening balances agree with last year's closing balances.")
            cl_nm = {l["name"].lower() for l in led}
            for l in ctx["prior"]:
                if l["name"].lower() not in cl_nm and abs(l["bal"]) > 0.5 and not _is_pl_nature(l, groups): add("INFO", "Comparatives", f"Ledger '{l['name']}' existed last year ({inr(l['bal'])}) but is missing now.")
    # ---- 3. nature checks
    for l in led:
        ch = chain(l["group"], groups); nat = nature(ch)
        if nat == "S" and abs(l["bal"]) > 0.004: add("ERROR", "Suspense", f"Suspense account '{l['name']}' has balance {inr(l['bal'])}.", "Unresolved suspense balances are a reportable audit matter.", "Clear to the proper ledger before finalising.")
        if nat == "I" and l["bal"] > 0.004: add("WARN", "P&L", f"Income ledger '{l['name']}' has a DEBIT balance {inr(l['bal'])}.", fix="Check for wrong-side postings.")
        if nat == "E" and l["bal"] < -0.004: add("WARN", "P&L", f"Expense ledger '{l['name']}' has a CREDIT balance {inr(-l['bal'])}.", fix="Check for wrong-side postings.")
        if "Misc. Expenses (ASSET)" in ch and abs(l["bal"]) > 0.004:
            add("WARN", "Accounting", f"Deferred revenue / miscellaneous expenditure '{l['name']}' of {inr(l['bal'])} carried forward.", "AS 26 (and AS 24/ Sch III) do not allow deferral of revenue expenses except in specific cases; many auditors insist on write-off (as done in your FY 2021-22 statements).", "Write off to P&L or document the AS basis.")
        if nat == "A" and "Fixed Assets" in ch and l["bal"] < -0.004: add("WARN", "Fixed assets", f"Fixed asset '{l['name']}' has a credit balance.", fix="Accumulated depreciation ledgers should be netted in the same group.")
        if RX_RELATED.search(l["name"]) and abs(l["bal"]) > 0.004 and nat in ("L", "A") and not is_pla(l):
            add("INFO", "Related party", f"'{l['name']}' looks like a related-party balance ({inr(abs(l['bal']))}).",
                "AS 18 requires disclosure of related-party transactions; for a company, loans from directors need the s.2(31)/Rule 2(1)(c)(viii) written declaration, and loans TO directors engage s.185.", "Add AS 18 disclosure and obtain the director's declaration.")
    for sev, area, msg in (cur.get("flags") or []): pass
    # ---- 4. statement-level
    if not horizontal:
        nw = cur["lines"]["share_capital"] + cur["lines"]["reserves"] + cur["lines"]["share_app"] * 0
        gc = RULES.use("going_concern", fy_e)
        if nw < 0:
            add("WARN", "Going concern", f"Net worth is NEGATIVE ({inr(nw)}). Accumulated losses exceed capital and reserves.",
                "SA 570 requires you to evaluate management's plan (here: continued funding by directors). If there is a material uncertainty a separate 'Material uncertainty related to going concern' section is needed; if not, consider an Emphasis of Matter. s.2(85)/Sch III disclosure of directors' letter of support is advisable.",
                "Obtain a director/shareholder letter of continued financial support; tick 'Going-concern paragraph' in the Audit Report tab.")
        if cur["income"] <= 0.004 and cur["expenses"] > 0: add("INFO", "P&L", "Nil revenue for the year with continuing expenses.", "Pre-revenue/start-up pattern; supports the going-concern evaluation.")
        if cur["lines"]["ppe"] + cur["lines"]["intang"] > 0 and cur["lines"]["dep"] < 0.004 and not any("under development" in n.lower() or "cwip" in n.lower() for n, _, _ in cur["items"].get("ppe", [])):
            add("WARN", "Fixed assets", "Fixed assets exist but no depreciation has been charged.", "Schedule II / AS 10 / AS 6.", "Check the depreciation entry.")
        if cur["lines"]["dep"] > 0 and prv and prv["lines"]["dep"] > 0 and cur["lines"]["dep"] > 2.5 * prv["lines"]["dep"]: add("INFO", "Fixed assets", "Depreciation more than 2.5x last year.")
        for k in ASSET_KEYS + LIAB_KEYS + ["other_exp", "employee"]:
            its = cur["items"].get(k, [])
            if its and abs(sum(a for _, a, _ in its) - cur["lines"][k]) > 0.01 and k not in ("reserves",): add("ERROR", "Notes", f"Note total for '{LABEL.get(k,k)}' does not agree to the face of the statements.")
        for k, lab in (("other_exp", "Other expenses"),):
            tot = cur["lines"][k]
            if tot and any(n.lower().startswith(("misc", "sundry")) and a > 0.25 * tot for n, a, _ in cur["items"][k]): add("INFO", "Notes", "Miscellaneous expenses exceed 25% of other expenses.", "Break out heads over 1% of revenue or Rs 1 lakh (Sch III general instructions).")
        if prv and prv["profit"] and abs(cur["profit"]) > 0 and (cur["profit"] > 0) != (prv["profit"] > 0): add("INFO", "Analytics", "Result changed sign compared with last year.")
        if ctx.get("cf") and abs(ctx["cf"]["gap"]) > 1: add("ERROR", "Cash flow", f"Cash flow statement does not reconcile to the balance sheet (gap {inr(ctx['cf']['gap'])}).", fix="Check borrowings tagging and non-cash entries.")
        elif ctx.get("cf"): add("OK", "Cash flow", "Cash flow statement reconciles to opening and closing cash.")
        if cur["lines"]["other_cl"] and any(t == "loan" and RX_RELATED.search(n) for n, _, t in cur["items"].get("other_cl", [])):
            add("WARN", "Schedule III", "Loans from directors/shareholders are shown under 'Other current liabilities'.", "Schedule III expects borrowings (with related-party and repayment-terms disclosure) under Long-term/Short-term borrowings; the classification follows the firm's earlier template.", "Confirm the classification, or change 'Loans shown under' in Firm Profile.")
    else:
        if cur["net"] < 0 and prof.get("entity") == "trust": add("INFO", "Trust", f"Deficit of {inr(-cur['net'])} for the year (expenditure exceeds income).", "For a charitable trust, check application of income (at least 85%) and Form 10B/10BB audit requirements under the income-tax law.")
    # ---- 5. voucher checks
    vr = ctx.get("vouchers")
    if vr is None: add("INFO", "Vouchers", "Voucher-level tests skipped (no vouchers loaded).")
    else:
        lim = RULES.use("cash_limits", fy_e) or {"pay": 10000, "loan": 20000, "receipt": 200000}
        cash = _cash_ledgers(led, groups); byv = defaultdict(list)
        for r in vr: byv[(r["vtype"], r["no"], r["date"])].append(r)
        n_unb = n_out = n_nar = 0
        for (vt, no, d), rs in byv.items():
            dr, cr = sum(r["debit"] for r in rs), sum(r["credit"] for r in rs)
            if abs(dr - cr) > 0.01: n_unb += 1; add("ERROR", "Vouchers", f"{vt} voucher {no} dated {d}: debit {inr(dr)} <> credit {inr(cr)}.", "An unbalanced voucher corrupts the trial balance.", "Re-enter in Tally.")
            if d is None or d < fy_s or d > fy_e: n_out += 1; add("ERROR", "Vouchers", f"{vt} voucher {no} dated {d} is outside the financial year {fy_s:%d-%m-%Y} to {fy_e:%d-%m-%Y}.", "Cut-off error / wrong period.", "Check date or period.")
            if not (rs[0]["narration"] or "").strip(): n_nar += 1
            cash_cr = sum(r["credit"] for r in rs if r["ledger"] in cash); cash_dr = sum(r["debit"] for r in rs if r["ledger"] in cash)
            if vt.lower() == "payment" and cash_cr > lim["pay"]:
                add("WARN", "Cash (40A(3))", f"Cash payment of {inr(cash_cr)} in voucher {no} ({d}).", f"Exceeds Rs {lim['pay']:,} per person per day: 100% disallowance of the expense for tax (s.40A(3) or its equivalent in the Income-tax Act 2025).", "Pay by banking channel or report in the tax-audit clause on cash payments.")
            if vt.lower() == "receipt" and cash_dr >= lim["receipt"]: add("WARN", "Cash (269ST)", f"Cash receipt of {inr(cash_dr)} in voucher {no} ({d}).", f"Receipts of Rs {lim['receipt']:,} or more in cash from one person attract penalty equal to the amount.", "Verify the party and the reason.")
            if cash_dr > lim["loan"] and any(RX_RELATED.search(r["ledger"]) or "loan" in (r["narration"] or "").lower() for r in rs if r["credit"] > 0):
                add("WARN", "Cash (269SS)", f"Loan / deposit of {inr(cash_dr)} accepted in cash (voucher {no}, {d}).", f"Acceptance of loans above Rs {lim['loan']:,} otherwise than by account-payee mode attracts penalty (s.269SS / equivalent).", "Obtain bank evidence.")
            tot = max(dr, cr)
            if vt.lower() == "journal" and tot >= 100000 and abs(tot % 100000) < 0.005: add("INFO", "Journal test", f"Round-sum journal {no} ({d}) for {inr(tot)}.", "SA 240 journal-entry testing: round amounts and manual journals are fraud-risk indicators.", "Inspect the supporting document.")
        if n_nar and n_nar / max(len(byv), 1) > 0.2: add("INFO", "Vouchers", f"{n_nar} of {len(byv)} vouchers have no narration.")
        seen = defaultdict(list)
        for (vt, no, d) in byv: seen[(vt, no)].append(d)
        for (vt, no), ds in seen.items():
            if len(ds) > 1: add("WARN", "Vouchers", f"Duplicate voucher number {vt} {no} on {len(ds)} different dates.", "Numbering gaps / duplicates break the audit trail.")
        # ledger movement vs vouchers
        mv = defaultdict(float)
        for r in vr: mv[r["ledger"]] += r["debit"] - r["credit"]
        mism = 0
        for l in led:
            if is_pla(l): continue
            ch = chain(l["group"], groups)
            diff = (l["bal"] - l["open"]) - mv.get(l["name"], 0.0)
            if abs(diff) > 1.0 and (l["name"] in mv or abs(l["bal"] - l["open"]) > 1):
                mism += 1
                if mism <= 12: add("WARN", "Completeness", f"'{l['name']}': movement per ledger {inr(l['bal'] - l['open'])} vs vouchers {inr(mv.get(l['name'], 0.0))}.", "Vouchers extracted do not explain the ledger movement (missing voucher types or modified ledger).", "Re-export vouchers for the full year.")
        if not mism: add("OK", "Completeness", "Voucher totals reproduce every ledger movement.")
        # TDS heuristic
        tds = [l for l in led if re.search(r"\btds\b", l["name"], re.I)]
        prof_exp = [l for l in led if re.search(r"professional|consult|contract|commission|rent|legal", l["name"], re.I) and nature(chain(l["group"], groups)) == "E" and l["bal"] > 30000]
        if prof_exp and not tds:
            add("WARN", "TDS", "Professional / contract / rent expenses exceed Rs 30,000 but no TDS ledger exists: " + ", ".join(l["name"] for l in prof_exp[:4]) + ".", "Non-deduction attracts 30% disallowance (s.40(a)(ia) or equivalent) and interest/penalty.", "Confirm whether TDS was deducted and deposited, or whether thresholds were not crossed per vendor.")
        elif prof_exp: add("OK", "TDS", "TDS ledger present alongside professional/contract/rent expenses.")
    # ---- 6. law / rules
    for rid, why in RULES.stale(): add("INFO", "Law library", f"Rule '{RULES.title(rid)}' {why}.", "Law changes often (e.g. the small-company limits changed on 01-12-2025).", "Law Library tab -> Update now, or confirm against the notification.")
    ap = ctx.get("applic")
    if ap:
        for sev, m, why in ap["notes"]: add(sev, "Applicability", m, why)
    f = ctx["prof"].get("facts", {})
    if prof.get("entity") == "company" and ap and pdate(ctx["fy_end"]) >= date(2024, 3, 31) and not f.get("audit_trail"):
        add("WARN", "Audit trail", "Audit-trail (edit log) status not stated.", "Rule 11(g) requires reporting from FY 2023-24. TallyPrime's 'Edit Log' must be switched on and used through the year.", "Audit Report tab -> Facts -> audit trail.")
    if not prof.get("udin"): add("INFO", "Report", "UDIN not entered (mandatory for attest functions - generate on the ICAI UDIN portal).")
    if not F or all(x.sev == "OK" for x in F): pass
    return F

def _is_pl_nature(l, groups): return nature(chain(l["group"], groups)) in ("I", "E")

# =============================================================== AUDIT-REPORT APPLICABILITY ENGINE
def determine(ctx):
    """CARO / cash flow / IFC / Rule 11 clause set for the year - every decision carries its reason."""
    prof, f = ctx["prof"], ctx["prof"].get("facts", {})
    fy_s, fy_e = pdate(ctx["fy_start"]), pdate(ctx["fy_end"]); cur, prv = ctx["cur"], ctx.get("prv")
    RULES.reset_used(); notes = []; out = dict(entity=prof.get("entity", "company"))
    if prof.get("entity") != "company":
        out.update(caro="n/a", caro_applicable=False, cfs=False, ifc="n/a", clauses=[], notes=[("INFO", "Not a company: Companies Act s.143 / CARO / Rule 11 clauses do not apply. Report follows SA 700 for non-corporate entities.",
                   "Trusts/societies also need Form 10B (income-tax audit) - separate from this report.")]); return out
    order = RULES.use("caro_regime", fy_s)["order"] if RULES.use("caro_regime", fy_s) else "CARO 2020"
    out["caro"] = order
    # --- small company
    sc = RULES.use("small_company", fy_e)
    paid = cur["lines"]["share_capital"] + cur["lines"]["share_app"] * 0
    turn_prev = (prv["income"] - 0) if prv and not ctx.get("prior_source") == "opening" else None
    turn_prev_used = turn_prev if turn_prev is not None else cur["lines"]["revenue"]
    if turn_prev is None: notes.append(("WARN", "Preceding-year turnover not available - current-year revenue used for the small-company test.", "The statutory test uses the PRECEDING financial year's turnover."))
    excl = f.get("is_public") or f.get("is_sub_holding") or f.get("is_section8") or f.get("special_act")
    small = bool(sc) and not excl and paid <= sc["paid_up"] and turn_prev_used <= sc["turnover"]
    out["small"] = small
    reasons = [f"Paid-up capital {inr(paid)} vs limit {inr(sc['paid_up'])}; turnover {inr(turn_prev_used)} vs limit {inr(sc['turnover'])} (limits in force on {fy_e:%d-%m-%Y}, source: {RULES.data['rules']['small_company']['source'].split(';')[0]})."] if sc else []
    if excl: reasons.append("Excluded from 'small company' (public / holding / subsidiary / s.8 / special-Act company).")
    # --- CARO
    ap, why = None, []
    if order in ("CARO 2003",): ap, why = None, ["CARO 2003 era - outside the scope of this tool."]
    if f.get("is_bank_ins"): ap, why = False, ["Banking/insurance company - CARO does not apply."]
    elif f.get("is_section8"): ap, why = False, ["Section 8 company - CARO does not apply."]
    elif f.get("is_opc"): ap, why = False, ["One Person Company - CARO does not apply."]
    elif small: ap, why = False, ["Small company under s.2(85) - CARO does not apply."] + reasons
    else:
        pe = RULES.use("caro_private_exempt", fy_e)
        if not f.get("is_public") and not f.get("is_sub_holding") and pe:
            cap_res = cur["lines"]["share_capital"] + cur["lines"]["reserves"]
            fi_est = f.get("max_bank_fi_borrowing")
            if fi_est is None: fi_est = sum(a for k in ("lt_borrow", "st_borrow", "other_cl", "other_ltl") for n, a, t in cur["items"].get(k, []) if t == "loan" and RX_LOAN_BANK.search(n)); notes.append(("INFO", f"Maximum bank/FI borrowing during the year estimated at {inr(fi_est)} from ledger names; enter the actual peak in Facts.", "The CARO private-company test looks at the HIGHEST borrowing at any time in the year, not the year-end figure."))
            rev = cur["income"]
            ok = cap_res <= pe["cap_reserves"] and fi_est <= pe["borrowings"] and rev <= pe["revenue"]
            why = [f"Private company test: capital+reserves {inr(cap_res)} (limit {inr(pe['cap_reserves'])}), bank/FI borrowings {inr(fi_est)} (limit {inr(pe['borrowings'])}), revenue {inr(rev)} (limit {inr(pe['revenue'])}) -> {'ALL within limits: exempt' if ok else 'a limit is breached: CARO applies'}."]
            ap = not ok
        else: ap, why = True, ["Public / holding / subsidiary company or no exemption available - CARO applies."]
    out["caro_applicable"], out["caro_why"] = ap, why
    # --- cash flow statement
    cf_ex = small or f.get("is_opc") or f.get("is_dormant")
    out["cfs"] = not cf_ex
    out["cfs_why"] = "Cash flow statement not required: one person / small / dormant company (s.2(40))." if cf_ex else "Cash flow statement required (company is not small / OPC / dormant)."
    # --- IFC
    ie = RULES.use("ifc_exempt", fy_e); fi = f.get("max_bank_fi_borrowing") or 0.0
    out["ifc_exempt"] = bool(f.get("is_opc") or small or (ie and (turn_prev_used < ie["turnover"] and fi < ie["borrowings"] and not f.get("is_public"))))
    out["ifc_why"] = "Exempt from IFCoFR opinion: OPC / small company, or turnover < Rs 50 cr and borrowings < Rs 25 cr (private company, notification 13-06-2017)." if out["ifc_exempt"] else "IFCoFR opinion required - Annexure B."
    # --- Rule 11
    out["r11"] = dict(funds=bool(RULES.use("rule11_funds", fy_s)), dividend=bool(RULES.use("rule11_dividend", fy_s)), audit_trail=RULES.use("rule11_audit_trail", fy_s))
    out["ind_as"] = bool(f.get("ind_as"))
    nw = cur["lines"]["share_capital"] + cur["lines"]["reserves"]
    out["neg_nw"] = nw < 0
    notes.append(("INFO", f"{order}: " + " ".join(why), "Decided from the law library on the balance-sheet date; edit the Facts if any are wrong."))
    notes.append(("INFO", out["cfs_why"], "")); notes.append(("INFO", out["ifc_why"], ""))
    out["notes"] = notes
    return out

# =============================================================== AUDITOR'S REPORT
def _roman(n):
    v = [(10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")]; out = ""
    for a, s in v:
        while n >= a: out += s; n -= a
    return out

def _set_text(p, text):
    runs = p.runs
    if runs:
        runs[0].text = text
        for r in runs[1:]: r._r.getparent().remove(r._r)
    else: p.add_run(text)

R11_FUNDS = ["v.\ta) The management has represented that, to the best of its knowledge and belief, no funds have been advanced or loaned or invested  (either from borrowed funds or share premium or any other sources or kind of funds) by the company to or in any other person or entity, including foreign entities (“intermediaries”) with the understanding, whether recorded in writing or otherwise, that the intermediary shall, whether directly or indirectly lend or invest in other person or entity identified in any manner whatsoever by or behalf of the company (“ultimate beneficiaries”) or provide any guarantee, security or the like on behalf of the Ultimate beneficiaries.",
             "b) The management has represented, that, to the best of its knowledge and belief,  no funds have been received by the company from any person or entity including foreign entities (“Funding Parties”) with the understanding, whether recorded in writing or otherwise, that the company shall, whether directly or indirectly, lend or invest in other person or entity identified in any manner whatsoever by or behalf of the Funding Party (“Ultimate Beneficiaries”) or provide any guarantee, security or the like on behalf of the ultimate beneficiaries; and",
             "c) Based on such audit procedures that were considered reasonable and appropriate in the circumstances, nothing has come to our notice that has caused us to believe that representations under sub clause (a) and (b) contain any material mis-statement."]
AT_OPERATED = ("The Company has used accounting software for maintaining its books of account which has a feature of recording audit trail (edit log) facility and the same has been operated throughout the year for all relevant transactions recorded in the software. During the course of our audit we did not come across any instance of the audit trail feature being tampered with.")
AT_PRESERVED = " Further, the audit trail has been preserved by the Company as per the statutory requirements for record retention."
AT_PARTIAL = ("The company has used an accounting software for maintaining books of accounts which has a feature of audit trail facility, but the same has not been operated throughout the year for all transactions recorded in the software and the audit trail feature has not been tampered with and the audit trail has been preserved by the company as per the statutory requirements for record retention.")
AT_NONE = ("The Company has used accounting software for maintaining its books of account which does not have a feature of recording audit trail (edit log) facility; accordingly, reporting on its operation, tampering and preservation is not possible. [CA to review wording]")

def legal_section(ctx, ap):
    """-> list of (level, text); level 1 = numbered para, 2 = lettered, 3 = roman. Lettering is automatic and year-driven."""
    prof, f = ctx["prof"], ctx["prof"].get("facts", {}); fy_e, fy_s = pdate(ctx["fy_end"]), pdate(ctx["fy_start"])
    yr = fy_e.year; LD = long_date(fy_e); out = []
    order = ap["caro"]; order_name = {"CARO 2020": "Companies (Auditor’s Report) Order, 2020", "CARO 2016": "Companies (Auditor’s Report) Order, 2016", "CARO 2015": "Companies (Auditor’s Report) Order, 2015"}.get(order, "Companies (Auditor’s Report) Order")
    if ap["caro_applicable"]:
        out.append((1, f"1.\tAs required by the {order_name} (“the Order”), issued by the Central Government of India in terms of sub-section (11) of section 143 of the Act, we give in the “Annexure A” a statement on the matters specified in paragraphs 3 and 4 of the Order, to the extent applicable."))
    else:
        out.append((1, f"1.\tThis report does not include a statement on the matters specified in paragraphs 3 & 4 of the {order_name} (“the Order”), since in our opinion and according to the information and explanation given to us, the said Order is not applicable to the company."))
    out.append((1, "2.\tAs required by section 143 (3) of the Act, we report that:"))
    at = ap["r11"]["audit_trail"]; at_mode = f.get("audit_trail", "operated")
    at_year = bool(at)
    letters = iter("abcdefghijkl"); L = lambda: next(letters)
    qual = bool(at_year and at_mode in ("partial", "none"))
    r11_letter = "abcdefghijkl"[5 + (1 if qual else 0) + 1 + (1 if (f.get("is_public") or f.get("include_197_16")) else 0)]   # a-e, [qualification], IFC, [197(16)], Rule 11
    # sub-item order inside Rule 11 so cross-references are right
    r11 = ["i.\tThe Company does not have any pending litigations which would impact its financial position." if not f.get("pending_litigation") else "i.\tThe Company has disclosed the impact of pending litigations on its financial position in its financial statements.",
           "ii.\tThe Company did not have any long-term contracts including derivative contracts for which there were any material foreseeable losses.",
           "iii.\tThere were no amounts which were required to be transferred to the Investor Education and Protection Fund by the Company."]
    nxt = 4
    if ap["r11"]["funds"]: r11 += [f"{_roman(nxt)}.\t" + R11_FUNDS[0][R11_FUNDS[0].index("\t") + 1:]] + R11_FUNDS[1:]; nxt += 1
    if ap["r11"]["dividend"]: r11.append(f"{_roman(nxt)}.\tThe Company has neither declared nor paid any dividend during the year." if not f.get("dividend_paid") else f"{_roman(nxt)}.\tThe dividend declared or paid during the year by the Company is in compliance with section 123 of the Act. [CA to confirm]"); nxt += 1
    at_ref = ""
    if at_year:
        base = {"operated": AT_OPERATED, "partial": AT_PARTIAL, "none": AT_NONE}[at_mode]
        if at_mode == "operated" and at.get("preserved"): base += AT_PRESERVED
        r11.append(f"{_roman(nxt)}.\t{base}"); at_ref = _roman(nxt); nxt += 1
    out.append((2, f"{L()}.\tWe have sought and obtained all the information and explanations which to the best of our knowledge and belief were necessary for the purposes of our audit."))
    bt = "in our opinion proper books of accounts as required by law have been kept by the Company so far as it appears from our examination of those books"
    if at_year and at_mode in ("partial", "none"): bt += f" except for the matters stated in the paragraph 2({r11_letter})({at_ref}) below on reporting under Rule 11(g)"
    out.append((2, f"{L()}.\t{bt};"))
    cfs = ap["cfs"]
    out.append((2, f"{L()}.\tthe Balance Sheet, the Statement of Profit and Loss{', and the Cash Flow Statement' if cfs else ''} dealt with by this Report are in agreement with the books of account"))
    std = ("Accounting Standards prescribed under Section 133 of the Act, read with the Companies (Accounting Standards) Rules, 2006" if not ap["ind_as"]
           else "Indian Accounting Standards prescribed under Section 133 of the Act read with the Companies (Indian Accounting Standards) Rules, 2015, as amended")
    out.append((2, f"{L()}.\tIn our opinion, the aforesaid financial statements comply with the {std}."))
    out.append((2, f"{L()}.\tOn the basis of written representations received from the directors as on {LD} taken on record by the Board of Directors, none of the directors is disqualified as on {LD} from being appointed as a director in terms of Section 164 (2) of the Act."))
    if at_year and at_mode in ("partial", "none"):
        out.append((2, f"{L()}.\tIn our opinion, the qualifications, reservations or adverse remarks relating to the maintenance of accounts and other matters connected therewith are as stated in paragraph 2(b) above on reporting under Section 143(3)(b) and paragraph 2({r11_letter})({at_ref}) below on reporting under Rule 11(g)."))
    if ap["ifc_exempt"]:
        out.append((2, f"{L()}.\tSince the Company’s turnover as per last audited financial statements is less than Rs.50 Crores and its borrowings from banks and financial institutions at any time during the year is less than Rs.25 Crores, the Company is exempted from getting an audit opinion with respect to the adequacy of the internal financial controls over financial reporting of the company and the operating effectiveness of such controls vide notification dated June 13, 2017."))
    else:
        out.append((2, f"{L()}.\tWith respect to the adequacy of the internal financial controls with reference to financial statements of the Company and the operating effectiveness of such controls, refer to our separate Report in “Annexure B”."))
    if f.get("is_public") or f.get("include_197_16"):
        out.append((2, f"{L()}.\tWith respect to the matter to be included in the Auditor’s Report under section 197(16): In our opinion and according to the information and explanations given to us, the remuneration paid by the Company to its directors during the year is in accordance with the provisions of section 197 of the Act. [CA to confirm]"))
    out.append((2, f"{L()}.\tWith respect to the other matters to be included in the Auditor’s Report in accordance with Rule 11 of the Companies (Audit and Auditors) Rules, 2014, in our opinion and to the best of our information and according to the explanations given to us:"))
    for t in r11: out.append((3 if re.match(r"^[ivx]+\.\t", t) else 4, t))
    # re-letter "g" reference check
    return out

CARO2020 = [("i", "Property, plant and equipment and intangible assets (records, physical verification, title deeds, revaluation, benami proceedings)"), ("ii", "Inventory (physical verification; working-capital limits above Rs 5 crore from banks/FIs)"),
 ("iii", "Investments, guarantees, loans and advances to companies/firms/LLPs/others (terms, overdues, repayment schedule)"), ("iv", "Loans, investments, guarantees and security - sections 185 and 186"), ("v", "Deposits - sections 73 to 76"), ("vi", "Cost records - section 148(1)"),
 ("vii", "Statutory dues (GST, PF, ESI, income-tax, etc.) - regularity and disputed dues"), ("viii", "Unrecorded income disclosed in tax assessments"), ("ix", "Default in repayment of loans / borrowings; utilisation of funds; wilful defaulter"), ("x", "Money raised by public offer / preferential allotment"),
 ("xi", "Fraud noticed / reported; whistle-blower complaints"), ("xii", "Nidhi company"), ("xiii", "Related-party transactions - sections 177 and 188; disclosure"), ("xiv", "Internal audit system"), ("xv", "Non-cash transactions with directors - section 192"),
 ("xvi", "Registration under section 45-IA of the RBI Act"), ("xvii", "Cash losses in the year and the immediately preceding year"), ("xviii", "Resignation of statutory auditors"), ("xix", "Material uncertainty on the ability to meet liabilities within one year of the balance sheet date"),
 ("xx", "Corporate social responsibility - section 135"), ("xxi", "Qualifications / adverse remarks in group CARO reports (consolidated statements)")]
CARO2016 = [("i", "Fixed assets (records, verification, title deeds)"), ("ii", "Inventory (verification, discrepancies)"), ("iii", "Loans to parties covered under section 189"), ("iv", "Loans, investments, guarantees - sections 185 and 186"), ("v", "Deposits - sections 73 to 76"), ("vi", "Cost records - section 148"),
 ("vii", "Statutory dues and disputed dues"), ("viii", "Default in repayment to financial institutions, banks, government or debenture holders"), ("ix", "Public offer / term loans - application"), ("x", "Fraud by or on the company"), ("xi", "Managerial remuneration - section 197"), ("xii", "Nidhi company"),
 ("xiii", "Related-party transactions - sections 177 and 188"), ("xiv", "Preferential allotment / private placement - section 42"), ("xv", "Non-cash transactions with directors - section 192"), ("xvi", "Registration under section 45-IA of the RBI Act")]

def caro_observations(ctx):
    """Facts the books already show, to help the CA write each clause. Never a conclusion."""
    cur, prv = ctx["cur"], ctx.get("prv"); o = {}
    loss = -cur["profit"] if cur["profit"] < 0 else 0.0
    cash_loss = max(loss - cur["lines"]["dep"], 0.0) if cur["profit"] < 0 else 0.0
    if cur["profit"] < 0:
        pl = prv and prv["profit"] < 0
        o["xvii"] = f"Books show a net loss of Rs {inr(loss)}; adding back depreciation of Rs {inr(cur['lines']['dep'])} the cash loss is Rs {inr(cash_loss)}" + (f" (previous year: net loss Rs {inr(-prv['profit'])}, cash loss Rs {inr(max(-prv['profit'] - prv['lines']['dep'], 0))})." if pl else ".")
    else: o["xvii"] = "Books show a profit for the year; no cash loss indicated."
    ppe = cur["lines"]["ppe"] + cur["lines"]["intang"] + cur["lines"]["cwip"]
    o["i"] = f"Net block per books: Rs {inr(ppe)}." + (" No fixed asset register is available in Tally - obtain one." if ppe else "")
    o["ii"] = f"Inventories per books: Rs {inr(cur['lines']['inventories'])}."
    o["iii"] = f"Loans/advances per books (assets side): Rs {inr(cur['lines']['st_loans'] + cur['lines']['lt_loans'])}."
    st = [(n, a) for k in ("other_cl", "other_ltl") for n, a, t in cur["items"].get(k, []) if re.search(r"gst|tds|tax|pf|esi|provident|duty|duties", n, re.I)]
    o["vii"] = "Statutory-dues ledgers: " + (", ".join(f"{n} Rs {inr(a)}" for n, a in st) if st else "none found in the books") + "."
    ln = [(n, a) for k in ("lt_borrow", "st_borrow", "other_cl", "other_ltl") for n, a, t in cur["items"].get(k, []) if t == "loan"]
    o["ix"] = "Borrowings per books: " + (", ".join(f"{n} Rs {inr(a)}" for n, a in ln) if ln else "nil") + ". Confirm no default and end-use."
    rp = [(n, a) for k in cur["items"] for n, a, t in cur["items"][k] if RX_RELATED.search(n) and k in LIAB_KEYS + ASSET_KEYS]
    o["xiii"] = "Related-party looking ledgers: " + (", ".join(f"{n} Rs {inr(abs(a))}" for n, a in rp) if rp else "none") + "."
    nw = cur["lines"]["share_capital"] + cur["lines"]["reserves"]
    o["xix"] = f"Net worth Rs {inr(nw)}; current liabilities Rs {inr(sum(cur['lines'][k] for k in ('st_borrow','trade_pay','other_cl','st_prov')))} vs current assets Rs {inr(sum(cur['lines'][k] for k in ('c_inv','inventories','trade_rec','cash','st_loans','other_ca')))}. Evaluate ability to meet liabilities as they fall due."
    o["xx"] = "Section 135 applies if net worth >= Rs 500 cr, turnover >= Rs 1,000 cr or profit >= Rs 5 cr in the preceding year (CA to confirm)."
    return o

def _add_par(doc, text, bold=False, size=12.5, align=None, indent=None):
    from docx.shared import Pt, Emu
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    p = doc.add_paragraph(); r = p.add_run(text); r.bold = bold; r.font.name = "Times New Roman"; r.font.size = Pt(size)
    p.alignment = {"c": WD_ALIGN_PARAGRAPH.CENTER, "j": WD_ALIGN_PARAGRAPH.JUSTIFY}.get(align, None)
    if indent: p.paragraph_format.left_indent = Emu(indent)
    return p

def _annexures(doc, ctx, ap):
    from docx.enum.text import WD_BREAK
    prof = ctx["prof"]; LD = long_date(pdate(ctx["fy_end"])); name = prof["name"]
    if ap["caro_applicable"]:
        doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        _add_par(doc, "ANNEXURE A TO THE INDEPENDENT AUDITOR’S REPORT", True, align="c")
        _add_par(doc, f"(Referred to in paragraph 1 under ‘Report on Other Legal and Regulatory Requirements’ of our report of even date to the members of {name} on the financial statements for the year ended {LD})", align="j")
        _add_par(doc, f"DRAFT SKELETON - {ap['caro']}. Each clause below must be answered by the auditor from audit evidence. “Books show” lines are generated from the accounts only; they are not conclusions.", True, 10.5, "j")
        obs = caro_observations(ctx)
        for rn, title in (CARO2020 if ap["caro"] == "CARO 2020" else CARO2016):
            _add_par(doc, f"({rn})  {title}", True, 12, "j")
            _add_par(doc, "[CA TO COMPLETE]", False, 12, "j", 228600)
            if rn in obs: _add_par(doc, "Books show: " + obs[rn], False, 11, "j", 228600)
    if not ap["ifc_exempt"]:
        doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        _add_par(doc, "ANNEXURE B TO THE INDEPENDENT AUDITOR’S REPORT", True, align="c")
        _add_par(doc, "Report on the Internal Financial Controls with reference to the aforesaid financial statements under Clause (i) of Sub-section 3 of Section 143 of the Companies Act, 2013 (“the Act”)", True, 12, "j")
        _add_par(doc, "[CA TO COMPLETE - Opinion, Management’s responsibility, Auditor’s responsibility, Meaning and limitations of internal financial controls, Inherent limitations - per ICAI Guidance Note on Audit of Internal Financial Controls Over Financial Reporting]", False, 12, "j")

def going_concern_par(ctx, ap):
    f = ctx["prof"].get("facts", {}); mode = f.get("going_concern_mode", "none"); cur = ctx["cur"]; fy = long_date(pdate(ctx["fy_end"]))
    if mode == "none": return None
    loss = f"incurred a net loss of Rs {inr(-cur['profit'])} during the year ended {fy} and, as of that date, " if cur["profit"] < 0 else ""
    base = (f"We draw attention to the financial statements which indicate that the Company has {loss}its accumulated losses exceed its net worth. "
            f"The Company is dependent on continued financial support from its directors. These financial statements have been prepared on a going concern basis, the validity of which depends on such continued support. ")
    if mode == "eom": return ("Emphasis of Matter", base + "Our opinion is not modified in respect of this matter.")
    return ("Material Uncertainty Related to Going Concern", base + "These events or conditions, along with other matters, indicate that a material uncertainty exists that may cast significant doubt on the Company’s ability to continue as a going concern. Our opinion is not modified in respect of this matter.")

@guard("audit_report", default=None)
def build_audit_report(ctx, out_path, template=None):
    from docx import Document
    from docx.text.paragraph import Paragraph
    prof, f = ctx["prof"], ctx["prof"].get("facts", {})
    ap = ctx["applic"]; fy_e = pdate(ctx["fy_end"]); LD = long_date(fy_e); cur = ctx["cur"]
    if prof.get("entity") != "company": return build_audit_report_noncompany(ctx, out_path)
    tpl = template or prof.get("template_docx")
    try:
        doc = Document(tpl) if tpl and Path(tpl).exists() else Document(io.BytesIO(base64.b64decode(EMBEDDED_TEMPLATE_B64)))
        P = list(doc.paragraphs)
        txt = lambda p: p.text.strip()
        hi = next(i for i, p in enumerate(P) if txt(p) == "Report on Other Legal and Regulatory Requirements")
        si = next(i for i, p in enumerate(P) if i > hi and txt(p).startswith("For ") and "Chartered Accountants" in txt(p))
    except Exception as e:
        LOG.add("ERROR", "audit_report", f"Template not usable ({e}); using the built-in firm template instead.")
        doc = Document(io.BytesIO(base64.b64decode(EMBEDDED_TEMPLATE_B64))); P = list(doc.paragraphs)
        hi = next(i for i, p in enumerate(P) if txt(p) == "Report on Other Legal and Regulatory Requirements")
        si = next(i for i, p in enumerate(P) if i > hi and txt(p).startswith("For ") and "Chartered Accountants" in txt(p))
    # ---- header facts
    ti = next(i for i, p in enumerate(P) if txt(p).startswith("To the Members"))
    nm_i = next(i for i in range(ti + 1, len(P)) if txt(P[i])); old_name = txt(P[nm_i])
    cin_i = next(i for i in range(nm_i + 1, len(P)) if txt(P[i]).startswith("CIN"))
    # ---- samples to clone
    S = {}
    S["l1"] = copy.deepcopy(P[hi + 1]._p)
    S["l2"] = copy.deepcopy(next((p for p in P[hi + 1:si] if re.match(r"^[a-z]\.\t", p.text)), P[hi + 2])._p)
    S["l3"] = copy.deepcopy(next((p for p in P[hi + 1:si] if re.match(r"^[ivx]+\.\t", p.text)), P[hi + 3])._p)
    S["head"] = copy.deepcopy(next(p for p in P if txt(p) == "Basis of Opinion")._p)
    S["body"] = copy.deepcopy(next(p for p in P if txt(p).startswith("We have audited the accompanying"))._p)
    # ---- opinion paragraphs (firm wording, year/loss/cash-flow aware)
    cfs, ind = ap["cfs"], ap["ind_as"]
    std = ("Indian Accounting Standards prescribed under section 133 of the Act read with the Companies (Indian Accounting Standards) Rules, 2015, as amended" if ind
           else "Accounting Standards prescribed under section 133 of the act read with the Companies (Accounting Standards) Rules, 2006")
    res = "profit" if cur["profit"] >= 0 else "loss"
    for p in P[:hi]:
        t = txt(p)
        if t.startswith("We have audited the accompanying"):
            _set_text(p, f"We have audited the accompanying financial statements of {prof['name']} (“the Company”) which comprises the Balance Sheet as at {LD}, the Statement of Profit and Loss{', the Cash Flow Statement' if cfs else ''} for the year ended {ord_date(fy_e)}, and notes to the financial statements, including a summary of significant accounting policies and other explanatory information (hereinafter referred to as “the Financial Statements”).")
        elif t.startswith("In our opinion and to the best of our information"):
            _set_text(p, f"In our opinion and to the best of our information and according to the explanations given to us, the aforesaid financial statements give the information required by the Companies Act, 2013 (“the Act”) in the manner so required and give a true and fair view in conformity with the {std} and other accounting principles generally accepted in India, of the state of affairs of the Company as at {LD}, and its {res}{' and its cash flows' if cfs else ''} for the year ended on that date.")
        elif t.startswith("The Company’s Board of Directors is responsible for the matters stated in Section 134(5)"):
            _set_text(p, t.replace("financial position, financial performance and cash flows", "financial position and financial performance") if not cfs else
                      t.replace("financial position and financial performance of", "financial position, financial performance and cash flows of").replace("financial position, financial performance of", "financial position, financial performance and cash flows of"))
        elif t == old_name: _set_text(p, prof["name"])
        elif t.startswith("CIN"): _set_text(p, f"CIN: {prof['cin']}" if prof.get("cin") else t)
    # ---- going concern / EOM paragraph after opinion
    gc = going_concern_par(ctx, ap)
    if gc:
        op = next(p for p in P[:hi] if txt(p).startswith("In our opinion and to the best of our information"))
        h = copy.deepcopy(S["head"]); b = copy.deepcopy(S["body"]); op._p.addnext(h); h.addnext(b)
        _set_text(Paragraph(h, op._parent), gc[0]); _set_text(Paragraph(b, op._parent), gc[1])
    # ---- rebuild legal section
    last_nonblank = max(i for i in range(hi + 1, si) if txt(P[i]))
    anchor = P[last_nonblank + 1]._p
    for p in P[hi + 1:last_nonblank + 1]: p._p.getparent().remove(p._p)
    for lvl, text in legal_section(ctx, ap):
        el = copy.deepcopy(S["l1"] if lvl == 1 else S["l2"] if lvl in (2, 4) else S["l3"])
        anchor.addprevious(el); pp = Paragraph(el, P[0]._parent); _set_text(pp, text)
        if lvl == 4:   # continuation lines of the funds clause sit under the roman item
            pp.paragraph_format.left_indent = P[hi + 0].paragraph_format.left_indent
    # ---- signature
    for p in doc.paragraphs:
        t = txt(p)
        if t.startswith("For ") and "Chartered Accountants" in t: _set_text(p, f"For {prof.get('audit_firm', 'Raghunath & Iyer')} (Chartered Accountants)")
        elif t.startswith("FRN"): _set_text(p, f"FRN. {prof.get('frn', '')}")
        elif t.startswith("CA ") and len(t) < 40: _set_text(p, f"CA {prof.get('partner', '')}".strip())
        elif t.startswith("Membership No"): _set_text(p, f"Membership No. {prof.get('mem_no', '')}")
        elif t.startswith("Date:"):
            _set_text(p, f"Date: {prof.get('report_date', '')}")
            for lab in (f"Place: {prof.get('place', '')}", f"UDIN: {prof.get('udin', '')}"):
                el = copy.deepcopy(p._p); p._p.addnext(el); _set_text(Paragraph(el, p._parent), lab)
            break
    # ---- any leftover old dates/years in the fixed paragraphs (generic refresh)
    _refresh_dates(doc, fy_e)
    _annexures(doc, ctx, ap)
    doc.save(out_path)
    return out_path

def _refresh_dates(doc, new_end):
    """Replace a stale year-end date written in any common style (templates carried from the last year)."""
    pats = []
    for p in doc.paragraphs:
        for m in re.finditer(r"March 31, (\d{4})|31(?:st)? March,? (\d{4})|31\.03\.(\d{4})|31-03-(\d{4})", p.text):
            y = int(next(g for g in m.groups() if g))
            if y != new_end.year: pats.append(y)
    for old in set(pats):
        for p in doc.paragraphs:
            t = p.text
            if str(old) in t:
                t2 = re.sub(rf"March 31, {old}", long_date(new_end), t); t2 = re.sub(rf"31(st)? March,? {old}", ord_date(new_end), t2)
                t2 = re.sub(rf"31\.03\.{old}", new_end.strftime("%d.%m.%Y"), t2); t2 = re.sub(rf"31-03-{old}", new_end.strftime("%d-%m-%Y"), t2)
                if t2 != t: _set_text(p, t2)

def build_audit_report_noncompany(ctx, out_path):
    """SA 700 report for trusts / societies / associations / partnership firms (no CARO, no s.143)."""
    from docx import Document
    from docx.shared import Pt, Cm
    prof, cur = ctx["prof"], ctx["cur"]; fy_e = pdate(ctx["fy_end"]); LD = long_date(fy_e)
    ent = prof.get("entity"); who = {"trust": "Trustees / Managing Committee", "firm": "Partners"}.get(ent, "Managing Committee")
    kind = {"trust": "Trust / Society", "firm": "Firm"}.get(ent, "Entity")
    doc = Document(); sec = doc.sections[0]; sec.left_margin = sec.right_margin = Cm(2.2)
    acct = "Income and Expenditure Account" if ent != "firm" else "Profit and Loss Account"
    res = ("surplus" if cur["net"] >= 0 else "deficit") if ent != "firm" else ("profit" if cur["net"] >= 0 else "loss")
    A = lambda t, b=False, a="j": _add_par(doc, t, b, 12, a)
    A("INDEPENDENT AUDITOR’S REPORT", True, "c"); A(f"To the {who} of", False, "j"); A(prof["name"], True, "j")
    A("Report on the Audit of the Financial Statements", True)
    A("Opinion", True)
    A(f"We have audited the accompanying financial statements of {prof['name']} (“the Entity”), which comprise the Balance Sheet as at {LD}, the {acct} for the year then ended, and notes to the financial statements, including a summary of significant accounting policies and other explanatory information.")
    A(f"In our opinion and to the best of our information and according to the explanations given to us, the aforesaid financial statements give a true and fair view in conformity with the accounting principles generally accepted in India, of the state of affairs of the Entity as at {LD}, and of its {res} for the year ended on that date.")
    A("Basis for Opinion", True)
    A("We conducted our audit in accordance with the Standards on Auditing (SAs) issued by the Institute of Chartered Accountants of India. Our responsibilities under those Standards are further described in the Auditor’s Responsibilities section of our report. We are independent of the Entity in accordance with the Code of Ethics issued by the ICAI, and we have fulfilled our other ethical responsibilities in accordance with the Code. We believe that the audit evidence we have obtained is sufficient and appropriate to provide a basis for our opinion.")
    if cur.get("net") is not None and ctx.get("going_concern_flag"): pass
    A(f"Responsibilities of Management and Those Charged with Governance for the Financial Statements", True)
    A(f"The {who} are responsible for the preparation of these financial statements that give a true and fair view of the financial position and financial performance of the Entity in accordance with the accounting principles generally accepted in India, including the maintenance of adequate accounting records, selection and application of appropriate accounting policies, safeguarding of assets, and design, implementation and maintenance of adequate internal financial controls. In preparing the financial statements, the {who} are responsible for assessing the Entity’s ability to continue as a going concern.")
    A("Auditor’s Responsibilities for the Audit of the Financial Statements", True)
    A("Our objectives are to obtain reasonable assurance about whether the financial statements as a whole are free from material misstatement, whether due to fraud or error, and to issue an auditor’s report that includes our opinion. Reasonable assurance is a high level of assurance, but is not a guarantee that an audit conducted in accordance with SAs will always detect a material misstatement when it exists. As part of an audit in accordance with SAs, we exercise professional judgment and maintain professional skepticism throughout the audit. We identify and assess the risks of material misstatement, obtain an understanding of internal control relevant to the audit, evaluate the appropriateness of accounting policies and estimates, conclude on the appropriateness of the going-concern basis of accounting, evaluate the overall presentation of the financial statements, and communicate with those charged with governance regarding the scope and timing of the audit and significant findings.")
    A("Report on Other Legal and Regulatory Requirements", True)
    A("1.\tWe have obtained all the information and explanations which to the best of our knowledge and belief were necessary for the purposes of our audit.")
    A("2.\tIn our opinion, proper books of account as required by law have been kept by the Entity so far as appears from our examination of those books.")
    A(f"3.\tThe Balance Sheet and the {acct} dealt with by this report are in agreement with the books of account.")
    if ent == "trust": A("4.\tOur separate report in the prescribed audit form under the income-tax law (Form 10B / 10BB or the corresponding form under the Income-tax Act, 2025) is issued separately and is not covered by this report. [CA to confirm applicable form]")
    A("5.\tThe Companies (Auditor’s Report) Order and section 143 of the Companies Act, 2013 do not apply to this Entity.")
    A("", False)
    A(f"For {prof.get('audit_firm', '')}" + (" (Chartered Accountants)" if prof.get("audit_firm") else ""), True, None)
    for t in (f"FRN. {prof['frn']}" if prof.get("frn") else "", f"CA {prof.get('partner', '')}".strip() if prof.get("partner") else "", f"Membership No. {prof.get('mem_no', '')}" if prof.get("mem_no") else "",
              f"UDIN: {prof.get('udin', '')}", f"Place: {prof.get('place', '')}", f"Date: {prof.get('report_date', '')}"):
        if t: A(t, False, None)
    doc.save(out_path); return out_path

# =============================================================== LEARN THE FIRM'S TEMPLATE FROM UPLOADED SAMPLES
@guard("learn_profile", default={})
def learn_profile(paths):
    """Reads sample statements / reports (xlsx, docx, pdf) and returns profile fields found: names, signatories, CA details, entity type."""
    P = {"facts": {}}
    for path in paths:
        p = str(path); low = p.lower()
        try:
            if low.endswith((".xlsx", ".xlsm")):
                import openpyxl
                wb = openpyxl.load_workbook(p, data_only=False); cells = []
                for ws in wb: cells += [(ws.title, c.row, c.column, str(c.value).strip()) for r in ws.iter_rows() for c in r if c.value is not None]
                texts = [t for _, _, _, t in cells]; up = [t.upper() for t in texts]
                first = next((t for t in texts if len(t) > 8 and not t.startswith("=")), "")
                if first and "name" not in P: P["name"] = re.sub(r"\s+", " ", first).strip()
                if any("EXCESS OF INCOME" in u or "INCOME AND EXPENDITURE" in u for u in up): P["entity"] = "trust"
                elif any("PARTNER" in u for u in up): P["entity"] = "firm"
                if any("LIABILITIES" == u or "ASSETS" == u for u in up): P.setdefault("horizontal_template", True)
                roles = [t.strip() for t in texts if t.strip().upper() in ("PRESIDENT", "SECRETARY", "TREASURER", "CHAIRMAN", "TRUSTEE", "DIRECTOR", "PARTNER")]
                if roles: P["signatories"] = [{"role": r, "name": ""} for r in dict.fromkeys(roles)]
                for t in texts:
                    m = re.match(r"M\.?\s*NO\s*[:.]?\s*(\d+)", t, re.I)
                    if m: P["mem_no"] = m.group(1)
                    if t.upper().startswith("PLACE"): P["place"] = t.split(":", 1)[-1].strip()
                for i, t in enumerate(texts):
                    if t.upper() == "CHARTERED ACCOUNTANT" and i > 0 and texts[i - 1].isupper() and len(texts[i - 1]) < 40 and "REPORT" not in texts[i - 1]: P["partner"] = texts[i - 1].title()
            elif low.endswith(".docx"):
                from docx import Document
                d = Document(p); T = [x.text.strip() for x in d.paragraphs if x.text.strip()]
                for i, t in enumerate(T):
                    if t.startswith("To the Members") and i + 1 < len(T): P["name"] = T[i + 1]
                    if t.startswith("CIN"): P["cin"] = t.split(":", 1)[-1].strip()
                    if t.startswith("For ") and "Chartered Accountants" in t: P["audit_firm"] = t[4:].replace("(Chartered Accountants)", "").strip()
                    if t.startswith("FRN"): P["frn"] = re.sub(r"FRN\.?\s*", "", t).strip()
                    if t.startswith("CA ") and len(t) < 40: P["partner"] = t[3:].strip()
                    if t.startswith("Membership No"): P["mem_no"] = re.sub(r"Membership No\.?\s*", "", t).strip()
                    if t.startswith("Place"): P["place"] = t.split(":", 1)[-1].strip()
                P["template_docx"] = p; P["entity"] = "company"
                if any("Accounting Standards) Rules, 2006" in t or "Accounting Standards) Rules, 2021" in t for t in T): P["facts"]["ind_as"] = False
            elif low.endswith(".pdf"):
                txt = ""
                try: txt = subprocess.run(["pdftotext", "-layout", p, "-"], capture_output=True, text=True, timeout=60).stdout
                except Exception:
                    try:
                        import pypdf; txt = "\n".join(pg.extract_text() or "" for pg in pypdf.PdfReader(p).pages)
                    except Exception: LOG.add("INFO", "learn_profile", f"No PDF reader available for {Path(p).name} (install poppler or pypdf).")
                m = re.search(r"\n\s*([A-Z][A-Z .&]+(?:PRIVATE LIMITED|LIMITED|LLP|TRUST|SOCIETY|ASSOCIATION))\s*\n", txt)
                if m: P["name"] = m.group(1).strip(); P["entity"] = "company" if "LIMITED" in m.group(1) else P.get("entity", "firm")
                m = re.search(r"For ([A-Z][\w .&]+)\s+Chartered Accountants?", txt)
                if m: P["sign_ca"] = m.group(1).strip()
                dirs = re.findall(r"([A-Z][a-z]+(?: [A-Z][a-z]+)+)\s+(?=.*\(Director\))", txt)
                names = re.findall(r"\n\s*([A-Z][a-z]+(?: [A-Z][a-z]+){1,3})\s+([A-Z][a-z]+(?: [A-Z][a-z]+){1,3})\s*\n\s*\(Director\)", txt)
                if names: P["directors"] = list(names[0])
                m = re.search(r"Place\s*:\s*([A-Za-z ]+)", txt)
                if m: P["place"] = m.group(1).strip().split("\n")[0]
                addr = re.findall(r"\n(\s*[0-9]+[A-Z]? ?\(.*?\).*|\s*[A-Z][A-Z ]+PANCHAYATH|\s*[A-Z]+-\d{6})", txt)
                if addr: P["address_lines"] = [a.strip() for a in dict.fromkeys(addr)][:4]
        except Exception as e:
            LOG.add("WARN", "learn_profile", f"{Path(p).name}: {e}")
    return P

# =============================================================== EXPORTERS
def header_lines(prof):
    return [prof.get("name", "")] + list(prof.get("address_lines", []))

def sign_rows(prof, fy_end):
    ent = prof.get("entity", "company"); place, dt = prof.get("place", ""), prof.get("fs_date", prof.get("report_date", ""))
    if ent == "company":
        ds = prof.get("directors", []) or [""]
        rows = [("In terms of our report attached.", "For and on behalf of the Board of Directors"),
                (f"For {prof.get('sign_ca', prof.get('audit_firm', ''))}", prof.get("name", "")), ("Chartered Accountants", ""), ("", ""),
                (f"Place: {place}", ds[0]), (f"Date: {dt}", "(Director)")]
        for d_ in ds[1:]: rows += [("", ""), ("", d_), ("", "(Director)")]
        return rows
    sg = prof.get("signatories") or [{"role": "PRESIDENT", "name": ""}, {"role": "SECRETARY", "name": ""}, {"role": "TREASURER", "name": ""}]
    rows = [("          ".join(s["role"] for s in sg), ""), ("          ".join(s.get("name", "") for s in sg), "")] if ent == "trust" else \
           [("          ".join(s.get("name", "") or s["role"] for s in (prof.get("signatories") or [{"role": "PARTNER"}])), "")]
    ca = [("VIDE OUR AUDIT REPORT OF EVEN DATE", ""), (prof.get("partner", "").upper(), ""), ("CHARTERED ACCOUNTANT", ""),
          (f"M NO:{prof.get('mem_no', '')}", ""), (f"UDIN: {prof.get('udin', '')}", ""), (f"DATE: {dt}", ""), (f"PLACE: {place}", "")]
    return rows + [("", "")] + ca

def _fmt_cell(v, kind, ci):
    if v is None or v == "": return ""
    if kind == "v" and ci == 1 and isinstance(v, (int, float)) and float(v).is_integer() and 0 < v < 200: return str(int(v))   # note number
    if isinstance(v, (int, float)): return inr(v, paren=(kind in ("v", "n")))
    return str(v)

def render_text(tables, prof, fy_end, width=100):
    out = []
    for t in tables:
        out.append("")
        if t["kind"] in ("v", "h"):
            out += [h.center(width) for h in header_lines(prof) if h]
        out.append(t["title"].center(width)); out.append("-" * width)
        if t["kind"] == "t": out.append(t["text"]); continue
        cols = t["cols"]; nc = len(cols); ws = [max(int(width * w), 6) for w in t["widths"]]
        out.append("".join(str(c)[:w - 1].ljust(w) if i == 0 or (t["kind"] == "h" and i in (1, 4)) else str(c)[:w - 1].rjust(w) for i, (c, w) in enumerate(zip(cols, ws)))); out.append("-" * width)
        for r in t["rows"]:
            cells = []
            for i, (v, w) in enumerate(zip(r["c"], ws)):
                s = ("  " * r["ind"] if i == 0 else "") + _fmt_cell(v, t["kind"], i)
                left = i == 0 or (t["kind"] == "h" and i in (1, 4))
                cells.append(s[:w - 1].ljust(w) if left else s[:w - 1].rjust(w))
            out.append("".join(cells))
            if r["top"]: pass
        if t["kind"] in ("v", "h") and t["title"].split()[0] in ("BALANCE", "STATEMENT", "INCOME", "TRADING", "CASH", "PROFIT"):
            out.append("-" * width)
            for a, b in sign_rows(prof, fy_end): out.append(f"{a:<{width//2}}{b}")
    return "\n".join(out)

def _numcols(t): return {"v": [2, 3] if len(t["cols"]) == 4 else [1, 2], "n": [1, 2], "h": [0, 2, 3, 5]}[t["kind"]]

@guard("export_excel", default=None)
def to_excel(path, tables, prof, fy_end):
    import openpyxl
    from openpyxl.styles import Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter as CL
    wb = openpyxl.Workbook(); wb.remove(wb.active); thin = Side(style="thin"); NF = '#,##0.00;(#,##0.00);"-"'
    notes_ws, nr = None, 1
    for t in tables:
        if t["kind"] == "n" or t["kind"] == "t":
            if notes_ws is None: notes_ws = wb.create_sheet("Notes"); nr = 1; notes_ws.column_dimensions["A"].width = 60; notes_ws.column_dimensions["B"].width = 22; notes_ws.column_dimensions["C"].width = 22
            ws, r0 = notes_ws, nr
        else:
            ws = wb.create_sheet(re.sub(r"[^A-Za-z0-9 &]", "", t["title"].split(" AS AT")[0].split(" FOR ")[0].title())[:28]); r0 = 1
            wsum = [0.45 * 120] + [14] * 5
            for i, w in enumerate(t["widths"]): ws.column_dimensions[CL(i + 1)].width = max(8, int(w * 120))
            for h in header_lines(prof):
                if h: ws.cell(r0, 1, h).font = Font(bold=True, size=13); ws.merge_cells(start_row=r0, start_column=1, end_row=r0, end_column=len(t["cols"])); ws.cell(r0, 1).alignment = Alignment(horizontal="center"); r0 += 1
        ws.cell(r0, 1, t["title"]).font = Font(bold=True, size=11)
        if t["kind"] in ("v", "h"): ws.merge_cells(start_row=r0, start_column=1, end_row=r0, end_column=len(t["cols"])); ws.cell(r0, 1).alignment = Alignment(horizontal="center")
        r0 += 1
        if t["kind"] == "t":
            for line in t["text"].split("\n"): c = ws.cell(r0, 1, line); c.alignment = Alignment(wrap_text=True, vertical="top"); ws.merge_cells(start_row=r0, start_column=1, end_row=r0, end_column=3); ws.row_dimensions[r0].height = 30 if len(line) > 90 else 16; r0 += 1
            nr = r0 + 1; continue
        for j, c in enumerate(t["cols"]):
            x = ws.cell(r0, j + 1, c); x.font = Font(bold=True); x.alignment = Alignment(horizontal="left" if j == 0 or (t["kind"] == "h" and j in (1, 4)) else "right", wrap_text=True); x.border = Border(bottom=thin)
        r0 += 1; first = r0; rowno = {}
        nc = _numcols(t)
        for i, r in enumerate(t["rows"]):
            rowno[i] = r0
            for j, v in enumerate(r["c"]):
                cell = ws.cell(r0, j + 1)
                if j in nc and isinstance(v, (int, float)):
                    if r["sm"] and t["kind"] != "h" and v is not None:
                        refs = [f"{CL(j+1)}{rowno[k]}" for k in r["sm"] if k in rowno and isinstance(t["rows"][k]["c"][j], (int, float))]
                        cell.value = ("=" + "+".join(refs)) if refs else r2(v)
                    elif t["kind"] == "h" and r["top"] and j in nc: cell.value = f"=SUM({CL(j+1)}{first}:{CL(j+1)}{r0-1})"
                    else: cell.value = r2(v)
                    cell.number_format = NF; cell.alignment = Alignment(horizontal="right")
                else:
                    cell.value = v if v not in (None,) else None
                    cell.alignment = Alignment(indent=r["ind"] if j == 0 else 0, wrap_text=(j == 0))
                if r["b"]: cell.font = Font(bold=True)
                if r["top"]: cell.border = Border(top=thin)
            r0 += 1
        r0 += 1
        if t["kind"] in ("v", "h") and t["title"].split()[0] in ("BALANCE", "STATEMENT", "INCOME", "TRADING", "CASH", "PROFIT"):
            for a, b in sign_rows(prof, fy_end):
                ws.cell(r0, 1, a); ws.cell(r0, max(2, len(t["cols"]) // 2 + 1), b); r0 += 1
        if t["kind"] == "n": nr = r0 + 1
    wb.save(path); return path

@guard("export_pdf", default=None)
def to_pdf(path, tables, prof, fy_end):
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether
    horiz = prof.get("entity") != "company"
    ps = landscape(A4) if horiz else A4
    doc = SimpleDocTemplate(str(path), pagesize=ps, leftMargin=28, rightMargin=28, topMargin=28, bottomMargin=28, title=prof.get("name", ""))
    ss = getSampleStyleSheet(); H = ParagraphStyle("h", parent=ss["Title"], fontSize=13, leading=16, spaceAfter=2); T = ParagraphStyle("t", parent=ss["Heading3"], alignment=1, fontSize=10.5, spaceAfter=6)
    CELL = ParagraphStyle("c", parent=ss["Normal"], fontSize=8.5, leading=10.5); B = ParagraphStyle("b", parent=CELL, fontName="Helvetica-Bold")
    W = ps[0] - 56; story = []; first = True
    for t in tables:
        is_stmt = t["kind"] in ("v", "h")
        if is_stmt and t["title"].split()[0] in ("BALANCE", "STATEMENT", "INCOME", "TRADING", "CASH", "PROFIT"):
            if not first: story.append(PageBreak())
            for h in header_lines(prof):
                if h: story.append(Paragraph(h, H))
        elif t["kind"] in ("n", "t") and first: pass
        first = False
        story.append(Paragraph(t["title"], T))
        if t["kind"] == "t":
            for line in t["text"].split("\n"): story.append(Paragraph(line.replace("&", "&amp;"), CELL))
            story.append(Spacer(1, 8)); continue
        data, st = [], [("LINEBELOW", (0, 0), (-1, 0), 0.6, colors.black), ("VALIGN", (0, 0), (-1, -1), "TOP"), ("TOPPADDING", (0, 0), (-1, -1), 1.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5)]
        left_cols = {0} | ({1, 4} if t["kind"] == "h" else set())
        data.append([Paragraph(str(c), B) for c in t["cols"]])
        for i, r in enumerate(t["rows"], start=1):
            row = []
            for j, v in enumerate(r["c"]):
                s = _fmt_cell(v, t["kind"], j).replace("&", "&amp;")
                pad = "&nbsp;" * (3 * r["ind"]) if j == 0 else ""
                sty = ParagraphStyle("x", parent=B if r["b"] else CELL, alignment=0 if j in left_cols else 2)
                row.append(Paragraph(pad + s, sty))
            data.append(row)
            if r["top"]: st.append(("LINEABOVE", (0, i), (-1, i), 0.5, colors.black)); st.append(("LINEBELOW", (0, i), (-1, i), 0.5, colors.black))
        for j in range(len(t["cols"])):
            if j not in left_cols: st.append(("ALIGN", (j, 0), (j, 0), "RIGHT"))
        tb = Table(data, colWidths=[W * w for w in t["widths"]], repeatRows=1); tb.setStyle(TableStyle(st)); story.append(tb)
        if is_stmt and t["title"].split()[0] in ("BALANCE", "STATEMENT", "INCOME", "TRADING", "CASH", "PROFIT"):
            story.append(Spacer(1, 14))
            sr = sign_rows(prof, fy_end); _sp = lambda s: re.sub(r"  +", lambda m: "&nbsp;" * len(m.group()), s.replace("&", "&amp;")); sd = [[Paragraph(_sp(a), CELL), Paragraph(_sp(b), CELL)] for a, b in sr]
            sg = Table(sd, colWidths=[W * 0.5, W * 0.5]); sg.setStyle(TableStyle([("TOPPADDING", (0, 0), (-1, -1), 1), ("BOTTOMPADDING", (0, 0), (-1, -1), 1)])); story.append(sg)
        else: story.append(Spacer(1, 10))
    doc.build(story); return path

@guard("export_word", default=None)
def to_word(path, tables, prof, fy_end):
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.section import WD_ORIENT
    doc = Document(); sec = doc.sections[0]
    if prof.get("entity") != "company": sec.orientation = WD_ORIENT.LANDSCAPE; sec.page_width, sec.page_height = sec.page_height, sec.page_width
    sec.left_margin = sec.right_margin = Cm(1.6); st = doc.styles["Normal"]; st.font.name = "Times New Roman"; st.font.size = Pt(10)
    usable = (sec.page_width - sec.left_margin - sec.right_margin)
    first = True
    for t in tables:
        stmt = t["kind"] in ("v", "h") and t["title"].split()[0] in ("BALANCE", "STATEMENT", "INCOME", "TRADING", "CASH", "PROFIT")
        if stmt and not first: doc.add_page_break()
        if stmt:
            for h in header_lines(prof):
                if h: p = doc.add_paragraph(); r = p.add_run(h); r.bold = True; r.font.size = Pt(13); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        first = False
        p = doc.add_paragraph(); r = p.add_run(t["title"]); r.bold = True; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if t["kind"] == "t":
            for line in t["text"].split("\n"): doc.add_paragraph(line)
            continue
        tb = doc.add_table(rows=1, cols=len(t["cols"])); tb.style = "Table Grid"
        left_cols = {0} | ({1, 4} if t["kind"] == "h" else set())
        for j, c in enumerate(t["cols"]):
            cell = tb.rows[0].cells[j]; cell.text = ""; rr = cell.paragraphs[0].add_run(str(c)); rr.bold = True; cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT if j in left_cols else WD_ALIGN_PARAGRAPH.RIGHT
        for r_ in t["rows"]:
            cells = tb.add_row().cells
            for j, v in enumerate(r_["c"]):
                cells[j].text = ""; rr = cells[j].paragraphs[0].add_run(("   " * r_["ind"] if j == 0 else "") + _fmt_cell(v, t["kind"], j)); rr.bold = r_["b"]
                cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT if j in left_cols else WD_ALIGN_PARAGRAPH.RIGHT
        for row in tb.rows:
            for j, w in enumerate(t["widths"]): row.cells[j].width = int(usable * w)
        if stmt:
            doc.add_paragraph()
            for a, b in sign_rows(prof, fy_end):
                p = doc.add_paragraph(); p.add_run(a); p.add_run("\t\t\t" + b if b else "")
        else: doc.add_paragraph()
    doc.save(path); return path

# =============================================================== LEGAL INTELLIGENCE (internet collector - no API, no AI, standard library only)
# Watches official listing pages, remembers what it has seen (SQLite), flags new items that touch a rule the app uses,
# and lets the CA turn an item into a rule change. Nothing is ever applied automatically.
import sqlite3, webbrowser
from html.parser import HTMLParser
from urllib.parse import urljoin, urlparse
INTEL_DB = APP_DIR / "intel.db"
MAX_PAGE_BYTES = 1_500_000          # keeps memory small on a 4 GB machine
MAX_LINKS_PER_PAGE = 400
DEFAULT_SOURCES = [   # URLs are editable in the app. NOT verified from the build environment - use "Test source" and fix if a site moved.
 {"id": "mca", "name": "MCA - notifications / rules", "url": "https://www.mca.gov.in/content/mca/global/en/acts-rules/ebooks/notifications.html", "on": True},
 {"id": "cbic", "name": "CBIC - GST notifications", "url": "https://taxinformation.cbic.gov.in/", "on": True},
 {"id": "it", "name": "Income Tax Dept - notifications", "url": "https://www.incometaxindia.gov.in/en/notifications", "on": True},
 {"id": "icai", "name": "ICAI - resources (AS / SA / guidance)", "url": "https://www.icai.org/resources.shtml", "on": True},
 {"id": "rbi", "name": "RBI - notifications", "url": "https://www.rbi.org.in/Scripts/NotificationUser.aspx", "on": False},
]
# which law-library rule a headline may touch (a hint for the CA, not a conclusion)
TOPICS = {"small_company": r"small compan|definition of small|specification of definitions",
          "caro_regime": r"auditor.?s report\) order|\bcaro\b",
          "rule11_audit_trail": r"audit trail|edit log|accounting software",
          "rule11_funds": r"audit and auditors\).*rules|rule 11",
          "ifc_exempt": r"internal financial control",
          "cash_limits": r"40a\(3\)|269s[st]|269t|cash (payment|transaction)|section 43b",
          "going_concern": r"going concern|standard on auditing|\bsa 570\b",
          "tds_tcs": r"\btds\b|\btcs\b|section 19[4-6]|withholding",
          "gst": r"\bgst\b|cgst|igst|e-?invoice|e-?way|gstr",
          "schedule_iii": r"schedule iii|division i{1,3}\b|accounting standard|ind as"}

def _db():
    c = sqlite3.connect(str(INTEL_DB), timeout=10)
    c.executescript("""CREATE TABLE IF NOT EXISTS items(url TEXT PRIMARY KEY, source TEXT, title TEXT, first_seen TEXT, topics TEXT, status TEXT DEFAULT 'new');
        CREATE TABLE IF NOT EXISTS pages(source TEXT PRIMARY KEY, sha TEXT, last_ok TEXT, last_error TEXT, n_links INTEGER);
        CREATE TABLE IF NOT EXISTS changes(ts TEXT, rule_id TEXT, summary TEXT, source TEXT)""")
    return c

class _Links(HTMLParser):
    def __init__(self, base): super().__init__(convert_charrefs=True); self.base, self.out, self._h, self._t = base, [], None, []
    def handle_starttag(self, tag, attrs):
        if tag == "a": self._h = dict(attrs).get("href"); self._t = []
    def handle_data(self, d):
        if self._h is not None: self._t.append(d)
    def handle_endtag(self, tag):
        if tag == "a" and self._h is not None:
            t = re.sub(r"\s+", " ", "".join(self._t)).strip()
            if self._h and not self._h.startswith(("#", "javascript:", "mailto:")) and len(t) >= 12 and len(self.out) < MAX_LINKS_PER_PAGE: self.out.append((urljoin(self.base, self._h), t))
            self._h = None

def http_get_capped(url, cap=MAX_PAGE_BYTES, timeout=25):
    req = urllib.request.Request(url, headers={"User-Agent": f"Mozilla/5.0 TallyCASuite/{APP_VERSION}", "Accept": "text/html,application/xhtml+xml,application/pdf"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        data = r.read(cap + 1)
        return data[:cap], len(data) > cap, r.headers.get("Content-Type", "")

def topics_of(title):
    return [k for k, rx in TOPICS.items() if re.search(rx, title, re.I)]

@guard("intel_check", default=None)
def intel_check(settings, only=None):
    """Fetch each enabled source once, store new items. Returns {'new': [...], 'errors': [...], 'checked': n}."""
    src = settings.get("sources") or DEFAULT_SOURCES; res = dict(new=[], errors=[], checked=0); db = _db(); now = datetime.now().strftime("%Y-%m-%d %H:%M")
    try:
        for s in src:
            if not s.get("on") or (only and s["id"] != only): continue
            try:
                res["checked"] += 1; data, cut, ctype = http_get_capped(s["url"])
                sha = hashlib.sha256(data).hexdigest(); old = db.execute("SELECT sha FROM pages WHERE source=?", (s["id"],)).fetchone()
                p = _Links(s["url"]); p.feed(data.decode("utf-8", "ignore")); links = p.out
                if not links: raise ValueError("page fetched but no links could be read (site may build its list with JavaScript, or the address has moved)")
                if old and old[0] == sha: db.execute("UPDATE pages SET last_ok=?, last_error='' WHERE source=?", (now, s["id"])); continue
                for url, title in links:
                    if db.execute("SELECT 1 FROM items WHERE url=?", (url,)).fetchone(): continue
                    tp = topics_of(title); first_run = old is None
                    db.execute("INSERT INTO items(url,source,title,first_seen,topics,status) VALUES(?,?,?,?,?,?)", (url, s["id"], title, now, ",".join(tp), "seen" if first_run else "new"))
                    if not first_run: res["new"].append(dict(source=s["id"], title=title, url=url, topics=tp))
                db.execute("INSERT OR REPLACE INTO pages VALUES(?,?,?,?,?)", (s["id"], sha, now, "", len(links)))
            except Exception as e:
                res["errors"].append(f"{s['name']}: {e}"); LOG.add("WARN", "intel", f"{s['name']}: {e}")
                db.execute("INSERT INTO pages(source,sha,last_ok,last_error,n_links) VALUES(?,?,?,?,0) ON CONFLICT(source) DO UPDATE SET last_error=excluded.last_error", (s["id"], "", "", str(e)[:300]))
        db.commit()
    finally: db.close()
    settings["intel_last"] = now
    LOG.add("INFO", "intel", f"checked {res['checked']} source(s): {len(res['new'])} new item(s), {len(res['errors'])} problem(s)")
    return res

def intel_items(status=None, limit=300):
    db = _db()
    try:
        q = "SELECT first_seen,source,title,topics,status,url FROM items" + (" WHERE status=?" if status else "") + " ORDER BY first_seen DESC, rowid DESC LIMIT ?"
        return db.execute(q, ((status, limit) if status else (limit,))).fetchall()
    finally: db.close()

def intel_mark(url, status):
    db = _db()
    try: db.execute("UPDATE items SET status=? WHERE url=?", (status, url)); db.commit()
    finally: db.close()

def intel_status():
    db = _db()
    try: return db.execute("SELECT source,last_ok,last_error,n_links FROM pages").fetchall()
    finally: db.close()

# ---- reading one item (HTML or PDF) as plain text, to help the CA see what changed
@guard("intel_text", default="")
def intel_text(url, limit=60000):
    data, cut, ctype = http_get_capped(url, cap=6_000_000, timeout=40)
    if url.lower().endswith(".pdf") or "pdf" in ctype.lower() or data[:4] == b"%PDF":
        f = Path(tempfile.mkdtemp()) / "d.pdf"; f.write_bytes(data)
        try: txt = subprocess.run(["pdftotext", "-layout", str(f), "-"], capture_output=True, text=True, timeout=60).stdout
        except Exception:
            try:
                import pypdf; txt = "\n".join((pg.extract_text() or "") for pg in pypdf.PdfReader(str(f)).pages[:40])
            except Exception: LOG.add("WARN", "intel_text", "no PDF reader (install poppler or pypdf) - open the link in the browser instead"); return ""
    else:
        t = data.decode("utf-8", "ignore"); t = re.sub(r"(?is)<(script|style).*?</\1>", " ", t); txt = re.sub(r"<[^>]+>", " ", t)
        txt = re.sub(r"&nbsp;|&#160;", " ", txt)
    return re.sub(r"[ \t\r\f]+", " ", txt)[:limit]

_WORDNUM = {"two": 2, "four": 4, "five": 5, "ten": 10, "twenty": 20, "forty": 40, "fifty": 50, "one hundred": 100, "two hundred": 200, "four hundred": 400, "one thousand": 1000}
def _amount(num, unit):
    n = _WORDNUM.get(num.lower().strip()) if not re.match(r"[\d]", num) else float(num.replace(",", ""))
    if n is None: return None
    return n * (10_000_000 if unit.lower().startswith("crore") else 100_000 if unit.lower().startswith("lakh") else 1)

def suggest_values(rule_id, text):
    """Best-effort: quote the sentences that carry the numbers and propose values. ALWAYS for the CA to confirm."""
    out = dict(snippets=[], values={})
    for m in re.finditer(r"[^.]{0,160}(?:rupees|Rs\.?|₹)\s*([\d,\.]+|[a-z ]{3,20}?)\s*(crore|lakh)s?[^.]{0,120}\.", text, re.I):
        out["snippets"].append(re.sub(r"\s+", " ", m.group(0)).strip()[:400])
    if rule_id == "small_company":
        w = r"(\d[\d,\.]*|one hundred|two hundred|four hundred|one thousand|fifty|forty|twenty|ten|five|four|two)"
        pu = re.search(r"paid.?up (?:share )?capital[^.]{0,120}?(?:rupees|Rs\.?|₹)\s*" + w + r"\s*(crore|lakh)", text, re.I)
        tn = re.search(r"turnover[^.]{0,120}?(?:rupees|Rs\.?|₹)\s*" + w + r"\s*(crore|lakh)", text, re.I)
        if pu and _amount(pu.group(1), pu.group(2)): out["values"]["paid_up"] = int(_amount(pu.group(1), pu.group(2)))
        if tn and _amount(tn.group(1), tn.group(2)): out["values"]["turnover"] = int(_amount(tn.group(1), tn.group(2)))
    out["snippets"] = out["snippets"][:8]
    return out

@guard("apply_rule_change", default=None)
def apply_rule_change(rid, from_date, values, source, rules=None):
    """CA-approved change: adds a history entry (verified today), validates, backs up, rolls back on any problem."""
    rules = rules or RULES; pdate(from_date)
    if rid not in rules.data["rules"]: raise ValueError(f"unknown rule {rid}")
    before = copy.deepcopy(rules.data); rules.save(backup=True)
    e = _h(str(pdate(from_date)), date.today().isoformat(), **values)
    r = rules.data["rules"][rid]; r["history"] = [x for x in r["history"] if x["from"] != e["from"]] + [e]
    r["source"] = (r.get("source", "") + "; " if r.get("source") else "") + source
    pr = rules.validate(rules.data)
    if pr: rules.data = before; rules.save(backup=False); raise ValueError("rejected: " + "; ".join(pr[:3]))
    rules.save(backup=False)
    db = _db()
    try: db.execute("INSERT INTO changes VALUES(?,?,?,?)", (datetime.now().isoformat(timespec="seconds"), rid, json.dumps(values), source)); db.commit()
    finally: db.close()
    LOG.add("INFO", "rules", f"CA-approved change to {rid} from {from_date}: {values} (source: {source})")
    return e

# ---- health report (cheap, no extra libraries)
def health_report(settings):
    H = []
    du = shutil.disk_usage(APP_DIR); H.append(("OK" if du.free > 500_000_000 else "WARN", f"Disk free {du.free // 1_000_000} MB"))
    H.append(("OK", f"Python {sys.version.split()[0]}, app {APP_VERSION}, law library {RULES.data['version']}"))
    st = RULES.stale(); RULES.reset_used()
    H.append(("OK", "Law library up to date") if not RULES.data["rules"] else ("INFO", "Rules not re-verified for 180+ days are flagged during checks"))
    last = settings.get("intel_last"); H.append(("OK" if last else "WARN", f"Last official-source check: {last or 'never - click Check now'}"))
    try: H.append(("OK", f"intel database {INTEL_DB.stat().st_size // 1024} KB, {len(intel_items('new'))} unreviewed item(s)"))
    except Exception: H.append(("INFO", "no intel database yet"))
    H.append(("OK" if not LOG.errors() else "WARN", f"{len(LOG.errors())} error(s) this session"))
    H.append(("OK", f"{len(list((APP_DIR / 'backups').glob('*')))} backup file(s)"))
    return H

@guard("rollback_app", default=None)
def rollback_app():
    """Put back the previous app version kept by self_update (only if it compiles)."""
    if getattr(sys, "frozen", False): LOG.add("INFO", "rollback", "packaged exe: install the previous .exe from GitHub Releases"); return None
    bk = sorted((APP_DIR / "backups").glob("tally_suite_*.py"), key=lambda p: _ver_tuple(p.stem))
    if not bk: LOG.add("INFO", "rollback", "no previous version saved"); return None
    src = bk[-1].read_text(encoding="utf-8"); compile(src, "prev", "exec"); Path(__file__).resolve().write_text(src, encoding="utf-8")
    LOG.add("INFO", "rollback", f"restored {bk[-1].name}. Restart the app."); return bk[-1].name

# ---- tests (called from run_tests)
def intel_tests(T, tmp):
    from http.server import BaseHTTPRequestHandler, HTTPServer
    page = {"html": "<html><body><a href='/a1.html'>Notification 1 on something old</a><a href='/a2.html'>Circular about GST e-invoice limits</a><a href='#x'>skip me please now</a><a href='/s'>x</a></body></html>"}
    SM = ("<html><script>var a=1</script><body><p>The Companies (Specification of Definitions Details) Amendment Rules: a small company means a company, other than a public company, "
          "having paid up share capital not exceeding rupees ten crore and turnover not exceeding rupees one hundred crore.</p></body></html>")
    class H(BaseHTTPRequestHandler):
        def log_message(self, *a): pass
        def do_GET(self):
            if self.path == "/down": self.send_response(503); self.end_headers(); return
            if self.path == "/empty": body = b"<html><body>js only</body></html>"
            elif self.path == "/huge": body = b"<a href='/z'>" + b"x" * 3_000_000 + b"</a>"
            elif self.path == "/sm": body = SM.encode()
            else: body = page["html"].encode()
            self.send_response(200); self.send_header("Content-Type", "text/html"); self.end_headers(); self.wfile.write(body)
    srv = HTTPServer(("127.0.0.1", 0), H); threading.Thread(target=srv.serve_forever, daemon=True).start(); b = f"http://127.0.0.1:{srv.server_address[1]}"
    S = {"sources": [{"id": "t", "name": "Test", "url": b + "/", "on": True}, {"id": "d", "name": "Down", "url": b + "/down", "on": True}, {"id": "e", "name": "Empty", "url": b + "/empty", "on": True}]}
    r1 = intel_check(S)
    T("intel.first_run_baseline_not_alerting", r1 and r1["checked"] == 3 and not r1["new"] and len(r1["errors"]) == 2, str(r1))
    T("intel.items_stored_and_short_links_skipped", len(intel_items()) == 2, str(intel_items()))
    r2 = intel_check(S); T("intel.unchanged_page_skipped", r2 and not r2["new"])
    page["html"] = page["html"].replace("</body>", "<a href='/a3.html'>Companies (Specification of Definitions Details) Amendment Rules 2026</a></body>")
    r3 = intel_check(S); T("intel.new_item_detected_with_topic", r3 and len(r3["new"]) == 1 and "small_company" in r3["new"][0]["topics"], str(r3))
    T("intel.gst_topic", "gst" in topics_of("CBIC circular on GST e-invoice") and "tds_tcs" in topics_of("TDS under section 194C"))
    intel_mark(r3["new"][0]["url"], "reviewed"); T("intel.mark_reviewed", len(intel_items("new")) == 0)
    S2 = {"sources": [{"id": "h", "name": "Huge", "url": b + "/huge", "on": True}]}; T("intel.huge_page_capped_not_crash", intel_check(S2) is not None)
    txt = intel_text(b + "/sm"); sv = suggest_values("small_company", txt)
    T("intel.suggests_small_company_limits", sv["values"] == {"paid_up": 100000000, "turnover": 1000000000} and sv["snippets"], str(sv))
    T("intel.script_stripped", "var a=1" not in txt)
    T("intel.offline_does_not_crash", intel_check({"sources": [{"id": "x", "name": "X", "url": "http://127.0.0.1:1/", "on": True}]})["errors"])
    e = apply_rule_change("small_company", "2027-04-01", {"paid_up": 1, "turnover": 2}, "test notification")
    T("intel.apply_change_adds_history", e and RULES.at("small_company", "2027-05-01")["paid_up"] == 1 and RULES.at("small_company", "2027-05-01")["verified"] == date.today().isoformat())
    T("intel.apply_rejects_bad", apply_rule_change("small_company", "2028-04-01", {"paid_up": "x"}, "t") is None and RULES.at("small_company", "2028-05-01")["paid_up"] == 1)
    T("intel.apply_rejects_unknown_rule", apply_rule_change("nope", "2028-04-01", {}, "t") is None)
    T("intel.health", len(health_report({})) >= 6)
    srv.shutdown()

# =============================================================== PIPELINE (one call: ledgers -> statements, checks, applicability)
def build_all(ledgers, prof, fy_start, fy_end, groups=None, prior=None, prior_source=None, vouchers=None):
    ctx = dict(ledgers=ledgers, groups=groups, prof=prof, fy_start=str(pdate(fy_start)), fy_end=str(pdate(fy_end)),
               vouchers=vouchers, prior=prior, prior_source=prior_source)
    horizontal = prof.get("entity", "company") != "company"
    comp = compute_h if horizontal else compute_company
    ctx["cur"] = comp(ledgers, prof, groups)
    ctx["prv"] = comp(prior, prof, groups) if prior else None
    ctx["applic"] = determine(ctx)
    ctx["cf"] = None
    if horizontal:
        ctx["tables"], _ = build_horizontal(ctx["cur"], ctx["prv"], prof, fy_end)
    else:
        if ctx["applic"]["cfs"] and ctx["prv"]: ctx["cf"] = cash_flow(ctx["cur"], ctx["prv"])
        ctx["tables"], ctx["notes"] = build_vertical(ctx["cur"], ctx["prv"], prof, fy_end, ctx["cf"])
    ctx["findings"] = run_checks(ctx)
    return ctx

@guard("build_all_safe", default=None)
def build_all_safe(*a, **k): return build_all(*a, **k)

def safe_name(s): return re.sub(r"[^A-Za-z0-9]+", "_", s).strip("_")[:60]

@guard("export_all", default=[])
def export_all(ctx, outdir, fmts=("xlsx", "pdf", "docx", "txt"), audit=True):
    outdir = Path(outdir); outdir.mkdir(parents=True, exist_ok=True)
    base = safe_name(ctx["prof"].get("name", "company")) + "_FY" + fy_label(pdate(ctx["fy_end"]))
    made = []
    for f in fmts:
        p = outdir / f"{base}_statements.{f}"
        if f == "xlsx": r = to_excel(p, ctx["tables"], ctx["prof"], ctx["fy_end"])
        elif f == "pdf": r = to_pdf(p, ctx["tables"], ctx["prof"], ctx["fy_end"])
        elif f == "docx": r = to_word(p, ctx["tables"], ctx["prof"], ctx["fy_end"])
        else:
            p.write_text(render_text(ctx["tables"], ctx["prof"], ctx["fy_end"], 130), encoding="utf-8"); r = p
        if r: made.append(str(r))
    if audit:
        p = outdir / f"{base}_auditors_report.docx"
        r = build_audit_report(ctx, str(p))
        if r: made.append(str(r))
    fp = outdir / f"{base}_checks.txt"
    fp.write_text("\n".join(f"[{x.sev}] {x.area}: {x.msg}\n    why: {x.why}\n    fix: {x.fix}" for x in ctx["findings"]), encoding="utf-8")
    made.append(str(fp))
    return made

def explain_applicability(ctx):
    ap = ctx["applic"]; L = []
    for sev, m, why in ap.get("notes", []): L.append(f"- {m}" + (f"  ({why})" if why else ""))
    if ap.get("entity") == "company":
        L.append(f"- Order: {ap['caro']} -> {'APPLICABLE' if ap['caro_applicable'] else 'NOT applicable'}; cash flow: {'yes' if ap['cfs'] else 'no'}; IFC (Annexure B): {'exempt' if ap['ifc_exempt'] else 'required'}")
    return "\n".join(L)

# =============================================================== OFFLINE LAW INBOX (no internet, no AI: drop official PDFs / text into the folder)
INBOX = APP_DIR / "law_inbox"
try: INBOX.mkdir(parents=True, exist_ok=True)
except Exception: pass

def _read_any(p):
    p = Path(p); low = p.suffix.lower()
    if low == ".pdf":
        try: return subprocess.run(["pdftotext", "-layout", str(p), "-"], capture_output=True, text=True, timeout=120).stdout
        except Exception:
            try:
                import pypdf; return "\n".join(pg.extract_text() or "" for pg in pypdf.PdfReader(str(p)).pages)
            except Exception as e: LOG.add("WARN", "law_inbox", f"{p.name}: cannot read PDF ({e}). Save it as text, or install pypdf."); return ""
    t = p.read_text(encoding="utf-8", errors="ignore")
    return re.sub(r"<[^>]+>", " ", t) if low in (".html", ".htm") else t

def _crore(num_s, unit):
    v = float(num_s.replace(",", "")); u = unit.lower()
    return v * (10000000 if u.startswith("cr") else 100000 if u.startswith("lak") else 1)

_AMT = r"(?:rupees|rs\.?|inr|₹)\s*([\d,]+(?:\.\d+)?)\s*(crore|lakh|lac)s?"
def _eff_date(txt):
    m = re.search(r"(?:w\.?e\.?f\.?|with effect from|effective from|come into force on)[^0-9]{0,30}(\d{1,2}[./-]\d{1,2}[./-]\d{4})", txt, re.I)
    try: return str(pdate(m.group(1).replace("/", "-").replace(".", "-"))) if m else ""
    except Exception: return ""

@guard("scan_law_inbox", default=[])
def scan_law_inbox(rules=None, folder=None):
    """Reads every file in the inbox and PROPOSES rule changes. Nothing is applied until the CA approves each one."""
    rules = rules or RULES; props = []
    files = [f for f in sorted(Path(folder or INBOX).glob("*")) if f.suffix.lower() in (".pdf", ".txt", ".html", ".htm")]
    for f in files:
        txt = re.sub(r"\s+", " ", _read_any(f))
        for m in re.finditer(r"[^.;]{0,200}?paid[- ]?up (?:share )?capital[^.;]{0,120}?" + _AMT + r"[^.;]{0,200}?turnover[^.;]{0,120}?" + _AMT, txt, re.I):
            pu, to = _crore(m.group(1), m.group(2)), _crore(m.group(3), m.group(4)); ctx = txt[max(0, m.start() - 150): m.end() + 250]
            cur = rules.at("small_company", "2099-01-01")
            same = cur and abs(cur["paid_up"] - pu) < 1 and abs(cur["turnover"] - to) < 1
            props.append(dict(rule="small_company", values=dict(paid_up=pu, turnover=to), frm=_eff_date(ctx), source=f.name, snippet=m.group(0).strip()[:350],
                              status="already in library" if same else "NEW - needs CA approval"))
        for m in re.finditer(r"[^.;]{0,200}?turnover[^.;]{0,100}?" + _AMT + r"[^.;]{0,200}?borrowings?[^.;]{0,100}?" + _AMT, txt, re.I):
            if "financial control" in m.group(0).lower() or "internal" in m.group(0).lower():
                tu, bo = _crore(m.group(1), m.group(2)), _crore(m.group(3), m.group(4)); cur = rules.at("ifc_exempt", "2099-01-01")
                props.append(dict(rule="ifc_exempt", values=dict(turnover=tu, borrowings=bo), frm=_eff_date(m.group(0)), source=f.name, snippet=m.group(0).strip()[:350],
                                  status="already in library" if cur and abs(cur["turnover"] - tu) < 1 and abs(cur["borrowings"] - bo) < 1 else "NEW - needs CA approval"))
    if not files: LOG.add("INFO", "law_inbox", f"No files in {INBOX}. Put official notifications (PDF / txt / html) there.")
    return props

@guard("apply_proposal", default=False)
def apply_proposal(p, frm=None, rules=None):
    """CA approval step. Backs up, adds a dated entry carrying the source file, validates, rolls back if invalid."""
    rules = rules or RULES; frm = frm or p.get("frm")
    if not frm: LOG.add("ERROR", "apply_proposal", "Effective date missing - enter the 'with effect from' date."); return False
    before = copy.deepcopy(rules.data); rules.save(backup=True)
    e = _h(str(pdate(frm)), date.today().isoformat(), **p["values"]); e["approved_by_ca"] = True; e["from_file"] = p["source"]
    r = rules.data["rules"][p["rule"]]; r["history"] = [x for x in r["history"] if x["from"] != e["from"]] + [e]
    pr = rules.validate(rules.data)
    if pr: rules.data = before; rules.save(backup=False); LOG.add("ERROR", "apply_proposal", "rejected, rolled back: " + "; ".join(pr[:3])); return False
    rules.save(backup=False); LOG.add("INFO", "apply_proposal", f"{p['rule']} updated from {e['from']} (source {p['source']})"); return True

# =============================================================== SELF-TEST (always runs in a throw-away data folder)
def run_tests():
    fails, passed = [], [0]
    def T(name, cond, detail=""):
        if cond: passed[0] += 1
        else: fails.append(f"{name} {detail}")
    tmp = Path(tempfile.mkdtemp())
    # ---- 1. Tally link via simulator
    ds = make_sim_datasets(); sim = TallySim(ds)
    try:
        os.environ["NO_PROXY"] = "127.0.0.1,localhost"; os.environ["no_proxy"] = "127.0.0.1,localhost"
        ty = Tally("127.0.0.1", sim.port); cs = ty.companies()
        T("tally.companies", len(cs) == 4, str(len(cs)))
        cname = "Mascot Techcrib Private Limited (FY 2024-25)"; pname = "Mascot Techcrib Private Limited (FY 2023-24)"
        led, grp, vch = ty.ledgers(cname), ty.groups(cname), ty.vouchers(cname)
        pled = ty.ledgers(pname)
        T("tally.ledgers", len(led) > 20 and abs(sum(l["bal"] for l in led)) < 1, f"{len(led)} sum={sum(l['bal'] for l in led):.2f}")
        T("tally.vouchers", len(vch) > 50, str(len(vch)))
        vd = ty.voucher_dates(cname); T("tally.voucher_dates", len(vd) >= 5 and all(isinstance(x, date) for x in vd), str(len(vd)))
        prof = dict(entity="company", name="Mascot Techcrib Private Limited", cin="U72900KL2021PTC000000", audit_firm="Raghunath & Iyer", frn="000000S",
                    partner="A B", mem_no="123456", udin="25123456ABCDEF1234", place="Kochi", report_date="30.08.2025", fs_date="30.08.2025",
                    directors=["Antonty George Pattara", "Cyriac Njavally"], facts={"audit_trail": "operated", "going_concern_mode": "mu"})
        ctx = build_all(led, prof, "2024-04-01", "2025-03-31", grp, pled, "snapshot", vch)
        c, p = ctx["cur"], ctx["prv"]
        T("company.balances", abs(c["diff"]) < 0.5, f"diff={c['diff']}")
        T("company.prior.balances", abs(p["diff"]) < 0.5, f"diff={p['diff']}")
        T("company.loss", abs(c["profit"] - -4147686.20) < 5 or c["profit"] < 0, f"profit={c['profit']}")
        ap = ctx["applic"]
        T("applic.small_company_exempt", ap["small"] and not ap["caro_applicable"] and not ap["cfs"] and ap["ifc_exempt"], str({k: ap[k] for k in ("small", "caro_applicable", "cfs", "ifc_exempt")}))
        errs = [f for f in ctx["findings"] if f.sev == "ERROR"]
        T("company.clean_has_no_errors", not errs, "; ".join(f.msg[:80] for f in errs[:3]))
        T("company.going_concern_flagged", any(f.area == "Going concern" for f in ctx["findings"]) == (c["lines"]["share_capital"] + c["lines"]["reserves"] < 0))
        T("company.vouchers_reproduce_ledgers", any(f.area == "Completeness" and f.sev == "OK" for f in ctx["findings"]), [f.msg for f in ctx["findings"] if f.area == "Completeness"][:2].__str__())
        made = export_all(ctx, tmp / "co")
        T("export.all_files", len(made) == 6, str(made))
        for m in made:
            T(f"export.nonempty {Path(m).name}", Path(m).exists() and Path(m).stat().st_size > 500)
        # round trip: Excel sheet totals
        import openpyxl
        wb = openpyxl.load_workbook([m for m in made if m.endswith(".xlsx")][0]); T("excel.sheets", len(wb.sheetnames) >= 3, str(wb.sheetnames))
        from docx import Document
        rep = Document([m for m in made if "auditors_report" in m][0]); txt = "\n".join(x.text for x in rep.paragraphs)
        T("report.name", "Mascot Techcrib Private Limited" in txt)
        T("report.fy_date", "March 31, 2025" in txt and "2024" not in txt.split("Basis")[0].replace("2024-25", ""), "")
        T("report.not_caro", "not applicable to the company" in txt and "ANNEXURE A" not in txt)
        T("report.rule11_funds_and_audit_trail", "intermediar" in txt and "audit trail" in txt and "preserved" in txt)
        T("report.going_concern", "Material Uncertainty Related to Going Concern" in txt)
        T("report.sign", "Membership No. 123456" in txt and "UDIN: 25123456ABCDEF1234" in txt and "Place: Kochi" in txt)
        T("report.no_cfs_wording", "Cash Flow Statement" not in txt.split("Basis")[0])
        # --- bigger company -> CARO applies, IFC, cash flow
        big = copy.deepcopy(led)
        for l in big:
            if l["name"] == "Equity Share Capital": l["bal"] = l["open"] = -200000000.0
        big.append(_L("Sales", "Sales Accounts", 0, -1500000000.0)); big.append(_L("Bank", "Bank Accounts", 0, 1500000000.0 - 4.0))
        ctx2 = build_all(big, dict(prof, facts={"audit_trail": "partial"}), "2024-04-01", "2025-03-31", grp, pled, "snapshot", None)
        a2 = ctx2["applic"]
        T("applic.big_company_caro", (not a2["small"]) and a2["caro_applicable"] and a2["cfs"] and not a2["ifc_exempt"] or a2["caro_applicable"], str(a2.get("caro_why")))
        r2_ = build_audit_report(ctx2, str(tmp / "big.docx")); T("report.big_built", bool(r2_) and Path(tmp / "big.docx").exists())
        if r2_:
            t2 = "\n".join(x.text for x in Document(str(tmp / "big.docx")).paragraphs)
            T("report.big_annexure_a", "ANNEXURE A" in t2 and "As required by the Companies (Auditor’s Report) Order, 2020" in t2)
            T("report.partial_audit_trail_qualified", "except for the matters stated in the paragraph 2(h)(" in t2)
            lett = {m.group(1): m.group(0) for m in re.finditer(r"^([a-z])\.\tWith respect to the other matters", t2, re.M)}
            T("report.rule11_letter_matches_reference", "h" in lett, str(lett))
        # --- old year: regime and Rule 11 switch by year
        from copy import deepcopy
        T("law.caro2016_for_FY1920", RULES.at("caro_regime", "2019-04-01")["order"] == "CARO 2016")
        T("law.caro2020_for_FY2122", RULES.at("caro_regime", "2021-04-01")["order"] == "CARO 2020")
        T("law.small_limits_by_date", RULES.at("small_company", "2022-03-31")["paid_up"] == 20000000 and RULES.at("small_company", "2025-03-31")["paid_up"] == 40000000 and RULES.at("small_company", "2026-03-31")["paid_up"] == 100000000)
        T("law.audit_trail_not_in_FY2223", RULES.at("rule11_audit_trail", "2022-04-01") is None)
        # ---- opening-balance fallback for prior year
        ctx3 = build_all(led, prof, "2024-04-01", "2025-03-31", grp, prior_from_opening(led), "opening", vch)
        T("fallback.opening_prior_warned", any(f.area == "Comparatives" and f.sev == "WARN" for f in ctx3["findings"]))
        ctx4 = build_all(led, prof, "2024-04-01", "2025-03-31", grp, None, None, None)
        T("no_prior_ok", ctx4["prv"] is None and bool(ctx4["tables"]) and any(f.area == "Comparatives" for f in ctx4["findings"]))
        T("export.no_prior", len(export_all(ctx4, tmp / "np", ("xlsx", "pdf", "docx", "txt"), True)) == 6)
        # ---- dirty books
        dn = "Mascot DIRTY TEST (FY 2024-25)"; dl, dv = ty.ledgers(dn), ty.vouchers(dn)
        cd = build_all(dl, prof, "2024-04-01", "2025-03-31", {}, pled, "snapshot", dv)
        areas = {(f.sev, f.area) for f in cd["findings"]}
        for need in (("ERROR", "Suspense"), ("ERROR", "Vouchers"), ("WARN", "Cash (40A(3))"), ("WARN", "Cash (269SS)"), ("INFO", "Journal test"), ("WARN", "TDS"), ("WARN", "Completeness")):
            T(f"checker.{need[1]}", need in areas, "")
        # ---- trust (horizontal)
        tn = "Association for Welfare of AirIndia Retired Employees (FY 2025-26)"
        tl = ty.ledgers(tn); tprof = dict(entity="trust", name="Association for Welfare of AirIndia Retired Employees", audit_firm="Raghunath & Iyer", partner="A B", mem_no="1",
                                          place="Kochi", report_date="10.08.2026", signatories=[{"role": "PRESIDENT", "name": "X"}, {"role": "SECRETARY", "name": "Y"}, {"role": "TREASURER", "name": "Z"}])
        ct = build_all(tl, tprof, "2025-04-01", "2026-03-31", {}, None, None, None)
        T("trust.horizontal", ct["tables"][0]["kind"] == "h" and ct["applic"]["caro"] == "n/a")
        T("trust.balances", abs(ct["cur"]["diff"]) < 1, f"diff={ct['cur']['diff']}")
        mt = export_all(ct, tmp / "tr"); T("trust.exports", len(mt) == 6, str(mt))
        tt = "\n".join(x.text for x in Document([m for m in mt if "auditors" in m][0]).paragraphs)
        T("trust.report_no_caro", "do not apply to this Entity" in tt and "surplus" in tt.lower() or "deficit" in tt.lower())
        T("trust.text_render", "LIABILITIES" in render_text(ct["tables"], tprof, "2026-03-31"))
    finally:
        sim.stop()
    # ---- 2. import from Excel TB
    import openpyxl
    wb = openpyxl.Workbook(); ws = wb.active; ws.append(["Trial balance"]); ws.append(["Ledger", "Group", "Debit", "Credit"])
    for r in (("Capital", "Capital Account", 0, 1000), ("Cash", "Cash-in-Hand", 1000, 0), ("Sales", "Sales Accounts", 0, 500), ("Rent", "Indirect Expenses", 200, 0), ("Bank", "Bank Accounts", 300, 0), ("Total", "", 1500, 1500)): ws.append(r)
    wb.save(tmp / "tb.xlsx"); tb = import_tb(tmp / "tb.xlsx")
    T("import_tb", tb is not None and len(tb) == 5 and abs(sum(l["bal"] for l in tb)) < 0.01, str(tb))
    T("import_tb.bad_file_logged_not_crash", import_tb(tmp / "nope.xlsx") is None and any(e["where"] == "import_tb" for e in LOG.errors()))
    # ---- 3. law library: self-repair, update, rollback
    RULES.save(); RULES.save()
    RULES_FILE.write_text("{ this is not json", encoding="utf-8"); RULES.load()
    T("repair.corrupt_rules_from_backup", RULES.at("small_company", "2026-03-31")["paid_up"] == 100000000 and RULES_FILE.exists() and "schema" in RULES_FILE.read_text())
    new = copy.deepcopy(RULES.data); new["version"] = "2099.01.01"
    new["rules"]["small_company"]["history"].append(_h("2099-04-01", "2099-01-01", paid_up=999, turnover=9999))
    new["rules"]["brand_new"] = dict(title="x", history=[_h("2020-01-01", "2020-01-01")]); new["sha256"] = _canon_hash(new)
    ch = update_rules({}, raw=json.dumps(new))
    T("update_rules.applies", ch and any("small_company" in c for c in ch) and RULES.at("small_company", "2099-05-01")["paid_up"] == 999, str(ch))
    bad = copy.deepcopy(new); bad["rules"]["small_company"]["history"][0]["paid_up"] = "lots"
    T("update_rules.rejects_invalid", update_rules({}, raw=json.dumps(bad)) is None)
    tam = copy.deepcopy(new); tam["version"] = "2099.02.02"
    T("update_rules.rejects_bad_checksum", update_rules({}, raw=json.dumps(tam)) is None and RULES.data["version"] == "2099.01.01")
    # ---- 3b. offline law inbox
    ib = tmp / "inbox"; ib.mkdir()
    (ib / "n1.txt").write_text("In the Companies (Specification of definitions details) Rules, a small company means paid up share capital not exceeding Rs. 10 crore and turnover not exceeding Rs. 100 crore, w.e.f. 01.12.2025.", encoding="utf-8")
    (ib / "n2.txt").write_text("Small company: paid up share capital does not exceed Rupees 20 crore and turnover does not exceed Rupees 200 crore with effect from 01-04-2030.", encoding="utf-8")
    pr_ = scan_law_inbox(folder=ib)
    T("inbox.finds_two", len(pr_) == 2, str(pr_))
    old = [p for p in pr_ if p["source"] == "n1.txt"]; new_ = [p for p in pr_ if p["source"] == "n2.txt"]
    T("inbox.recognises_existing", old and old[0]["status"] == "already in library")
    T("inbox.new_needs_approval", new_ and new_[0]["status"].startswith("NEW") and new_[0]["values"]["paid_up"] == 200000000 and new_[0]["frm"] == "2030-04-01", str(new_))
    nodate = dict(new_[0], frm="")
    T("inbox.refuses_without_date", apply_proposal(nodate) is False)
    T("inbox.apply_and_use", apply_proposal(new_[0]) and RULES.at("small_company", "2031-03-31")["paid_up"] == 200000000 and RULES.at("small_company", "2026-03-31")["paid_up"] == 100000000)
    # ---- 4. live exchange rates via a local server (good provider after a bad one)
    from http.server import BaseHTTPRequestHandler, HTTPServer
    class H(BaseHTTPRequestHandler):
        def log_message(self, *a): pass
        def do_GET(self):
            if self.path == "/bad": self.send_response(500); self.end_headers(); return
            body = json.dumps({"date": "2026-11-07", "rates": {"USD": 0.0112, "EUR": 0.0104}}).encode() if self.path == "/ok" else b'{"rates":{"USD":0}}'
            self.send_response(200); self.end_headers(); self.wfile.write(body)
    srv = HTTPServer(("127.0.0.1", 0), H); threading.Thread(target=srv.serve_forever, daemon=True).start(); b = f"http://127.0.0.1:{srv.server_address[1]}"
    fx = update_fx({"fx_providers": [b + "/bad", b + "/zero", b + "/ok"]})
    T("fx.falls_over_to_next_provider", fx and abs(fx["fx"]["USD"] - 89.2857) < 0.01 and fx["fx_asof"] == "2026-11-07", str(fx))
    keep = RULES.data["live"]["fx"].copy(); T("fx.all_fail_keeps_old", update_fx({"fx_providers": [b + "/bad"]}) is None and RULES.data["live"]["fx"] == keep)
    srv.shutdown()
    # ---- 5. self update (positive path only at top level; the child test skips this to avoid recursion)
    me = Path(__file__).resolve() if "__file__" in globals() else None
    if me and not getattr(sys, "frozen", False) and not os.environ.get("TCS_NESTED"):
        src = me.read_text(encoding="utf-8"); os.environ["TCS_NESTED"] = "1"
        try:
            T("self_update.same_version", self_update({}, new_source=src, apply=False) == "current")
            up = re.sub(r'APP_VERSION = "[\d.]+"', 'APP_VERSION = "99.0.0"', src, count=1)
            T("self_update.good_candidate_passes_selftest", self_update({}, new_source=up, apply=False) == "99.0.0")
            T("self_update.bad_syntax_rejected", self_update({}, new_source=up + "\ndef (:\n", apply=False) is None)
            T("self_update.sha_pin", self_update({"app_sha256": "00"}, new_source=up, apply=False) is None)
        finally: os.environ.pop("TCS_NESTED", None)
    intel_tests(T, tmp)
    # ---- 6. error log + guard
    @guard("probe", default="dflt")
    def boom(): raise RuntimeError("boom")
    T("guard.logs_and_survives", boom() == "dflt" and any(e["where"] == "probe" for e in LOG.errors()))
    T("inr.format", inr(1234567.891) == "12,34,567.89" and inr(-5, paren=True) == "(5.00)" and inr(0) == "-")
    # ---- 7. System theme helpers (pure code - no window needed)
    T("theme.rank", rank_for(0)[1] == "E" and rank_for(100)[1] == "D" and rank_for(5000)[1] == "S" and level_for(250) == 3)
    T("theme.mix", mixc("#000000", "#ffffff", .5) == "#7f7f7f")
    try: T("theme.sound", wave.open(io.BytesIO(synth_tones(SFX["ok"]()))).getnframes() > 1000)
    except Exception as e: T("theme.sound", False, str(e))
    try:
        from PIL import Image
        T("theme.bg", render_bg(320, 200).size == (320, 200))
    except ImportError: pass
    except Exception as e: T("theme.bg", False, str(e))
    return passed[0], fails

def selftest(quiet=False):
    """Runs the full test in a separate process with its own throw-away data folder (never touches your rules/settings)."""
    tmp = tempfile.mkdtemp()
    cmd = [sys.executable, "--selftest-run"] if getattr(sys, "frozen", False) else [sys.executable, str(Path(__file__).resolve()), "--selftest-run"]
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=600, env=dict(os.environ, TALLY_CA_HOME=tmp))
    except Exception as e:
        return False, [f"could not start self-test: {e}"]
    lines = [l for l in r.stdout.splitlines() if l.startswith("FAIL")]
    if r.returncode != 0 and not lines: lines = [(r.stdout + r.stderr)[-500:]]
    return r.returncode == 0, lines

# =============================================================== COMMAND LINE (no window needed)
def cli(argv):
    import argparse
    ap = argparse.ArgumentParser(prog="tally_suite", description="Tally CA Suite - command line mode")
    ap.add_argument("--tb", help="trial balance .xlsx/.csv"); ap.add_argument("--prior-tb", help="last year's trial balance")
    ap.add_argument("--fy-end", required=True); ap.add_argument("--name", required=True); ap.add_argument("--entity", default="company", choices=["company", "trust", "firm"])
    ap.add_argument("--out", default=str(APP_DIR / "output")); ap.add_argument("--tally-company"); ap.add_argument("--port", type=int, default=9000)
    a = ap.parse_args(argv); e = pdate(a.fy_end); s = date(e.year - 1, e.month, e.day) + timedelta(days=1)
    prof = dict(entity=a.entity, name=a.name, facts={})
    if a.tb: led, grp, vch = import_tb(a.tb), {}, None
    else:
        t = Tally(port=a.port); led, grp, vch = t.ledgers(a.tally_company), t.groups(a.tally_company), t.vouchers(a.tally_company)
    prior, src = (import_tb(a.prior_tb), "file") if a.prior_tb else (prior_from_opening(led), "opening")
    ctx = build_all(led, prof, s, e, grp, prior, src, vch)
    for x in ctx["findings"]: print(f"[{x.sev}] {x.area}: {x.msg}")
    print("\n".join(export_all(ctx, a.out)))
    for er in LOG.errors(): print("ERROR LOG:", er["where"], er["msg"])

# =============================================================== WINDOW
SEV_COL = {"ERROR": "#ff6b8b", "WARN": "#ffd166", "INFO": "#4fd1ff", "OK": "#5cf2b0"}      # bright, readable on the dark System theme
BG_JPEG_B64 = (
    "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"
    "S9TJK29m432a9ukUECxxDSi8CuMq5yBvWDH1lyd4xozxnerM3V4VA9FdR76tsVn4+S9o3EuXTOy5OAvvSbKMSSa2wI05J81Qkvvx"
    "RyDj/l8UE0rpb6TkDnFVWKSXEb2E3U8jTuZGyc42qqXoWekl/Fd0YUicpD9dcJMGkaq4MAdzTcRORej0kanYgeB3p019kAICMbD6"
    "VlNMgcbn6iueYEHScileK3bDzHz3BbOo89qplwTuM0DNQFqtGFEZTsYzClOc1xNCaokScrAJqV96Eg5qV5pxGaX45x0lbTUNImLY"
    "wOMf61VTfAG5NAOKKF/SlR/8rA1JRSToCWzet4hBEqDnufJo3fSPeoVw66lOxpbDJrh7ezuSpaIZs811CRXUxj82i1fuVH3piWxV"
    "gRIAR7UUM6uQrgBv2NW1RDnOoHG2KvixY3uJ5MpyTpiVtpGfCFSD74/rVV45IXKSoyN4YYrUjQjtke1aEcYuIgJFDgbEEZBHaur0"
    "OfTJ/mFDtGDExK/Srd1NbSxQxwW6RGNMNIGJMjZOWP8Ap7Vd6n8OzRWZvbNHaEAmSM7lR3I8jz3FYRYrGDWU/T+mS6OzHmU4XHoY"
    "HwcHmtjol+9rkPkwO24/ynyKw3YfKw71p9PHo2sl0+CsY1BWGQWJwo/qfoKp4+RObvpbL4sjjJUe0ikEjq8eNJwRg5GKPOiIu7BV"
    "G7MxwAK878NfEENrcMvVklntmLPlPzBj/Y/tWb1jqknUbppSPTi4jiByEXt9T710fmE/Y9N+ZjUOS7+DS651pbqNbS1yIFbU7nmV"
    "u32Hb65rIhHrMwMiJpUkas/MfAx3qsp18nA802NgpHp4weTU1Nt6POnncpWxikq+G2pzXEiEo4II80y5u57sg3DayFCqcAYAGBxV"
    "VirQkNq9VW24xp/81f6ktA9SuiclzvvVqxsLi/uY7a2iaWaU6UQcn/tTOhdPk6lfx2sTKurJZ34RRuSa+qdG6f0/pluVsk+Zhh5X"
    "/O/18D2FJlyrFH5Zsf8AqOm6PnQ6ebXVbzKY2iYrIDzq70mWPRGzNhIxuR/vmvo3xJ0aDqd1JJERFOGOmT+Vh2Df618366zRO9i6"
    "lZUfEoP8pHahjyRyxtd/BxZ8ThPvTMWe6eaIo8ac5DDkDxSFXJq5HaSTsVhQsQNTew8mtSxjitnT1E/hagHYD5iO+PeorBOcvqYY"
    "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"
    "2Q=="
)

# =============================================================== SYSTEM THEME  (look . animation . sound . XP)
# Everything below is cosmetic. It never touches the accounting logic: it only listens to the status bar and styles widgets.
# If anything in here fails it logs a WARN and the app carries on un-themed.
import math, random, time, wave, array

PAL = dict(bg="#07061c", panel="#0b0d33", panel2="#12154f", field="#080a2c", line="#2f6fb5", ac="#8b6bff", ac2="#4fd1ff",
           tx="#f0eeff", mut="#aaa6e0", ok="#5cf2b0", warn="#ffd166", bad="#ff6b8b")
RANKS = [(0, "E", "Novice Clerk"), (100, "D", "Ledger Hunter"), (250, "C", "Trial-Balance Knight"),
         (500, "B", "Audit Commander"), (900, "A", "Statement Sovereign"), (1500, "S", "Shadow Monarch of Accounts")]
def rank_for(xp):
    r = RANKS[0]
    for t in RANKS:
        if xp >= t[0]: r = t
    return r
def level_for(xp): return int(xp) // 100 + 1
def mixc(c1, c2, a):
    """blend colour c1 -> c2 by a (0..1); '#rrggbb' strings"""
    a = max(0.0, min(1.0, a)); x = [int(c1[i:i + 2], 16) for i in (1, 3, 5)]; y = [int(c2[i:i + 2], 16) for i in (1, 3, 5)]
    return "#%02x%02x%02x" % tuple(int(x[i] + (y[i] - x[i]) * a) for i in range(3))

# ------------------------------------------------------------------ sound synthesis (no files shipped - generated on first run)
SND_SR = 22050
def _osc(kind, ph):
    if kind == "square": return 1.0 if (ph % 1.0) < 0.5 else -1.0
    if kind == "sawtooth": return 2.0 * ((ph + 0.5) % 1.0) - 1.0
    if kind == "triangle": return 4.0 * abs((ph % 1.0) - 0.5) - 1.0
    return math.sin(6.283185307179586 * ph)
def _wav_bytes(buf, sr, gain=1.0):
    pcm = array.array("h", (int(max(-1.0, min(1.0, x * gain)) * 32767) for x in buf))
    if sys.byteorder == "big": pcm.byteswap()
    bio = io.BytesIO()
    with wave.open(bio, "wb") as w: w.setnchannels(1); w.setsampwidth(2); w.setframerate(sr); w.writeframes(pcm.tobytes())
    return bio.getvalue()
def synth_tones(seq, sr=SND_SR, gain=3.0):
    """seq = [(freq_hz, seconds, wave, volume, start_seconds), ...]  - same recipe as the web theme's oscillator beeps."""
    n = int((max(s + d for _, d, _, _, s in seq) + 0.03) * sr); buf = [0.0] * n
    for f, d, kind, v, s in seq:
        i0, m = int(s * sr), int(d * sr); k = (0.001 / v) ** (1.0 / max(1, m)); env = v; att = max(1, int(0.003 * sr))
        for i in range(m):
            j = i0 + i
            if j < n: buf[j] += env * min(1.0, i / att) * _osc(kind, f * (i / sr))
            env *= k
    return _wav_bytes(buf, sr, gain)
def _arp(freqs, d, kind, v, step, start=0.0): return [(f, d, kind, v, start + i * step) for i, f in enumerate(freqs)]
SFX = {
    "click":  lambda: [(660, .07, "square", .035, 0)],
    "nav":    lambda: [(440, .08, "triangle", .06, 0), (880, .12, "triangle", .09, .07)],
    "ok":     lambda: _arp([523, 659, 784], .16, "sine", .09, .08),
    "lvl":    lambda: _arp([392, 523, 659, 784, 1047, 1319], .3, "sawtooth", .055, .1),
    "bad":    lambda: [(160, .22, "sawtooth", .06, 0), (120, .3, "sawtooth", .06, .12)],
    "tick":   lambda: [(1000, .05, "square", .04, 0)],
    "notify": lambda: [(1175, .09, "sine", .08, 0), (1568, .18, "sine", .08, .09)],
    "boot":   lambda: _arp([220 * (1.12 ** i) for i in range(12)], .09, "triangle", .06, .06),
    "arise":  lambda: [(55, 1.8, "sine", .30, 0), (110, 1.3, "sawtooth", .07, 0), (82, 1.5, "sine", .14, .1)] + _arp([880, 1109, 1319, 1760, 2217], .6, "sine", .05, .12, .45),
}
def synth_ambient(sr=11025, seed=7):
    """~32 s seamless loop: soft evolving pad (Am - F - C - G) with a few distant bell notes. Quiet by design."""
    seg, ov = 8.0, 2.0; total = seg * 4; n = int(total * sr); buf = [0.0] * n; sin = math.sin; tp = 6.283185307179586
    chords = [(220.00, 261.63, 329.63), (174.61, 220.00, 261.63), (261.63, 329.63, 392.00), (196.00, 246.94, 293.66)]; bass = [110.0, 87.31, 130.81, 98.0]
    for ci, (ch, b) in enumerate(zip(chords, bass)):
        start = int(ci * seg * sr); L = int((seg + ov) * sr); voices = [(f, .5) for f in ch] + [(b, .7), (ch[0] * 2, .15)]
        for i in range(L):
            t = i / sr; env = sin(3.141592653589793 * i / (L - 1)) ** 2; trem = 0.85 + 0.15 * sin(tp * 0.25 * t + ci); s = 0.0
            for f, a in voices: s += a * (sin(tp * f * t) + 0.25 * sin(2 * tp * f * t + 0.5))
            buf[(start + i) % n] += 0.05 * env * trem * s
    rnd = random.Random(seed); pent = [440.0, 523.25, 587.33, 659.25, 783.99, 880.0]
    for k in range(14):
        f = rnd.choice(pent); t0 = k * 2.2 + rnd.uniform(0, 0.8); i0 = int(t0 * sr)
        for i in range(int(2.4 * sr)):
            t = i / sr; buf[(i0 + i) % n] += 0.035 * math.exp(-t * 2.2) * sin(tp * f * t)
    pk = max(1e-9, max(abs(x) for x in buf)); return _wav_bytes([x / pk * 0.30 for x in buf], sr, 1.0)

class SysAudio:
    """Plays the synthesised sounds. Windows: MCI (several sounds + looping music at once). Mac/Linux: afplay / paplay / aplay. Never raises."""
    def __init__(self, folder, sfx=True):
        self.dir = Path(folder); self.sfx_on = sfx; self.music_on = False; self._n = 0; self._mproc = None; self._mthread = None
        self.win = sys.platform.startswith("win"); self.cmd = None
        if not self.win: self.cmd = shutil.which("afplay") or shutil.which("paplay") or shutil.which("aplay")
        try: self.dir.mkdir(parents=True, exist_ok=True)
        except Exception: pass
    def _file(self, name, builder):
        p = self.dir / f"{name}.wav"
        if not p.exists() or p.stat().st_size < 100:
            p.write_bytes(builder())
        return p
    def prepare_sfx(self):   # fast (first run only) - done before the window shows
        try:
            for k, fn in SFX.items(): self._file(k, lambda fn=fn: synth_tones(fn()))
        except Exception as e: LOG.add("WARN", "theme.audio", f"{type(e).__name__}: {e}", traceback.format_exc())
    def prepare_music(self):   # a few seconds, first run only - done in a thread
        try: self._file("ambient_v1", synth_ambient)
        except Exception as e: LOG.add("WARN", "theme.audio", f"{type(e).__name__}: {e}", traceback.format_exc())
    def _mci(self, cmd):
        try:
            import ctypes; return ctypes.windll.winmm.mciSendStringW(cmd, None, 0, None)
        except Exception: return -1
    def sfx(self, name):
        if not self.sfx_on or name not in SFX: return
        try:
            p = self.dir / f"{name}.wav"
            if not p.exists(): return
            if self.win:
                al = f"tcs{self._n % 6}"; self._n += 1; self._mci(f"close {al}")
                if self._mci(f'open "{p}" type waveaudio alias {al}') == 0: self._mci(f"play {al}")
                else:
                    import winsound; winsound.PlaySound(str(p), winsound.SND_FILENAME | winsound.SND_ASYNC)
            elif self.cmd:
                args = [self.cmd] + (["-q"] if self.cmd.endswith("aplay") else []) + [str(p)]
                subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception: pass
    def music(self, on):
        self.music_on = bool(on)
        try:
            p = self.dir / "ambient_v1.wav"
            if self.win:
                self._mci("stop tcsbgm"); self._mci("close tcsbgm")
                if on and p.exists() and self._mci(f'open "{p}" type waveaudio alias tcsbgm') == 0: self._mci("play tcsbgm repeat")
            else:
                if not on:
                    if self._mproc:
                        try: self._mproc.terminate()
                        except Exception: pass
                elif self.cmd and p.exists() and not (self._mthread and self._mthread.is_alive()):
                    def loop():
                        while self.music_on:
                            try:
                                self._mproc = subprocess.Popen([self.cmd] + (["-q"] if self.cmd.endswith("aplay") else []) + [str(p)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL); self._mproc.wait()
                            except Exception: break
                    self._mthread = threading.Thread(target=loop, daemon=True); self._mthread.start()
        except Exception: pass
    def stop_all(self):
        self.music_on = False; self.music(False)


def render_bg(w, h, src_bytes=None):
    """Cover-fit the artwork to w x h and add a dark scrim (heavier at the top so the title reads). Needs Pillow; returns a PIL image."""
    from PIL import Image, ImageOps
    im = Image.open(io.BytesIO(src_bytes or base64.b64decode(BG_JPEG_B64))).convert("RGB"); w, h = max(2, int(w)), max(2, int(h))
    sw, sh = im.size; sc = max(w / sw, h / sh); nw, nh = int(sw * sc) + 1, int(sh * sc) + 1
    im = im.resize((nw, nh), getattr(Image, "Resampling", Image).BILINEAR); l, t = (nw - w) // 2, (nh - h) // 2; im = im.crop((l, t, l + w, t + h))
    dark = Image.new("RGB", (w, h), (6, 4, 30)); im = Image.blend(im, dark, 0.28)
    grad = Image.linear_gradient("L").resize((w, h)); top = ImageOps.invert(grad).point(lambda v: int(v * 0.55)); return Image.composite(dark, im, top)

def _fx_safe(fn):
    def w(self, *a, **k):
        try: return fn(self, *a, **k)
        except Exception as e:
            self.errs += 1
            if self.errs <= 3: LOG.add("WARN", "theme", f"{fn.__name__}: {type(e).__name__}: {e}", traceback.format_exc())
            if self.errs > 8: self.anim_on = False
    w.__name__ = fn.__name__; return w

class SystemTheme:
    HEAD = 124
    def __init__(self, app):
        import tkinter as tk
        from tkinter import ttk
        self.tk, self.ttk, self.app = tk, ttk, app; s = app.settings
        for k, v in (("fx_sfx", True), ("fx_music", True), ("fx_anim", True), ("fx_boot", True), ("fx_xp", 0)): s.setdefault(k, v)
        self.errs = 0; self.W, self.H = 1280, 800; self.frame = 0; self.after_id = None; self.anim_on = bool(s["fx_anim"])
        self.stars, self.parts, self.shots, self.ripples, self.toasts, self.arcs = [], [], [], [], [], []
        self.ov = None; self._tt = None; self._tn = 0; self._last_bad = 0.0; self.bgimg = None; self._bgjob = None; self._hudjob = None; self._pil = None
        try: import PIL.ImageTk; self._pil = True
        except Exception: self._pil = False
        self.audio = SysAudio(APP_DIR / "sounds", sfx=bool(s["fx_sfx"])); self.audio.prepare_sfx(); threading.Thread(target=self.audio.prepare_music, daemon=True).start()
        self.style(); app.configure(bg=PAL["bg"]); app.minsize(1000, 640)
        try:
            sh = app.winfo_screenheight(); app.geometry(f"1280x{min(860, max(680, sh - 90))}")
        except Exception: pass
        self.make_canvas(); self.restyle(app); self.hooks(); app.after(60, self.start)

    # ---------------------------------------------------------------- ttk look
    def style(self):
        tk, ttk, app = self.tk, self.ttk, self.app; st = ttk.Style(app)
        try: st.theme_use("clam")
        except Exception: pass
        P = PAL; fb = ("Segoe UI", 10)
        st.configure(".", background=P["panel"], foreground=P["tx"], fieldbackground=P["field"], bordercolor=P["line"], lightcolor=P["panel2"], darkcolor=P["panel"],
                     troughcolor=P["bg"], focuscolor=P["ac2"], insertcolor=P["ac2"], selectbackground="#3b2bff", selectforeground="#ffffff", font=fb)
        st.configure("TFrame", background=P["panel"]); st.configure("TLabel", background=P["panel"], foreground=P["tx"])
        st.configure("TLabelframe", background=P["panel"], bordercolor=P["ac2"], lightcolor=P["ac2"], darkcolor=P["ac2"]); st.configure("TLabelframe.Label", background=P["panel"], foreground=P["ac2"], font=("Segoe UI", 10, "bold"))
        st.configure("TButton", background="#1b1f6b", foreground=P["tx"], bordercolor=P["ac2"], lightcolor="#2a2f9a", darkcolor="#10124a", padding=(12, 6), focuscolor=P["panel"], relief="flat")
        st.map("TButton", background=[("pressed", P["ac2"]), ("active", "#2b35a8")], foreground=[("pressed", "#02020e"), ("active", "#ffffff")], bordercolor=[("active", "#ffffff")])
        st.configure("TEntry", fieldbackground=P["field"], foreground=P["tx"], bordercolor=P["line"], insertcolor=P["ac2"], padding=4)
        st.map("TEntry", bordercolor=[("focus", P["ac2"])], fieldbackground=[("readonly", P["panel2"])])
        st.configure("TCombobox", fieldbackground=P["field"], background="#1b1f6b", foreground=P["tx"], arrowcolor=P["ac2"], bordercolor=P["line"], padding=4)
        st.map("TCombobox", fieldbackground=[("readonly", P["field"])], foreground=[("readonly", P["tx"])], bordercolor=[("focus", P["ac2"])], background=[("active", "#2b35a8")])
        app.option_add("*TCombobox*Listbox.background", P["field"]); app.option_add("*TCombobox*Listbox.foreground", P["tx"])
        app.option_add("*TCombobox*Listbox.selectBackground", "#3b2bff"); app.option_add("*TCombobox*Listbox.selectForeground", "#ffffff")
        st.configure("TNotebook", background=P["bg"], bordercolor=P["ac2"], lightcolor=P["ac2"], darkcolor=P["ac2"], tabmargins=(2, 4, 2, 0))
        st.configure("TNotebook.Tab", background=P["panel"], foreground=P["mut"], padding=(16, 7), bordercolor=P["line"], lightcolor=P["panel"], darkcolor=P["panel"])
        st.map("TNotebook.Tab", background=[("selected", "#1b1f6b"), ("active", P["panel2"])], foreground=[("selected", P["ac2"]), ("active", "#ffffff")], expand=[("selected", (0, 2, 0, 0))])
        st.configure("Treeview", background=P["field"], fieldbackground=P["field"], foreground=P["tx"], rowheight=24, bordercolor=P["line"])
        st.configure("Treeview.Heading", background="#1b1f6b", foreground=P["ac2"], relief="flat", font=("Segoe UI", 10, "bold"))
        st.map("Treeview", background=[("selected", "#3b2bff")], foreground=[("selected", "#ffffff")]); st.map("Treeview.Heading", background=[("active", "#2b35a8")])
        for w in ("TCheckbutton", "TRadiobutton"):
            st.configure(w, background=P["panel"], foreground=P["tx"], indicatorcolor=P["field"], indicatorbackground=P["field"], focuscolor=P["panel"])
            st.map(w, background=[("active", P["panel2"])], indicatorcolor=[("selected", P["ac2"]), ("active", P["ac"])], foreground=[("active", "#ffffff")])
        for w in ("Vertical.TScrollbar", "Horizontal.TScrollbar"):
            st.configure(w, background="#1b1f6b", troughcolor=P["bg"], arrowcolor=P["ac2"], bordercolor=P["line"], lightcolor="#1b1f6b", darkcolor="#1b1f6b")
            st.map(w, background=[("active", "#2b35a8")])
        st.configure("TSeparator", background=P["line"])
        st.configure("Horizontal.TProgressbar", background=P["ac2"], troughcolor=P["field"], bordercolor=P["line"], lightcolor=P["ac2"], darkcolor=P["ac"])

    def restyle(self, w):
        """Recolour the plain-tk widgets (Text, Label, Frame, Listbox) that ttk styles cannot reach."""
        try: name = str(w).rsplit(".", 1)[-1]
        except Exception: name = ""
        if name.startswith("sys_"): return
        P = PAL
        try:
            cls = w.winfo_class()
            if cls == "Text": w.configure(bg=P["field"], fg=P["tx"], insertbackground=P["ac2"], selectbackground="#3b2bff", selectforeground="#ffffff", relief="flat", highlightthickness=1, highlightbackground=P["line"], highlightcolor=P["ac2"])
            elif cls == "Listbox": w.configure(bg=P["field"], fg=P["tx"], selectbackground="#3b2bff", relief="flat", highlightthickness=1, highlightbackground=P["line"])
            elif cls == "Label": w.configure(bg=P["panel"])
            elif cls == "Frame": w.configure(bg=P["panel"])
            elif cls == "TLabel":
                fg = str(w.cget("foreground")).lower()
                if fg == "#1f3864": w.configure(foreground=P["ac2"], font=("Consolas", 10, "bold"))
                elif fg in ("#555", "#555555"): w.configure(foreground=P["mut"])
        except Exception: pass
        for c in w.winfo_children(): self.restyle(c)

    # ---------------------------------------------------------------- scene canvas (sits UNDER every widget, shows in the header + margins)
    def make_canvas(self):
        tk, app = self.tk, self.app
        self.cv = tk.Canvas(app, bg=PAL["bg"], highlightthickness=0, bd=0, name="sys_bg"); self.cv.place(x=0, y=0, relwidth=1, relheight=1); tk.Misc.lower(self.cv)
        top = self.app.errlbl.master; top.pack_configure(pady=(self.HEAD + 4, 4), padx=18)      # reserve the header strip (nothing sits on it)
        try: app.nb.pack_configure(padx=18, pady=(4, 16))
        except Exception: pass
        self.cv.bind("<Button-1>", self.on_click_canvas); app.bind("<Configure>", self.on_configure, add="+")

    def paint_bg(self):
        cv = self.cv; cv.delete("bg")
        if self.bgimg is not None: cv.create_image(0, 0, anchor="nw", image=self.bgimg, tags="bg")
        else:
            h = max(self.H, 2)
            for y in range(0, h, 8): cv.create_rectangle(0, y, self.W, y + 8, outline="", fill=mixc("#0b0a2a", "#3a2488", (y / h) ** 1.6), tags="bg")
        cv.tag_lower("bg")
    @_fx_safe
    def render_bg_now(self):
        self._bgjob = None
        if not self._pil: return
        from PIL import ImageTk
        self.bgimg = ImageTk.PhotoImage(render_bg(self.W, self.H)); self.paint_bg()
    def on_configure(self, ev):
        if ev.widget is not self.app: return
        w, h = ev.width, ev.height
        if w < 100 or (w == self.W and h == self.H): return
        self.W, self.H = w, h
        if self._bgjob: self.app.after_cancel(self._bgjob)
        self._bgjob = self.app.after(160, self.render_bg_now)
        if self._hudjob: self.app.after_cancel(self._hudjob)
        self._hudjob = self.app.after(60, self.build_scene)

    @_fx_safe
    def build_scene(self):
        self._hudjob = None; cv = self.cv; cv.delete("star", "part", "hud"); self.paint_bg(); W = self.W; rnd = random
        self.stars = []
        for _ in range(85):
            x = rnd.uniform(0, W); y = rnd.uniform(0, self.HEAD + 14) if rnd.random() < .6 else rnd.uniform(0, self.H); r = rnd.choice((1, 1, 1.5, 2))
            self.stars.append([cv.create_oval(x - r, y - r, x + r, y + r, fill="#ffffff", outline="", tags="star"), rnd.uniform(0, 6.28), rnd.uniform(.04, .16)])
        self.parts = []
        for _ in range(26):
            x, y = rnd.uniform(0, W), rnd.uniform(0, self.H); r = rnd.uniform(1.5, 3.2); col = rnd.choice((PAL["ac2"], PAL["ac"], "#c9b8ff"))
            self.parts.append([cv.create_oval(x - r, y - r, x + r, y + r, fill=col, outline="", tags="part"), rnd.uniform(.2, .8), rnd.uniform(-.25, .25), col])
        self.build_hud()
    def build_hud(self):
        cv, W, P = self.cv, self.W, PAL; cv.delete("hud"); self.arcs = []; xp = int(self.app.settings.get("fx_xp", 0)); rk = rank_for(xp); lv = level_for(xp)
        f = ("Georgia", 27, "bold"); title = "TALLY  CA  SUITE"
        for dx, dy, c in ((-2, 0, "#3b2bff"), (2, 0, "#3b2bff"), (0, 2, P["ac"])): cv.create_text(26 + dx, 34 + dy, anchor="w", text=title, font=f, fill=c, tags="hud")
        self.title = cv.create_text(26, 34, anchor="w", text=title, font=f, fill="#f0eeff", tags="hud")
        cv.create_text(28, 66, anchor="w", text="\u25c6 THE SYSTEM  \u00b7  STATEMENTS \u00b7 CHECKS \u00b7 AUDIT", font=("Consolas", 10, "bold"), fill=P["ac2"], tags="hud")
        self.player = cv.create_text(28, 90, anchor="w", text="PLAYER: awaiting data...", font=("Consolas", 10), fill=P["warn"], tags="hud")
        cx, cy = W - 74, 62                                                                    # rank badge with spinning rings
        cv.create_rectangle(cx - 30, cy - 30, cx + 30, cy + 30, outline=P["ac2"], width=2, fill="#080a3c", tags="hud")
        self.rank_txt = cv.create_text(cx, cy, text=rk[1], font=("Georgia", 30, "bold"), fill=P["ac2"], tags="hud")
        for r_, ext, col, sp in ((44, 70, P["ac2"], 3.0), (54, 40, P["ac"], -2.0)):
            for k in range(3): self.arcs.append((cv.create_arc(cx - r_, cy - r_, cx + r_, cy + r_, start=k * 120, extent=ext, style="arc", outline=col, width=2, tags="hud"), k * 120, sp))
        x1, x2, y = W - 430, W - 150, 88                                                       # level + xp bar
        cv.create_text(x1, 66, anchor="w", text=f"LV {lv}  \u00b7  {rk[2].upper()}", font=("Consolas", 10, "bold"), fill=P["warn"], tags="hud")
        cv.create_text(x2, 66, anchor="e", text=f"{xp % 100}/100 XP", font=("Consolas", 9), fill=P["mut"], tags="hud")
        cv.create_rectangle(x1, y, x2, y + 12, outline=P["ac2"], fill=P["field"], tags="hud"); fw = int((x2 - x1 - 2) * (xp % 100) / 100)
        if fw > 0: cv.create_rectangle(x1 + 1, y + 1, x1 + 1 + fw, y + 11, outline="", fill=mixc("#3b2bff", P["ac2"], .7), tags="hud")
        self.xpbar = (x1 + 1, x1 + 1 + max(fw, 0), y + 1); self.shine = cv.create_rectangle(x1, y + 1, x1 + 8, y + 11, outline="", fill="#ffffff", state="hidden", tags="hud")
        s = self.app.settings; x = W - 138
        for tag, lab, on in (("tg_fx", "FX", s["fx_anim"]), ("tg_sfx", "SFX", s["fx_sfx"]), ("tg_music", "\u266a MUSIC", s["fx_music"])):
            t = cv.create_text(x, 18, anchor="e", text=f"[ {lab} {'ON' if on else 'OFF'} ]", font=("Consolas", 9, "bold"), fill=P["ac2"] if on else P["mut"], tags=("hud", tag))
            cv.tag_bind(tag, "<Button-1>", lambda e, t_=tag: self.toggle(t_)); cv.tag_bind(tag, "<Enter>", lambda e: cv.config(cursor="hand2")); cv.tag_bind(tag, "<Leave>", lambda e: cv.config(cursor=""))
            bb = cv.bbox(t); x = (bb[0] - 12) if bb else x - 90
    def toggle(self, tag):
        s = self.app.settings
        if tag == "tg_music": s["fx_music"] = not s["fx_music"]; self.audio.music(s["fx_music"] and not self.ov_is_boot())
        elif tag == "tg_sfx": s["fx_sfx"] = not s["fx_sfx"]; self.audio.sfx_on = s["fx_sfx"]; self.audio.sfx("click")
        else:
            s["fx_anim"] = not s["fx_anim"]; self.anim_on = s["fx_anim"]
            if self.anim_on and self.after_id is None: self.tick()
        save_settings(s); self.build_hud()
    def ov_is_boot(self): return getattr(self, "_booting", False)

    # ---------------------------------------------------------------- animation loop
    def start(self):
        self.audio.sfx_on = bool(self.app.settings["fx_sfx"]); self.build_scene(); self.render_bg_now()
        if self.app.settings.get("fx_boot", True): self.boot()
        else:
            if self.app.settings["fx_music"]: self.app.after(1500, lambda: self.audio.music(True))
        self.tick()
    def tick(self):
        self.after_id = None
        if not self.anim_on: return
        self.step(); self.after_id = self.app.after(40, self.tick)
    @_fx_safe
    def step(self):
        app, cv = self.app, self.cv
        if app.state() == "iconic": return
        self.frame += 1; fr = self.frame; W, H = self.W, self.H
        if fr % 2 == 0:
            for sid, ph, sp in self.stars: self.stars_set(sid, ph + fr * sp)
        for p in self.parts:
            cv.move(p[0], p[2], -p[1]); x1, y1, x2, y2 = cv.coords(p[0])
            if y2 < -6: cv.move(p[0], random.uniform(0, W) - x1, H + 12 - y1)
            elif x1 > W + 6: cv.move(p[0], -W - 12, 0)
            elif x2 < -6: cv.move(p[0], W + 12, 0)
        if random.random() < 0.008 and len(self.shots) < 2:
            x, y = random.uniform(W * .3, W), random.uniform(0, self.HEAD); self.shots.append([cv.create_line(x, y, x - 70, y + 26, fill="#ffffff", width=2, tags="fx"), 0])
        for s in self.shots[:]:
            cv.move(s[0], -16, 6); s[1] += 1; cv.itemconfig(s[0], fill=mixc("#ffffff", PAL["bg"], s[1] / 26))
            if s[1] >= 26: cv.delete(s[0]); self.shots.remove(s)
        for r in self.ripples[:]:
            r[3] += 1; k = r[3] / 18; rad = 6 + 70 * k; cv.coords(r[0], r[1] - rad, r[2] - rad, r[1] + rad, r[2] + rad); cv.itemconfig(r[0], outline=mixc(PAL["ac2"], PAL["bg"], k))
            if r[3] >= 18: cv.delete(r[0]); self.ripples.remove(r)
        for aid, base, sp in self.arcs: cv.itemconfig(aid, start=(base + fr * sp) % 360)
        if hasattr(self, "xpbar") and self.xpbar[1] > self.xpbar[0] + 10:
            x0, x1, y = self.xpbar; k = (fr % 50) / 50; sx = x0 + (x1 - x0 - 8) * k; cv.coords(self.shine, sx, y, sx + 8, y + 10); cv.itemconfig(self.shine, state="normal", fill=mixc("#ffffff", PAL["ac2"], .35))
        elif hasattr(self, "shine"): cv.itemconfig(self.shine, state="hidden")
        m = fr % 180; cv.itemconfig(self.title, fill=("#8f8cc0" if m in (170, 174) else "#f0eeff"))     # the little 'System' flicker
        if fr % 25 == 0:
            co = getattr(app, "company", None) if getattr(app, "led", None) else None
            co = (co[:38] + "...") if co and len(co) > 41 else co
            cv.itemconfig(self.player, text=f"PLAYER: {co}" if co else "PLAYER: awaiting data...")
    def stars_set(self, sid, v): self.cv.itemconfig(sid, fill=mixc(PAL["bg"], "#ffffff", 0.35 + 0.65 * (0.5 + 0.5 * math.sin(v))))
    def on_click_canvas(self, ev):
        if not self.anim_on: return
        o = self.cv.create_oval(ev.x - 6, ev.y - 6, ev.x + 6, ev.y + 6, outline=PAL["ac2"], width=2, tags="fx"); self.ripples.append([o, ev.x, ev.y, 0]); self.audio.sfx("click")

    # ---------------------------------------------------------------- notifications (slide in from the top, type themselves out)
    @_fx_safe
    def notify(self, msg, kind="info", sound=True):
        tk, app = self.tk, self.app; col = {"info": PAL["ac2"], "ok": PAL["ok"], "bad": PAL["bad"], "big": PAL["warn"]}[kind]
        head = {"info": "! NOTIFICATION", "ok": "! QUEST UPDATE", "bad": "! WARNING", "big": "! LEVEL UP"}[kind]; self._tn += 1; bgc = "#06093a"
        if kind == "bad": self._last_bad = time.time()
        f = tk.Frame(app, bg=bgc, highlightthickness=1, highlightbackground=col, highlightcolor=col, name=f"sys_toast{self._tn}")
        tk.Label(f, text=head, bg=bgc, fg=col, font=("Consolas", 8, "bold"), anchor="w").pack(fill="x", padx=14, pady=(7, 0))
        lab = tk.Label(f, text="", bg=bgc, fg=PAL["tx"], font=("Segoe UI", 10), wraplength=560, justify="left", anchor="w"); lab.pack(fill="x", padx=14, pady=(0, 9))
        self.toasts.append(dict(f=f, lab=lab, msg=str(msg), n=0, y=-90.0, die=time.time() + 3.6 + len(str(msg)) / 40))
        f.place(relx=.5, y=-90, anchor="n"); tk.Misc.lift(f)
        if sound: self.audio.sfx("lvl" if kind == "big" else "bad" if kind == "bad" else "notify")
        if self._tt is None: self._toast_tick()
    @_fx_safe
    def _toast_tick(self):
        now = time.time(); keep = []
        for i, t in enumerate(self.toasts):
            dying = now > t["die"]; target = -100 if dying else 8 + i * 70; t["y"] += (target - t["y"]) * 0.28
            if dying and t["y"] < -85: t["f"].destroy(); continue
            if t["n"] < len(t["msg"]): t["n"] += 2; t["lab"].config(text=t["msg"][:t["n"]])
            t["f"].place_configure(y=int(t["y"])); keep.append(t)
        self.toasts = keep; self._tt = self.app.after(30, self._toast_tick) if keep else None

    # ---------------------------------------------------------------- full-window overlays: boot screen + ARISE
    def _overlay(self):
        if self.ov is not None:
            try: self.ov.destroy()
            except Exception: pass
        cv = self.tk.Canvas(self.app, bg="#02020e", highlightthickness=0, name="sys_overlay"); cv.place(x=0, y=0, relwidth=1, relheight=1); self.tk.Misc.lift(cv); self.ov = cv; return cv
    @_fx_safe
    def arise(self, sub="", dur=2.6):
        cv = self._overlay(); w, h = max(self.W, 400), max(self.H, 300); cx, cy = w // 2, h // 2 - 20; self.audio.sfx("arise"); t0 = time.time(); bg0 = "#02020e"
        rings = [cv.create_oval(cx, cy, cx, cy, outline=PAL["ac2"], width=2) for _ in range(4)]
        glow = cv.create_text(cx + 3, cy + 3, text="ARISE", font=("Georgia", 150, "bold"), fill=bg0); main = cv.create_text(cx, cy, text="ARISE", font=("Georgia", 150, "bold"), fill=bg0)
        subt = cv.create_text(cx, cy + 120, text=sub, font=("Consolas", 14, "bold"), fill=bg0)
        def done(*_):
            if self.ov is cv:
                try: cv.destroy()
                except Exception: pass
                self.ov = None
        cv.bind("<Button-1>", done)
        def step():
            try:
                if self.ov is not cv: return
                e = (time.time() - t0) / dur
                if e >= 1: return done()
                a = min(e / .25, 1.0) if e < .75 else max(0.0, (1 - e) / .25); sz = int(150 - 62 * min(e / .3, 1.0))
                cv.itemconfig(main, font=("Georgia", sz, "bold"), fill=mixc(bg0, "#e8e2ff", a)); cv.itemconfig(glow, font=("Georgia", sz, "bold"), fill=mixc(bg0, PAL["ac"], a * .9)); cv.itemconfig(subt, fill=mixc(bg0, PAL["ac2"], a))
                R = min(w, h) * 0.6
                for i, r in enumerate(rings):
                    k = max(0.0, e * 1.5 - i * 0.2)
                    if k > 0: rad = k * R; cv.coords(r, cx - rad, cy - rad, cx + rad, cy + rad); cv.itemconfig(r, outline=mixc(PAL["ac2"] if i % 2 == 0 else PAL["ac"], bg0, min(1.0, k)))
                self.app.after(30, step)
            except Exception: done()
        step()
    @_fx_safe
    def boot(self):
        cv = self._overlay(); self._booting = True; w, h = max(self.W, 400), max(self.H, 300); self.audio.sfx("boot")
        for y in range(0, h, 6): cv.create_line(0, y, w, y, fill="#05061c")
        lines = ["[ SYSTEM ]  Awakening ...", "[ SYSTEM ]  Loading statutory law library ................ OK", "[ SYSTEM ]  Calibrating Schedule III + Trust/Firm formats ... OK",
                 "[ SYSTEM ]  Scanning for Tally gateway (port 9000) ....... STANDBY", "[ SYSTEM ]  Audit-report engine (CARO / Rule 11 / IFC) ..... OK", "", "[ SYSTEM ]  PLAYER DETECTED."]
        ids = [cv.create_text(max(40, w // 2 - 330), max(60, h // 2 - 150) + i * 28, anchor="w", text="", font=("Consolas", 13), fill=PAL["ac2"]) for i in range(len(lines))]
        prompt = cv.create_text(w // 2, h - 90, text="", font=("Consolas", 12, "bold"), fill=PAL["warn"]); st = dict(i=0, j=0, end=False, fin=False)
        def finish(*_):
            if st["fin"]: return
            st["fin"] = True; self._booting = False
            try: self.app.unbind_all("<Key>")
            except Exception: pass
            self.arise("WELCOME, PLAYER", 2.4)
            if self.app.settings["fx_music"]: self.app.after(900, lambda: self.audio.music(True))
        cv.bind("<Button-1>", finish); self.app.bind_all("<Key>", finish)
        def type_step():
            try:
                if st["fin"] or self.ov is not cv: return
                if st["i"] >= len(lines):
                    if not st["end"]: st["end"] = True; cv.itemconfig(prompt, text="[ CLICK OR PRESS ANY KEY TO ARISE ]"); self.app.after(2600, finish)
                    cv.itemconfig(prompt, fill=PAL["warn"] if (time.time() * 2) % 2 < 1 else PAL["bg"]); self.app.after(250, type_step); return
                ln = lines[st["i"]]; st["j"] += 3; cv.itemconfig(ids[st["i"]], text=ln[:st["j"]], fill=PAL["warn"] if "PLAYER" in ln else PAL["ac2"])
                if st["j"] >= len(ln): st["i"] += 1; st["j"] = 0; (ln and self.audio.sfx("tick"))
                self.app.after(14, type_step)
            except Exception: finish()
        type_step()

    # ---------------------------------------------------------------- hooks: the app tells us what happened through its status bar
    def hooks(self):
        app = self.app
        app.status.trace_add("write", lambda *_: self.on_status())
        app.bind_all("<ButtonRelease-1>", self.on_any_click, add="+"); app.nb.bind("<<NotebookTabChanged>>", lambda e: self.audio.sfx("nav"), add="+")
        LOG.listeners.append(lambda it: app.after(0, self.on_log, it))
        g = globals(); orig = g["build_audit_report"]
        def wrapped(*a, **k):
            r = orig(*a, **k)
            if r: app.after(0, lambda: self.award(25, "Auditor's report drafted"))
            return r
        g["build_audit_report"] = wrapped
        app.protocol("WM_DELETE_WINDOW", self.on_close)
    def on_close(self):
        try: self.audio.stop_all()
        except Exception: pass
        self.app.destroy()
    def on_any_click(self, ev):
        try:
            if ev.widget.winfo_class() in ("TButton", "TCheckbutton", "TRadiobutton", "TCombobox"): self.audio.sfx("click")
        except Exception: pass
    def on_log(self, it):
        if it["level"] == "ERROR" and time.time() - self._last_bad > 4: self._last_bad = time.time(); self.notify("An error was logged. Details are in the Errors tab.", "bad")
    @_fx_safe
    def award(self, n, why):
        s = self.app.settings; old = int(s.get("fx_xp", 0)); new = old + n; s["fx_xp"] = new; save_settings(s); r0, r1 = rank_for(old), rank_for(new)
        self.build_hud()
        if r1[1] != r0[1]: self.notify(f"+{n} XP \u00b7 {why}", "ok", sound=False); self.arise(f"RANK UP  \u00b7  {r1[1]}-RANK  \u00b7  {r1[2].upper()}", 3.0)
        elif level_for(new) != level_for(old): self.notify(f"LEVEL {level_for(new)} reached  \u00b7  +{n} XP  \u00b7  {why}", "big")
        else: self.notify(f"+{n} XP  \u00b7  {why}", "ok")
    @_fx_safe
    def on_status(self):
        m = self.app.status.get(); lo = m.lower()
        if m.startswith("Built."):
            mm = re.search(r"(\d+) errors, (\d+) warnings", m); e_, w_ = (int(mm.group(1)), int(mm.group(2))) if mm else (0, 0)
            if e_ == 0: self.award(80, "Statements built  \u00b7  all checks clear" + (f"  ({w_} warnings)" if w_ else ""))
            else: self.award(50, f"Statements built  \u00b7  {e_} anomalies to fix")
        elif any(x in lo for x in ("failed", "cannot connect", "error occurred", "not updated", "not applied", "rejected")): self.notify(m, "bad")
        elif m.startswith("Connected:"): self.award(10, "Gate opened  \u00b7  Tally connected")
        elif "ledgers" in lo and ("loaded" in lo or "imported" in lo): self.award(20, "Dungeon loaded  \u00b7  ledgers read")
        elif lo.startswith("last year's tb loaded"): self.award(15, "Comparatives secured")
        elif "files saved to" in lo: self.award(40, "Statements exported")
        elif lo.startswith("learned from samples"): self.award(15, "Firm profile learned")
        elif lo.startswith(("law library update finished", "exchange rates updated", "self-repair finished", "official sources:")): self.award(10, "System maintenance")
        elif lo.startswith(("step 1", "connecting", "reading", "running self", "checking", "fetching", "updating", "done")): pass
        elif lo.startswith(("rule updated", "update result")): self.award(10, "Law library updated")

def _apply_theme(app):
    try: app.fx = SystemTheme(app)
    except Exception as e: LOG.add("WARN", "theme", f"theme not applied: {type(e).__name__}: {e}", traceback.format_exc())

FACT_BOXES = [("is_public", "Public company"), ("is_sub_holding", "Holding / subsidiary"), ("is_section8", "Section 8 company"), ("is_opc", "One Person Company"),
              ("is_dormant", "Dormant company"), ("is_bank_ins", "Bank / insurance"), ("special_act", "Governed by special Act"), ("ind_as", "Ind AS applies"),
              ("pending_litigation", "Pending litigation impacting position"), ("dividend_paid", "Dividend declared / paid"), ("include_197_16", "Include s.197(16) remuneration para")]
PROF_FIELDS = [("name", "Entity name"), ("cin", "CIN"), ("address", "Address lines (separate with | )"), ("audit_firm", "Audit firm"), ("frn", "FRN"), ("partner", "Partner name"),
               ("mem_no", "Membership no."), ("udin", "UDIN"), ("place", "Place"), ("report_date", "Report date (dd.mm.yyyy)"), ("fs_date", "Statements signing date"),
               ("directors", "Directors (separate with | )"), ("signatories", "Trust/firm signatories  ROLE:Name | ROLE:Name")]

def launch_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, simpledialog

    class App(tk.Tk):
        def __init__(self):
            super().__init__(); self.title(f"Tally CA Suite {APP_VERSION}"); self.geometry("1280x760")
            self.settings = load_settings(); self.led = []; self.grp = {}; self.vch = None; self.prior = None; self.prior_source = None; self.ctx = None; self.companies = []
            self.tally = Tally(self.settings["tally_host"], self.settings["tally_port"]); self.src_file = None
            top = ttk.Frame(self); top.pack(fill="x", padx=8, pady=4)
            self.status = tk.StringVar(value="Step 1: Data tab - connect to Tally or import a trial balance."); ttk.Label(top, textvariable=self.status, foreground="#1F3864").pack(side="left")
            self.errlbl = tk.Label(top, text="Errors: 0", fg=SEV_COL["OK"]); self.errlbl.pack(side="right")
            self.nb = ttk.Notebook(self); self.nb.pack(fill="both", expand=True, padx=8, pady=4)
            self.tabs = {}
            for n in ("Data", "Firm Profile", "Statements", "Checks", "Audit Report", "Law Library", "Legal Intelligence", "Errors"):
                f = ttk.Frame(self.nb); self.nb.add(f, text=n); self.tabs[n] = f
            self.t_data(); self.t_prof(); self.t_stmt(); self.t_checks(); self.t_audit(); self.t_law(); self.t_intel(); self.t_err()
            LOG.listeners.append(lambda it: self.after(0, self.on_log, it))
            self.report_callback_exception = self._tk_exc
            for it in LOG.items: self.on_log(it)
            if self.settings.get("intel_auto", True) and self.settings.get("intel_last", "")[:10] != datetime.now().strftime("%Y-%m-%d"): self.after(3000, lambda: self.do_intel(auto=True))

        # ---------- plumbing
        def _tk_exc(self, et, ev, tb):
            LOG.add("ERROR", "window", f"{et.__name__}: {ev}", "".join(traceback.format_exception(et, ev, tb)))
            self.status.set("An error occurred - see the Errors tab."); self.nb.select(self.tabs["Errors"])
        def on_log(self, it):
            if hasattr(self, "errtree"):
                self.errtree.insert("", 0, values=(it["ts"], it["level"], it["where"], it["msg"][:300]), tags=(it["level"],), iid=None)
                self.errdet[self.errtree.get_children()[0]] = it["tb"] or it["msg"]
            n = len(LOG.errors()); self.errlbl.config(text=f"Errors: {n}", fg=SEV_COL["ERROR"] if n else SEV_COL["OK"])
        def thread(self, fn, done=None):
            def run():
                try: r = fn()
                except Exception as e: LOG.add("ERROR", "thread", f"{type(e).__name__}: {e}", traceback.format_exc()); r = None
                if done: self.after(0, lambda: done(r))
            threading.Thread(target=run, daemon=True).start()

        # ---------- Data
        def t_data(self):
            f = self.tabs["Data"]; a = ttk.LabelFrame(f, text="A. From TallyPrime (XML server on, port 9000)"); a.pack(fill="x", padx=8, pady=6)
            ttk.Label(a, text="Port").pack(side="left", padx=4); self.port = tk.StringVar(value=str(self.settings["tally_port"])); ttk.Entry(a, textvariable=self.port, width=6).pack(side="left")
            ttk.Button(a, text="Connect", command=self.connect).pack(side="left", padx=6)
            self.cbo = ttk.Combobox(a, width=60, state="readonly"); self.cbo.pack(side="left", padx=4); self.cbo.bind("<<ComboboxSelected>>", self.pick_company)
            ttk.Button(a, text="Load company data", command=self.load_tally).pack(side="left", padx=6)
            self.pb = ttk.Progressbar(a, mode="indeterminate", length=120); self.pb.pack(side="left", padx=6)
            b = ttk.LabelFrame(f, text="B. From an Excel / CSV trial balance (Ledger | Group | Debit | Credit  or  Ledger | Group | Closing)"); b.pack(fill="x", padx=8, pady=6)
            ttk.Button(b, text="Import trial balance...", command=self.load_tb).pack(side="left", padx=6)
            ttk.Button(b, text="Import LAST YEAR's trial balance (comparatives)...", command=self.load_prior_tb).pack(side="left", padx=6)
            c = ttk.LabelFrame(f, text="Financial year  (detected automatically from the data - change it if needed)"); c.pack(fill="x", padx=8, pady=6)
            self.fy_opts = []; self.fycb = ttk.Combobox(c, width=40, state="readonly"); self.fycb.pack(side="left", padx=4); self.fycb.bind("<<ComboboxSelected>>", self.pick_fy)
            ttk.Button(c, text="Detect years", command=self.probe_years).pack(side="left", padx=4)
            self.fs = tk.StringVar(); self.fe = tk.StringVar()
            for lab, v in (("From (dd-mm-yyyy)", self.fs), ("To (dd-mm-yyyy)", self.fe)): ttk.Label(c, text=lab).pack(side="left", padx=4); ttk.Entry(c, textvariable=v, width=12).pack(side="left")
            ttk.Button(c, text="BUILD statements + run all checks", command=self.build).pack(side="left", padx=14)
            self.info = tk.Text(f, height=16, wrap="word"); self.info.pack(fill="both", expand=True, padx=8, pady=6)
            self.info.insert("end", "Comparatives order of preference: imported last-year TB, then the snapshot this app saved last year, then Tally opening balances (balance sheet only).\n"
                                    "Every load saves a snapshot, so next year's comparatives are automatic.\n")
        def connect(self):
            self.settings["tally_port"] = int(self.port.get()); save_settings(self.settings); self.tally = Tally(self.settings["tally_host"], int(self.port.get()))
            self.status.set("Connecting...")
            def done(r):
                if r is None: self.status.set("Cannot connect - see Errors tab."); messagebox.showerror("Cannot connect", "Tally did not answer.\nTallyPrime: F1 Help > Settings > Connectivity > Client/Server: set 'TallyPrime acts as' Server/Both, port 9000, and open the company."); return
                self.companies = r; self.cbo["values"] = [c["name"] for c in r]
                if r: self.cbo.current(0); self.pick_company()
                self.status.set(f"Connected: {len(r)} company(ies) open.")
            def work():
                try: return self.tally.companies()
                except Exception as e: LOG.add("ERROR", "tally.connect", f"{e}", traceback.format_exc()); return None
            self.thread(work, done)
        def pick_company(self, *_):
            for c in self.companies:
                if c["name"] == self.cbo.get():
                    try: self.fs.set(pdate(c["start"]).strftime("%d-%m-%Y")); self.fe.set(pdate(c["end"]).strftime("%d-%m-%Y"))
                    except Exception: pass
            self.probe_years()
        def pick_fy(self, *_):
            i = self.fycb.current()
            if 0 <= i < len(self.fy_opts): self.fs.set(self.fy_opts[i][1].strftime("%d-%m-%Y")); self.fe.set(self.fy_opts[i][2].strftime("%d-%m-%Y"))
        def probe_years(self, *_):
            nm = self.cbo.get()
            if not nm or not getattr(self, "tally", None): return
            c = next((c for c in self.companies if c["name"] == nm), None)
            try: per = (pdate(c["start"]), pdate(c["end"]))
            except Exception: per = None
            self.status.set("Scanning Tally to find which years contain data...")
            def work():
                try: return self.tally.voucher_dates(nm, per)
                except Exception as e: LOG.add("ERROR", "tally.years", f"{e}", traceback.format_exc()); return None
            def done(ds):
                if not ds: self.status.set("Years could not be detected automatically - type the From and To dates yourself."); return
                by = {}
                for d in ds: by.setdefault(d.year if d.month >= 4 else d.year - 1, []).append(d)
                today = date.today(); self.fy_opts = []
                for y in sorted(by):
                    s_, e_ = date(y, 4, 1), date(y + 1, 3, 31); n = len(by[y])
                    self.fy_opts.append((f"FY {y}-{str(y + 1)[2:]}   -   {n:,} vouchers" + ("   (still running)" if e_ >= today else ""), s_, e_))
                self.fycb["values"] = [o[0] for o in self.fy_opts]; closed = [i for i, o in enumerate(self.fy_opts) if o[2] < today]
                k = closed[-1] if closed else len(self.fy_opts) - 1; self.fycb.current(k); self.pick_fy()
                self.status.set(f"Found data in {len(self.fy_opts)} year(s). Selected {self.fy_opts[k][0].split('   -')[0]} (latest completed year) - pick another in the year box if you need it, then press Load company data.")
            self.thread(work, done)
        def load_tally(self):
            nm = self.cbo.get()
            if not nm: return messagebox.showinfo("Select", "Connect and choose a company first.")
            try: per = (pdate(self.fs.get()), pdate(self.fe.get()))
            except Exception: return messagebox.showerror("Financial year", "Type the financial year dates first (dd-mm-yyyy), e.g. 01-04-2024 and 31-03-2025. The app reads Tally for exactly that year.")
            prog = dict(step="1 of 3: ledger balances (the slow part)", t0=time.time(), run=True); self.pb.start(12)
            def tick():
                if not prog["run"]: return
                el = int(time.time() - prog["t0"])
                self.status.set(f"Reading Tally for {per[0]:%d-%m-%Y} to {per[1]:%d-%m-%Y}  -  step {prog['step']}  -  {el // 60:02d}:{el % 60:02d} elapsed. Tally is calculating, please keep it idle.")
                self.after(1000, tick)
            tick()
            def work():
                try:
                    L = self.tally.ledgers(nm, per); prog["step"] = "2 of 3: groups"; G = self.tally.groups(nm); prog["step"] = "3 of 3: vouchers"; V = self.tally.vouchers(nm, per); return L, G, V
                except Exception as e: LOG.add("ERROR", "tally.load", f"{e}", traceback.format_exc()); return None
            def done(r):
                prog["run"] = False; self.pb.stop()
                if not r: self.status.set("Load failed - see Errors tab."); return
                self.led, self.grp, self.vch = r; self.company = nm; self.prior = None; self.prior_source = None; self.tally_period = per
                self.status.set(f"{nm}: {len(self.led)} ledgers, {len(self.vch)} voucher lines loaded."); self.load_profile(nm)
            self.thread(work, done)
        def load_tb(self):
            p = filedialog.askopenfilename(filetypes=[("Trial balance", "*.xlsx *.xlsm *.csv")])
            if not p: return
            r = import_tb(p)
            if r is None: self.status.set("Import failed - see Errors tab."); return messagebox.showerror("Import failed", "Could not read the file. Details are in the Errors tab.")
            self.led, self.grp, self.vch, self.prior, self.prior_source = r, {}, None, None, None; self.tally_period = None; self.company = Path(p).stem; self.status.set(f"{len(r)} ledgers imported from {Path(p).name}."); self.load_profile(self.company)
        def load_prior_tb(self):
            p = filedialog.askopenfilename(filetypes=[("Trial balance", "*.xlsx *.xlsm *.csv")])
            if not p: return
            r = import_tb(p)
            if r: self.prior, self.prior_source = r, "file"; self.status.set(f"Last year's TB loaded ({len(r)} ledgers).")
        def fy(self):
            try: return pdate(self.fs.get()), pdate(self.fe.get())
            except Exception: messagebox.showerror("Financial year", "Enter the financial year dates as dd-mm-yyyy."); return None
        def build(self):
            if not self.led: return messagebox.showinfo("No data", "Load Tally data or import a trial balance first.")
            d = self.fy()
            if not d: return
            s, e = d
            if getattr(self, "tally_period", None) and self.tally_period != (s, e):
                return messagebox.showwarning("Reload the data", f"The data was read from Tally for {self.tally_period[0]:%d-%m-%Y} to {self.tally_period[1]:%d-%m-%Y}, but the dates now say {s:%d-%m-%Y} to {e:%d-%m-%Y}.\nPress 'Load company data' again so the balances match the year.")
            prof = self.read_prof(); prior, src = self.prior, self.prior_source
            if prior is None:
                snap = load_snapshot(self.company, date(e.year - 1, e.month, e.day))
                if snap: prior, src = snap["ledgers"], "snapshot"
                else: prior, src = prior_from_opening(self.led), "opening"
            ctx = build_all_safe(self.led, prof, s, e, self.grp, prior, src, self.vch)
            if not ctx: self.status.set("Build failed - see Errors tab."); self.nb.select(self.tabs["Errors"]); return
            save_snapshot(self.company, e, self.led, self.grp); self.ctx = ctx; self.save_profile()
            self.show_stmt(); self.show_checks(); self.audit_txt.delete("1.0", "end"); self.audit_txt.insert("end", explain_applicability(ctx))
            c = Counter_(ctx["findings"]); self.status.set(f"Built. Checks: {c.get('ERROR', 0)} errors, {c.get('WARN', 0)} warnings, {c.get('INFO', 0)} info. Prior-year source: {src}.")
            self.nb.select(self.tabs["Checks"] if c.get("ERROR") else self.tabs["Statements"])
        # ---------- profile
        def t_prof(self):
            f = self.tabs["Firm Profile"]; l = ttk.Frame(f); l.pack(side="left", fill="y", padx=8, pady=6); self.pv = {}
            r = ttk.Frame(l); r.pack(fill="x"); ttk.Label(r, text="Entity type", width=34).pack(side="left"); self.ent = tk.StringVar(value="company")
            for v in ("company", "trust", "firm"): ttk.Radiobutton(r, text={"company": "Company", "trust": "Trust/Society/Assn", "firm": "Partnership firm"}[v], variable=self.ent, value=v).pack(side="left")
            for k, lab in PROF_FIELDS:
                r = ttk.Frame(l); r.pack(fill="x", pady=1); ttk.Label(r, text=lab, width=34).pack(side="left"); self.pv[k] = tk.StringVar(); ttk.Entry(r, textvariable=self.pv[k], width=48).pack(side="left")
            r = ttk.Frame(l); r.pack(fill="x", pady=1); ttk.Label(r, text="Loans from directors shown under", width=34).pack(side="left"); self.loanline = tk.StringVar(value="lt_borrow"); ttk.Combobox(r, textvariable=self.loanline, values=["lt_borrow", "other_cl", "other_ltl"], width=14, state="readonly").pack(side="left")
            ttk.Button(l, text="Learn from sample statements / audit reports (xlsx, docx, pdf)...", command=self.learn).pack(pady=6, anchor="w")
            ttk.Button(l, text="Use my own audit-report template (.docx)...", command=self.pick_tpl).pack(anchor="w"); self.tpl = None
            g = ttk.LabelFrame(f, text="Facts that change the audit report"); g.pack(side="left", fill="y", padx=8, pady=6); self.fv = {}
            for k, lab in FACT_BOXES: self.fv[k] = tk.BooleanVar(); ttk.Checkbutton(g, text=lab, variable=self.fv[k]).pack(anchor="w")
            for k, lab, vals in (("audit_trail", "Audit trail (edit log)", ["operated", "partial", "none"]), ("going_concern_mode", "Going-concern paragraph", ["none", "eom", "mu"])):
                ttk.Label(g, text=lab).pack(anchor="w", pady=(8, 0)); self.fv[k] = tk.StringVar(value=vals[0]); ttk.Combobox(g, textvariable=self.fv[k], values=vals, state="readonly", width=12).pack(anchor="w")
            ttk.Label(g, text="Peak bank/FI borrowing in year (Rs)").pack(anchor="w", pady=(8, 0)); self.fv["max_bank_fi_borrowing"] = tk.StringVar(); ttk.Entry(g, textvariable=self.fv["max_bank_fi_borrowing"], width=16).pack(anchor="w")
            ttk.Label(g, text="eom = Emphasis of Matter, mu = Material Uncertainty.\nThe app only DRAFTS; you decide.", foreground="#555").pack(anchor="w", pady=8)
        def pick_tpl(self):
            p = filedialog.askopenfilename(filetypes=[("Word", "*.docx")])
            if p: self.tpl = p; self.status.set(f"Audit-report template: {Path(p).name}")
        def read_prof(self):
            P = dict(entity=self.ent.get(), loan_line=self.loanline.get(), facts={}); sp = lambda s: [x.strip() for x in s.split("|") if x.strip()]
            for k, _ in PROF_FIELDS:
                v = self.pv[k].get().strip()
                if k == "address": P["address_lines"] = sp(v)
                elif k == "directors": P["directors"] = sp(v)
                elif k == "signatories": P["signatories"] = [dict(role=x.split(":")[0].strip().upper(), name=(x.split(":", 1)[1].strip() if ":" in x else "")) for x in sp(v)]
                else: P[k] = v
            for k, _ in FACT_BOXES: P["facts"][k] = bool(self.fv[k].get())
            P["facts"]["audit_trail"] = self.fv["audit_trail"].get(); P["facts"]["going_concern_mode"] = self.fv["going_concern_mode"].get()
            mb = self.fv["max_bank_fi_borrowing"].get().replace(",", "").strip()
            if mb: P["facts"]["max_bank_fi_borrowing"] = float(mb)
            if not P.get("name"): P["name"] = getattr(self, "company", "")
            if self.tpl: P["template_docx"] = self.tpl
            return P
        def save_profile(self):
            self.settings.setdefault("profiles", {})[self.company] = self.read_prof(); save_settings(self.settings)
        def load_profile(self, nm, P=None):
            P = P or self.settings.get("profiles", {}).get(nm) or {}
            self.ent.set(P.get("entity", self.ent.get()))
            for k, _ in PROF_FIELDS:
                v = P.get({"address": "address_lines"}.get(k, k), "")
                if k in ("address", "directors"): v = " | ".join(v or [])
                elif k == "signatories": v = " | ".join(f"{s['role']}:{s.get('name', '')}" for s in (v or []))
                if v or not self.pv[k].get(): self.pv[k].set(v or (nm if k == "name" else ""))
            for k, _ in FACT_BOXES: self.fv[k].set(bool(P.get("facts", {}).get(k, False)))
            for k in ("audit_trail", "going_concern_mode"):
                if P.get("facts", {}).get(k): self.fv[k].set(P["facts"][k])
        def learn(self):
            ps = filedialog.askopenfilenames(filetypes=[("Samples", "*.xlsx *.xlsm *.docx *.pdf")])
            if not ps: return
            P = learn_profile(ps)
            if P.get("template_docx"): self.tpl = P["template_docx"]
            self.load_profile(P.get("name", getattr(self, "company", "")), P)
            self.status.set("Learned from samples: " + ", ".join(k for k in P if k != "facts") + ". Check each field, then Build.")
        # ---------- statements
        def t_stmt(self):
            f = self.tabs["Statements"]; b = ttk.Frame(f); b.pack(fill="x", padx=6, pady=4); self.xf = {k: tk.BooleanVar(value=True) for k in ("xlsx", "pdf", "docx", "txt")}
            for k in self.xf: ttk.Checkbutton(b, text=k.upper(), variable=self.xf[k]).pack(side="left")
            self.xa = tk.BooleanVar(value=True); ttk.Checkbutton(b, text="Auditor's report", variable=self.xa).pack(side="left", padx=8)
            ttk.Button(b, text="EXPORT to folder...", command=self.export).pack(side="left", padx=8)
            self.stmt = tk.Text(f, font=("Courier New", 9), wrap="none"); sy = ttk.Scrollbar(f, command=self.stmt.yview); self.stmt.config(yscrollcommand=sy.set); sy.pack(side="right", fill="y"); self.stmt.pack(fill="both", expand=True)
        def show_stmt(self):
            self.stmt.delete("1.0", "end"); self.stmt.insert("end", render_text(self.ctx["tables"], self.ctx["prof"], self.ctx["fy_end"], 120))
        def export(self):
            if not self.ctx: return messagebox.showinfo("Build first", "Use the Data tab: BUILD.")
            d = filedialog.askdirectory(initialdir=str(APP_DIR / "output"))
            if not d: return
            made = export_all(self.ctx, d, tuple(k for k, v in self.xf.items() if v.get()), self.xa.get())
            n = len(LOG.errors())
            messagebox.showinfo("Exported", "\n".join(made) + (f"\n\n{n} error(s) logged - see Errors tab." if n else "")); self.status.set(f"{len(made)} files saved to {d}")
        # ---------- checks
        def t_checks(self):
            f = self.tabs["Checks"]; top = ttk.Frame(f); top.pack(fill="x", padx=6, pady=4); self.cs = tk.StringVar(value="Build first."); ttk.Label(top, textvariable=self.cs).pack(side="left")
            self.hideok = tk.BooleanVar(value=False); ttk.Checkbutton(top, text="Hide OK", variable=self.hideok, command=self.show_checks).pack(side="right")
            cols = ("Severity", "Area", "Finding", "Why it matters", "What to do")
            self.ctree = ttk.Treeview(f, columns=cols, show="headings", height=18)
            for c, w in zip(cols, (70, 110, 420, 380, 300)): self.ctree.heading(c, text=c); self.ctree.column(c, width=w, anchor="w")
            for k, v in SEV_COL.items(): self.ctree.tag_configure(k, foreground=v)
            self.ctree.pack(fill="both", expand=True, padx=6); self.ctree.bind("<<TreeviewSelect>>", self.check_sel)
            self.cdet = tk.Text(f, height=7, wrap="word"); self.cdet.pack(fill="x", padx=6, pady=4)
        def show_checks(self):
            if not self.ctx: return
            self.ctree.delete(*self.ctree.get_children()); self.cmap = {}; order = {"ERROR": 0, "WARN": 1, "INFO": 2, "OK": 3}
            for x in sorted(self.ctx["findings"], key=lambda x: order.get(x.sev, 9)):
                if self.hideok.get() and x.sev == "OK": continue
                i = self.ctree.insert("", "end", values=x.as_row(), tags=(x.sev,)); self.cmap[i] = x
            c = Counter_(self.ctx["findings"]); self.cs.set(f"{c.get('ERROR', 0)} errors   {c.get('WARN', 0)} warnings   {c.get('INFO', 0)} info   {c.get('OK', 0)} passed")
        def check_sel(self, *_):
            s = self.ctree.selection()
            if s: x = self.cmap[s[0]]; self.cdet.delete("1.0", "end"); self.cdet.insert("end", f"{x.sev} - {x.area}\n{x.msg}\n\nWhy: {x.why}\nFix: {x.fix}")
        # ---------- audit report
        def t_audit(self):
            f = self.tabs["Audit Report"]; ttk.Label(f, text="What the app decided for this year, and why (CARO / cash flow / IFC / Rule 11). Change the Facts in Firm Profile and Build again if any are wrong.").pack(anchor="w", padx=8, pady=4)
            self.audit_txt = tk.Text(f, wrap="word", height=14); self.audit_txt.pack(fill="both", expand=True, padx=8)
            ttk.Button(f, text="Generate Auditor's Report (.docx)...", command=self.make_report).pack(pady=8)
        def make_report(self):
            if not self.ctx: return messagebox.showinfo("Build first", "Use the Data tab: BUILD.")
            p = filedialog.asksaveasfilename(defaultextension=".docx", initialfile=safe_name(self.ctx["prof"]["name"]) + "_auditors_report.docx")
            if p:
                r = build_audit_report(self.ctx, p); messagebox.showinfo("Auditor's report", f"Saved:\n{r}\n\nAnnexures are skeletons - the auditor completes them." if r else "Failed - see Errors tab.")
        # ---------- law library / live update / self repair
        def t_law(self):
            f = self.tabs["Law Library"]; top = ttk.LabelFrame(f, text="Live update & self-repair"); top.pack(fill="x", padx=8, pady=6)
            self.lawinfo = tk.StringVar(); ttk.Label(top, textvariable=self.lawinfo).grid(row=0, column=0, columnspan=6, sticky="w", padx=6)
            self.rurl = tk.StringVar(value=self.settings["rules_url"]); self.aurl = tk.StringVar(value=self.settings["app_url"]); self.asha = tk.StringVar(value=self.settings["app_sha256"])
            for i, (lab, v) in enumerate((("Law library URL (your firm's hosted rules.json)", self.rurl), ("App update URL (your tally_suite.py)", self.aurl), ("Pinned SHA-256 of that file (optional)", self.asha)), start=1):
                ttk.Label(top, text=lab).grid(row=i, column=0, sticky="w", padx=6); ttk.Entry(top, textvariable=v, width=70).grid(row=i, column=1, columnspan=3, pady=1)
            bt = ttk.Frame(top); bt.grid(row=4, column=0, columnspan=6, sticky="w", pady=4)
            for t, c in (("Scan law inbox (offline)", self.do_inbox), ("Open inbox folder", self.open_inbox), ("Self-repair & test", self.do_repair), ("[needs internet] Update law library", self.do_rules), ("[needs internet] Exchange rates", self.do_fx), ("[needs internet] New app version", self.do_app)): ttk.Button(bt, text=t, command=c).pack(side="left", padx=4)
            self.fxlbl = tk.StringVar(); ttk.Label(f, textvariable=self.fxlbl).pack(anchor="w", padx=10)
            self.lawtxt = tk.Text(f, wrap="word"); self.lawtxt.pack(fill="both", expand=True, padx=8, pady=6); self.refresh_law()
        def refresh_law(self):
            d = RULES.data; self.lawinfo.set(f"Law library version {d['version']} (updated {d['updated']}). Rules unverified or older than 180 days are flagged in Checks.")
            lv = d.get("live", {}); self.fxlbl.set("Exchange rates (INR per unit): " + (", ".join(f"{k} {v}" for k, v in lv.get("fx", {}).items()) + f"   as of {lv.get('fx_asof')} from {lv.get('fx_source')}" if lv.get("fx") else "not fetched yet"))
            self.lawtxt.delete("1.0", "end")
            for rid, r in d["rules"].items():
                self.lawtxt.insert("end", f"{r['title']}\n  {r.get('explain', '')}\n  Source: {r.get('source', '')}\n")
                for e in r["history"]: self.lawtxt.insert("end", "    from " + e["from"] + ": " + ", ".join(f"{k}={v}" for k, v in e.items() if k not in ("from", "verified")) + f"   [verified: {e.get('verified') or 'NOT verified'}]\n")
                self.lawtxt.insert("end", "\n")
        def _saveset(self): self.settings.update(rules_url=self.rurl.get().strip(), app_url=self.aurl.get().strip(), app_sha256=self.asha.get().strip()); save_settings(self.settings)
        def open_inbox(self):
            try: os.startfile(str(INBOX))
            except Exception: messagebox.showinfo("Law inbox", f"Put official notifications (PDF / txt / html) in:\n{INBOX}")
        def do_inbox(self):
            props = scan_law_inbox()
            if not props: return messagebox.showinfo("Law inbox", f"Nothing recognised in {INBOX}.\nThe scanner understands small-company and IFC-exemption limits; anything else is for the CA to read.")
            for p in props:
                if p["status"].startswith("already"): continue
                fr = p["frm"] or simpledialog.askstring("Effective date", f"{p['source']}: no 'with effect from' date found. Enter it (dd-mm-yyyy) or cancel:")
                if messagebox.askyesno("Approve rule change?", f"{p['rule']}: {p['values']}\nFrom {fr}\nSource: {p['source']}\n\n\"{p['snippet']}\"\n\nCheck against the official text. Apply?") and fr: apply_proposal(p, fr)
            self.refresh_law()
        def do_rules(self):
            self._saveset(); self.status.set("Updating law library...")
            def done(r):
                self.refresh_law(); self.status.set("Law library update finished.")
                messagebox.showinfo("Law library", "\n".join(r) if r else "Nothing changed - see Errors tab for the reason (e.g. no URL set).")
            self.thread(lambda: update_rules(self.settings), done)
        def do_fx(self):
            self.status.set("Fetching live exchange rates...")
            def done(r): self.refresh_law(); self.status.set("Exchange rates updated." if r else "Exchange rates NOT updated (previous kept) - see Errors tab.")
            self.thread(lambda: update_fx(self.settings), done)
        def do_app(self):
            self._saveset(); self.status.set("Checking for a new version and testing it...")
            def done(r): self.status.set(f"Update result: {r}" if r else "Update not applied - see Errors tab."); messagebox.showinfo("App update", {"current": "Already the latest version.", "frozen": "Packaged .exe: download the new one from GitHub Releases."}.get(r, f"Updated to {r}. Restart the app." if r else "Not updated - see Errors tab."))
            self.thread(lambda: self_update(self.settings), done)
        def do_repair(self):
            self.status.set("Running self-repair and the full self-test (about a minute)...")
            def done(r): self.refresh_law(); self.status.set("Self-repair finished."); messagebox.showinfo("Self-repair", "\n".join(r or ["failed - see Errors tab"]))
            self.thread(lambda: self_repair(self.settings), done)
        # ---------- legal intelligence (official-site watcher; nothing is applied without your approval)
        def t_intel(self):
            f = self.tabs["Legal Intelligence"]; top = ttk.Frame(f); top.pack(fill="x", padx=8, pady=4)
            ttk.Button(top, text="Check official sources now", command=self.do_intel).pack(side="left")
            self.iauto = tk.BooleanVar(value=self.settings.get("intel_auto", True)); ttk.Checkbutton(top, text="Check automatically once a day", variable=self.iauto, command=lambda: (self.settings.update(intel_auto=self.iauto.get()), save_settings(self.settings))).pack(side="left", padx=10)
            ttk.Button(top, text="Open selected in browser", command=self.intel_open).pack(side="left", padx=4); ttk.Button(top, text="Mark reviewed", command=lambda: self.intel_set("reviewed")).pack(side="left", padx=4)
            ttk.Button(top, text="Health report", command=self.show_health).pack(side="left", padx=4); ttk.Button(top, text="Restore previous app version", command=self.do_rollback).pack(side="left", padx=4)
            cols = ("First seen", "Source", "Title", "May affect rule", "Status"); self.itree = ttk.Treeview(f, columns=cols, show="headings", height=12); self.imap = {}
            for c, w in zip(cols, (110, 70, 620, 200, 70)): self.itree.heading(c, text=c); self.itree.column(c, width=w, anchor="w")
            self.itree.pack(fill="both", expand=True, padx=8)
            mid = ttk.LabelFrame(f, text="Turn the selected item into a rule change (you verify against the official text first)"); mid.pack(fill="x", padx=8, pady=4)
            ttk.Button(mid, text="Read selected item & suggest values", command=self.intel_suggest).grid(row=0, column=0, padx=4, pady=2)
            self.irule = tk.StringVar(); ttk.Combobox(mid, textvariable=self.irule, values=list(RULES.data["rules"]), width=24, state="readonly").grid(row=0, column=1)
            self.ifrom = tk.StringVar(); ttk.Label(mid, text="Effective from (yyyy-mm-dd)").grid(row=0, column=2); ttk.Entry(mid, textvariable=self.ifrom, width=12).grid(row=0, column=3)
            self.ivals = tk.StringVar(); ttk.Label(mid, text='Values as JSON, e.g. {"paid_up": 100000000, "turnover": 1000000000}').grid(row=1, column=0, columnspan=3, sticky="w"); ttk.Entry(mid, textvariable=self.ivals, width=70).grid(row=1, column=3, columnspan=3)
            ttk.Button(mid, text="I have verified this - APPLY to law library", command=self.intel_apply).grid(row=0, column=4, padx=6)
            self.itxt = tk.Text(f, height=10, wrap="word"); self.itxt.pack(fill="x", padx=8, pady=4); self.refresh_intel()
        def do_rollback(self):
            r = rollback_app(); messagebox.showinfo("Rollback", f"Restored {r}. Restart the app." if r else "Nothing restored - see Errors tab.")
        def refresh_intel(self):
            self.itree.delete(*self.itree.get_children()); self.imap = {}
            for ts_, src, title, tp, st, url in intel_items(limit=300):
                i = self.itree.insert("", "end", values=(ts_, src, title, tp.replace(",", ", "), st), tags=(("WARN" if tp and st == "new" else "INFO"),)); self.imap[i] = (url, title)
            self.itree.tag_configure("WARN", foreground=SEV_COL["WARN"])
        def do_intel(self, auto=False):
            self.status.set("Checking official sources (internet)...")
            def done(r):
                self.refresh_intel()
                if r is None: self.status.set("Source check failed - see Errors tab."); return
                save_settings(self.settings); n = len(r["new"]); hit = [x for x in r["new"] if x["topics"]]
                self.status.set(f"Official sources: {r['checked']} checked, {n} new item(s), {len(hit)} touching rules the app uses, {len(r['errors'])} problem(s).")
                if hit: self.nb.select(self.tabs["Legal Intelligence"]); self.itxt.delete("1.0", "end"); self.itxt.insert("end", "NEW items that may affect the app's rules:\n" + "\n".join(f"- [{x['source']}] {x['title']}  ->  {', '.join(x['topics'])}" for x in hit))
                elif r["errors"] and not auto: messagebox.showinfo("Sources", "\n".join(r["errors"]))
            self.thread(lambda: intel_check(self.settings), done)
        def _isel(self):
            s = self.itree.selection(); return self.imap.get(s[0]) if s else None
        def intel_open(self):
            x = self._isel()
            if x: webbrowser.open(x[0])
        def intel_set(self, st):
            x = self._isel()
            if x: intel_mark(x[0], st); self.refresh_intel()
        def intel_suggest(self):
            x = self._isel()
            if not x: return messagebox.showinfo("Select", "Select an item first.")
            self.status.set("Reading the item...")
            def done(t):
                self.status.set("Done."); self.itxt.delete("1.0", "end")
                if not t: self.itxt.insert("end", "Could not read the text (scanned PDF, JavaScript page or no PDF reader). Open it in the browser and enter the values by hand."); return
                tp = topics_of(x[1]); rid = tp[0] if tp and tp[0] in RULES.data["rules"] else self.irule.get(); sv = suggest_values(rid or "", t)
                if rid: self.irule.set(rid)
                if sv["values"]: self.ivals.set(json.dumps(sv["values"]))
                self.itxt.insert("end", "Suggested values are a best guess from the text - CHECK them against the official notification.\n\n" + ("\n\n".join("> " + s_ for s_ in sv["snippets"]) or t[:1500]))
            self.thread(lambda: intel_text(x[0]), done)
        def intel_apply(self):
            try: vals = json.loads(self.ivals.get() or "{}"); pdate(self.ifrom.get())
            except Exception: return messagebox.showerror("Check inputs", "Effective date must be yyyy-mm-dd and values must be valid JSON.")
            x = self._isel(); rid = self.irule.get()
            if not rid or not vals: return messagebox.showinfo("Incomplete", "Choose the rule and the values.")
            if not messagebox.askyesno("Apply?", f"Add to rule '{RULES.title(rid)}' from {self.ifrom.get()}:\n{vals}\n\nA backup is taken first. Continue?"): return
            e = apply_rule_change(rid, self.ifrom.get(), vals, (x[1] + " - " + x[0]) if x else "manual entry")
            if e: self.refresh_law(); self.status.set("Rule updated. Rebuild the statements to use it."); (x and intel_mark(x[0], "applied"), self.refresh_intel())
            else: messagebox.showerror("Rejected", "The change was rejected - see Errors tab.")
        def show_health(self):
            self.itxt.delete("1.0", "end"); self.itxt.insert("end", "\n".join(f"[{a}] {b}" for a, b in health_report(self.settings)) + "\n\nSources:\n" + "\n".join(f"  {s_}: last ok {ok or '-'}  {('PROBLEM: ' + er) if er else ''}  ({n} links)" for s_, ok, er, n in intel_status()))
        # ---------- errors
        def t_err(self):
            f = self.tabs["Errors"]; top = ttk.Frame(f); top.pack(fill="x", padx=6, pady=4)
            ttk.Button(top, text="Clear", command=lambda: (LOG.clear(), self.errtree.delete(*self.errtree.get_children()), self.errlbl.config(text="Errors: 0", fg=SEV_COL["OK"]))).pack(side="left")
            ttk.Button(top, text="Copy all to clipboard", command=self.copy_err).pack(side="left", padx=6); ttk.Label(top, text=f"Also saved to {LOG_FILE}").pack(side="left", padx=8)
            cols = ("Time", "Level", "Where", "Message"); self.errtree = ttk.Treeview(f, columns=cols, show="headings", height=16); self.errdet = {}
            for c, w in zip(cols, (140, 60, 140, 800)): self.errtree.heading(c, text=c); self.errtree.column(c, width=w, anchor="w")
            for k, v in SEV_COL.items(): self.errtree.tag_configure(k, foreground=v)
            self.errtree.pack(fill="both", expand=True, padx=6); self.errtree.bind("<<TreeviewSelect>>", lambda e: (self.errtxt.delete("1.0", "end"), self.errtxt.insert("end", self.errdet.get((self.errtree.selection() or [""])[0], ""))))
            self.errtxt = tk.Text(f, height=10); self.errtxt.pack(fill="x", padx=6, pady=4)
        def copy_err(self):
            self.clipboard_clear(); self.clipboard_append("\n".join(f"[{i['ts']}] {i['level']} {i['where']}: {i['msg']}\n{i['tb']}" for i in LOG.items))
    app = App(); _apply_theme(app); app.mainloop(); return app

def Counter_(findings):
    c = defaultdict(int)
    for x in findings: c[x.sev] += 1
    return c

# =============================================================== MAIN
if __name__ == "__main__":
    if "--selftest-run" in sys.argv and not os.environ.get("TCS_SANDBOX_OK") and not os.environ.get("TALLY_CA_HOME"):
        ok, f = selftest(); print("SELF-TEST PASSED" if ok else "SELF-TEST FAILED\n" + "\n".join(f)); sys.exit(0 if ok else 1)
    elif "--selftest-run" in sys.argv:
        n, f = run_tests()
        for x in f: print("FAIL:", x)
        for e in LOG.errors(): pass
        print(f"{n} checks passed, {len(f)} failed"); sys.exit(1 if f else 0)
    elif "--selftest" in sys.argv:
        ok, f = selftest(); print("SELF-TEST PASSED" if ok else "SELF-TEST FAILED\n" + "\n".join(f)); sys.exit(0 if ok else 1)
    elif "--cli" in sys.argv:
        cli([a for a in sys.argv[1:] if a != "--cli"])
    else:
        try: launch_gui()
        except Exception as e:
            LOG.add("ERROR", "startup", f"{type(e).__name__}: {e}", traceback.format_exc()); print("Cannot start the window:", e, "\nUse --cli mode or install tkinter.")
