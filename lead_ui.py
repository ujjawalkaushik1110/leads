#!/usr/bin/env python3
"""
Lead Scraper - local web interface.

    python lead_scraper.py ui            # opens http://127.0.0.1:8501

Everything the CLI does, from the browser: set up Ollama and choose the model, connect Google / Reddit,
fill in queries and options (every field has a hover tooltip), watch live progress, stop early, browse
results, download Excel/CSV. Forms are generated from the CLI's own options, so they never drift apart.

Security: binds to 127.0.0.1 only, checks the Host/Origin headers (blocks DNS-rebinding and cross-site
requests), and requires a random per-session token on every API call.
"""
import argparse
import contextlib
import io
import json
import os
import re
import secrets
import shutil
import subprocess
import sys
import threading
import time
import traceback
import webbrowser
from dataclasses import asdict
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

import httpx

import lead_scraper as ls

STATE = {"token": "", "port": 0, "out_dir": Path("."), "job": None, "pull": {"status": "idle"}, "auth": {}}
DEFAULT_SETTINGS = {"ollama_host": "http://127.0.0.1:11434", "ai_mode": "auto", "model": "", "ai_workers": 2}
SUGGESTED_MODELS = [
    ("qwen2.5:3b", "small & fast, ~2 GB download - for 4-8 GB RAM"),
    ("qwen2.5:7b", "best all-round balance, ~4.7 GB - for 8-16 GB RAM"),
    ("llama3.2:3b", "small, quick, ~2 GB"),
    ("llama3.1:8b", "strong general model, ~4.9 GB"),
    ("gemma3:4b", "compact Google model, ~3.3 GB"),
    ("mistral:7b", "reliable 7B, ~4.1 GB"),
    ("qwen2.5:14b", "more accurate, ~9 GB - needs 16 GB+ RAM"),
]
HIDDEN_FIELDS = {"help", "out", "ai", "model", "ollama_host", "ai_workers", "no_autostart", "pull"}   # controlled by the Setup tab
LABELS = {
    "queries": "Industries / niches (one per line)", "ask": "Describe what you want (AI plans the search)", "location": "Location",
    "max_leads": "Max leads", "depth": "Research depth", "icp": "Ideal customer (AI scores against this)", "outreach": "What you sell (AI writes outreach drafts)",
    "min_fit": "Minimum AI fit (0-10)", "ai_top": "Leads the AI should judge", "engines": "Search engines", "no_osm": "Skip OpenStreetMap",
    "no_harvest": "Skip directory/list harvesting", "harvest_limit": "Directory pages to harvest", "maps": "Include Google Maps (needs Google login)",
    "no_web": "Skip web search engines", "maps_max": "Max Google Maps places per query", "per_query": "Results per engine per query",
    "search_region": "Search region code", "include_filter": "Must mention (any of, comma-separated)", "exclude": "Must not mention (comma-separated)",
    "tlds": "Only these TLDs (e.g. in,com)", "exclude_domains": "Ignore these domains", "require_contact": "Only leads with email or phone",
    "min_score": "Minimum lead score", "targets": "Website URL(s) (one per line)", "max_pages": "Max pages per site", "extract": "AI insight fields (comma-separated)",
    "strict_host": "Stay on the exact host (no subdomains)", "no_sitemap": "Don't read sitemap.xml", "export_text": "Save full text of every page",
    "js": "Render JavaScript pages (needs Playwright)", "site_workers": "Parallel pages per site", "subs": "Subreddits (comma-separated, blank = all)",
    "days": "Look back (days)", "offer": "What you sell (AI rates posts by fit)", "max_posts": "Max posts per query", "reddit_user_agent": "Reddit User-Agent",
    "workers": "Parallel workers", "delay": "Delay per host (seconds)", "timeout": "Request timeout (seconds)", "region": "Phone country (ISO code)",
    "user_agent": "User-Agent", "pages_per_site": "Contact/about pages per company", "verify_mx": "Verify email domains (needs dnspython)", "resume": "Resume previous run",
}
PLACEHOLDERS = {"queries": "dental clinics\northodontists", "ask": "small dental clinics in Pune that offer implants", "location": "Mumbai",
                "icp": "Independent dental clinics with 1-5 dentists", "outreach": "website redesign for clinics", "targets": "https://example.com",
                "subs": "india, mumbai", "offer": "dental marketing services", "extract": "pricing plans, team members, refund policy",
                "include_filter": "implants, braces", "exclude": "jobs, university", "tlds": "in,com", "region": "IN"}
GROUPS = {
    "leads": ["Search", "Sources", "Filters", "AI targeting", "Advanced"],
    "crawl": ["Crawl", "Options", "Advanced"],
    "reddit": ["Search", "Advanced"],
}
GROUP_OF = {
    "leads": {**{k: "Search" for k in ("queries", "ask", "location", "max_leads", "depth")},
              **{k: "Sources" for k in ("engines", "no_osm", "no_harvest", "harvest_limit", "maps", "no_web", "maps_max", "per_query", "search_region")},
              **{k: "Filters" for k in ("include_filter", "exclude", "tlds", "exclude_domains", "require_contact", "min_score")},
              **{k: "AI targeting" for k in ("icp", "outreach", "min_fit", "ai_top")}},
    "crawl": {**{k: "Crawl" for k in ("targets", "max_pages", "depth", "extract")},
              **{k: "Options" for k in ("strict_host", "no_sitemap", "export_text", "js")}},
    "reddit": {k: "Search" for k in ("queries", "subs", "days", "offer", "max_posts", "min_score", "ai_top")},
}
MODE_INFO = {
    "leads": ("Find leads", "Search many engines at once (+ OpenStreetMap, Google Maps, directory lists), crawl each company site, extract contacts, score, export."),
    "crawl": ("Crawl a website", "Crawl entire websites: every page, contacts, documents, external links and an on-page issues audit."),
    "reddit": ("Reddit signals", "Find public posts where people ask for recommendations or help (buying intent). Needs a Reddit login."),
}


# =========================================================================== settings / spec
def settings_path():
    return ls.cred_path().parent / "ui_settings.json"


def load_settings():
    try:
        return {**DEFAULT_SETTINGS, **json.loads(settings_path().read_text())}
    except (OSError, json.JSONDecodeError):
        return dict(DEFAULT_SETTINGS)


def save_settings(d):
    cur = load_settings()
    for k in DEFAULT_SETTINGS:
        if k in d:
            cur[k] = d[k]
    if cur["ai_mode"] not in ("auto", "on", "off"):
        cur["ai_mode"] = "auto"
    cur["ai_workers"] = max(1, min(8, int(cur.get("ai_workers") or 2)))
    settings_path().write_text(json.dumps(cur, indent=2))
    return cur


def subparsers():
    ap = ls.build_parser()
    for a in ap._actions:
        if isinstance(a, ls.argparse._SubParsersAction):
            return a.choices
    return {}


def build_spec():
    import argparse
    subs, modes = subparsers(), {}
    for mode in ("leads", "crawl", "reddit"):
        sp, fields = subs[mode], []
        for a in sp._actions:
            if a.dest in HIDDEN_FIELDS:
                continue
            if not a.option_strings:
                kind = "lines"
            elif isinstance(a, argparse._StoreTrueAction):
                kind = "bool"
            elif a.choices:
                kind = "select"
            elif a.type is int:
                kind = "int"
            elif a.type is float:
                kind = "float"
            else:
                kind = "text"
            default = a.default if isinstance(a.default, (str, int, float, bool)) or a.default is None else ""
            f = {"dest": a.dest, "flag": (a.option_strings[-1] if a.option_strings else a.dest), "kind": kind,
                 "label": LABELS.get(a.dest, a.dest.replace("_", " ").capitalize()), "help": a.help or "",
                 "default": default, "choices": list(a.choices) if a.choices else None,
                 "placeholder": PLACEHOLDERS.get(a.dest, ""), "group": GROUP_OF[mode].get(a.dest, "Advanced")}
            if a.dest == "engines":
                f["kind"] = "engines"
            fields.append(f)
        order = {g: i for i, g in enumerate(GROUPS[mode])}
        fields.sort(key=lambda f: order.get(f["group"], 99))
        modes[mode] = {"title": MODE_INFO[mode][0], "description": MODE_INFO[mode][1], "groups": GROUPS[mode], "fields": fields}
    return {"modes": modes, "engines": ls.ALL_ENGINES}


def build_argv(mode, params, out_path):
    import argparse
    sp = subparsers()[mode]
    argv, pos = [mode], []
    for a in sp._actions:
        if a.dest in HIDDEN_FIELDS:
            continue
        v = params.get(a.dest)
        if not a.option_strings:
            lines = v if isinstance(v, list) else str(v or "").splitlines()
            pos += [x.strip() for x in lines if str(x).strip()]
            continue
        flag = a.option_strings[-1]
        if isinstance(a, argparse._StoreTrueAction):
            if v is True:
                argv.append(flag)
            continue
        if v in (None, ""):
            continue
        if a.type is int:
            v = int(float(v))
        elif a.type is float:
            v = float(v)
        else:
            v = str(v).strip()
            if not v:
                continue
        if a.choices and v not in a.choices:
            raise ValueError(f"{a.dest}: '{v}' is not one of {list(a.choices)}")
        argv.append(f"{flag}={v}")
    st = load_settings()
    argv += [f"--ai={st['ai_mode']}", f"--ollama-host={st['ollama_host']}", f"--ai-workers={st['ai_workers']}", f"--out={out_path}"]
    if st.get("model"):
        argv.append(f"--model={st['model']}")
    if pos:
        argv += ["--"] + pos
    return argv


def validate_argv(argv):
    err = io.StringIO()
    try:
        with contextlib.redirect_stderr(err):
            ls.build_parser().parse_args(argv)
    except SystemExit:
        msg = err.getvalue().strip().splitlines()
        raise ValueError(msg[-1] if msg else "invalid options")


# =========================================================================== Ollama helpers
def norm_host(h):
    h = (h or load_settings()["ollama_host"] or DEFAULT_SETTINGS["ollama_host"]).strip()
    if not re.match(r"^https?://", h):
        h = "http://" + h
    return h.rstrip("/")


def ram_gb():
    try:
        return round(os.sysconf("SC_PAGE_SIZE") * os.sysconf("SC_PHYS_PAGES") / 2 ** 30, 1)
    except (ValueError, OSError, AttributeError):
        pass
    try:   # Windows
        import ctypes

        class MS(ctypes.Structure):
            _fields_ = [("l", ctypes.c_ulong), ("m", ctypes.c_ulong), ("total", ctypes.c_ulonglong), ("a", ctypes.c_ulonglong),
                        ("b", ctypes.c_ulonglong), ("c", ctypes.c_ulonglong), ("d", ctypes.c_ulonglong), ("e", ctypes.c_ulonglong), ("f", ctypes.c_ulonglong)]
        ms = MS()
        ms.l = ctypes.sizeof(ms)
        ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(ms))
        return round(ms.total / 2 ** 30, 1)
    except Exception:
        return None


def recommend(ram):
    if not ram:
        return {"model": "qwen2.5:7b", "why": "Could not detect your RAM; qwen2.5:7b is a good default if you have 8 GB or more."}
    pick = "qwen2.5:3b" if ram < 8 else "qwen2.5:7b" if ram < 16 else "qwen2.5:14b" if ram < 32 else "qwen2.5:14b"
    return {"model": pick, "why": f"Your computer has about {ram} GB RAM. A model of this size should run comfortably (a GPU makes it faster)."}


def ollama_status(host):
    ram = ram_gb()
    out = {"host": host, "running": False, "models": [], "binary": bool(shutil.which("ollama")), "ram_gb": ram,
           "recommended": recommend(ram), "suggested": [{"name": n, "note": d} for n, d in SUGGESTED_MODELS]}
    try:
        r = httpx.get(host + "/api/tags", timeout=3)
        r.raise_for_status()
        out["running"] = True
        for m in r.json().get("models", []):
            d = m.get("details") or {}
            out["models"].append({"name": m["name"], "size_gb": round(m.get("size", 0) / 2 ** 30, 1), "params": d.get("parameter_size", ""),
                                  "quant": d.get("quantization_level", ""), "family": d.get("family", ""),
                                  "chat": not any(h in m["name"].lower() for h in ls.EMBED_HINTS)})
    except Exception as e:
        out["error"] = type(e).__name__
    if out["running"]:
        chat = [m["name"] for m in out["models"] if m["chat"]]
        out["auto_pick"] = next((m for p in ls.PREFERRED_MODELS for m in chat if m.lower().startswith(p)), chat[0] if chat else "")
    return out


def start_ollama(host):
    st = ollama_status(host)
    if st["running"]:
        return {"ok": True, "message": "Ollama is already running."}
    if not st["binary"]:
        return {"ok": False, "message": "Ollama is not installed. Download it from https://ollama.com/download, then press Check again."}
    try:
        subprocess.Popen(["ollama", "serve"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, start_new_session=True)
    except OSError as e:
        return {"ok": False, "message": f"Could not start Ollama: {e}"}
    for _ in range(24):
        time.sleep(0.5)
        if ollama_status(host)["running"]:
            return {"ok": True, "message": "Ollama started."}
    return {"ok": False, "message": "Started `ollama serve` but it is not answering yet - wait a few seconds and press Check again."}


def pull_worker(host, model):
    p = STATE["pull"] = {"status": "running", "model": model, "pct": 0, "msg": "starting...", "error": ""}
    try:
        with httpx.stream("POST", host + "/api/pull", json={"name": model, "stream": True}, timeout=None) as r:
            if r.status_code != 200:
                p.update(status="error", error=f"Ollama answered HTTP {r.status_code}")
                return
            for line in r.iter_lines():
                try:
                    j = json.loads(line)
                except json.JSONDecodeError:
                    continue
                if j.get("error"):
                    p.update(status="error", error=str(j["error"])[:200])
                    return
                p["msg"] = j.get("status", "")
                if j.get("total"):
                    p["pct"] = round(100 * j.get("completed", 0) / j["total"], 1)
        p.update(status="done", pct=100, msg="download complete")
    except Exception as e:
        p.update(status="error", error=f"{type(e).__name__}: {e}")


def test_model(host, model):
    t0 = time.time()
    try:
        r = httpx.post(host + "/api/chat", timeout=180, json={
            "model": model, "stream": False, "format": "json", "options": {"temperature": 0},
            "messages": [{"role": "user", "content": 'Reply with JSON only: {"ok": true, "word": "<one short English word>"}'}]})
        r.raise_for_status()
        txt = r.json().get("message", {}).get("content", "")
        parsed = ls.parse_json_loose(txt)
        return {"ok": bool(parsed), "seconds": round(time.time() - t0, 1), "output": txt[:200],
                "message": "Model answered with valid JSON." if parsed else "Model answered but not with valid JSON - try a larger model."}
    except Exception as e:
        return {"ok": False, "seconds": round(time.time() - t0, 1), "output": "", "message": f"{type(e).__name__}: {e}"[:200]}


# =========================================================================== accounts
def auth_status():
    c = ls.load_creds()
    g, r = c.get("google"), c.get("reddit")
    return {
        "google": ({"connected": True, "mode": g.get("mode"), "email": g.get("email", ""), "project": g.get("project_id", ""),
                    "key_tail": (g.get("api_key") or "")[-4:]} if g else {"connected": False}),
        "reddit": ({"connected": True, "mode": r.get("mode"), "username": r.get("username", ""), "client": (r.get("client_id") or "")[:4] + "..."}
                   if r else {"connected": False}),
        "pending": STATE["auth"],
    }


def auth_thread(kind, fn):
    STATE["auth"][kind] = {"state": "waiting", "message": "Complete the sign-in in the browser window that just opened..."}
    try:
        fn()
        STATE["auth"][kind] = {"state": "done", "message": "Connected."}
    except SystemExit as e:
        STATE["auth"][kind] = {"state": "error", "message": str(e.code)[:300]}
    except Exception as e:
        STATE["auth"][kind] = {"state": "error", "message": f"{type(e).__name__}: {e}"[:300]}


# =========================================================================== jobs
class Job:
    def __init__(self, mode, argv, out):
        self.id, self.mode, self.argv, self.out = secrets.token_hex(4), mode, argv, Path(out)
        self.status, self.lines, self.error, self.results, self.files = "running", [], "", None, []
        self.progress, self.cancelled, self.started, self.finished = {"stage": "Starting", "done": 0, "total": 0}, False, time.time(), None
        self.lock = threading.Lock()

    def add_log(self, line):
        with self.lock:
            self.lines.append(line)
            if len(self.lines) > 5000:
                del self.lines[:1000]
            self._progress(line)

    def _progress(self, line):
        p = self.progress
        m = re.match(r"^\[(\d+)/(\d+)\]", line)
        if m:
            p.update(stage="Crawling sites", done=int(m.group(1)), total=int(m.group(2)))
            return
        m = re.search(r"search progress (\d+)/(\d+)", line)
        if m:
            p.update(stage="Searching engines", done=int(m.group(1)), total=int(m.group(2)))
            return
        m = re.match(r"^\s*AI \[(\d+)/(\d+)\]", line)
        if m:
            p.update(stage="AI judging", done=int(m.group(1)), total=int(m.group(2)))
            return
        m = re.search(r"level \d+: \d+ pages fetched \(total (\d+)", line)
        if m:
            p.update(stage="Crawling pages", done=int(m.group(1)), total=0)
            return
        m = re.match(r"^(\d)(b)?/4 (.+?)\.*$", line)
        if m:
            p.update(stage=m.group(3).strip(), done=0, total=0)
        elif line.startswith("Reddit search:"):
            p.update(stage="Searching Reddit", done=0, total=0)
        elif line.startswith("AI:"):
            p.update(stage="AI", done=0, total=0)

    def snapshot(self, since):
        with self.lock:
            return {"id": self.id, "mode": self.mode, "status": self.status, "error": self.error, "progress": dict(self.progress),
                    "lines": self.lines[since:], "next": len(self.lines), "files": self.files, "results": self.results if self.status != "running" else None,
                    "seconds": round((self.finished or time.time()) - self.started, 1), "cancelled": self.cancelled}


def jsonable(mode, res):
    if res is None:
        return None
    if mode == "leads":
        return {"kind": "leads", "rows": [asdict(x) for x in res][:3000]}
    if mode == "reddit":
        return {"kind": "reddit", "rows": [asdict(x) for x in res][:3000]}
    pages, contacts = [], []
    for s in res:
        site = ls.reg_domain(ls.urlparse(s["start"]).netloc)
        for r in s["pages"][:3000]:
            pages.append({"site": site, **{k: r.get(k) for k in ("url", "status", "depth", "title", "h1", "meta", "words", "emails", "phones", "issues", "via", "parent")}})
        for val, c in s["contacts"].items():
            contacts.append({"site": site, "type": c["type"], "value": val, "found_on": c["first"], "pages": c["count"]})
        if s.get("ai"):
            for k, v in s["ai"].items():
                contacts.append({"site": site, "type": "AI: " + str(k), "value": json.dumps(v, ensure_ascii=False)[:600] if not isinstance(v, str) else v[:600],
                                 "found_on": "", "pages": 0})
    return {"kind": "crawl", "rows": pages, "contacts": contacts}


def run_job(job):
    ls.CANCEL.clear()
    ls.LOG_HOOK = job.add_log
    err = io.StringIO()
    try:
        with contextlib.redirect_stderr(err):
            res = ls.main(job.argv)
        job.results = jsonable(job.mode, res)
    except SystemExit as e:
        code = e.code
        job.error = "" if code in (0, None) else (str(code) if not isinstance(code, int) else (err.getvalue().strip().splitlines() or ["The run stopped early."])[-1])
    except Exception as e:
        job.error = f"{type(e).__name__}: {e}"
        job.add_log(traceback.format_exc(limit=3))
    finally:
        ls.LOG_HOOK = None
        stem = job.out.stem
        txt_dir = job.out.parent / f"{stem}_text"
        if txt_dir.is_dir():
            shutil.make_archive(str(txt_dir), "zip", str(txt_dir))
        job.files = sorted(p.name for p in job.out.parent.glob(stem + "*") if p.is_file())
        job.cancelled = ls.CANCEL.is_set()
        job.status = "error" if job.error else ("stopped" if job.cancelled else "done")
        job.finished = time.time()
        ls.CANCEL.clear()


def history():
    out = []
    for p in sorted(STATE["out_dir"].glob("*.xlsx"), key=lambda x: x.stat().st_mtime, reverse=True)[:60]:
        stem = p.stem
        out.append({"name": p.name, "mb": round(p.stat().st_size / 1e6, 2), "when": datetime.fromtimestamp(p.stat().st_mtime).strftime("%Y-%m-%d %H:%M"),
                    "files": sorted(q.name for q in p.parent.glob(stem + "*") if q.is_file())})
    return out


# =========================================================================== HTTP server
MIME = {".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".csv": "text/csv", ".jsonl": "application/x-ndjson",
        ".zip": "application/zip", ".txt": "text/plain"}


class Handler(BaseHTTPRequestHandler):
    server_version = "LeadUI/1.0"

    def log_message(self, *a):
        pass

    # ---- security
    def host_ok(self):
        host = (self.headers.get("Host") or "").lower()
        return host in (f"127.0.0.1:{STATE['port']}", f"localhost:{STATE['port']}")

    def origin_ok(self):
        o = self.headers.get("Origin")
        return not o or urlparse(o).hostname in ("127.0.0.1", "localhost")

    def token_ok(self, q):
        t = self.headers.get("X-UI-Token") or (q.get("t") or [""])[0]
        return bool(t) and secrets.compare_digest(t, STATE["token"])

    # ---- helpers
    def send_json(self, obj, code=200):
        b = json.dumps(obj).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(b)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(b)

    def err(self, msg, code=400):
        self.send_json({"error": msg}, code)

    def body(self):
        n = int(self.headers.get("Content-Length") or 0)
        if n > 1_000_000:
            raise ValueError("request too large")
        return json.loads(self.rfile.read(n) or b"{}")

    # ---- routing
    def do_GET(self):
        u = urlparse(self.path)
        q = parse_qs(u.query)
        if not self.host_ok() or not self.origin_ok():
            return self.err("forbidden host/origin", 403)
        if u.path == "/":
            b = INDEX_HTML.replace("__TOKEN__", STATE["token"]).encode()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(b)))
            self.send_header("Cache-Control", "no-store")
            self.send_header("X-Frame-Options", "DENY")
            self.send_header("Content-Security-Policy", "default-src 'self'; script-src 'unsafe-inline'; style-src 'unsafe-inline'; img-src 'self' data:; connect-src 'self'")
            self.end_headers()
            self.wfile.write(b)
            return
        if not self.token_ok(q):
            return self.err("missing or wrong token", 403)
        try:
            if u.path == "/download":
                return self.download((q.get("f") or [""])[0])
            if u.path == "/api/spec":
                return self.send_json(build_spec())
            if u.path == "/api/settings":
                return self.send_json(load_settings())
            if u.path == "/api/ollama/status":
                return self.send_json(ollama_status(norm_host((q.get("host") or [""])[0])))
            if u.path == "/api/ollama/pull_status":
                return self.send_json(STATE["pull"])
            if u.path == "/api/auth/status":
                return self.send_json(auth_status())
            if u.path == "/api/job":
                job = STATE["job"]
                return self.send_json(job.snapshot(int((q.get("since") or ["0"])[0])) if job else {"status": "none"})
            if u.path == "/api/history":
                return self.send_json({"files": history()})
        except Exception as e:
            return self.err(f"{type(e).__name__}: {e}", 500)
        self.err("not found", 404)

    def do_POST(self):
        u = urlparse(self.path)
        if not self.host_ok() or not self.origin_ok():
            return self.err("forbidden host/origin", 403)
        if not self.token_ok({}):
            return self.err("missing or wrong token", 403)
        try:
            b = self.body()
            return self.post(u.path, b)
        except ValueError as e:
            return self.err(str(e), 400)
        except Exception as e:
            return self.err(f"{type(e).__name__}: {e}", 500)

    def post(self, path, b):
        if path == "/api/settings":
            return self.send_json(save_settings(b))
        if path == "/api/ollama/start":
            return self.send_json(start_ollama(norm_host(b.get("host"))))
        if path == "/api/ollama/pull":
            model = str(b.get("model", "")).strip()
            if not re.match(r"^[A-Za-z0-9._\-/:]{1,100}$", model):
                raise ValueError("invalid model name")
            if STATE["pull"].get("status") == "running":
                raise ValueError("a download is already running")
            threading.Thread(target=pull_worker, args=(norm_host(b.get("host")), model), daemon=True).start()
            return self.send_json({"started": True})
        if path == "/api/ollama/test":
            return self.send_json(test_model(norm_host(b.get("host")), str(b.get("model", "")).strip()))
        if path == "/api/auth/google":
            if b.get("api_key"):
                c = ls.load_creds()
                c["google"] = {"mode": "key", "api_key": str(b["api_key"]).strip()}
                ls.save_creds(c)
                return self.send_json({"started": False, "connected": True})
            raw = b.get("client_json") or ""
            try:
                parsed = json.loads(raw)
                assert (parsed.get("installed") or parsed.get("web") or parsed).get("client_id")
            except Exception:
                raise ValueError("That file is not a valid Google OAuth client_secret.json")
            tmp = ls.cred_path().parent / "client_secret.json"
            tmp.write_text(raw)
            os.chmod(tmp, 0o600)
            threading.Thread(target=auth_thread, args=("google", lambda: ls.google_login(str(tmp), (b.get("project") or "").strip() or None)), daemon=True).start()
            return self.send_json({"started": True})
        if path == "/api/auth/reddit":
            cid = str(b.get("client_id", "")).strip()
            if not cid:
                raise ValueError("Reddit client id is required")
            args = (cid, str(b.get("secret", "")).strip(), str(b.get("username", "")).strip(), 8765, bool(b.get("app_only")))
            threading.Thread(target=auth_thread, args=("reddit", lambda: ls.reddit_login(*args)), daemon=True).start()
            return self.send_json({"started": True})
        if path == "/api/auth/logout":
            c = ls.load_creds()
            for k in (["google", "reddit"] if b.get("target") in (None, "all") else [b["target"]]):
                c.pop(k, None)
            ls.save_creds(c)
            STATE["auth"] = {}
            return self.send_json({"ok": True})
        if path == "/api/run":
            job = STATE["job"]
            if job and job.status == "running":
                raise ValueError("A run is already in progress - stop it first or wait for it to finish.")
            mode = b.get("mode")
            if mode not in ("leads", "crawl", "reddit"):
                raise ValueError("unknown mode")
            out = STATE["out_dir"] / f"{mode}_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            argv = build_argv(mode, b.get("params") or {}, out)
            validate_argv(argv)
            STATE["job"] = job = Job(mode, argv, out)
            threading.Thread(target=run_job, args=(job,), daemon=True).start()
            return self.send_json({"started": True, "id": job.id})
        if path == "/api/cancel":
            job = STATE["job"]
            if job and job.status == "running":
                ls.CANCEL.set()
                job.add_log("Stop requested - finishing in-flight work and exporting what was found...")
            return self.send_json({"ok": True})
        self.err("not found", 404)

    def download(self, name):
        if not re.match(r"^[A-Za-z0-9_.\- ]+$", name or ""):
            return self.err("bad file name", 400)
        p = (STATE["out_dir"] / name).resolve()
        if STATE["out_dir"].resolve() not in p.parents or not p.is_file() or p.suffix.lower() not in MIME:
            return self.err("not found", 404)
        data = p.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", MIME[p.suffix.lower()])
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Content-Disposition", f'attachment; filename="{p.name}"')
        self.end_headers()
        self.wfile.write(data)


def serve(port=8501, out_dir=None, open_browser=True):
    out = Path(out_dir or Path.home() / "lead_scraper_outputs").expanduser()
    out.mkdir(parents=True, exist_ok=True)
    srv = None
    for p in range(port, port + 20):
        try:
            STATE.update(token=secrets.token_urlsafe(24), out_dir=out, port=p)
            srv = ThreadingHTTPServer(("127.0.0.1", p), Handler)
            break
        except OSError:
            continue
    if not srv:
        sys.exit(f"No free port between {port} and {port + 19}.")
    srv.daemon_threads = True
    url = f"http://127.0.0.1:{STATE['port']}/"
    print(f"Lead Scraper web interface: {url}\nResults are saved in: {out}\nPress Ctrl+C to stop.", flush=True)
    if open_browser:
        threading.Timer(0.8, webbrowser.open, [url]).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")
    finally:
        srv.server_close()


# =========================================================================== the page
INDEX_HTML = r"""<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Lead Scraper</title>
<style>
:root{--bg:#f6f7f9;--card:#fff;--ink:#1c2330;--mute:#667085;--line:#e4e7ec;--brand:#1f3864;--accent:#2f6fed;--ok:#12805c;--warn:#b54708;--bad:#b42318;--chip:#eef2f7;--code:#0f172a}
@media (prefers-color-scheme:dark){:root{--bg:#0f1420;--card:#171d2b;--ink:#e6e9f0;--mute:#9aa4b8;--line:#2a3347;--brand:#8fb0ff;--accent:#6c9bff;--ok:#4cc79a;--warn:#f0b35c;--bad:#ff8a80;--chip:#222b3d;--code:#0a0e17}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:14px/1.45 system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif}
header{background:var(--card);border-bottom:1px solid var(--line);position:sticky;top:0;z-index:20}
.bar{max-width:1180px;margin:0 auto;padding:10px 18px;display:flex;align-items:center;gap:18px;flex-wrap:wrap}
.logo{font-weight:700;color:var(--brand);font-size:16px}
nav{display:flex;gap:4px}nav button{background:none;border:0;padding:8px 14px;border-radius:8px;color:var(--mute);font-weight:600;cursor:pointer;font-size:14px}
nav button.on{background:var(--chip);color:var(--ink)}
.chips{margin-left:auto;display:flex;gap:8px;flex-wrap:wrap}.chip{background:var(--chip);border-radius:99px;padding:3px 10px;font-size:12px;color:var(--mute);cursor:pointer}
.chip b{display:inline-block;width:8px;height:8px;border-radius:50%;background:var(--mute);margin-right:6px}.chip.ok b{background:var(--ok)}.chip.bad b{background:var(--bad)}
main{max-width:1180px;margin:0 auto;padding:18px}section{display:none}section.on{display:block}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px 18px;margin-bottom:16px}
.card h2{margin:0 0 4px;font-size:16px}.sub{color:var(--mute);margin:0 0 12px}
.row{display:flex;gap:10px;flex-wrap:wrap;align-items:center}.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:12px 16px}
label{display:block;font-weight:600;font-size:13px;margin-bottom:4px}
input[type=text],input[type=number],input[type=password],select,textarea{width:100%;padding:8px 10px;border:1px solid var(--line);border-radius:8px;background:var(--bg);color:var(--ink);font:inherit}
textarea{min-height:84px;resize:vertical}input:focus,select:focus,textarea:focus{outline:2px solid var(--accent);outline-offset:0}
.chk{display:flex;gap:8px;align-items:flex-start;padding:6px 0}.chk input{margin-top:3px}.chk label{margin:0;font-weight:500}
button.btn{background:var(--accent);color:#fff;border:0;border-radius:8px;padding:9px 16px;font-weight:600;cursor:pointer;font-size:14px}
button.btn.sec{background:var(--chip);color:var(--ink)}button.btn.bad{background:var(--bad)}button.btn:disabled{opacity:.5;cursor:not-allowed}
.note{font-size:12.5px;color:var(--mute)}.msg{padding:8px 12px;border-radius:8px;background:var(--chip);margin-top:10px;font-size:13px}.msg.ok{color:var(--ok)}.msg.bad{color:var(--bad)}.msg.warn{color:var(--warn)}
.tip{display:inline-block;width:15px;height:15px;line-height:15px;text-align:center;border-radius:50%;background:var(--chip);color:var(--mute);font-size:10px;font-weight:700;margin-left:6px;position:relative;cursor:help;vertical-align:middle}
.tip:hover::after,.tip:focus::after{content:attr(data-tip);position:absolute;left:50%;bottom:130%;transform:translateX(-30%);width:280px;background:var(--code);color:#e8ecf5;padding:8px 10px;border-radius:8px;font-size:12px;font-weight:400;line-height:1.4;z-index:50;white-space:normal;text-align:left;box-shadow:0 6px 20px rgba(0,0,0,.25)}
fieldset{border:1px solid var(--line);border-radius:10px;margin:0 0 14px;padding:10px 14px 14px}legend{font-weight:700;padding:0 6px;color:var(--brand)}
details>summary{cursor:pointer;font-weight:700;color:var(--brand);margin-bottom:10px}
.modes{display:flex;gap:8px;margin-bottom:10px;flex-wrap:wrap}.modes button{border:1px solid var(--line);background:var(--card);border-radius:10px;padding:8px 14px;cursor:pointer;font-weight:600;color:var(--ink)}
.modes button.on{border-color:var(--accent);background:var(--chip)}
.prog{height:10px;background:var(--chip);border-radius:99px;overflow:hidden;margin:8px 0}.prog i{display:block;height:100%;width:0;background:var(--accent);transition:width .3s}
.prog.ind i{width:35%;animation:ind 1.2s infinite ease-in-out}@keyframes ind{0%{margin-left:-35%}100%{margin-left:100%}}
pre#log{background:var(--code);color:#d6deeb;border-radius:10px;padding:12px;height:260px;overflow:auto;font:12px/1.5 ui-monospace,Menlo,Consolas,monospace;white-space:pre-wrap;margin:8px 0 0}
.tablewrap{overflow:auto;max-height:520px;border:1px solid var(--line);border-radius:10px}
table{border-collapse:collapse;width:100%;font-size:13px}th,td{padding:7px 10px;border-bottom:1px solid var(--line);text-align:left;white-space:nowrap;max-width:340px;overflow:hidden;text-overflow:ellipsis}
th{position:sticky;top:0;background:var(--card);cursor:pointer;user-select:none}tbody tr:hover{background:var(--chip);cursor:pointer}
a{color:var(--accent)}.pill{display:inline-block;padding:1px 8px;border-radius:99px;background:var(--chip);font-size:12px}.pill.g{color:var(--ok)}.pill.r{color:var(--bad)}
#drawer{position:fixed;right:0;top:0;bottom:0;width:min(460px,100%);background:var(--card);border-left:1px solid var(--line);box-shadow:-8px 0 30px rgba(0,0,0,.18);padding:16px;overflow:auto;transform:translateX(105%);transition:transform .2s;z-index:40}
#drawer.on{transform:none}#drawer dt{font-weight:700;margin-top:10px;font-size:12px;color:var(--mute);text-transform:uppercase;letter-spacing:.03em}#drawer dd{margin:2px 0 0;white-space:pre-wrap;word-break:break-word}
.bar2{display:flex;gap:10px;align-items:center;margin-bottom:10px;flex-wrap:wrap}.bar2 input{max-width:280px}
.empty{color:var(--mute);padding:26px;text-align:center}
</style></head><body>
<header><div class="bar"><span class="logo">Lead Scraper</span>
<nav><button data-tab="setup" class="on">1 &middot; Setup</button><button data-tab="run">2 &middot; Run</button><button data-tab="results">3 &middot; Results</button></nav>
<div class="chips"><span class="chip" id="chip-ollama"><b></b>Ollama</span><span class="chip" id="chip-google"><b></b>Google</span><span class="chip" id="chip-reddit"><b></b>Reddit</span></div></div></header>
<main>
<!-- ============ SETUP ============ -->
<section id="tab-setup" class="on">
 <div class="card"><h2>Local AI (Ollama) <span class="tip" tabindex="0" data-tip="Ollama runs open-source language models on your own computer: free, private, no API key. The scraper uses it to plan searches, judge how well each company fits, extract people and signals, and draft outreach.">?</span></h2>
  <p class="sub">Connect Ollama, choose the model, and (optionally) download a new one. Everything stays on your machine.</p>
  <div class="grid">
   <div><label>Ollama address <span class="tip" tabindex="0" data-tip="Where Ollama is listening. The default http://127.0.0.1:11434 is right unless you run Ollama on another machine or port.">?</span></label><input type="text" id="o-host"></div>
   <div><label>AI mode <span class="tip" tabindex="0" data-tip="Auto: use AI when Ollama is available, otherwise run without it. Always on: stop with an error if Ollama is missing. Off: never use AI.">?</span></label>
    <select id="o-mode"><option value="auto">Auto (use if available)</option><option value="on">Always on (require Ollama)</option><option value="off">Off</option></select></div>
   <div><label>Model <span class="tip" tabindex="0" data-tip="The language model used for all AI steps. 'Auto' picks the best installed chat model (qwen2.5, llama3.x, gemma3, mistral...). Embedding-only models are hidden. Bigger = smarter but slower.">?</span></label><select id="o-model"></select></div>
   <div><label>Parallel AI requests <span class="tip" tabindex="0" data-tip="How many AI calls run at once. Keep it at 1-2 on a laptop; raise it only if you have a strong GPU.">?</span></label><input type="number" id="o-workers" min="1" max="8"></div>
  </div>
  <div class="row" style="margin-top:12px"><button class="btn" id="o-check">Check connection</button><button class="btn sec" id="o-start" style="display:none">Start Ollama</button><button class="btn sec" id="o-test">Test model</button><a href="https://ollama.com/download" target="_blank" rel="noopener noreferrer" id="o-install" style="display:none">Install Ollama &rarr;</a></div>
  <div id="o-msg" class="msg">Checking...</div>
  <h3 style="margin:18px 0 6px;font-size:14px">Download a model <span class="tip" tabindex="0" data-tip="Downloads a model into Ollama (one-time, a few GB). Pick a suggestion or type any model name from ollama.com/library, e.g. qwen2.5:7b.">?</span></h3>
  <div id="o-rec" class="note"></div>
  <div class="row" style="margin-top:8px"><input type="text" id="o-pull" list="o-suggest" placeholder="qwen2.5:7b" style="max-width:300px"><datalist id="o-suggest"></datalist><button class="btn" id="o-pullbtn">Download</button></div>
  <div class="prog" id="o-prog" style="display:none"><i></i></div><div class="note" id="o-pullmsg"></div>
 </div>
 <div class="card"><h2>Accounts (optional) <span class="tip" tabindex="0" data-tip="Needed only for the Google Maps source and the Reddit tool. You sign in on Google's / Reddit's own page; this tool never sees your password. Logins are stored in a file only your user account can read.">?</span></h2>
  <p class="sub">Official APIs with your own login. Skip this if you only want web search.</p>
  <div class="grid">
   <div><h3 style="margin:0 0 8px;font-size:14px">Google (Maps / Places API) <span class="tip" tabindex="0" data-tip="Option A: sign in with your Google account using an OAuth 'Desktop app' client_secret.json from console.cloud.google.com (enable Places API (New) + billing on that project). Option B: paste a Maps Platform API key. Google's terms limit storing Places data - read the README.">?</span></h3>
    <div id="g-status" class="note"></div>
    <label style="margin-top:8px">client_secret.json <span class="tip" tabindex="0" data-tip="OAuth client of type 'Desktop app'. Download it from Google Cloud Console > APIs & Services > Credentials.">?</span></label><input type="file" id="g-file" accept=".json">
    <label style="margin-top:8px">Cloud project ID</label><input type="text" id="g-project" placeholder="my-project-123">
    <div class="row" style="margin:8px 0"><button class="btn" id="g-login">Sign in with Google</button></div>
    <label>&hellip;or API key</label><input type="password" id="g-key" placeholder="AIza..."><div class="row" style="margin-top:8px"><button class="btn sec" id="g-savekey">Save key</button><button class="btn sec" id="g-out">Disconnect</button></div>
    <div id="g-msg" class="msg" style="display:none"></div></div>
   <div><h3 style="margin:0 0 8px;font-size:14px">Reddit <span class="tip" tabindex="0" data-tip="Create an app at reddit.com/prefs/apps with redirect URI http://127.0.0.1:8765/callback. Reddit's sign-in page offers 'Continue with Google'. Without sign-in (app-only) you still get read access to public posts.">?</span></h3>
    <div id="r-status" class="note"></div>
    <label style="margin-top:8px">Client ID</label><input type="text" id="r-id"><label style="margin-top:8px">Secret (blank for installed-app type)</label><input type="password" id="r-secret">
    <label style="margin-top:8px">Your Reddit username <span class="tip" tabindex="0" data-tip="Reddit requires a descriptive User-Agent that includes your username.">?</span></label><input type="text" id="r-user">
    <div class="chk"><input type="checkbox" id="r-app"><label for="r-app">App-only access (skip user sign-in)</label></div>
    <div class="row"><button class="btn" id="r-login">Connect Reddit</button><button class="btn sec" id="r-out">Disconnect</button></div>
    <div id="r-msg" class="msg" style="display:none"></div></div>
  </div></div>
</section>
<!-- ============ RUN ============ -->
<section id="tab-run">
 <div class="modes" id="modes"></div><p class="sub" id="mode-desc"></p>
 <div class="card"><div id="form"></div>
  <div class="row"><button class="btn" id="run">Run</button><button class="btn bad" id="stop" style="display:none">Stop &amp; export what's found</button><button class="btn sec" id="reset">Reset form</button><span class="note" id="ai-note"></span></div></div>
 <div class="card" id="job-card" style="display:none"><h2 id="job-title">Progress</h2><div class="note" id="job-stage"></div><div class="prog" id="prog"><i></i></div><div id="job-msg"></div><pre id="log"></pre>
  <div class="row" style="margin-top:10px" id="job-files"></div></div>
</section>
<!-- ============ RESULTS ============ -->
<section id="tab-results">
 <div class="card"><div class="bar2"><h2 style="margin:0" id="res-title">Results</h2><input type="text" id="filter" placeholder="Filter rows..."><span class="note" id="res-count"></span><span id="res-files" class="row"></span></div>
  <div class="tablewrap"><table><thead id="thead"></thead><tbody id="tbody"></tbody></table></div><div class="empty" id="res-empty">Run something first - results appear here.</div>
  <div id="contacts-wrap" style="display:none"><h3 style="font-size:14px;margin:16px 0 6px">Contacts &amp; insights</h3><div class="tablewrap"><table><thead id="chead"></thead><tbody id="cbody"></tbody></table></div></div></div>
 <div class="card"><h2>Previous runs <span class="tip" tabindex="0" data-tip="Every run is saved as an Excel file (plus CSV/JSONL) in your results folder, so you can re-download them any time.">?</span></h2><div id="history" class="note">Loading...</div></div>
</section>
</main>
<aside id="drawer"><div class="row"><b id="d-title" style="flex:1"></b><button class="btn sec" id="d-close">Close</button></div><dl id="d-body"></dl></aside>
<script>
const TOKEN="__TOKEN__";
const $=(s,r=document)=>r.querySelector(s), $$=(s,r=document)=>[...r.querySelectorAll(s)];
const esc=s=>String(s==null?"":s).replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const safeUrl=u=>/^https?:\/\//i.test(u||"")?u:"";
async function api(path,body){const o={headers:{"X-UI-Token":TOKEN}};if(body!==undefined){o.method="POST";o.headers["Content-Type"]="application/json";o.body=JSON.stringify(body)}
 const r=await fetch(path,o);let j={};try{j=await r.json()}catch(e){}if(!r.ok)throw new Error(j.error||("HTTP "+r.status));return j}
function el(tag,attrs,...kids){const e=document.createElement(tag);for(const[k,v]of Object.entries(attrs||{})){if(k==="class")e.className=v;else if(k==="text")e.textContent=v;else if(v!==false&&v!=null)e.setAttribute(k,v===true?"":v)}for(const k of kids)if(k!=null)e.append(k);return e}
function setMsg(node,text,cls){node.style.display=text?"block":"none";node.textContent=text||"";node.className="msg "+(cls||"")}
function tip(text){return el("span",{class:"tip",tabindex:"0","data-tip":text,text:"?"})}

/* ---------- tabs ---------- */
function show(tab){$$("section").forEach(s=>s.classList.toggle("on",s.id==="tab-"+tab));$$("nav button").forEach(b=>b.classList.toggle("on",b.dataset.tab===tab));
 if(tab==="results")loadHistory();if(tab==="setup"){refreshOllama();refreshAuth()}}
$$("nav button").forEach(b=>b.onclick=()=>show(b.dataset.tab));$$(".chip").forEach(c=>c.onclick=()=>show("setup"));

/* ---------- settings + Ollama ---------- */
let SETTINGS={},OST={};
async function saveSettings(){SETTINGS=await api("/api/settings",{ollama_host:$("#o-host").value,ai_mode:$("#o-mode").value,model:$("#o-model").value,ai_workers:+$("#o-workers").value||2});aiNote()}
function aiNote(){const m=SETTINGS.ai_mode;$("#ai-note").textContent=m==="off"?"AI is off (change in Setup)":(OST.running?("AI: "+(SETTINGS.model||OST.auto_pick||"auto")):(m==="on"?"AI required but Ollama is not running":"AI unavailable - will run without it"))}
async function refreshOllama(){
 try{SETTINGS=await api("/api/settings");$("#o-host").value=SETTINGS.ollama_host;$("#o-mode").value=SETTINGS.ai_mode;$("#o-workers").value=SETTINGS.ai_workers;
  OST=await api("/api/ollama/status?host="+encodeURIComponent($("#o-host").value));
  const sel=$("#o-model");sel.textContent="";sel.append(el("option",{value:"",text:"Auto"+(OST.auto_pick?" (best installed: "+OST.auto_pick+")":"")}));
  OST.models.filter(m=>m.chat).forEach(m=>sel.append(el("option",{value:m.name,text:m.name+(m.params?" - "+m.params:"")+" ("+m.size_gb+" GB)"})));
  sel.value=SETTINGS.model||"";if(sel.value!==(SETTINGS.model||""))sel.value="";
  const dl=$("#o-suggest");dl.textContent="";OST.suggested.forEach(s=>dl.append(el("option",{value:s.name,label:s.note})));
  $("#o-rec").textContent="";const rec=OST.recommended;$("#o-rec").append(rec.why+" Suggested: ",el("a",{href:"#",text:rec.model}),". ");$("#o-rec a").onclick=e=>{e.preventDefault();$("#o-pull").value=rec.model};
  if(!$("#o-pull").value)$("#o-pull").value=rec.model;
  const chip=$("#chip-ollama");chip.className="chip "+(OST.running?"ok":"bad");chip.lastChild.textContent="Ollama"+(OST.running?" \u00b7 "+(SETTINGS.model||OST.auto_pick||"no model"):"");
  $("#o-start").style.display=(!OST.running&&OST.binary)?"":"none";$("#o-install").style.display=(!OST.running&&!OST.binary)?"":"none";
  const chat=OST.models.filter(m=>m.chat).length;
  if(!OST.running)setMsg($("#o-msg"),OST.binary?"Ollama is installed but not running. Press 'Start Ollama'.":"Ollama was not found. Install it, then press 'Check connection'.","warn");
  else if(!chat)setMsg($("#o-msg"),"Ollama is running but has no chat model yet. Download one below (suggested: "+rec.model+").","warn");
  else setMsg($("#o-msg"),"Connected to Ollama - "+chat+" chat model"+(chat>1?"s":"")+" available.","ok");
  aiNote()}catch(e){setMsg($("#o-msg"),"Could not reach the Lead Scraper server: "+e.message,"bad")}}
["o-host","o-mode","o-model","o-workers"].forEach(id=>$("#"+id).addEventListener("change",async()=>{await saveSettings();if(id==="o-host")refreshOllama()}));
$("#o-check").onclick=refreshOllama;
$("#o-start").onclick=async()=>{setMsg($("#o-msg"),"Starting Ollama...","");const r=await api("/api/ollama/start",{host:$("#o-host").value});await refreshOllama();setMsg($("#o-msg"),r.message,r.ok?"ok":"warn")};
$("#o-test").onclick=async()=>{const m=$("#o-model").value||OST.auto_pick;if(!m){setMsg($("#o-msg"),"Pick or download a model first.","warn");return}
 setMsg($("#o-msg"),"Testing "+m+" (the first call loads the model and can take a minute)...","");try{const r=await api("/api/ollama/test",{host:$("#o-host").value,model:m});setMsg($("#o-msg"),r.message+" ("+r.seconds+"s)"+(r.output?" Output: "+r.output:""),r.ok?"ok":"warn")}catch(e){setMsg($("#o-msg"),e.message,"bad")}};
let pullTimer=null;
$("#o-pullbtn").onclick=async()=>{const m=$("#o-pull").value.trim();if(!m)return;try{await api("/api/ollama/pull",{host:$("#o-host").value,model:m})}catch(e){$("#o-pullmsg").textContent=e.message;return}
 $("#o-prog").style.display="";clearInterval(pullTimer);pullTimer=setInterval(async()=>{const p=await api("/api/ollama/pull_status");$("#o-prog i").style.width=(p.pct||0)+"%";
  $("#o-pullmsg").textContent=p.status==="error"?("Error: "+p.error):((p.model||"")+" - "+(p.msg||"")+(p.pct?" "+p.pct+"%":""));
  if(p.status!=="running"){clearInterval(pullTimer);if(p.status==="done"){$("#o-model").value=p.model;await refreshOllama();$("#o-model").value=p.model;await saveSettings()}}},1000)};

/* ---------- accounts ---------- */
async function refreshAuth(){try{const a=await api("/api/auth/status");
 const g=a.google,r=a.reddit;$("#g-status").textContent=g.connected?(g.mode==="key"?"Connected with API key ..."+g.key_tail:"Signed in as "+(g.email||"your Google account")+" (project "+(g.project||"none")+")"):"Not connected";
 $("#r-status").textContent=r.connected?("Connected ("+r.mode+" access, client "+r.client+")"):"Not connected";
 $("#chip-google").className="chip "+(g.connected?"ok":"");$("#chip-reddit").className="chip "+(r.connected?"ok":"");
 for(const[k,id]of[["google","g-msg"],["reddit","r-msg"]]){const p=a.pending[k];if(p)setMsg($("#"+id),p.message,p.state==="done"?"ok":p.state==="error"?"bad":"warn")}
 if(Object.values(a.pending).some(p=>p.state==="waiting"))setTimeout(refreshAuth,1500)}catch(e){}}
$("#g-login").onclick=async()=>{const f=$("#g-file").files[0];if(!f){setMsg($("#g-msg"),"Choose your client_secret.json first.","warn");return}
 try{await api("/api/auth/google",{client_json:await f.text(),project:$("#g-project").value});setMsg($("#g-msg"),"A browser window opened - finish signing in with Google...","warn");setTimeout(refreshAuth,1200)}catch(e){setMsg($("#g-msg"),e.message,"bad")}};
$("#g-savekey").onclick=async()=>{try{await api("/api/auth/google",{api_key:$("#g-key").value});$("#g-key").value="";setMsg($("#g-msg"),"API key saved.","ok");refreshAuth()}catch(e){setMsg($("#g-msg"),e.message,"bad")}};
$("#g-out").onclick=async()=>{await api("/api/auth/logout",{target:"google"});setMsg($("#g-msg"),"Disconnected.","ok");refreshAuth()};
$("#r-login").onclick=async()=>{try{await api("/api/auth/reddit",{client_id:$("#r-id").value,secret:$("#r-secret").value,username:$("#r-user").value,app_only:$("#r-app").checked});
 setMsg($("#r-msg"),$("#r-app").checked?"Saved.":"A browser window opened - sign in on Reddit (it offers 'Continue with Google')...","warn");setTimeout(refreshAuth,1200)}catch(e){setMsg($("#r-msg"),e.message,"bad")}};
$("#r-out").onclick=async()=>{await api("/api/auth/logout",{target:"reddit"});setMsg($("#r-msg"),"Disconnected.","ok");refreshAuth()};

/* ---------- run form (generated from the CLI options) ---------- */
let SPEC=null,MODE="leads",VALUES={};
function loadVals(){try{VALUES=JSON.parse(localStorage.getItem("ls_values")||"{}")}catch(e){VALUES={}}}
function saveVals(){try{localStorage.setItem("ls_values",JSON.stringify(VALUES))}catch(e){}}
function field(f){const v=(VALUES[MODE]||{})[f.dest];const wrap=el("div",{class:f.kind==="lines"||f.kind==="engines"?"":""});
 const label=el("label",{},f.label,f.help?tip(f.help+"  ("+f.flag+")"):null);let input;
 const set=val=>{(VALUES[MODE]=VALUES[MODE]||{})[f.dest]=val;saveVals()};
 if(f.kind==="bool"){const id="f-"+f.dest;input=el("input",{type:"checkbox",id});input.checked=v===undefined?!!f.default:!!v;input.onchange=()=>set(input.checked);
  return el("div",{class:"chk"},input,el("label",{for:id},f.label,f.help?tip(f.help+"  ("+f.flag+")"):null))}
 if(f.kind==="engines"){const box=el("div",{class:"row"});const cur=v===undefined?"all":v;const on=cur==="all"?SPEC.engines:String(cur).split(",");
  SPEC.engines.forEach(n=>{const cb=el("input",{type:"checkbox","data-e":n});cb.checked=on.includes(n);cb.onchange=()=>{const sel=$$("input[data-e]",box).filter(x=>x.checked).map(x=>x.dataset.e);set(sel.length===SPEC.engines.length?"all":sel.join(","))};box.append(el("span",{class:"chk"},cb,el("label",{text:n})))});
  wrap.append(label,box);return wrap}
 if(f.kind==="lines"){input=el("textarea",{placeholder:f.placeholder||""});input.value=v===undefined?"":v;input.oninput=()=>set(input.value)}
 else if(f.kind==="select"){input=el("select");f.choices.forEach(c=>input.append(el("option",{value:c,text:String(c)})));input.value=v===undefined?f.default:v;input.onchange=()=>set(input.value)}
 else{input=el("input",{type:f.kind==="text"?"text":"number",placeholder:f.placeholder||"",step:f.kind==="float"?"any":null});input.value=v===undefined?(f.default==null?"":f.default):v;input.oninput=()=>set(input.value)}
 input.title=(f.help||"")+"  ("+f.flag+")";wrap.append(label,input);return wrap}
function renderForm(){const m=SPEC.modes[MODE];$("#mode-desc").textContent=m.description;const form=$("#form");form.textContent="";
 $$("#modes button").forEach(b=>b.classList.toggle("on",b.dataset.mode===MODE));
 for(const g of m.groups){const fs=m.fields.filter(f=>f.group===g);if(!fs.length)continue;
  const grid=el("div",{class:"grid"});fs.forEach(f=>{const n=field(f);if(f.kind==="lines")n.style.gridColumn="1/-1";grid.append(n)});
  if(g==="Advanced"){form.append(el("details",{},el("summary",{text:"Advanced options"}),grid))}else form.append(el("fieldset",{},el("legend",{text:g}),grid))}}
function params(){const out={};for(const f of SPEC.modes[MODE].fields){const v=(VALUES[MODE]||{})[f.dest];out[f.dest]=v===undefined?(f.kind==="engines"?"all":f.default):v}return out}
async function initRun(){SPEC=await api("/api/spec");loadVals();const bar=$("#modes");
 for(const[k,m]of Object.entries(SPEC.modes))bar.append(el("button",{"data-mode":k,text:m.title}));
 bar.onclick=e=>{const b=e.target.closest("button");if(!b)return;MODE=b.dataset.mode;renderForm()};renderForm()}
$("#reset").onclick=()=>{VALUES[MODE]={};saveVals();renderForm()};

/* ---------- running a job ---------- */
let poll=null,since=0,CUR=null;
function setRunning(on){$("#run").disabled=on;$("#stop").style.display=on?"":"none"}
$("#run").onclick=async()=>{try{await api("/api/run",{mode:MODE,params:params()})}catch(e){alert(e.message);return}
 since=0;$("#log").textContent="";$("#job-card").style.display="";$("#job-files").textContent="";setMsg($("#job-msg"),"");$("#job-title").textContent="Running: "+SPEC.modes[MODE].title;setRunning(true);clearInterval(poll);poll=setInterval(tick,700);tick()};
$("#stop").onclick=async()=>{await api("/api/cancel",{});$("#stop").disabled=true};
async function tick(){let j;try{j=await api("/api/job?since="+since)}catch(e){return}
 if(j.status==="none")return;since=j.next;const log=$("#log");if(j.lines.length){const stick=log.scrollTop+log.clientHeight>=log.scrollHeight-30;log.textContent+=j.lines.join("\n")+"\n";if(stick)log.scrollTop=log.scrollHeight}
 const p=j.progress||{},bar=$("#prog");$("#job-stage").textContent=(p.stage||"")+(p.total?" - "+p.done+"/"+p.total:p.done?" - "+p.done:"")+"  ("+j.seconds+"s)";
 if(p.total){bar.classList.remove("ind");bar.firstChild.style.width=Math.min(100,100*p.done/p.total)+"%"}else if(j.status==="running")bar.classList.add("ind");
 if(j.status!=="running"){clearInterval(poll);setRunning(false);$("#stop").disabled=false;bar.classList.remove("ind");bar.firstChild.style.width=j.status==="error"?"0":"100%";CUR=j;
  $("#job-title").textContent=j.status==="done"?"Finished":j.status==="stopped"?"Stopped (partial results exported)":"Stopped with an error";
  if(j.error)setMsg($("#job-msg"),j.error,"bad");else setMsg($("#job-msg"),j.status==="stopped"?"Stopped early - partial results are available.":"Done in "+j.seconds+"s.","ok");
  fileLinks($("#job-files"),j.files);if(j.results){renderResults(j.results,j.files);if(!j.error)$("#job-files").append(el("button",{class:"btn",text:"View results",onclick:()=>show("results")}))}}}
function fileLinks(node,files){node.textContent="";(files||[]).forEach(f=>node.append(el("a",{class:"btn sec",style:"text-decoration:none",href:"/download?f="+encodeURIComponent(f)+"&t="+TOKEN,text:"\u2b07 "+f.replace(/^.*\./,"").toUpperCase()+" ("+f+")"})))}

/* ---------- results ---------- */
const COLS={leads:[["company","Company"],["website","Website","url"],["email","Email"],["phone","Phone"],["score","Score"],["ai_fit","AI fit"],["ai_industry","Industry"],["status","Status"],["found_by","Found by"]],
 reddit:[["rank_score","Rank"],["intent","Intent"],["subreddit","Subreddit"],["title","Title","link:url"],["author","Author"],["upvotes","Upvotes"],["comments","Comments"],["ai_relevance","AI rel."]],
 crawl:[["url","URL","url"],["status","Status"],["depth","Depth"],["title","Title"],["words","Words"],["emails","Emails"],["issues","Issues"]]};
let ROWS=[],KIND="",SORT=null,FILTER="";
function cellNode(row,[key,,type]){const v=row[key];if(v===-1||v===""||v==null)return document.createTextNode("");
 if(type==="url"||(type||"").startsWith("link:")){const u=safeUrl(type==="url"?v:row[type.slice(5)]);if(u)return el("a",{href:u,target:"_blank",rel:"noopener noreferrer",text:String(type==="url"?v:v).replace(/^https?:\/\//,"")});return document.createTextNode(String(v))}
 if(key==="score"||key==="rank_score"||key==="ai_fit"){const n=+v,hi=key==="ai_fit"?n>=7:n>=70;return el("span",{class:"pill "+(hi?"g":""),text:String(v)})}
 if(key==="status"&&String(v)!=="ok"&&KIND!=="crawl")return el("span",{class:"pill r",text:String(v)});return document.createTextNode(String(v))}
function drawTable(){const cols=COLS[KIND]||[];const th=$("#thead");th.textContent="";const tr=el("tr");
 cols.forEach(c=>{const h=el("th",{text:c[1]+(SORT&&SORT.k===c[0]?(SORT.d>0?" \u25b2":" \u25bc"):"")});h.onclick=()=>{SORT=SORT&&SORT.k===c[0]?{k:c[0],d:-SORT.d}:{k:c[0],d:-1};drawTable()};tr.append(h)});th.append(tr);
 let rows=ROWS.filter(r=>!FILTER||JSON.stringify(r).toLowerCase().includes(FILTER));
 if(SORT)rows=[...rows].sort((a,b)=>{const x=a[SORT.k],y=b[SORT.k];return(typeof x==="number"&&typeof y==="number"?x-y:String(x||"").localeCompare(String(y||"")))*SORT.d});
 const tb=$("#tbody");tb.textContent="";rows.slice(0,1500).forEach(r=>{const t=el("tr");cols.forEach(c=>t.append(el("td",{},cellNode(r,c))));t.onclick=()=>detail(r);tb.append(t)});
 $("#res-count").textContent=rows.length+" row"+(rows.length===1?"":"s");$("#res-empty").style.display=ROWS.length?"none":""}
function renderResults(res,files){KIND=res.kind;ROWS=res.rows||[];SORT=KIND==="leads"?{k:"score",d:-1}:KIND==="reddit"?{k:"rank_score",d:-1}:null;FILTER="";$("#filter").value="";
 $("#res-title").textContent=KIND==="leads"?"Leads":KIND==="reddit"?"Reddit signals":"Crawled pages";fileLinks($("#res-files"),files);drawTable();
 const cw=$("#contacts-wrap");if(KIND==="crawl"&&(res.contacts||[]).length){cw.style.display="";const h=$("#chead");h.textContent="";const tr=el("tr");["Site","Type","Value","Found on","Pages"].forEach(x=>tr.append(el("th",{text:x})));h.append(tr);
  const b=$("#cbody");b.textContent="";res.contacts.slice(0,1000).forEach(c=>{const t=el("tr");[c.site,c.type,c.value,c.found_on,c.pages||""].forEach(x=>t.append(el("td",{text:String(x)})));b.append(t)})}else cw.style.display="none"}
$("#filter").oninput=e=>{FILTER=e.target.value.toLowerCase();drawTable()};
function detail(r){$("#d-title").textContent=r.company||r.title||r.url||"Details";const dl=$("#d-body");dl.textContent="";
 for(const[k,v]of Object.entries(r)){if(v===""||v==null||v===-1)continue;dl.append(el("dt",{text:k.replace(/_/g," ")}));const dd=el("dd");const u=safeUrl(String(v));if(u)dd.append(el("a",{href:u,target:"_blank",rel:"noopener noreferrer",text:String(v)}));else dd.textContent=String(v);
  if(k==="ai_outreach")dd.append(el("div",{},el("button",{class:"btn sec",text:"Copy draft",onclick:()=>navigator.clipboard&&navigator.clipboard.writeText(String(v))})));dl.append(dd)}
 $("#drawer").classList.add("on")}
$("#d-close").onclick=()=>$("#drawer").classList.remove("on");document.addEventListener("keydown",e=>{if(e.key==="Escape")$("#drawer").classList.remove("on")});
async function loadHistory(){try{const h=await api("/api/history");const box=$("#history");box.textContent="";if(!h.files.length){box.textContent="No runs yet.";return}
 h.files.forEach(f=>{const row=el("div",{class:"row",style:"margin:4px 0"},el("b",{text:f.name}),el("span",{class:"note",text:f.when+" \u00b7 "+f.mb+" MB"}));
  f.files.forEach(n=>row.append(el("a",{href:"/download?f="+encodeURIComponent(n)+"&t="+TOKEN,text:n.replace(/^.*\./,".")})));box.append(row)})}catch(e){}}

/* ---------- boot ---------- */
(async()=>{try{await initRun()}catch(e){$("#form").textContent="Could not load options: "+e.message}
 await refreshOllama();await refreshAuth();
 const j=await api("/api/job?since=0").catch(()=>({status:"none"}));if(j.status==="running"){$("#job-card").style.display="";setRunning(true);since=0;poll=setInterval(tick,700)}
 const st=await api("/api/ollama/status?host="+encodeURIComponent($("#o-host").value)).catch(()=>null);if(st&&!st.running||st&&!st.models.some(m=>m.chat))show("setup")})();
</script></body></html>
"""

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=8501)
    ap.add_argument("--no-browser", action="store_true")
    ap.add_argument("--out-dir", default=None)
    a = ap.parse_args()
    serve(a.port, a.out_dir, not a.no_browser)
