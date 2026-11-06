#!/usr/bin/env python3
"""
Local web UI for the M365 exporter.

Starts a small web server on 127.0.0.1 (this machine only) where you can fill in
a form, run an export (or the credential-free demo), watch live progress, and
browse / download the results. Standard library only; the export itself is run
by the existing run_m365_export.py / run_demo.py scripts.

Usage:
    python3 server.py                 # http://127.0.0.1:8365
    python3 server.py --port 9000
"""

import argparse
import html
import json
import mimetypes
import re
import subprocess
import sys
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, unquote, urlparse

ROOT = Path(__file__).resolve().parent
OUTPUT_DIRS = {"real": ROOT / "out", "demo": ROOT / "demo_out"}
EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+'-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$")
DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
MAX_LOG_LINES = 5000


class Job:
    """One export run at a time; output lines are kept for the UI to poll."""

    def __init__(self):
        self.lock = threading.Lock()
        self.proc = None
        self.lines = []
        self.status = "idle"  # idle | running | done | failed | stopped
        self.mode = None
        self.stopped = False

    def start(self, mode, cmd):
        with self.lock:
            if self.status == "running":
                raise RuntimeError("An export is already running.")
            self.lines = [f"$ {' '.join(cmd)}"]
            self.status, self.mode, self.stopped = "running", mode, False
            self.proc = subprocess.Popen(
                cmd, cwd=ROOT, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                text=True, bufsize=1, errors="replace",
            )
        threading.Thread(target=self._pump, args=(self.proc,), daemon=True).start()

    def _pump(self, proc):
        for line in proc.stdout:
            with self.lock:
                self.lines.append(line.rstrip("\n"))
                del self.lines[:-MAX_LOG_LINES]
        code = proc.wait()
        with self.lock:
            self.status = "stopped" if self.stopped else ("done" if code == 0 else "failed")
            self.lines.append(f"[process exited with code {code}]")

    def stop(self):
        with self.lock:
            if self.status == "running" and self.proc:
                self.stopped = True
                self.proc.terminate()

    def snapshot(self, since=0):
        with self.lock:
            return {"status": self.status, "mode": self.mode,
                    "total": len(self.lines), "lines": self.lines[since:]}


JOB = Job()


def build_command(form):
    """Validate form input and return (mode, argv). Raises ValueError on bad input."""
    mode = "demo" if form.get("mode") == "demo" else "real"
    users = [u.strip() for u in re.split(r"[,\s]+", form.get("users", "")) if u.strip()]
    if not users:
        raise ValueError("Enter at least one employee email.")
    for u in users:
        if not EMAIL_RE.match(u):
            raise ValueError(f"Not a valid email address: {u}")
    script = "run_demo.py" if mode == "demo" else "run_m365_export.py"
    cmd = [sys.executable, "-u", str(ROOT / script)]
    if mode == "real":
        admin = form.get("admin", "").strip()
        if not EMAIL_RE.match(admin):
            raise ValueError("Enter a valid admin email address.")
        cmd += ["--admin", admin, "--export-only", "--local-only"]
    for u in users:
        cmd += ["--only", u]
    since_days = form.get("since_days", "").strip()
    after = form.get("modified_after", "").strip()
    if since_days:
        if not since_days.isdigit() or int(since_days) < 1:
            raise ValueError("'Last N days' must be a positive whole number.")
        cmd += ["--since-days", since_days]
    elif after:
        if not DATE_RE.match(after):
            raise ValueError("Date must be in YYYY-MM-DD format.")
        cmd += ["--modified-after", after]
    return mode, cmd


def resolve_output(rel):
    """Map a '<real|demo>/relative/path' URL path to a file under an output dir (no traversal)."""
    parts = Path(unquote(rel)).parts
    if not parts or parts[0] not in OUTPUT_DIRS:
        return None
    base = OUTPUT_DIRS[parts[0]].resolve()
    target = base.joinpath(*parts[1:]).resolve()
    if target != base and base not in target.parents:
        return None
    return target if target.exists() else None


def browse_html(rel):
    target = resolve_output(rel) if rel else None
    esc = html.escape
    if not rel:
        items = []
        for key, d in OUTPUT_DIRS.items():
            label = "Real exports (out/)" if key == "real" else "Demo data (demo_out/)"
            items.append(f'<li><a href="/browse/{key}">{label}</a>'
                         f'{"" if d.exists() else " <em>(empty)</em>"}</li>')
        body = "<ul>" + "".join(items) + "</ul>"
    elif target is None:
        body = "<p>Not found.</p>"
    elif target.is_dir():
        crumbs = Path(unquote(rel)).parts
        rows = []
        if len(crumbs) > 1:
            up = "/".join(quote(c) for c in crumbs[:-1])
            rows.append(f'<li><a href="/browse/{up}">⬆ ..</a></li>')
        for p in sorted(target.iterdir(), key=lambda p: (p.is_file(), p.name.lower())):
            link = "/".join(quote(c) for c in (*crumbs, p.name))
            if p.is_dir():
                rows.append(f'<li>📁 <a href="/browse/{link}">{esc(p.name)}/</a></li>')
            else:
                rows.append(f'<li>📄 <a href="/file/{link}">{esc(p.name)}</a> '
                            f'<span class="sz">{p.stat().st_size:,} bytes</span></li>')
        body = f"<p><code>{esc('/'.join(crumbs))}</code></p><ul>{''.join(rows) or '<li><em>empty</em></li>'}</ul>"
    else:
        body = "<p>Not a folder.</p>"
    return page("Results", f'<p><a href="/">← Back</a></p><h2>Results</h2>{body}')


STYLE = """
body{font:15px/1.5 system-ui,sans-serif;max-width:860px;margin:2rem auto;padding:0 1rem;color:#1b1f24}
h1{margin-bottom:.2rem} .sub{color:#57606a;margin-top:0}
fieldset{border:1px solid #d0d7de;border-radius:8px;margin:1rem 0;padding:1rem}
label{display:block;margin:.6rem 0 .2rem;font-weight:600}
input[type=text],input[type=number],input[type=date]{width:100%;padding:.5rem;border:1px solid #d0d7de;border-radius:6px;box-sizing:border-box}
.row{display:flex;gap:1rem}.row>div{flex:1}
button{padding:.55rem 1.1rem;border:0;border-radius:6px;background:#0969da;color:#fff;font-size:15px;cursor:pointer}
button.stop{background:#cf222e} button:disabled{opacity:.5;cursor:default}
pre{background:#0d1117;color:#e6edf3;padding:1rem;border-radius:8px;height:340px;overflow:auto;white-space:pre-wrap}
.badge{padding:.1rem .6rem;border-radius:99px;background:#eaeef2;font-size:13px}
.err{color:#cf222e;font-weight:600} .sz{color:#57606a;font-size:13px} ul{line-height:1.9}
@media (prefers-color-scheme:dark){body{background:#0d1117;color:#e6edf3}fieldset,input{border-color:#30363d}
input{background:#161b22;color:#e6edf3}.badge{background:#21262d}.sub,.sz{color:#8b949e}}
"""


def page(title, inner, script=""):
    return (f"<!doctype html><html lang='en'><head><meta charset='utf-8'>"
            f"<meta name='viewport' content='width=device-width,initial-scale=1'>"
            f"<title>{html.escape(title)}</title><style>{STYLE}</style></head>"
            f"<body>{inner}<script>{script}</script></body></html>")


INDEX_BODY = """
<h1>M365 Exporter</h1>
<p class="sub">Local only — runs on this machine, nothing is uploaded anywhere.</p>
<form id="f">
<fieldset>
  <label><input type="radio" name="mode" value="demo" checked> Demo (fake data, no credentials)</label>
  <label><input type="radio" name="mode" value="real"> Real export (needs <code>.env</code> credentials)</label>
  <div id="adminrow" style="display:none"><label>Admin email (pre-flight check)</label>
  <input type="text" name="admin" placeholder="admin@yourcompany.com"></div>
  <label>Employee email(s) — separate several with commas or spaces</label>
  <input type="text" name="users" placeholder="employee@yourcompany.com" required>
  <div class="row">
    <div><label>Last N days</label><input type="number" name="since_days" min="1" placeholder="90"></div>
    <div><label>…or modified after</label><input type="date" name="modified_after"></div>
  </div>
</fieldset>
<button id="go" type="submit">Start export</button>
<button id="stop" class="stop" type="button" disabled>Stop</button>
<span id="status" class="badge">idle</span> <a href="/browse">Browse results →</a>
<p id="err" class="err"></p>
</form>
<pre id="log">Log output will appear here.</pre>
"""

INDEX_JS = """
const $=s=>document.querySelector(s);let seen=0,timer=null;
document.querySelectorAll('[name=mode]').forEach(r=>r.onchange=()=>{
  $('#adminrow').style.display=document.querySelector('[name=mode]:checked').value==='real'?'block':'none'});
async function poll(){
  const r=await (await fetch('/api/status?since='+seen)).json();
  if(r.total<seen){seen=0;$('#log').textContent='';return poll()}
  if(r.lines.length){const l=$('#log');if(seen===0)l.textContent='';
    l.textContent+=r.lines.join('\\n')+'\\n';l.scrollTop=l.scrollHeight;seen=r.total}
  $('#status').textContent=r.status;
  const run=r.status==='running';$('#go').disabled=run;$('#stop').disabled=!run;
  if(!run&&timer){clearInterval(timer);timer=null}
  if(run&&!timer)timer=setInterval(poll,1000);
}
$('#f').onsubmit=async e=>{e.preventDefault();$('#err').textContent='';
  const res=await fetch('/api/start',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},
    body:new URLSearchParams(new FormData($('#f')))});
  const j=await res.json();if(!res.ok){$('#err').textContent=j.error;return}
  seen=0;poll()};
$('#stop').onclick=()=>fetch('/api/stop',{method:'POST'}).then(poll);
poll();
"""


class Handler(BaseHTTPRequestHandler):
    server_version = "M365ExporterUI"

    def log_message(self, fmt, *args):
        pass

    def _send(self, code, body, ctype="text/html; charset=utf-8", extra=None):
        data = body if isinstance(body, bytes) else body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(data)

    def _json(self, code, obj):
        self._send(code, json.dumps(obj), "application/json")

    def _host_ok(self):
        # Defend against DNS-rebinding: only accept requests addressed to loopback.
        host = (self.headers.get("Host") or "").rsplit(":", 1)[0].strip("[]")
        return host in ("127.0.0.1", "localhost", "::1")

    def do_GET(self):
        if not self._host_ok():
            return self._send(403, "Forbidden")
        url = urlparse(self.path)
        if url.path == "/":
            return self._send(200, page("M365 Exporter", INDEX_BODY, INDEX_JS))
        if url.path == "/api/status":
            since = parse_qs(url.query).get("since", ["0"])[0]
            return self._json(200, JOB.snapshot(int(since) if since.isdigit() else 0))
        if url.path == "/browse" or url.path.startswith("/browse/"):
            return self._send(200, browse_html(url.path[len("/browse/"):]))
        if url.path.startswith("/file/"):
            target = resolve_output(url.path[len("/file/"):])
            if target is None or not target.is_file():
                return self._send(404, "Not found")
            ctype = mimetypes.guess_type(target.name)[0] or "application/octet-stream"
            return self._send(200, target.read_bytes(), ctype, {
                "Content-Disposition": f"attachment; filename*=UTF-8''{quote(target.name)}",
                "X-Content-Type-Options": "nosniff"})
        self._send(404, "Not found")

    def do_POST(self):
        # Reject cross-site requests: browsers always send Origin on cross-origin POSTs.
        origin = self.headers.get("Origin")
        if not self._host_ok() or (origin and urlparse(origin).hostname not in ("127.0.0.1", "localhost", "::1")):
            return self._send(403, "Forbidden")
        if self.path == "/api/stop":
            JOB.stop()
            return self._json(200, {"ok": True})
        if self.path == "/api/start":
            length = min(int(self.headers.get("Content-Length") or 0), 10_000)
            form = {k: v[0] for k, v in parse_qs(self.rfile.read(length).decode("utf-8", "replace")).items()}
            try:
                mode, cmd = build_command(form)
                JOB.start(mode, cmd)
            except (ValueError, RuntimeError) as e:
                return self._json(400, {"error": str(e)})
            return self._json(200, {"ok": True})
        self._send(404, "Not found")


def main():
    ap = argparse.ArgumentParser(description="Local web UI for the M365 exporter.")
    ap.add_argument("--port", type=int, default=8365)
    ap.add_argument("--no-browser", action="store_true", help="don't open a browser tab")
    args = ap.parse_args()
    httpd = ThreadingHTTPServer(("127.0.0.1", args.port), Handler)
    url = f"http://127.0.0.1:{args.port}/"
    print(f"M365 Exporter UI running at {url}  (Ctrl+C to quit)")
    if not args.no_browser:
        threading.Timer(0.5, lambda: webbrowser.open(url)).start()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down.")
        JOB.stop()
    finally:
        httpd.server_close()


if __name__ == "__main__":
    main()
