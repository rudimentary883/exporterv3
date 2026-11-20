"""
Missing check-ins and missing feedback  ->  one .txt file  (page /missing on the exporter)

  * Teams (not speakers) and judges that are NOT checked in, alphabetical (full or short team name)
      - the check-in of every speaker and judge is read from its own API link (".../speakers/<id>/checkin" and
        ".../adjudicators/<id>/checkin" -> {"checked": true/false, "timestamp": ...})
      - 3v3 / WSDC teams get a note:  "(all debaters not checked-in)"  /  "(2/3 debaters not checked in)"  /  "(1/3 debaters not checked in)"
        BP teams get no note.  Teams whose debaters are all checked in are left out.
  * Teams and judges with missing feedback:  "Name (missing N feedbacks)"
      - Tabbycat's API has no "missing feedback" list, so the numbers are read from the Feedback Progress page itself
        (the public page /<slug>/feedback/progress/ when it is public, otherwise the page saved from the admin site: Ctrl+S, upload the .html)
  * Teams and judges are always listed separately; optional: group by the number missing first, then alphabetical.

Needs the 2 lines at the bottom of app.py:
    from missing_lists import register_missing_lists
    register_missing_lists(app)
"""
import io
import json
import re
import time
from concurrent.futures import ThreadPoolExecutor
from urllib.parse import urlsplit

import requests

UA = "Mozilla/5.0 (compatible; TabbycatMissingLists/1.0; Render)"


# ------------------------------------------------------------------ API helpers
def _get_json(session, url, retries=3):
    for attempt in range(retries):
        try:
            resp = session.get(url, timeout=30)
        except requests.exceptions.RequestException as e:
            if attempt == retries - 1:
                raise RuntimeError(f"Could not reach {url}: {e}")
            time.sleep(1)
            continue
        if resp.status_code == 200:
            try:
                return resp.json()
            except ValueError:
                raise RuntimeError(f"{url} did not return JSON")
        if resp.status_code == 429:
            time.sleep(2 ** attempt)
            continue
        if resp.status_code in (401, 403):
            raise RuntimeError(f"HTTP {resp.status_code} on {url}: the token was refused - it must be an ADMIN account's token.")
        raise RuntimeError(f"HTTP {resp.status_code} on {url}")
    raise RuntimeError(f"Rate limited on {url}")


def _list(session, url):
    data = _get_json(session, url)
    if isinstance(data, dict) and isinstance(data.get("results"), list):
        items = list(data["results"])
        while data.get("next"):
            data = _get_json(session, data["next"])
            items += data.get("results", [])
        return items
    return data if isinstance(data, list) else []


def _id_of(ref):
    text = (ref.get("url") if isinstance(ref, dict) else str(ref or "")).rstrip("/")
    return text.rsplit("/", 1)[-1]


def guess_slug(url):
    parts = [p for p in urlsplit(url.strip()).path.split("/") if p]
    return parts[0] if parts and parts[0] not in ("_", "api", "admin") else ""


def resolve(url, slug, session):
    sp = urlsplit(url.strip())
    origin = f"{sp.scheme or 'https'}://{sp.netloc or sp.path.split('/')[0]}"
    slug = (slug or guess_slug(url)).strip()
    tournaments = _list(session, f"{origin}/api/v1/tournaments")
    slugs = [t.get("slug") for t in tournaments if isinstance(t, dict)]
    if slug not in slugs:
        if len(slugs) == 1:
            slug = slugs[0]
        else:
            raise ValueError("Which tournament? Add its slug. Found: " + ", ".join(map(str, slugs)))
    return origin, slug


def team_label(team, mode):
    order = (("short_name", "long_name", "reference") if mode == "short" else ("long_name", "short_name", "reference"))
    for key in order:
        if team.get(key):
            return str(team[key]).strip()
    return str(team.get("id", "?"))


def read_checkins(session, links, workers=10):
    """links: {key: checkin url} -> {key: True / False / None (could not be read)}"""
    def one(item):
        key, url = item
        try:
            data = _get_json(session, url)
            if isinstance(data, dict) and "checked" in data:
                return key, bool(data["checked"])
        except RuntimeError:
            pass
        return key, None
    with ThreadPoolExecutor(max_workers=workers) as pool:
        return dict(pool.map(one, links.items()))


def missing_checkins(url, token, slug, team_mode="long", session=None):
    s = session or requests.Session()
    s.headers.update({"Authorization": f"Token {token.strip()}", "Accept": "application/json", "User-Agent": UA})
    origin, slug = resolve(url, slug, s)
    base = f"{origin}/api/v1/tournaments/{slug}"
    teams = _list(s, base + "/teams")
    speakers = _list(s, base + "/speakers")
    adjudicators = _list(s, base + "/adjudicators")
    links = {}
    for sp in speakers:
        links[("s", _id_of(sp))] = ((sp.get("_links") or {}).get("checkin")) or (str(sp.get("url", "")).rstrip("/") + "/checkin")
    for a in adjudicators:
        links[("a", _id_of(a))] = ((a.get("_links") or {}).get("checkin")) or (str(a.get("url", "")).rstrip("/") + "/checkin")
    status = read_checkins(s, links)
    unread = sum(1 for v in status.values() if v is None)
    by_team = {}
    for sp in speakers:
        by_team.setdefault(_id_of(sp.get("team")), []).append(status.get(("s", _id_of(sp))))
    team_rows, no_speakers = [], 0
    for t in teams:
        st = by_team.get(_id_of(t), [])
        if not st:
            no_speakers += 1
            continue
        missing = sum(1 for v in st if v is not True)            # not checked in (or unreadable) counts as missing
        if missing:
            team_rows.append({"name": team_label(t, team_mode), "missing": missing, "size": len(st)})
    judge_rows = [{"name": (a.get("name") or "").strip()} for a in adjudicators if status.get(("a", _id_of(a))) is not True]
    return {"teams": team_rows, "judges": judge_rows, "slug": slug, "origin": origin, "unread": unread,
            "no_speakers": no_speakers, "n_teams": len(teams), "n_judges": len(adjudicators)}


# ------------------------------------------------------------------ Feedback Progress page
def parse_feedback_progress(html):
    """-> {'teams': [{'short','full','owed'}], 'judges': [{'name','owed'}]} from the tablesData of the Feedback Progress page."""
    i = html.find("tablesData")
    k = html.find("[", i) if i >= 0 else -1
    if k < 0:
        raise ValueError("This does not look like a Tabbycat Feedback Progress page (no table data found).")
    try:
        tables, _ = json.JSONDecoder().raw_decode(html[k:])
    except ValueError:
        raise ValueError("The table data of the Feedback Progress page could not be read.")
    out = {"teams": [], "judges": []}
    for t in tables:
        keys = [c.get("key") for c in t.get("head", [])]
        if "owed" not in keys:
            continue
        o = keys.index("owed")
        is_team = keys[0] == "team" or "team" in str(t.get("title", "")).lower()
        for row in t.get("data", []):
            cell, owed_cell = row[0], row[o]
            try:
                owed = int(owed_cell.get("sort") if owed_cell.get("sort") is not None else owed_cell.get("text"))
            except (TypeError, ValueError):
                continue
            text = re.sub(r"<[^>]+>", "", str(cell.get("text", ""))).strip()
            if is_team:
                full = ((cell.get("popover") or {}).get("title")) or text
                out["teams"].append({"short": text, "full": str(full).strip(), "owed": owed})
            else:
                out["judges"].append({"name": text, "owed": owed})
    if not out["teams"] and not out["judges"]:
        raise ValueError("No 'From Teams' / 'From Adjudicators' tables were found on that page.")
    return out


def fetch_feedback_page(url, slug):
    sp = urlsplit(url.strip())
    origin = f"{sp.scheme or 'https'}://{sp.netloc or sp.path.split('/')[0]}"
    slug = (slug or guess_slug(url)).strip()
    if not slug:
        raise ValueError("The tournament slug is needed to open the public Feedback Progress page.")
    page = f"{origin}/{slug}/feedback/progress/"
    try:
        resp = requests.get(page, headers={"User-Agent": UA}, timeout=30)
    except requests.exceptions.RequestException as e:
        raise ValueError(f"Could not open {page}: {e}")
    if resp.status_code != 200 or "login" in resp.url.lower() or "tablesData" not in resp.text:
        raise ValueError(f"The public Feedback Progress page ({page}) is not available (HTTP {resp.status_code}). "
                         "Either turn on the public feedback progress page in the tab's settings, or open the admin page "
                         "(Feedback > Overview... > Feedback Progress), press Ctrl+S to save it, and upload that .html file here.")
    return resp.text


# ------------------------------------------------------------------ text output
def _alpha(items, key):
    return sorted(items, key=lambda x: (key(x).casefold(), key(x)))


def _group_lines(items, group_key, group_title, line, name_key, desc):
    groups = {}
    for it in items:
        groups.setdefault(group_key(it), []).append(it)
    lines = []
    for gk in sorted(groups, reverse=desc):
        members = _alpha(groups[gk], name_key)
        lines.append(f"-- {group_title(gk)} ({len(members)}) --")
        lines += [line(m) for m in members]
        lines.append("")
    return lines[:-1] if lines else lines


def checkin_suffix(row, fmt):
    if fmt == "bp":
        return ""
    if row["missing"] >= row["size"]:
        return " (all debaters not checked-in)"
    return f" ({row['missing']}/{row['size']} debaters not checked in)"


def build_text(checkins=None, feedback=None, fmt="3v3", arrange="alpha", desc=False, team_mode="long"):
    out = []

    def section(title, lines, count):
        out.append(f"{title} ({count})")
        out.extend(lines if lines else ["(none)"])
        out.append("")

    if checkins is not None:
        if checkins.get("unread"):
            out.append(f"NOTE: {checkins['unread']} check-in(s) could not be read from the API and are counted as NOT checked in.")
            out.append("")
        trows = checkins["teams"]
        line = lambda r: r["name"] + checkin_suffix(r, fmt)
        if arrange == "category":
            lines = _group_lines(trows, lambda r: (r["missing"] / r["size"], r["missing"], r["size"]),
                                 lambda k: ("All debaters not checked in" if k[1] >= k[2] else f"Missing {k[1]}/{k[2]} check-ins"),
                                 line, lambda r: r["name"], desc)
        else:
            lines = [line(r) for r in _alpha(trows, lambda r: r["name"])]
        section("TEAMS NOT CHECKED IN", lines, len(trows))
        section("JUDGES NOT CHECKED IN", [r["name"] for r in _alpha(checkins["judges"], lambda r: r["name"])], len(checkins["judges"]))
    if feedback is not None:
        pick = (lambda r: r["short"]) if team_mode == "short" else (lambda r: r["full"])
        trows = [r for r in feedback["teams"] if r["owed"] > 0]
        jrows = [r for r in feedback["judges"] if r["owed"] > 0]
        tline = lambda r: f"{pick(r)} (missing {r['owed']} feedbacks)"
        jline = lambda r: f"{r['name']} (missing {r['owed']} feedbacks)"
        if arrange == "category":
            tl = _group_lines(trows, lambda r: r["owed"], lambda n: f"Missing {n} feedbacks", tline, pick, desc)
            jl = _group_lines(jrows, lambda r: r["owed"], lambda n: f"Missing {n} feedbacks", jline, lambda r: r["name"], desc)
        else:
            tl = [tline(r) for r in _alpha(trows, pick)]
            jl = [jline(r) for r in _alpha(jrows, lambda r: r["name"])]
        section("TEAMS WITH MISSING FEEDBACK", tl, len(trows))
        section("JUDGES WITH MISSING FEEDBACK", jl, len(jrows))
    return "\n".join(out).rstrip() + "\n"


# ------------------------------------------------------------------ web page
PAGE = """<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Missing check-ins and feedback</title><style>body{font-family:-apple-system,Segoe UI,Roboto,sans-serif;max-width:680px;margin:20px auto;padding:0 16px;background:#f5f5f5}
.card{background:#fff;border-radius:8px;padding:20px;box-shadow:0 1px 3px rgba(0,0,0,.1)}label{display:block;font-weight:600;margin:14px 0 4px}
label.i{font-weight:400;margin:6px 0}input[type=text],input[type=url],select,input[type=file]{width:100%;padding:9px;border:1px solid #ccc;border-radius:6px;box-sizing:border-box}
button{background:#28a745;color:#fff;border:0;padding:12px 24px;border-radius:6px;font-size:15px;cursor:pointer;margin-top:16px}
.err{background:#f8d7da;color:#721c24;padding:10px;border-radius:6px;margin-bottom:12px}.hint{font-size:12px;color:#888;margin:4px 0}h3{margin:22px 0 0}</style></head><body>
<p><a href="/">&larr; exporter</a></p><div class="card"><h2>Missing check-ins and missing feedback &rarr; .txt</h2>
{% if error %}<div class="err">{{ error }}</div>{% endif %}
<form method="post" enctype="multipart/form-data">
<label>Tab site URL</label><input type="url" name="url" value="{{ f.url }}" placeholder="https://yoursite.calicotab.com">
<label>Tournament slug (optional)</label><input type="text" name="slug" value="{{ f.slug }}"><p class="hint">Leave blank if the URL already ends with the slug or the site has one tournament.</p>
<label>Admin API token</label><input type="text" name="token" autocomplete="off"><p class="hint">Needed for check-ins. Not needed if you only export missing feedback.</p>
<h3>What to export</h3>
<label class="i"><input type="checkbox" name="what_checkins" {{ 'checked' if f.what_checkins else '' }}> Teams and judges NOT checked in</label>
<label class="i"><input type="checkbox" name="what_feedback" {{ 'checked' if f.what_feedback else '' }}> Teams and judges with MISSING FEEDBACK</label>
<label>Format (for the note next to the teams)</label>
<select name="fmt"><option value="3v3" {{ 'selected' if f.fmt=='3v3' else '' }}>3v3 / WSDC &mdash; add "(2/3 debaters not checked in)"</option>
<option value="bp" {{ 'selected' if f.fmt=='bp' else '' }}>BP &mdash; no note next to the teams</option></select>
<label>Team names</label><select name="team_mode"><option value="long" {{ 'selected' if f.team_mode=='long' else '' }}>Full name</option><option value="short" {{ 'selected' if f.team_mode=='short' else '' }}>Short name</option></select>
<label>Arrange</label><select name="arrange"><option value="alpha" {{ 'selected' if f.arrange=='alpha' else '' }}>Alphabetical</option>
<option value="category" {{ 'selected' if f.arrange=='category' else '' }}>Group by number missing first, then alphabetical</option></select>
<label>Group order</label><select name="order"><option value="asc" {{ 'selected' if f.order=='asc' else '' }}>Fewest missing first (1/3, 2/3, all)</option><option value="desc" {{ 'selected' if f.order=='desc' else '' }}>Most missing first</option></select>
<p class="hint">Teams and judges are always listed separately.</p>
<h3>Missing feedback: where to read it</h3>
<p class="hint">Tabbycat's API cannot list missing feedback, so the numbers are read from the Feedback Progress page.</p>
<label>Saved Feedback Progress page (.html, optional)</label><input type="file" name="fb_file" accept=".html,.htm">
<p class="hint">Admin site: Feedback &rarr; Progress, press Ctrl+S, upload that file. Leave empty to read the public page <code>/slug/feedback/progress/</code> instead (only works when it is public).</p>
<button type="submit">Download .txt</button></form></div></body></html>"""


def register_missing_lists(app):
    from flask import Response, render_template_string, request

    @app.route("/missing", methods=["GET", "POST"])
    def missing_page():
        if request.method == "GET":
            f = {"url": "", "slug": "", "what_checkins": True, "what_feedback": False, "fmt": "3v3", "team_mode": "long",
                 "arrange": "alpha", "order": "asc"}
            return render_template_string(PAGE, f=f, error=None)
        form = request.form
        f = {"url": form.get("url", ""), "slug": form.get("slug", ""), "what_checkins": bool(form.get("what_checkins")),
             "what_feedback": bool(form.get("what_feedback")), "fmt": "bp" if form.get("fmt") == "bp" else "3v3",
             "team_mode": "short" if form.get("team_mode") == "short" else "long",
             "arrange": "category" if form.get("arrange") == "category" else "alpha",
             "order": "desc" if form.get("order") == "desc" else "asc"}
        try:
            if not (f["what_checkins"] or f["what_feedback"]):
                raise ValueError("Tick at least one thing to export.")
            if not f["url"].strip():
                raise ValueError("The tab site URL is needed.")
            checkins = feedback = None
            slug = f["slug"]
            if f["what_checkins"]:
                if not form.get("token", "").strip():
                    raise ValueError("The admin API token is needed to read the check-ins.")
                checkins = missing_checkins(f["url"], form.get("token", ""), f["slug"], f["team_mode"])
                slug = checkins["slug"]
            if f["what_feedback"]:
                up = request.files.get("fb_file")
                html = up.read().decode("utf-8", errors="ignore") if up and up.filename else fetch_feedback_page(f["url"], slug)
                feedback = parse_feedback_progress(html)
            text = build_text(checkins, feedback, f["fmt"], f["arrange"], f["order"] == "desc", f["team_mode"])
        except (ValueError, RuntimeError) as e:
            return render_template_string(PAGE, f=f, error=str(e))
        name = (slug or "tab") + "_missing_" + "_".join(k for k, v in (("checkins", f["what_checkins"]), ("feedback", f["what_feedback"])) if v) + ".txt"
        headers = {"Content-Disposition": f"attachment; filename={name}"}
        if checkins:
            headers["X-Unreadable-Checkins"] = str(checkins["unread"])
        return Response(text.encode("utf-8"), mimetype="text/plain; charset=utf-8", headers=headers)
