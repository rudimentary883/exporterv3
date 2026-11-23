"""
Team awards + Speaker awards  ->  .csv  (page /awards-full on the exporter)

TEAM AWARDS   columns  Team, Award, Category, Institution, Country
  * every BREAKING team of EVERY break category (one row per category it broke in)
  * Award    = the last outround the team reached in that category, e.g. "Quarterfinals", "Semifinals", "Finals"
               (the round's name without the category name in front: "Open Pre-Quarterfinals" -> "Pre-Quarterfinals").
               A team is "in" a round when it is in a debate of that round and the draw is Confirmed / Released (option: drafts too).
  * Category = the break category;  Institution = the team's institution (full name or code);  Country = the institution's Region on the tab

SPEAKER AWARDS   columns  Speaker, Award, Category, Team, Institution, Country
  * for the Open tab and for every speaker category: the top N speakers of that tab, ties included
    Award = "Overall Best Speaker" (1st), "2nd Best Speaker", "3rd Best Speaker" ...;  tied speakers "Joint 2nd Best Speaker" ...
  * N = 10 for every tab, with exceptions typed as  11-and-Under=4, Northeast Asia=3
  * not included yet: "Finals Best Speaker" and "Outstanding Speaker" (see the notes in the answer)

Read-only: nothing is written to the tab.   Needs the 2 lines at the bottom of app.py:
    from awards_full import register_awards_full
    register_awards_full(app, TabbycatAPI, _collect_award_rows, pick_team_name, normalize_format)
"""
import csv
import io
import re
import zipfile


# ------------------------------------------------------------------ small helpers
def _tail(ref):
    text = (ref.get("url") if isinstance(ref, dict) else str(ref or "")).rstrip("/")
    return text.rsplit("/", 1)[-1]


def _ord(n):
    if 10 <= n % 100 <= 20:
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suffix}"


def speaker_award_label(rank, tied):
    base = "Overall Best Speaker" if rank == 1 else f"{_ord(rank)} Best Speaker"
    return ("Joint " + base) if tied else base


def round_award(round_name, category_name):
    """'Open Pre-Quarterfinals' + 'Open' -> 'Pre-Quarterfinals'"""
    name = (round_name or "").strip()
    cat = (category_name or "").strip()
    if cat and name.casefold().startswith(cat.casefold()):
        rest = name[len(cat):].strip(" -:\u2013\u2014")
        if rest:
            return rest
    return name


def parse_top_n(text, default):
    """'11-and-Under=4, Northeast Asia=3'  ->  ({'11-and-under': 4, 'northeast asia': 3}, default)"""
    out = {}
    for part in re.split(r"[,;\n]+", str(text or "")):
        if "=" in part:
            name, _, n = part.rpartition("=")
            if name.strip() and n.strip().isdigit():
                out[name.strip().casefold()] = int(n.strip())
    return out


def _institutions(api):
    data = api._get_list(f"{api.base_url}/api/v1/institutions")
    if isinstance(data, dict) and "_error" in data:
        api._log(f"Institutions could not be read: {data.get('_error')} status={data.get('_status')}")
        return {}, {}
    by_id, region_by_name = {}, {}
    for inst in data:
        if isinstance(inst, dict):
            if inst.get("id") is not None:
                by_id[str(inst["id"])] = inst
            by_id[_tail(inst)] = inst
            if inst.get("name"):
                region_by_name[str(inst["name"]).strip().casefold()] = (inst.get("region") or "").strip()
    return by_id, region_by_name


def _inst_fields(team, by_id, mode):
    inst = by_id.get(_tail(team.get("institution"))) if team and team.get("institution") else None
    if not inst:
        return "", ""
    label = (inst.get("code") if mode == "code" else inst.get("name")) or inst.get("name") or inst.get("code") or ""
    return str(label).strip(), (inst.get("region") or "").strip()


def _csv(header, rows):
    out = io.StringIO()
    w = csv.writer(out, lineterminator="\r\n")
    w.writerow(header)
    w.writerows(rows)
    return out.getvalue()


# ------------------------------------------------------------------ team awards
def build_team_awards(api, team_name_source="reference", inst_mode="name", include_drafts=False, furthest_first=True):
    notes = []
    cats = api.get_break_categories()
    if not cats:
        return None, ["No break categories were found on the tab."]
    rounds = api._get_list(api._url("/rounds"))
    if isinstance(rounds, dict):
        return None, [f"The rounds could not be read: {rounds.get('_error')} status={rounds.get('_status')}"]
    teams = {str(t.get("id")): t for t in api.get_teams()}
    teams.update({_tail(t): t for t in teams.values()})
    by_id, _ = _institutions(api)
    # outrounds of each break category, in creation order
    out_rounds = {}
    for r in sorted(rounds, key=lambda r: r.get("seq") or 0):
        if (r.get("stage") == "E") if "stage" in r else bool(r.get("break_category")):
            out_rounds.setdefault(_tail(r.get("break_category")), []).append(r)
    rows, n_unplaced = [], 0
    for cat in cats:
        cid, cname = _tail(cat.get("url") or cat.get("id")), cat.get("name") or cat.get("slug") or ""
        entries = api.get_breaking_teams(cid)
        breaking = [e for e in entries if isinstance(e, dict) and e.get("break_rank") is not None]
        # round index of every team in every outround of this category
        reached = {}
        used_rounds = []
        for idx, r in enumerate(out_rounds.get(cid, [])):
            if r.get("draw_status") not in ("C", "R") and not include_drafts:
                continue
            pairing_url = ((r.get("_links") or {}).get("pairing")) or ((r.get("url") or "").rstrip("/") + "/pairings")
            debates = api._get_list(pairing_url)
            if isinstance(debates, dict):
                notes.append(f"{cname}: round {r.get('name') or r.get('seq')} could not be read ({debates.get('_error')}).")
                continue
            used_rounds.append(r)
            for d in debates:
                for t in d.get("teams") or []:
                    tid = _tail(t.get("team") if isinstance(t, dict) else t)
                    reached[tid] = idx
        names = {idx: round_award(r.get("name") or r.get("abbreviation"), cname) for idx, r in enumerate(out_rounds.get(cid, []))}
        cat_rows = []
        for e in breaking:
            tid = _tail(e.get("team"))
            team = teams.get(tid)
            if team is None:
                notes.append(f"{cname}: team {tid} is not in the teams list and was skipped.")
                continue
            idx = reached.get(tid)
            if idx is None:
                n_unplaced += 1
            inst_label, region = _inst_fields(team, by_id, inst_mode)
            from_team = api_team_name(team, team_name_source)
            cat_rows.append((idx if idx is not None else -1, from_team, [from_team, names.get(idx, "") if idx is not None else "", cname, inst_label, region]))
        cat_rows.sort(key=lambda x: (-x[0] if furthest_first else x[0], x[1].casefold()))
        rows += [r[2] for r in cat_rows]
        if breaking and not out_rounds.get(cid):
            notes.append(f"{cname}: no outrounds exist for this category yet - its teams have an empty Award.")
    if n_unplaced:
        notes.append(f"{n_unplaced} breaking team(s) were in no confirmed/released outround draw - their Award is empty.")
    return rows, notes


def api_team_name(team, source):
    order = {"short": ("short_name", "long_name", "reference"), "long": ("long_name", "short_name", "reference")}.get(
        source, ("reference", "short_name", "long_name"))
    for key in order:
        if team.get(key):
            return str(team[key]).strip()
    return str(team.get("id", "?"))


# ------------------------------------------------------------------ speaker awards
def build_speaker_awards(api, collect_rows, fmt, team_name_source, inst_mode, default_n, exceptions, only_tabs=None, reverse=False):
    notes = []
    cats = api.get_speaker_categories()
    if cats is None:
        return None, ["The speaker categories could not be read."]
    tabs = [("open", "Open")] + [((c.get("slug") or _tail(c.get("url"))), c.get("name") or c.get("slug") or "") for c in cats]
    if only_tabs:
        wanted = {t.casefold() for t in only_tabs}
        tabs = [t for t in tabs if t[0].casefold() in wanted or t[1].casefold() in wanted]
    _, region_by_name = _institutions(api)
    cache, rows = {}, []
    for slug, label in tabs:
        n = exceptions.get(label.casefold(), exceptions.get(slug.casefold(), default_n))
        got, error, info = collect_rows(api, slug, fmt, n, "upper", "joint", team_name_source, cache, inst_mode)
        if error:
            notes.append(f"{label}: skipped ({error[:160]})")
            continue
        part = []
        for r in got:
            region = region_by_name.get((r.get("institution") or "").strip().casefold(), "")
            part.append([r["speaker"], speaker_award_label(r["rank_num"], r["tied"]), label, r["team"], r.get("institution", ""), region])
        rows += part[::-1] if reverse else part
    return rows, notes


# ------------------------------------------------------------------ web page
PAGE = """<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Team awards and speaker awards</title><style>body{font-family:-apple-system,Segoe UI,Roboto,sans-serif;max-width:680px;margin:20px auto;padding:0 16px;background:#f5f5f5}
.card{background:#fff;border-radius:8px;padding:20px;box-shadow:0 1px 3px rgba(0,0,0,.1)}label{display:block;font-weight:600;margin:14px 0 4px}label.i{font-weight:400;margin:6px 0}
input[type=text],input[type=url],select,textarea{width:100%;padding:9px;border:1px solid #ccc;border-radius:6px;box-sizing:border-box}
button{background:#28a745;color:#fff;border:0;padding:12px 24px;border-radius:6px;font-size:15px;cursor:pointer;margin-top:16px}
.err{background:#f8d7da;color:#721c24;padding:10px;border-radius:6px;margin-bottom:12px}.hint{font-size:12px;color:#888;margin:4px 0}</style></head><body>
<p><a href="/">&larr; exporter</a></p><div class="card"><h2>Team awards and speaker awards &rarr; .csv</h2>
{% if error %}<div class="err">{{ error }}</div>{% endif %}
<form method="post"><label>Tab site URL</label><input type="url" name="url" value="{{ f.url }}" placeholder="https://yoursite.calicotab.com">
<label>Tournament slug</label><input type="text" name="slug" value="{{ f.slug }}"><p class="hint">As in the tab's address, e.g. sisdc2026 (use _ on single-tournament sites).</p>
<label>Admin API token</label><input type="text" name="token" autocomplete="off">
<label>Format</label><select name="fmt"><option value="wsdc" {{ 'selected' if f.fmt=='wsdc' else '' }}>WSDC</option><option value="3v3" {{ 'selected' if f.fmt=='3v3' else '' }}>3v3</option><option value="bp" {{ 'selected' if f.fmt=='bp' else '' }}>BP</option></select>
<label>Files</label>
<label class="i"><input type="checkbox" name="what" value="team" {{ 'checked' if 'team' in f.what else '' }}> Team awards (every breaking team, last outround reached)</label>
<label class="i"><input type="checkbox" name="what" value="speaker" {{ 'checked' if 'speaker' in f.what else '' }}> Speaker awards (Open tab + every speaker category)</label>
<label>Team names</label><select name="team_name"><option value="reference" {{ 'selected' if f.team_name=='reference' else '' }}>Team name only (reference)</option><option value="short" {{ 'selected' if f.team_name=='short' else '' }}>Short name</option><option value="long" {{ 'selected' if f.team_name=='long' else '' }}>Full name</option></select>
<label>Institution</label><select name="inst"><option value="name" {{ 'selected' if f.inst=='name' else '' }}>Full name</option><option value="code" {{ 'selected' if f.inst=='code' else '' }}>Code</option></select>
<label>Speaker awards: top N for every tab</label><input type="text" name="top_n" value="{{ f.top_n }}"><p class="hint">Ties at the cut-off are included.</p>
<label>Exceptions</label><input type="text" name="exceptions" value="{{ f.exceptions }}" placeholder="11-and-Under=4, Northeast Asia=3, Central Asia=3">
<label>Only these speaker tabs (optional)</label><input type="text" name="only_tabs" value="{{ f.only_tabs }}" placeholder="empty = Open + every speaker category">
<label>Order</label><select name="order"><option value="first" {{ 'selected' if f.order=='first' else '' }}>Best / furthest first</option><option value="last" {{ 'selected' if f.order=='last' else '' }}>Last place first (for announcing)</option></select>
<label class="i"><input type="checkbox" name="drafts" {{ 'checked' if f.drafts else '' }}> count outround draws that are not confirmed / released yet</label>
<button type="submit">Download</button></form></div></body></html>"""


def register_awards_full(app, TabbycatAPI, collect_rows, pick_team_name, normalize_format):
    from flask import Response, render_template_string, request

    @app.route("/awards-full", methods=["GET", "POST"])
    def awards_full_page():
        form = request.form
        f = {"url": form.get("url", ""), "slug": form.get("slug", ""), "fmt": form.get("fmt", "wsdc"),
             "what": form.getlist("what") if request.method == "POST" else ["team", "speaker"],
             "team_name": form.get("team_name", "reference"), "inst": form.get("inst", "name"), "top_n": form.get("top_n", "10") or "10",
             "exceptions": form.get("exceptions", ""), "only_tabs": form.get("only_tabs", ""), "order": form.get("order", "first"),
             "drafts": bool(form.get("drafts"))}
        if request.method == "GET":
            return render_template_string(PAGE, f=f, error=None)
        try:
            if not f["url"].strip() or not f["slug"].strip():
                raise ValueError("The tab site URL and the tournament slug are needed.")
            if not form.get("token", "").strip():
                raise ValueError("The admin API token is needed.")
            if not f["what"]:
                raise ValueError("Tick at least one file.")
            if not f["top_n"].isdigit() or int(f["top_n"]) < 1:
                raise ValueError("Top N must be a number, e.g. 10.")
            api = TabbycatAPI(f["url"], form.get("token", ""), f["slug"])
            files, notes = {}, []
            reverse = f["order"] == "last"
            if "team" in f["what"]:
                rows, n = build_team_awards(api, f["team_name"], f["inst"], f["drafts"], furthest_first=not reverse)
                notes += n
                if rows is None:
                    raise ValueError(" ".join(n))
                files["team_awards.csv"] = _csv(["Team", "Award", "Category", "Institution", "Country"], rows)
            if "speaker" in f["what"]:
                only = [t.strip() for t in re.split(r"[,;]", f["only_tabs"]) if t.strip()]
                rows, n = build_speaker_awards(api, collect_rows, normalize_format(f["fmt"]), f["team_name"], f["inst"], int(f["top_n"]),
                                               parse_top_n(f["exceptions"], int(f["top_n"])), only, reverse)
                notes += n
                if rows is None:
                    raise ValueError(" ".join(n))
                files["speaker_awards.csv"] = _csv(["Speaker", "Award", "Category", "Team", "Institution", "Country"], rows)
        except (ValueError, RuntimeError) as e:
            return render_template_string(PAGE, f=f, error=str(e))
        if len(files) == 1 and not notes:
            name, text = next(iter(files.items()))
            return Response(text.encode("utf-8"), mimetype="text/csv; charset=utf-8",
                            headers={"Content-Disposition": f"attachment; filename={f['slug']}_{name}"})
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            for name, text in files.items():
                z.writestr(name, text.encode("utf-8"))
            if notes:
                z.writestr("notes.txt", "\n".join(notes) + "\n")
        return Response(buf.getvalue(), mimetype="application/zip",
                        headers={"Content-Disposition": f"attachment; filename={f['slug']}_awards.zip"})
