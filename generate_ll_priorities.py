#!/usr/bin/env python3
"""
LL Team Priorities Dashboard Generator — LAUSD Brand-Aligned
Generates a printable HTML dashboard from the Gantt Activities tab.
Aligned with LAUSD Brand Guide (Jan 2026): Poppins typography, navy primary (#00237A),
secondary palette of green/orange/gold, Open Sans for body copy.
"""

import pandas as pd
from datetime import datetime
from pathlib import Path
import sys
import re

# ---------- BRAND ASSET LOADING ----------
def load_svg_asset(svg_path, namespace=None):
    """Load an LAUSD brand SVG (wordmark, seal, or combo) and return clean inline markup.
    Strips XML declaration and SVG generator comments so it can be embedded directly into HTML.

    If namespace is provided, the SVG's internal CSS classes (.cls-1, .cls-2, etc.) are
    prefixed so multiple SVGs from the same Adobe Illustrator export source can coexist
    in one HTML document without style collisions.
    """
    if not svg_path or not Path(svg_path).exists():
        return None
    try:
        with open(svg_path, 'r', encoding='utf-8') as f:
            svg = f.read()
        # Strip <?xml ... ?> declaration
        svg = re.sub(r'<\?xml[^?]*\?>\s*', '', svg)
        # Strip HTML comments (e.g. Adobe Illustrator generator notes)
        svg = re.sub(r'<!--[\s\S]*?-->', '', svg)
        # Strip leading/trailing whitespace
        svg = svg.strip()

        # Namespace CSS classes to prevent collisions when embedding multiple SVGs
        if namespace:
            # Rename "cls-N" → "{namespace}-cls-N" in BOTH the <style> definitions
            # AND every class="cls-N" attribute on shapes
            def rename_class(match):
                return f'{namespace}-cls-{match.group(1)}'
            # Handle .cls-N in style blocks (e.g. ".cls-1{fill:#009fe2;}")
            svg = re.sub(r'\.cls-(\d+)', lambda m: f'.{namespace}-cls-{m.group(1)}', svg)
            # Handle class="cls-N" or class="cls-N cls-M" on elements
            def rename_in_class_attr(m):
                classes = m.group(1).split()
                renamed = [f'{namespace}-{c}' if c.startswith('cls-') else c for c in classes]
                return f'class="{" ".join(renamed)}"'
            svg = re.sub(r'class="([^"]+)"', rename_in_class_attr, svg)
            # Also handle .st0, .st1 (Illustrator alternative naming for some exports)
            svg = re.sub(r'\.st(\d+)', lambda m: f'.{namespace}-st{m.group(1)}', svg)
            svg = re.sub(r'class="(st\d+(?:\s+st\d+)*)"',
                         lambda m: 'class="' + ' '.join(f'{namespace}-{c}' for c in m.group(1).split()) + '"',
                         svg)

        return svg
    except Exception as e:
        print(f"⚠️  Could not load SVG at {svg_path}: {e}")
        return None


# Backwards-compatible alias
load_wordmark = load_svg_asset


# ---------- LAUSD BRAND TOKENS (from official Brand Guide, Jan 2026) ----------
LAUSD = {
    # Primary
    "navy": "#00237A",          # Primary — dominant
    "blue_bright": "#0089FF",   # Primary bright blue
    "white": "#FFFFFF",
    # Secondary
    "green": "#00602D",
    "gold": "#FF9C00",
    "orange": "#FF4D00",
    "red": "#FF0000",
    # Tertiary blues
    "blue_005": "#005FD3",
    "blue_009": "#009DDB",
    "blue_5B": "#5B91D8",
    "blue_6F": "#6FAFFF",
    "blue_ABD": "#ABDEFF",
    "blue_B3D": "#B3D4FF",
    "blue_D0D": "#D0DDFF",
    "blue_DCF": "#DCF0FF",
    # Tertiary other
    "green_light": "#7BC2B2",
    "green_mid": "#008063",
    "green_bright": "#008E47",
    "peach": "#FFD1AD",
    "orange_light": "#FF7D4A",
    "coral": "#FF5050",
    "yellow": "#FCCC00",
    "yellow_light": "#FFF3B6",
    # Neutrals
    "ink": "#333F49",
    "slate": "#4E6267",
    "gray_mid": "#80898C",
    "gray": "#9EA3A4",
    "gray_light": "#C0C1C3",
    "gray_lightest": "#E1E1E5",
}

# Workstream colors — each pulled from LAUSD tertiary palette, navy stays primary
WORKSTREAM_COLORS = {
    "CTE advising":                              LAUSD["blue_B3D"],
    "WBL plan & partnerships":                   LAUSD["green_light"],
    "Funding & grants":                          LAUSD["peach"],
    "PBL & high quality instruction":            LAUSD["blue_ABD"],
    "Data, systems & improvement":               LAUSD["yellow_light"],
    "Work readiness":                            LAUSD["blue_DCF"],
    "Professional learning & adult capacity":    LAUSD["blue_D0D"],
}

# Darker variants for the workstream label cell border-left accent
WORKSTREAM_ACCENTS = {
    "CTE advising":                              LAUSD["blue_005"],
    "WBL plan & partnerships":                   LAUSD["green_mid"],
    "Funding & grants":                          LAUSD["orange"],
    "PBL & high quality instruction":            LAUSD["blue_009"],
    "Data, systems & improvement":               LAUSD["gold"],
    "Work readiness":                            LAUSD["blue_5B"],
    "Professional learning & adult capacity":    LAUSD["navy"],
}

MONTHS = {1:"July", 2:"August", 3:"September", 4:"October", 5:"November", 6:"December",
          7:"January", 8:"February", 9:"March", 10:"April", 11:"May", 12:"June"}

WORKSTREAM_ORDER = [
    "CTE advising",
    "WBL plan & partnerships",
    "Funding & grants",
    "PBL & high quality instruction",
    "Data, systems & improvement",
    "Work readiness",
    "Professional learning & adult capacity",
]

TEAM_SYMBOLS = {"Coaches": "🌱", "Coordinators": "⚙️", "Both": "👥"}


def read_gantt(path):
    df = pd.read_excel(path, sheet_name="Activities", dtype=str)
    return df.dropna(how="all")


def parse_months(row):
    custom = row.get("Custom months", "")
    if pd.notna(custom) and str(custom).strip():
        try:
            return [int(m.strip()) for m in str(custom).split(",")]
        except Exception:
            pass
    try:
        start = int(row.get("Start month", 0))
        end = int(row.get("End month", 0))
        if start > 0 and end > 0:
            return list(range(start, end + 1)) if start <= end else list(range(start, 13)) + list(range(1, end + 1))
    except Exception:
        pass
    return []


def effort_dots(effort):
    if pd.isna(effort) or not str(effort).strip():
        return ""
    try:
        return "●" * int(effort)
    except Exception:
        return ""


def team_symbol(team):
    if pd.isna(team):
        return ""
    return TEAM_SYMBOLS.get(str(team).strip(), "")


def build_data(df):
    data = {ws: {m: [] for m in range(1, 13)} for ws in WORKSTREAM_ORDER}
    effort = {"Coaches": {}, "Coordinators": {}, "Overall": {}}

    for _, row in df.iterrows():
        ws = str(row.get("Workstream", "")).strip()
        activity = str(row.get("Activity", "")).strip()
        if not activity or activity == "nan" or ws not in data:
            continue

        eff_raw = str(row.get("Effort (1–4)", "")).strip()
        milestone = str(row.get("Milestone?", "")).strip().lower() == "yes"
        who = str(row.get("Who", "")).strip()
        link = str(row.get("Link", "")).strip()
        notes = str(row.get("Notes", "")).strip()

        for m in parse_months(row):
            if 1 <= m <= 12:
                data[ws][m].append({
                    "activity": activity,
                    "effort": effort_dots(eff_raw),
                    "milestone": milestone,
                    "who_symbol": team_symbol(who),
                    "who": who,
                    "link": link if link and link.lower() != "nan" else "",
                    "notes": notes if notes and notes.lower() != "nan" else "",
                })

                if not milestone:
                    try:
                        e = int(eff_raw) if eff_raw and eff_raw.lower() != "nan" else 0
                        if who in ("Coaches", "Both"):
                            effort["Coaches"][m] = effort["Coaches"].get(m, 0) + e
                        if who in ("Coordinators", "Both"):
                            effort["Coordinators"][m] = effort["Coordinators"].get(m, 0) + e
                        effort["Overall"][m] = effort["Overall"].get(m, 0) + e
                    except Exception:
                        pass

    return data, effort


def heat_color(val, maxv):
    """Map an effort load value 0..maxv to a color from LAUSD palette (white -> navy)."""
    if maxv == 0 or val == 0:
        return LAUSD["white"]
    ratio = min(val / maxv, 1.0)
    if ratio < 0.25:
        return LAUSD["blue_DCF"]
    elif ratio < 0.5:
        return LAUSD["blue_ABD"]
    elif ratio < 0.75:
        return LAUSD["blue_5B"]
    else:
        return LAUSD["blue_005"]


def generate_html(data, effort, source_file, wordmark_svg=None, seal_svg=None):
    today = datetime.now().strftime("%B %d, %Y")
    max_effort = max([max(v.values()) if v else 0 for v in effort.values()] + [1])

    css = f"""
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700;800&family=Open+Sans:wght@400;500;600;700&display=swap');

    :root {{
      --navy: {LAUSD['navy']};
      --blue-bright: {LAUSD['blue_bright']};
      --ink: {LAUSD['ink']};
      --slate: {LAUSD['slate']};
      --gray: {LAUSD['gray_mid']};
      --gray-light: {LAUSD['gray_lightest']};
      --gold: {LAUSD['gold']};
      --green: {LAUSD['green']};
      --orange: {LAUSD['orange']};
      --white: #FFFFFF;
    }}

    * {{ margin: 0; padding: 0; box-sizing: border-box; }}

    html, body {{
      font-family: 'Open Sans', Arial, sans-serif;
      color: var(--ink);
      background: #F4F6F9;
      font-size: 13px;
      line-height: 1.4;
    }}

    h1, h2, h3, h4, .display {{
      font-family: 'Poppins', Arial, sans-serif;
      color: var(--navy);
      font-weight: 700;
      letter-spacing: -0.01em;
    }}

    .page {{
      max-width: 1500px;
      margin: 20px auto;
      background: var(--white);
      box-shadow: 0 2px 12px rgba(0,35,122,0.08);
    }}

    /* ============ HEADER BAND ============ */
    .header-band {{
      background: var(--navy);
      color: var(--white);
      padding: 28px 40px 24px;
      display: flex;
      align-items: center;
      gap: 28px;
    }}
    .header-seal {{
      flex-shrink: 0;
      width: 96px;
      height: 96px;
      display: block;
      filter: drop-shadow(0 2px 6px rgba(0,0,0,0.25));
    }}
    .header-seal svg {{
      width: 100%;
      height: 100%;
      display: block;
    }}
    .header-center {{
      flex: 1;
      min-width: 0;
    }}
    .header-center h1 {{
      color: var(--white);
      font-family: 'Poppins', sans-serif;
      font-weight: 800;
      font-size: 30px;
      line-height: 1.1;
      letter-spacing: -0.02em;
    }}
    .header-center .sub {{
      font-family: 'Open Sans', sans-serif;
      font-size: 14px;
      margin-top: 6px;
      color: var(--white);
      opacity: 0.92;
    }}
    .header-right {{
      text-align: right;
      flex-shrink: 0;
    }}
    .lausd-wordmark {{
      width: 160px;
      height: auto;
      display: block;
      margin-left: auto;
      filter: drop-shadow(0 1px 2px rgba(0,0,0,0.15));
    }}
    .lausd-wordmark svg {{
      width: 100%;
      height: auto;
      display: block;
    }}
    .header-right .meta {{
      font-family: 'Open Sans', sans-serif;
      font-size: 11px;
      margin-top: 10px;
      opacity: 0.85;
      color: var(--white);
      text-align: right;
    }}

    /* ============ DEPARTMENT BAR ============ */
    .dept-bar {{
      background: var(--blue-bright);
      color: var(--white);
      padding: 8px 40px;
      font-family: 'Poppins', sans-serif;
      font-weight: 600;
      font-size: 12px;
      letter-spacing: 0.08em;
      text-transform: uppercase;
    }}

    /* ============ INTRO / LEGEND ============ */
    .intro {{
      padding: 24px 40px 8px;
    }}
    .intro h2 {{
      font-size: 18px;
      margin-bottom: 6px;
    }}
    .intro p {{
      color: var(--slate);
      font-size: 13px;
      max-width: 900px;
    }}

    .legend {{
      display: grid;
      grid-template-columns: repeat(2, 1fr);
      gap: 24px;
      padding: 20px 40px;
      margin: 16px 40px 24px;
      background: #F8FAFD;
      border-left: 4px solid var(--navy);
      border-radius: 0 4px 4px 0;
    }}
    .legend-section h4 {{
      font-size: 11px;
      letter-spacing: 0.12em;
      text-transform: uppercase;
      color: var(--navy);
      margin-bottom: 10px;
      font-family: 'Poppins', sans-serif;
      font-weight: 700;
    }}
    .legend-row {{ display: flex; flex-wrap: wrap; gap: 12px 18px; font-size: 12px; align-items: center; }}
    .legend-item {{ display: flex; align-items: center; gap: 6px; }}
    .legend-swatch {{
      width: 14px; height: 14px; border-radius: 3px;
      border: 1px solid rgba(0,35,122,0.12);
    }}

    /* ============ MAIN TABLE ============ */
    .table-wrap {{
      padding: 0 40px 24px;
      overflow-x: auto;
    }}
    table.grid {{
      width: 100%;
      border-collapse: separate;
      border-spacing: 0;
      font-size: 10.5px;
      table-layout: fixed;
    }}
    table.grid thead th {{
      background: var(--navy);
      color: var(--white);
      font-family: 'Poppins', sans-serif;
      font-weight: 600;
      font-size: 11px;
      padding: 10px 8px;
      text-align: left;
      letter-spacing: 0.04em;
      text-transform: uppercase;
      border-right: 1px solid rgba(255,255,255,0.12);
      vertical-align: middle;
    }}
    table.grid thead th:last-child {{ border-right: none; }}
    table.grid thead th.month {{ text-align: center; width: 7.3%; }}
    table.grid thead th.workstream-col {{ width: 11.5%; }}

    table.grid td {{
      border-bottom: 1px solid var(--gray-light);
      border-right: 1px solid var(--gray-light);
      padding: 6px 5px;
      vertical-align: top;
      background: var(--white);
    }}
    table.grid tr:last-child td {{ border-bottom: 1px solid var(--gray); }}

    td.workstream-label {{
      font-family: 'Poppins', sans-serif;
      font-weight: 600;
      font-size: 11.5px;
      color: var(--navy);
      padding: 12px 10px;
      vertical-align: middle;
      line-height: 1.25;
    }}

    .activity-card {{
      padding: 6px 7px;
      border-radius: 3px;
      margin-bottom: 5px;
      font-size: 10px;
      line-height: 1.3;
      border-left: 2px solid var(--navy);
    }}
    .activity-card:last-child {{ margin-bottom: 0; }}
    .activity-card .title {{
      font-weight: 600;
      color: var(--ink);
      display: block;
    }}
    .activity-card .meta {{
      display: flex;
      align-items: center;
      gap: 6px;
      margin-top: 3px;
      font-size: 10px;
    }}
    .activity-card .dots {{
      color: var(--navy);
      letter-spacing: 1px;
      font-size: 9px;
    }}
    .activity-card .symbol {{ font-size: 11px; }}
    .activity-card a.link {{
      display: inline-block;
      color: var(--navy);
      text-decoration: none;
      font-weight: 600;
      font-size: 9.5px;
      margin-top: 2px;
    }}
    .activity-card a.link:hover {{ text-decoration: underline; }}

    .activity-card.milestone {{
      background: var(--navy) !important;
      border-left: 2px solid var(--gold);
      color: var(--white);
    }}
    .activity-card.milestone .title,
    .activity-card.milestone .dots,
    .activity-card.milestone a.link {{ color: var(--white); }}
    .activity-card.milestone::before {{
      content: "★ ";
      color: var(--gold);
      font-size: 11px;
    }}

    /* ============ SUMMARY TABLE ============ */
    .summary-section {{ padding: 8px 40px 36px; }}
    .summary-section h2 {{
      font-size: 20px;
      margin-bottom: 4px;
    }}
    .summary-section .sub {{
      color: var(--slate);
      font-size: 12px;
      margin-bottom: 16px;
    }}
    table.summary {{
      width: 100%;
      border-collapse: separate;
      border-spacing: 0;
      font-family: 'Poppins', sans-serif;
      font-size: 12px;
    }}
    table.summary thead th {{
      background: var(--navy);
      color: var(--white);
      padding: 10px 8px;
      font-weight: 600;
      font-size: 11px;
      letter-spacing: 0.05em;
      text-align: center;
      text-transform: uppercase;
      border-right: 1px solid rgba(255,255,255,0.12);
    }}
    table.summary thead th:first-child {{ text-align: left; padding-left: 14px; }}
    table.summary tbody td {{
      padding: 12px 8px;
      text-align: center;
      font-weight: 600;
      color: var(--navy);
      border-right: 1px solid var(--gray-light);
      border-bottom: 1px solid var(--gray-light);
    }}
    table.summary tbody td.team-name {{
      text-align: left;
      padding-left: 14px;
      background: #F8FAFD;
      color: var(--navy);
      font-weight: 700;
    }}

    /* ============ FOOTER ============ */
    .footer {{
      background: var(--navy);
      color: var(--white);
      padding: 18px 40px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      font-family: 'Open Sans', sans-serif;
      font-size: 11px;
    }}
    .footer .tagline {{
      font-family: 'Poppins', sans-serif;
      font-weight: 700;
      letter-spacing: 0.1em;
      font-size: 10px;
      text-transform: uppercase;
      opacity: 0.85;
    }}
    .footer a.print-btn {{
      background: var(--gold);
      color: var(--navy);
      padding: 8px 16px;
      border-radius: 3px;
      font-family: 'Poppins', sans-serif;
      font-weight: 700;
      font-size: 11px;
      letter-spacing: 0.05em;
      text-transform: uppercase;
      text-decoration: none;
      cursor: pointer;
    }}

    /* ============ PRINT ============ */
    @media print {{
      body {{ background: var(--white); }}
      .page {{ box-shadow: none; margin: 0; max-width: 100%; }}
      .footer a.print-btn {{ display: none; }}
      @page {{ size: landscape; margin: 0.4in; }}
      table.grid {{ page-break-inside: avoid; }}
      .summary-section {{ page-break-before: auto; }}
      .activity-card a.link {{
        font-size: 8.5px;
        color: var(--navy) !important;
      }}
    }}
    """

    parts = []
    # Build wordmark markup — official SVG if available, text fallback otherwise
    if wordmark_svg:
        wordmark_html = f'<div class="lausd-wordmark">{wordmark_svg}</div>'
    else:
        wordmark_html = ('<div class="lausd-wordmark" style="font-family:Poppins,sans-serif;'
                         'font-weight:800;font-size:26px;color:#fff;letter-spacing:-0.02em;'
                         'line-height:1;">LAUSD<span style="display:block;font-size:11px;'
                         'font-weight:700;letter-spacing:0.18em;margin-top:2px;opacity:0.85;">'
                         'UNIFIED · READY FOR THE WORLD</span></div>')

    # Build seal markup — official SVG if available, omit otherwise
    seal_html = f'<div class="header-seal">{seal_svg}</div>' if seal_svg else ''

    parts.append(f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Linked Learning Team Priorities | LAUSD</title>
<style>{css}</style>
</head>
<body>
<div class="page">

  <!-- HEADER BAND -->
  <div class="header-band">
    {seal_html}
    <div class="header-center">
      <h1>Linked Learning Team Priorities</h1>
      <div class="sub">2026–2027 School Year · Team Capacity &amp; Workstream Calendar</div>
    </div>
    <div class="header-right">
      {wordmark_html}
      <div class="meta">Updated {today}</div>
    </div>
  </div>

  <!-- DEPARTMENT BAR -->
  <div class="dept-bar">
    Division of Instruction · CTE–Linked Learning Department
  </div>

  <!-- INTRO -->
  <div class="intro">
    <h2>Annual Workstream Overview</h2>
    <p>This dashboard shows what the Linked Learning team is focused on each month of the school year. Use it to plan initiatives, anticipate capacity, and align departmental priorities. The summary at the bottom shows total team effort load by month.</p>
  </div>

  <!-- LEGEND -->
  <div class="legend">
    <div class="legend-section">
      <h4>Workstreams</h4>
      <div class="legend-row">""")

    for ws in WORKSTREAM_ORDER:
        parts.append(f'<div class="legend-item"><span class="legend-swatch" style="background:{WORKSTREAM_COLORS[ws]};border-left:3px solid {WORKSTREAM_ACCENTS[ws]};"></span> {ws}</div>')

    parts.append(f"""
      </div>
    </div>
    <div class="legend-section">
      <h4>Team &amp; Effort</h4>
      <div class="legend-row">
        <div class="legend-item">🌱 Coaches</div>
        <div class="legend-item">⚙️ Coordinators</div>
        <div class="legend-item">👥 Both</div>
        <div class="legend-item"><span style="color:{LAUSD['navy']};letter-spacing:1px;">●</span> = light · <span style="color:{LAUSD['navy']};letter-spacing:1px;">●●●●</span> = peak effort</div>
        <div class="legend-item"><span style="color:{LAUSD['gold']};">★</span> Milestone / hard deadline</div>
      </div>
    </div>
  </div>

  <!-- MAIN GRID TABLE -->
  <div class="table-wrap">
    <table class="grid">
      <thead>
        <tr>
          <th class="workstream-col">Workstream</th>""")
    for m in range(1, 13):
        parts.append(f'<th class="month">{MONTHS[m]}</th>')
    parts.append('</tr></thead><tbody>')

    for ws in WORKSTREAM_ORDER:
        bg = WORKSTREAM_COLORS[ws]
        accent = WORKSTREAM_ACCENTS[ws]
        parts.append(f'<tr>')
        parts.append(f'<td class="workstream-label" style="background:{bg};border-left:5px solid {accent};">{ws}</td>')

        for m in range(1, 13):
            items = data[ws][m]
            parts.append('<td>')
            for it in items:
                klass = "activity-card milestone" if it["milestone"] else "activity-card"
                style = "" if it["milestone"] else f"background:{bg};border-left-color:{accent};"
                parts.append(f'<div class="{klass}" style="{style}">')
                parts.append(f'<span class="title">{it["activity"]}</span>')
                parts.append('<div class="meta">')
                if it["who_symbol"]:
                    parts.append(f'<span class="symbol" title="{it["who"]}">{it["who_symbol"]}</span>')
                if it["effort"]:
                    parts.append(f'<span class="dots">{it["effort"]}</span>')
                parts.append('</div>')
                if it["link"]:
                    parts.append(f'<a href="{it["link"]}" target="_blank" class="link">↗ Open link</a>')
                parts.append('</div>')
            parts.append('</td>')
        parts.append('</tr>')

    parts.append('</tbody></table></div>')

    # --- SUMMARY ---
    parts.append(f"""
  <!-- SUMMARY -->
  <div class="summary-section">
    <h2>Team Capacity Load by Month</h2>
    <div class="sub">Total effort points across all active workstreams. Higher numbers indicate heavier capacity demands on the team.</div>
    <table class="summary">
      <thead>
        <tr><th>Team</th>""")
    for m in range(1, 13):
        parts.append(f'<th>{MONTHS[m][:3]}</th>')
    parts.append('</tr></thead><tbody>')

    for team in ("Coaches", "Coordinators", "Overall"):
        parts.append(f'<tr><td class="team-name">{team}</td>')
        for m in range(1, 13):
            v = effort[team].get(m, 0)
            color = heat_color(v, max_effort)
            text_color = LAUSD["white"] if v / max_effort > 0.6 else LAUSD["navy"]
            parts.append(f'<td style="background:{color};color:{text_color};">{v}</td>')
        parts.append('</tr>')
    parts.append('</tbody></table></div>')

    # --- FOOTER ---
    parts.append(f"""
  <div class="footer">
    <div>
      <div class="tagline">Ready for the World</div>
      <div style="opacity:0.75;margin-top:4px;font-size:10px;">Generated {today} · Source: LL Team Gantt (Activities tab)</div>
    </div>
    <a class="print-btn" href="javascript:window.print()">Print / Save as PDF</a>
  </div>

</div>
</body>
</html>""")

    return "".join(parts)


if __name__ == "__main__":
    gantt = sys.argv[1] if len(sys.argv) > 1 else "LL_Team_Gantt_Heatmap_050126.xlsm"
    out = sys.argv[2] if len(sys.argv) > 2 else "LL_Team_Priorities_Dashboard.html"
    wordmark_path = sys.argv[3] if len(sys.argv) > 3 else "LAUSD_wordmark_RGB.svg"
    seal_path = sys.argv[4] if len(sys.argv) > 4 else "LAUSD_seal.svg"

    print(f"📂 Reading {gantt}…")
    df = read_gantt(gantt)
    print("🔧 Processing activities…")
    data, effort = build_data(df)

    wordmark = load_svg_asset(wordmark_path, namespace='wm')
    if wordmark:
        print(f"🏷️  Embedding official LAUSD wordmark from {wordmark_path}")
    else:
        print(f"⚠️  Wordmark file not found at {wordmark_path} — using text fallback")

    seal = load_svg_asset(seal_path, namespace='sl')
    if seal:
        print(f"🛡️  Embedding official LAUSD seal from {seal_path}")
    else:
        print(f"ℹ️  Seal file not found at {seal_path} — proceeding without seal")

    print("🎨 Generating brand-aligned HTML…")
    html = generate_html(data, effort, gantt, wordmark_svg=wordmark, seal_svg=seal)
    with open(out, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"✅ Dashboard ready: {out}")
