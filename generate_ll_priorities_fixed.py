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
    "CTE Pathway, Scheduling & Funding":                 LAUSD["blue_B3D"],
    "Instructional Design, PBL & Defenses":             LAUSD["blue_ABD"],
    "Pathway Strategy, Data & Certification":           LAUSD["yellow_light"],
    "Professional Learning & Leadership":               LAUSD["blue_D0D"],
    "WBL Planning, Partnerships & Career Exposure":     LAUSD["green_light"],
    "Internships & Work Readiness":                     LAUSD["peach"],
    "WBL Operations, Logistics & Reporting":            LAUSD["blue_DCF"],
}

# Darker variants for the workstream label cell border-left accent
WORKSTREAM_ACCENTS = {
    "CTE Pathway, Scheduling & Funding":                 LAUSD["blue_005"],
    "Instructional Design, PBL & Defenses":             LAUSD["blue_009"],
    "Pathway Strategy, Data & Certification":           LAUSD["gold"],
    "Professional Learning & Leadership":               LAUSD["navy"],
    "WBL Planning, Partnerships & Career Exposure":     LAUSD["green_mid"],
    "Internships & Work Readiness":                     LAUSD["orange"],
    "WBL Operations, Logistics & Reporting":            LAUSD["blue_5B"],
}

MONTHS = {1:"July", 2:"August", 3:"September", 4:"October", 5:"November", 6:"December",
          7:"January", 8:"February", 9:"March", 10:"April", 11:"May", 12:"June"}

WORKSTREAM_ORDER = [
    "CTE Pathway, Scheduling & Funding",
    "Instructional Design, PBL & Defenses",
    "Pathway Strategy, Data & Certification",
    "Professional Learning & Leadership",
    "WBL Planning, Partnerships & Career Exposure",
    "Internships & Work Readiness",
    "WBL Operations, Logistics & Reporting",
]

TEAM_SYMBOLS = {"Coaches": "🌱", "Coordinators": "⚙️", "Both": "👥"}


def read_gantt(path):
    """Read the Activities tab, automatically detecting the header row.

    The Gantt file has a title and instructions above the actual column headers,
    so we scan the first 10 rows for the one containing 'Workstream' and use that
    as the header row.
    """
    # First read raw with no header to find the right row
    df_raw = pd.read_excel(path, sheet_name="Activities", dtype=str, header=None)

    header_row = None
    for i in range(min(10, len(df_raw))):
        row_values = [str(v).strip() if pd.notna(v) else "" for v in df_raw.iloc[i].tolist()]
        if "Workstream" in row_values:
            header_row = i
            break

    if header_row is None:
        raise ValueError(
            "Could not find header row in Activities tab. "
            "Expected to find a row containing 'Workstream' within the first 10 rows. "
            f"First 3 rows seen: {df_raw.head(3).to_string()}"
        )

    print(f"   • Header row detected at row {header_row + 1} (1-indexed)")

    # Re-read with the correct header row
    df = pd.read_excel(path, sheet_name="Activities", dtype=str, header=header_row)
    df = df.dropna(how="all")

    print(f"   • Loaded {len(df)} data rows")
    print(f"   • Columns: {list(df.columns)}")

    return df


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

    processed = 0
    skipped_no_ws = 0
    skipped_no_activity = 0
    unknown_workstreams = set()

    for _, row in df.iterrows():
        ws = str(row.get("Workstream", "")).strip()
        activity = str(row.get("Activity", "")).strip()

        if not activity or activity == "nan":
            skipped_no_activity += 1
            continue
        if not ws or ws == "nan":
            skipped_no_ws += 1
            continue
        if ws not in data:
            unknown_workstreams.add(ws)
            continue

        eff_raw = str(row.get("Effort (1–4)", "")).strip()
        milestone = str(row.get("Milestone?", "")).strip().lower() == "yes"
        who = str(row.get("Who", "")).strip()
        link = str(row.get("Link", "")).strip()
        notes = str(row.get("Notes", "")).strip()

        months_for_row = parse_months(row)
        if not months_for_row:
            continue

        for m in months_for_row:
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
        processed += 1

    total_activity_instances = sum(len(data[ws][m]) for ws in data for m in data[ws])
    print(f"   • Processed {processed} activity rows into {total_activity_instances} month-instances")
    if skipped_no_activity:
        print(f"   • Skipped {skipped_no_activity} rows with no activity name")
    if skipped_no_ws:
        print(f"   • Skipped {skipped_no_ws} rows with no workstream")
    if unknown_workstreams:
        print(f"   ⚠️  Found {len(unknown_workstreams)} unknown workstream(s) — these were skipped:")
        for ws in sorted(unknown_workstreams):
            print(f"      - '{ws}'")
        print(f"   ℹ️  Known workstreams: {WORKSTREAM_ORDER}")

    if processed == 0:
        raise ValueError(
            "No activities were processed! Check that the Activities tab has data rows "
            "and that the 'Workstream' column values match the expected names exactly."
        )

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
    /* When split into halves, give months more room */
    table.grid-half {{
      font-size: 11.5px;
    }}
    table.grid-half thead th.month {{ width: 13.75%; }}
    table.grid-half thead th.workstream-col {{ width: 17.5%; }}
    table.grid-half td {{ padding: 8px 7px; }}

    /* Half-section wrapping each semester table */
    .half-section {{
      padding-top: 8px;
    }}
    .half-title {{
      font-family: 'Poppins', sans-serif;
      font-weight: 700;
      font-size: 16px;
      color: var(--navy);
      margin: 0 40px 12px;
      padding: 10px 16px;
      background: linear-gradient(90deg, var(--navy) 0%, var(--navy) 4px, transparent 4px);
      padding-left: 18px;
      border-bottom: 2px solid var(--navy);
    }}

    /* Compact second-page header (visible only when printing) */
    .page-two-header {{ display: none; }}
    .header-band-compact {{
      padding: 18px 40px 16px;
    }}
    .header-band-compact .header-seal {{
      width: 70px;
      height: 70px;
    }}
    .header-band-compact h1 {{
      font-size: 22px !important;
    }}
    .header-band-compact .lausd-wordmark {{
      width: 130px;
    }}
    .page-tag {{
      font-family: 'Poppins', sans-serif;
      font-weight: 400;
      font-size: 14px;
      opacity: 0.7;
      letter-spacing: 0.02em;
    }}

    /* Forced print page break between halves */
    .page-break {{
      display: block;
      height: 1px;
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
      .activity-card a.link {{
        font-size: 8.5px;
        color: var(--navy) !important;
      }}
      /* Force a fresh page before the second-half table */
      .page-break {{
        page-break-after: always;
        break-after: page;
      }}
      /* Show the compact header when printing the second page */
      .page-two-header {{ display: block; }}
      /* On the second printed page, the intro/legend was on page 1 — start fresh */
      .summary-section {{
        page-break-before: always;
        break-before: page;
      }}
    }}

    /* On screen, hide the print-only second header to avoid visual duplication */
    @media screen {{
      .page-two-header {{ display: none; }}
      .page-break {{ display: none; }}
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
    <p>This dashboard shows what the Linked Learning team is focused on each month of the school year, split into two semesters for easier reading and printing. Use it to plan initiatives, anticipate capacity, and align departmental priorities. The team capacity summary at the end shows total effort load by month.</p>
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

  <!-- MAIN GRID TABLE — FIRST HALF: JULY THROUGH DECEMBER -->
  <div class="half-section">
    <h3 class="half-title">First Semester · July – December 2026</h3>
    <div class="table-wrap">
      <table class="grid grid-half">
        <thead>
          <tr>
            <th class="workstream-col">Workstream</th>""")
    for m in range(1, 7):  # months 1-6 (Jul-Dec)
        parts.append(f'<th class="month">{MONTHS[m]}</th>')
    parts.append('</tr></thead><tbody>')

    for ws in WORKSTREAM_ORDER:
        bg = WORKSTREAM_COLORS[ws]
        accent = WORKSTREAM_ACCENTS[ws]
        parts.append(f'<tr>')
        parts.append(f'<td class="workstream-label" style="background:{bg};border-left:5px solid {accent};">{ws}</td>')

        for m in range(1, 7):
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
    parts.append('</tbody></table></div></div>')

    # --- PAGE BREAK + SECOND HALF HEADER ---
    parts.append(f"""
  <!-- PAGE 2 BREAK -->
  <div class="page-break"></div>

  <!-- COMPACT PAGE 2 HEADER (visible only on print + as a divider on screen) -->
  <div class="page-two-header">
    <div class="header-band header-band-compact">
      {seal_html}
      <div class="header-center">
        <h1>Linked Learning Team Priorities <span class="page-tag">· Page 2</span></h1>
        <div class="sub">Second Semester · January – June 2027</div>
      </div>
      <div class="header-right">
        {wordmark_html}
      </div>
    </div>
  </div>

  <!-- MAIN GRID TABLE — SECOND HALF: JANUARY THROUGH JUNE -->
  <div class="half-section">
    <h3 class="half-title">Second Semester · January – June 2027</h3>
    <div class="table-wrap">
      <table class="grid grid-half">
        <thead>
          <tr>
            <th class="workstream-col">Workstream</th>""")
    for m in range(7, 13):  # months 7-12 (Jan-Jun)
        parts.append(f'<th class="month">{MONTHS[m]}</th>')
    parts.append('</tr></thead><tbody>')

    for ws in WORKSTREAM_ORDER:
        bg = WORKSTREAM_COLORS[ws]
        accent = WORKSTREAM_ACCENTS[ws]
        parts.append(f'<tr>')
        parts.append(f'<td class="workstream-label" style="background:{bg};border-left:5px solid {accent};">{ws}</td>')

        for m in range(7, 13):
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
    parts.append('</tbody></table></div></div>')

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
    gantt = sys.argv[1] if len(sys.argv) > 1 else "LL_Team_Gantt_Heatmap_Revised_Categories.xlsx"
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
