#!/usr/bin/env python3
"""
Builds the password-protected, interactive UCSF Lynch Lab Method Compound Registry
& Standards Inventory Dashboard (Lynch_Lab_X500R_Method_Dashboard.html)
integrating:
  1. CDHR_X500R_drug_list.xlsx (234 X500R method compounds, RTs, LODs, Mix 8 recipe)
  2. CURRENT STANDARDS INVENTORY 2026 (1).xlsx (891 Current Stock + 105 Old Orders vials)
"""

import datetime
import json
import os
import re
import zipfile
import xml.etree.ElementTree as ET

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
X500R_XLSX = os.path.join(BASE_DIR, "CDHR_X500R_drug_list.xlsx")
INV_XLSX = (
    os.path.join(BASE_DIR, "CURRENT STANDARDS INVENTORY 2026 (1).xlsx")
    if os.path.exists(os.path.join(BASE_DIR, "CURRENT STANDARDS INVENTORY 2026 (1).xlsx"))
    else os.path.join(BASE_DIR, "CURRENT_STANDARDS_INVENTORY_2026.xlsx")
)
OUT_HTML = os.path.join(BASE_DIR, "Lynch_Lab_X500R_Method_Dashboard.html")


def parse_xlsx_with_styles_and_comments(xlsx_path):
    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(xlsx_path, "r") as z:
        shared_strings = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root.findall("s:si", ns):
                shared_strings.append(
                    "".join(t.text or "" for t in si.findall(".//s:t", ns))
                )

        fills = []
        xfs = []
        if "xl/styles.xml" in z.namelist():
            styles = ET.fromstring(z.read("xl/styles.xml"))
            for f in styles.findall(".//s:fills/s:fill", ns):
                pf = f.find("s:patternFill", ns)
                if pf is not None:
                    fg = pf.find("s:fgColor", ns)
                    fills.append(
                        (
                            pf.attrib.get("patternType"),
                            fg.attrib if fg is not None else {},
                        )
                    )
                else:
                    fills.append((None, {}))
            for xf in styles.findall(".//s:cellXfs/s:xf", ns):
                fid = int(xf.attrib.get("fillId", 0))
                xfs.append(fills[fid] if fid < len(fills) else (None, {}))

        wb = ET.fromstring(z.read("xl/workbook.xml"))
        sheets = [sh.attrib.get("name") for sh in wb.findall(".//s:sheet", ns)]

        parsed_sheets = {}
        sheet_styles = {}
        sheet_comments = {}

        for idx, sh_name in enumerate(sheets, 1):
            sh_file = f"xl/worksheets/sheet{idx}.xml"
            if sh_file not in z.namelist():
                continue

            # Parse threaded comments if present
            comm_map = {}
            tc_file = f"xl/threadedComments/threadedComment{idx}.xml"
            if tc_file in z.namelist():
                tc_root = ET.fromstring(z.read(tc_file))
                for el in tc_root.iter():
                    if el.tag.endswith("threadedComment"):
                        ref = el.attrib.get("ref", "")
                        txt = "".join(
                            t.text or ""
                            for t in el.iter()
                            if t.tag.endswith("text")
                        ).strip()
                        if ref and txt:
                            comm_map[ref] = txt

            ws = ET.fromstring(z.read(sh_file))
            rows_out = []
            styles_out = []
            for row in ws.findall(".//s:row", ns):
                r_dict = {}
                s_dict = {}
                r_num = row.attrib.get("r", "")
                for c in row.findall("s:c", ns):
                    ref = c.attrib.get("r", "")
                    col = re.sub(r"\d+", "", ref)
                    t = c.attrib.get("t", "")
                    s_idx = int(c.attrib.get("s", 0))
                    v_el = c.find("s:v", ns)
                    is_el = c.find("s:is", ns)
                    val = None
                    if t == "s" and v_el is not None and v_el.text is not None:
                        val = shared_strings[int(v_el.text)]
                    elif t == "inlineStr" and is_el is not None:
                        val = "".join(
                            x.text or "" for x in is_el.findall(".//s:t", ns)
                        )
                    elif v_el is not None and v_el.text is not None:
                        val = v_el.text
                    r_dict[col] = val
                    if s_idx < len(xfs):
                        s_dict[col] = xfs[s_idx]
                    if ref in comm_map:
                        r_dict[f"_comment_{col}"] = comm_map[ref]
                r_dict["_row_num"] = r_num
                rows_out.append(r_dict)
                styles_out.append(s_dict)
            parsed_sheets[sh_name] = rows_out
            sheet_styles[sh_name] = styles_out
            sheet_comments[sh_name] = comm_map
    return parsed_sheets, sheet_styles


def clean_num_str(val, decimals=2):
    if val is None:
        return ""
    s = str(val).strip()
    if s in ("", "#REF!", "None", "N/A"):
        return "" if s in ("#REF!", "None") else s
    try:
        f = float(s)
        if f.is_integer():
            return str(int(f))
        return f"{f:.{decimals}f}".rstrip("0").rstrip(".")
    except ValueError:
        return s


def convert_excel_date(val):
    if val is None:
        return ""
    s = str(val).strip()
    if not s:
        return ""
    if s.isdigit() and 30000 <= int(s) <= 65000:
        dt = datetime.date(1899, 12, 30) + datetime.timedelta(days=int(s))
        return dt.isoformat()
    if s.lower() == "early may 2026":
        return "2026-05-01"
    return s


def norm_chem_name(s):
    if not s:
        return ""
    s = str(s).lower().strip()
    # Extract both main name and parenthetical alias if present
    s = re.sub(r"\(suitable.*?\)", "", s)
    for tok in [
        " hydrochloride",
        " dihydrochloride",
        " hcl",
        " oxalate",
        " maleate",
        " dimaleate",
        " tartrate",
        " bitartrate",
        " sulfate",
        " hemisulfate",
        " fumarate",
        " citrate",
        " hydrobromide",
        " hbr",
        " hydroiodide",
        " hi",
        " mesylate",
        " besylate",
        " phosphate",
        " sodium",
        " potassium",
        " calcium",
        " free base",
        " solution",
        "(±)-",
        "(+)-",
        "(-)-",
        "(r)-",
        "(s)-",
        "±-",
        "d,l-",
        "dl-",
        "alpha-",
        "α-",
        "beta-",
        "β-",
        "delta9-",
        "delta-9-",
        "δ9-",
        "Δ9-",
    ]:
        s = s.replace(tok, "")
    return re.sub(r"[^a-z0-9]+", "", s)


def extract_candidate_keys(raw_name, synonyms=""):
    keys = set()
    if not raw_name:
        return keys
    base = norm_chem_name(raw_name)
    if base:
        keys.add(base)
    # Strip parenthetical text and also add parenthetical text separately
    no_paren = norm_chem_name(re.sub(r"\(.*?\)", "", raw_name))
    if no_paren:
        keys.add(no_paren)
    for m in re.findall(r"\((.*?)\)", raw_name):
        for part in re.split(r"[,;/]", m):
            k = norm_chem_name(part)
            if len(k) >= 3:
                keys.add(k)
    if synonyms:
        for part in re.split(r"[,;/]", str(synonyms)):
            k = norm_chem_name(part)
            if len(k) >= 3:
                keys.add(k)
    return keys


def build_dataset():
    x500_sheets, _ = parse_xlsx_with_styles_and_comments(X500R_XLSX)
    inv_sheets, inv_styles = parse_xlsx_with_styles_and_comments(INV_XLSX)

    # 1. Parse Standards Inventory (Current Stock + Old Orders)
    inventory_vials = []
    today_iso = datetime.date.today().isoformat()

    for sheet_label, start_row in [("Current Stock", 2), ("Old Orders", 1)]:
        rows = inv_sheets.get(sheet_label, [])
        styles = inv_styles.get(sheet_label, [])
        for idx in range(start_row, len(rows)):
            r = rows[idx]
            st = styles[idx] if idx < len(styles) else {}
            sname = (r.get("A") or "").strip()
            loc = (r.get("C") or "").strip()
            if not sname and not loc:
                continue

            status_raw = (r.get("B") or "").strip().upper()
            if status_raw in ("OPEN", "OPENED"):
                vial_status = "OPEN"
            elif status_raw == "UNOPENED":
                vial_status = "UNOPENED"
            elif status_raw:
                vial_status = status_raw
            else:
                vial_status = "OPEN" if sheet_label == "Current Stock" else "ARCHIVED"

            # Detect color-coded possession / flags from cell styles
            checked_out_by = (r.get("AC") or "").strip()
            style_flags = set()
            for col_k, (f_type, f_attr) in st.items():
                if f_type == "solid":
                    rgb = f_attr.get("rgb", "")
                    theme = f_attr.get("theme", "")
                    tint = str(f_attr.get("tint", ""))
                    if rgb == "FFD7BCE3":
                        checked_out_by = checked_out_by or "Gillian"
                    elif theme == "4" and "0.599" in tint:
                        checked_out_by = checked_out_by or "Emilio"
                    elif rgb == "FFFFFF00":
                        style_flags.add("Empty")
                    elif theme == "2":
                        style_flags.add("Past Exp Highlight")

            exp_raw = (r.get("J") or "").strip()
            exp_date = convert_excel_date(exp_raw)
            is_expired = False
            if re.match(r"^\d{4}-\d{2}-\d{2}$", exp_date) and exp_date < today_iso:
                is_expired = True

            rec_date = convert_excel_date(r.get("K"))
            open_date = convert_excel_date(r.get("M"))
            has_tag = (r.get("N") or "").strip() if sheet_label == "Current Stock" else ""
            notes_raw = (r.get("O") if sheet_label == "Current Stock" else r.get("N")) or ""
            notes_list = [notes_raw.strip()] if notes_raw.strip() else []
            for k_col, v_comm in r.items():
                if k_col.startswith("_comment_") and v_comm:
                    notes_list.append(f"Note ({k_col.replace('_comment_', '')}): {v_comm}")
            if "Empty" in style_flags:
                vial_status = "EMPTY"

            vial_id = (
                f"INV-{len(inventory_vials)+1:04d}"
                if sheet_label == "Current Stock"
                else f"OLD-{len(inventory_vials)+1:04d}"
            )
            inventory_vials.append(
                {
                    "vial_id": vial_id,
                    "sheet": sheet_label,
                    "standard_name": sname or "(Unnamed Standard)",
                    "vial_status": vial_status,
                    "checked_out_by": checked_out_by,
                    "location": loc,
                    "concentration": (r.get("D") or "").strip(),
                    "solvent": (r.get("E") or "").strip(),
                    "manufacturer": (r.get("F") or "").strip(),
                    "catalog_no": clean_num_str(r.get("G"), 0),
                    "lot_no": clean_num_str(r.get("H"), 0),
                    "cas_no": (r.get("I") or "").strip(),
                    "exp_date": exp_date,
                    "is_expired": is_expired,
                    "receive_date": rec_date,
                    "initial_vol_mass": clean_num_str(r.get("L"), 1),
                    "open_date": open_date,
                    "has_tag": has_tag,
                    "notes": " | ".join(notes_list),
                    "linked_methods": [],
                    "linked_compound_ids": [],
                    "is_custom": False,
                }
            )

    # Index Current Stock & Old Orders by location and by normalized chemical keys
    vials_by_loc = {}
    vials_by_key = {}
    for v in inventory_vials:
        loc_u = v["location"].strip().upper()
        if loc_u and loc_u not in ("TBD", "NEW", "UNKNOWN", "?"):
            vials_by_loc.setdefault(loc_u, []).append(v)
        for k in extract_candidate_keys(v["standard_name"]):
            vials_by_key.setdefault(k, []).append(v)

    # 2. Parse X500R Method Compounds and Cross-Link with Standards Inventory
    main_rows = x500_sheets.get("updated_260929", [])
    compounds = []

    for idx, r in enumerate(main_rows[1:], 1):
        cname = (r.get("F") or "").strip()
        if not cname:
            continue

        raw_a = (r.get("A") or "").strip()
        raw_b = (r.get("B") or "").strip()
        raw_c = (r.get("C") or "").strip()
        raw_d = (r.get("D") or "").strip()

        if raw_b == "1":
            status = "Active"
        elif raw_b == "0":
            status = "Inactive / Pending"
        elif "removed" in raw_b.lower():
            status = "Removed"
        else:
            status = "Inactive / Pending"

        notes_parts = []
        if raw_d:
            notes_parts.append(raw_d)
        if raw_b not in ("1", "0", ""):
            notes_parts.append(f"Method status note: {raw_b}")
        if raw_a not in ("1", "0", ""):
            notes_parts.append(f"Library note: {raw_a}")
        if (r.get("AN") or "").strip():
            notes_parts.append(f"Reason: {(r.get('AN') or '').strip()}")

        in_lib = "Yes" if raw_a == "1" else ("No" if raw_a == "0" else "Needs Review")
        date_added = convert_excel_date(raw_c)

        mono_mass = clean_num_str(r.get("K"), 4)
        mh_mz = ""
        try:
            if mono_mass:
                mh_mz = f"{float(mono_mass) + 1.007276:.4f}"
        except ValueError:
            mh_mz = ""

        box_loc = (r.get("P") or "").strip()
        stock_conc = (r.get("Q") or "").strip()
        stock_solv = (r.get("R") or "").strip()
        synonyms = (r.get("O") or "").strip()

        # Match against inventory vials (prefer Current Stock over Old Orders, exact location match or name match)
        matched_vials = []
        seen_vids = set()
        cand_keys = extract_candidate_keys(cname, synonyms)

        # First check name matches
        for ck in cand_keys:
            for v in vials_by_key.get(ck, []):
                if v["vial_id"] not in seen_vids:
                    matched_vials.append(v)
                    seen_vids.add(v["vial_id"])

        # Also check exact box location match if box_loc is specific (like A-05, B-42, CAY2-37, D-11)
        if box_loc and box_loc.upper() in vials_by_loc:
            for v in vials_by_loc[box_loc.upper()]:
                if v["vial_id"] not in seen_vids:
                    v_keys = extract_candidate_keys(v["standard_name"])
                    # Add if chemical name shares prefix/substring or if no name match was found yet
                    if (v_keys & cand_keys) or not matched_vials or len(vials_by_loc[box_loc.upper()]) == 1:
                        matched_vials.append(v)
                        seen_vids.add(v["vial_id"])

        # Sort matched vials so Current Stock + Checked-Out + OPEN/UNOPENED come first
        matched_vials.sort(
            key=lambda v: (
                0 if v["sheet"] == "Current Stock" else 1,
                0 if v.get("checked_out_by") else 1,
                0 if v["vial_status"] in ("OPEN", "UNOPENED") else 1,
            )
        )

        comp_id = f"X500R-{idx:03d}"
        comp_id_7600 = f"7600-{idx:03d}"
        primary_vial = matched_vials[0] if matched_vials else None
        checked_out_holder = next((mv["checked_out_by"] for mv in matched_vials if mv.get("checked_out_by")), "")

        for mv in matched_vials:
            if "X500R (CDHR)" not in mv["linked_methods"]:
                mv["linked_methods"].append("X500R (CDHR)")
            if "7600 (CDS W93)" not in mv["linked_methods"]:
                mv["linked_methods"].append("7600 (CDS W93)")
            if comp_id not in mv["linked_compound_ids"]:
                mv["linked_compound_ids"].append(comp_id)
            if comp_id_7600 not in mv["linked_compound_ids"]:
                mv["linked_compound_ids"].append(comp_id_7600)

        inv_enriched = False
        if primary_vial:
            if not box_loc or box_loc.upper() in ("NEW", "TBD"):
                if primary_vial["location"]:
                    box_loc = f"{primary_vial['location']} (Inventory)"
                    inv_enriched = True
            if not stock_conc and primary_vial["concentration"]:
                stock_conc = primary_vial["concentration"]
            if not stock_solv and primary_vial["solvent"]:
                stock_solv = primary_vial["solvent"]

        raw_e = (r.get("E") or "").strip()
        in_lib_7600 = "No" if ("missing" in raw_e.lower() or raw_e.lower() in ("no", "0", "false", "")) else "Yes"
        notes_7600 = list(notes_parts)
        if raw_e and raw_e.lower() not in ("yes", "no", "1", "0"):
            notes_7600.append(f"7600_CID note: {raw_e}")

        rt_x500r_val = clean_num_str(r.get("H"), 2)
        rt_7600_val = clean_num_str(r.get("G"), 2)

        base_entry = {
            "id": comp_id,
            "method": "X500R (CDHR)",
            "compound": cname,
            "status": status,
            "in_library": in_lib,
            "date_added": date_added,
            "date_first_w93": "",
            "date_last_rt_update": "2026-07-01" if rt_x500r_val else "",
            "checked_7600": raw_e,
            "rt_x500r_jul26": rt_x500r_val,
            "rt_x500r_dec25": clean_num_str(r.get("M"), 2),
            "rt_7600_jul26": rt_7600_val,
            "primary_mix": clean_num_str(r.get("I"), 0),
            "secondary_mix": (r.get("J") or "").strip(),
            "monoisotopic_mass": mono_mass,
            "mh_plus_mz": mh_mz,
            "formula": (r.get("L") or "").strip(),
            "isomers": (r.get("N") or "").strip(),
            "synonyms": synonyms,
            "box_location": box_loc,
            "box_position": (r.get("S") or "").strip(),
            "stock_conc": stock_conc,
            "stock_solvent": stock_solv,
            "manufacturer": primary_vial["manufacturer"] if primary_vial else "",
            "catalog_no": primary_vial["catalog_no"] if primary_vial else "",
            "lot_no": primary_vial["lot_no"] if primary_vial else "",
            "cas_no": primary_vial["cas_no"] if primary_vial else "",
            "vial_status": primary_vial["vial_status"] if primary_vial else "",
            "vial_exp_date": primary_vial["exp_date"] if primary_vial else "",
            "checked_out_by": checked_out_holder,
            "linked_vial_ids": [mv["vial_id"] for mv in matched_vials],
            "inv_enriched": inv_enriched,
            "inchikey": (r.get("T") or "").strip(),
            "lod_7600_u_swath": clean_num_str(r.get("U"), 1),
            "lod_7600_u_ida": clean_num_str(r.get("V"), 1),
            "lod_7600_s_ida": clean_num_str(r.get("AC"), 1),
            "lod_5600_u_ida": clean_num_str(r.get("W"), 1),
            "lod_brown_u": clean_num_str(r.get("X"), 1),
            "lod_x500r_u_p1": clean_num_str(r.get("Y"), 1),
            "lod_x500r_u_p2": clean_num_str(r.get("Z"), 1),
            "x500r_u_agree": (r.get("AA") or "").strip(),
            "lod_x500r_u_consensus": clean_num_str(r.get("AB"), 1),
            "lod_brown_s": clean_num_str(r.get("AD"), 1),
            "lod_x500r_s_p1": clean_num_str(r.get("AE"), 1),
            "lod_x500r_s_p2": clean_num_str(r.get("AF"), 1),
            "x500r_s_agree": (r.get("AG") or "").strip(),
            "lod_x500r_s_consensus": clean_num_str(r.get("AH"), 1),
            "me_7600_i": clean_num_str(r.get("AJ"), 1),
            "me_7600_s": clean_num_str(r.get("AK"), 1),
            "me_5600_u": clean_num_str(r.get("AL"), 1),
            "include_flag": clean_num_str(r.get("AM"), 0),
            "comments": " | ".join(notes_parts),
            "submitted_by": "Initial Import (CDHR_X500R_drug_list.xlsx)",
            "updated_at": "2026-09-29",
            "is_custom": False,
        }
        compounds.append(base_entry)

        entry_7600 = dict(base_entry)
        entry_7600.update(
            {
                "id": comp_id_7600,
                "method": "7600 (CDS W93)",
                "in_library": in_lib_7600,
                "date_first_w93": "",
                "date_last_rt_update": "2026-07-01" if rt_7600_val else "",
                "primary_mix": "",
                "secondary_mix": "",
                "comments": " | ".join(notes_7600),
            }
        )
        compounds.append(entry_7600)

    # Parse Sheet 2 (urine_LOD outstanding), Sheet 3 (serum_LOD outstanding), Sheet 4 (Mix 8)
    urine_sheet = x500_sheets.get("urine_LOD", [])
    urine_outstanding = []
    for r in urine_sheet[1:]:
        if r.get("E"):
            urine_outstanding.append(
                {
                    "compound": (r.get("E") or "").strip(),
                    "lod_note": clean_num_str(r.get("F"), 1),
                    "comment": (r.get("G") or "").strip() or "Outstanding Urine LOD validation",
                }
            )

    serum_sheet = x500_sheets.get("serum_LOD", [])
    serum_outstanding = []
    for r in serum_sheet[1:]:
        if r.get("G"):
            serum_outstanding.append(
                {
                    "compound": (r.get("G") or "").strip(),
                    "comment": "Outstanding Serum LOD validation",
                }
            )

    mix8_sheet = x500_sheets.get("Mix 8 and high LOD", [])
    mix8_rows = []
    for r in mix8_sheet[1:]:
        if r.get("C"):
            mix8_rows.append(
                {
                    "compound": (r.get("C") or "").strip(),
                    "stock_ug_ml": clean_num_str(r.get("D"), 1),
                    "final_ug_ml": clean_num_str(r.get("E"), 1),
                    "vol_ul_per_ml": clean_num_str(r.get("F"), 1),
                    "validated_phase1": (r.get("H") or "no").strip(),
                }
            )

    return {
        "compounds": compounds,
        "inventory": inventory_vials,
        "urine_outstanding": urine_outstanding,
        "serum_outstanding": serum_outstanding,
        "mix8_recipe": mix8_rows,
        "source_workbook": "CDHR_X500R_drug_list.xlsx + CURRENT STANDARDS INVENTORY 2026 (1).xlsx",
        "generated_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
    }


HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>UCSF Lynch Lab — Method Compound Registry & Standards Inventory Dashboard</title>
<style>
  :root {
    --ucsf-navy: #052049;
    --ucsf-blue: #007CBE;
    --ucsf-teal: #00A3AD;
    --bg-main: #f4f7fb;
    --card-bg: #ffffff;
    --border-col: #dce3ee;
    --text-main: #1e293b;
    --text-muted: #64748b;
    --green-bg: #dcfce7;
    --green-fg: #166534;
    --amber-bg: #fef3c7;
    --amber-fg: #92400e;
    --red-bg: #fee2e2;
    --red-fg: #991b1b;
    --purple-bg: #f3e8ff;
    --purple-fg: #6b21a8;
    --blue-bg: #dbeafe;
    --blue-fg: #1e40af;
  }
  * { box-sizing: border-box; }
  body {
    margin: 0;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    background: var(--bg-main);
    color: var(--text-main);
    font-size: 13.5px;
  }

  /* PASSWORD LOCK OVERLAY */
  #auth-overlay {
    position: fixed;
    inset: 0;
    background: linear-gradient(135deg, #052049 0%, #0b3b75 55%, #007CBE 100%);
    display: flex;
    align-items: center;
    justify-content: center;
    z-index: 99999;
    padding: 20px;
  }
  .auth-card {
    background: #ffffff;
    width: 100%;
    max-width: 430px;
    border-radius: 14px;
    padding: 32px 28px;
    box-shadow: 0 24px 48px rgba(0, 0, 0, 0.35);
    text-align: center;
  }
  .auth-logo {
    width: 56px;
    height: 56px;
    border-radius: 12px;
    background: var(--ucsf-navy);
    color: #fff;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    font-weight: 800;
    font-size: 20px;
    margin-bottom: 14px;
    letter-spacing: 0.5px;
  }
  .auth-title {
    margin: 0 0 6px 0;
    font-size: 20px;
    font-weight: 800;
    color: var(--ucsf-navy);
  }
  .auth-subtitle {
    margin: 0 0 22px 0;
    font-size: 13px;
    color: var(--text-muted);
    line-height: 1.45;
  }
  .auth-input {
    width: 100%;
    padding: 11px 14px;
    font-size: 14px;
    border: 1.5px solid #cbd5e1;
    border-radius: 8px;
    margin-bottom: 12px;
    outline: none;
    transition: border-color 0.15s;
  }
  .auth-input:focus {
    border-color: var(--ucsf-blue);
    box-shadow: 0 0 0 3px rgba(0, 124, 190, 0.15);
  }
  .auth-btn {
    width: 100%;
    padding: 11px 16px;
    font-size: 14px;
    font-weight: 700;
    color: #fff;
    background: var(--ucsf-navy);
    border: none;
    border-radius: 8px;
    cursor: pointer;
    transition: background 0.15s;
  }
  .auth-btn:hover { background: var(--ucsf-blue); }
  .auth-err {
    color: #dc2626;
    font-size: 12.5px;
    font-weight: 600;
    margin-top: 10px;
    min-height: 18px;
  }

  /* HEADER & METHOD TABS */
  .top-header {
    background: linear-gradient(90deg, #052049 0%, #09356e 65%, #007CBE 100%);
    color: #fff;
    padding: 16px 24px 0 24px;
    box-shadow: 0 2px 10px rgba(5, 32, 73, 0.18);
  }
  .header-row {
    display: flex;
    align-items: center;
    justify-content: space-between;
    flex-wrap: wrap;
    gap: 14px;
    padding-bottom: 14px;
  }
  .brand-group {
    display: flex;
    align-items: center;
    gap: 14px;
  }
  .brand-badge {
    background: rgba(255, 255, 255, 0.14);
    border: 1px solid rgba(255, 255, 255, 0.25);
    padding: 6px 11px;
    border-radius: 8px;
    font-weight: 800;
    font-size: 13px;
    letter-spacing: 0.6px;
  }
  .brand-title {
    margin: 0;
    font-size: 20px;
    font-weight: 800;
    letter-spacing: -0.2px;
  }
  .brand-sub {
    margin: 3px 0 0 0;
    font-size: 12.5px;
    color: #cbd5e1;
  }
  .header-actions {
    display: flex;
    align-items: center;
    gap: 8px;
    flex-wrap: wrap;
  }
  .btn {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    padding: 8px 13px;
    border-radius: 7px;
    font-size: 12.5px;
    font-weight: 600;
    cursor: pointer;
    border: 1px solid transparent;
    transition: all 0.15s;
    text-decoration: none;
    white-space: nowrap;
  }
  .btn-primary {
    background: #10b981;
    color: #fff;
    border-color: #059669;
    box-shadow: 0 2px 6px rgba(16, 185, 129, 0.3);
  }
  .btn-primary:hover { background: #059669; }
  .btn-teal {
    background: #0d9488;
    color: #fff;
    border-color: #0f766e;
  }
  .btn-teal:hover { background: #0f766e; }
  .btn-glass {
    background: rgba(255, 255, 255, 0.12);
    color: #fff;
    border-color: rgba(255, 255, 255, 0.25);
  }
  .btn-glass:hover { background: rgba(255, 255, 255, 0.22); }
  .btn-outline {
    background: #fff;
    color: var(--ucsf-navy);
    border-color: #cbd5e1;
  }
  .btn-outline:hover {
    background: #f1f5f9;
    border-color: #94a3b8;
  }
  .btn-blue {
    background: var(--ucsf-blue);
    color: #fff;
  }
  .btn-blue:hover { background: #0369a1; }

  /* METHOD TAB BAR */
  .method-tabs {
    display: flex;
    align-items: center;
    gap: 6px;
    overflow-x: auto;
    padding-top: 4px;
  }
  .method-tab {
    padding: 10px 18px;
    border-radius: 8px 8px 0 0;
    background: rgba(255, 255, 255, 0.1);
    color: #e2e8f0;
    font-weight: 700;
    font-size: 13px;
    cursor: pointer;
    border: none;
    display: inline-flex;
    align-items: center;
    gap: 8px;
    transition: all 0.15s;
  }
  .method-tab:hover {
    background: rgba(255, 255, 255, 0.2);
    color: #fff;
  }
  .method-tab.active {
    background: var(--bg-main);
    color: var(--ucsf-navy);
    box-shadow: 0 -2px 8px rgba(0, 0, 0, 0.1);
  }
  .tab-pill {
    background: rgba(15, 23, 42, 0.18);
    color: inherit;
    padding: 2px 7px;
    border-radius: 999px;
    font-size: 11px;
    font-weight: 700;
  }
  .method-tab.active .tab-pill {
    background: #dbeafe;
    color: #1e40af;
  }

  /* MAIN CONTAINER */
  .container {
    max-width: 1720px;
    margin: 0 auto;
    padding: 18px 24px 48px 24px;
  }

  /* KPI ROW */
  .kpi-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(210px, 1fr));
    gap: 12px;
    margin-bottom: 16px;
  }
  .kpi-card {
    background: var(--card-bg);
    border: 1px solid var(--border-col);
    border-radius: 10px;
    padding: 14px 16px;
    box-shadow: 0 1px 3px rgba(15, 23, 42, 0.04);
  }
  .kpi-label {
    font-size: 11.5px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.4px;
    color: var(--text-muted);
    margin-bottom: 6px;
  }
  .kpi-val {
    font-size: 24px;
    font-weight: 800;
    color: var(--ucsf-navy);
    line-height: 1.1;
  }
  .kpi-sub {
    font-size: 12px;
    color: var(--text-muted);
    margin-top: 5px;
  }

  /* SUB-NAVIGATION PILLS */
  .subnav-bar {
    display: flex;
    align-items: center;
    justify-content: space-between;
    flex-wrap: wrap;
    gap: 10px;
    margin-bottom: 14px;
  }
  .subnav-group {
    display: inline-flex;
    background: #e2e8f0;
    padding: 4px;
    border-radius: 9px;
    gap: 4px;
  }
  .subnav-btn {
    border: none;
    background: transparent;
    padding: 7px 14px;
    border-radius: 6px;
    font-size: 12.5px;
    font-weight: 700;
    color: #475569;
    cursor: pointer;
  }
  .subnav-btn.active {
    background: #fff;
    color: var(--ucsf-navy);
    box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
  }

  /* FILTER & EXPORT TOOLBAR */
  .toolbar-card {
    background: var(--card-bg);
    border: 1px solid var(--border-col);
    border-radius: 10px;
    padding: 14px 16px;
    margin-bottom: 14px;
    display: flex;
    flex-direction: column;
    gap: 12px;
  }
  .toolbar-row {
    display: flex;
    align-items: center;
    justify-content: space-between;
    flex-wrap: nowrap;
    gap: 8px;
    overflow-x: auto;
  }
  .filter-group {
    display: flex;
    align-items: center;
    flex-wrap: nowrap;
    gap: 6px;
    flex: 1 1 auto;
    min-width: 0;
  }
  .search-box {
    flex: 1 1 170px;
    min-width: 135px;
    max-width: 240px;
    padding: 6px 9px;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    font-size: 12px;
    outline: none;
  }
  .search-box:focus {
    border-color: var(--ucsf-blue);
    box-shadow: 0 0 0 2px rgba(0, 124, 190, 0.15);
  }
  .select-ctrl, .date-ctrl {
    padding: 5px 7px;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    font-size: 11.5px;
    background: #fff;
    color: var(--text-main);
    max-width: 155px;
    text-overflow: ellipsis;
  }
  .ctrl-label {
    font-size: 11px;
    font-weight: 700;
    color: var(--text-muted);
    margin-right: 0;
    white-space: nowrap;
  }

  /* REAL-TIME CHANGE / SUGGESTION INPUT BAR */
  .realtime-bar {
    background: linear-gradient(90deg, #f8fafc 0%, #f0f9ff 100%);
    border: 1px solid #bae6fd;
    border-radius: 8px;
    padding: 8px 12px;
    display: flex;
    flex-direction: column;
    gap: 5px;
  }
  .realtime-bar-row {
    display: flex;
    align-items: center;
    gap: 8px;
    flex-wrap: nowrap;
  }
  .realtime-badge {
    background: var(--ucsf-navy);
    color: #fff;
    font-size: 11px;
    font-weight: 800;
    padding: 4px 8px;
    border-radius: 6px;
    white-space: nowrap;
    letter-spacing: 0.2px;
  }
  .realtime-input {
    flex: 1 1 auto;
    min-width: 180px;
    padding: 6px 10px;
    border: 1px solid #7dd3fc;
    border-radius: 6px;
    font-size: 12px;
    background: #fff;
    color: var(--text-main);
    outline: none;
  }
  .realtime-input:focus {
    border-color: var(--ucsf-blue);
    box-shadow: 0 0 0 2px rgba(0, 124, 190, 0.2);
  }
  .cmd-example-chip {
    display: inline-block;
    background: #e0f2fe;
    color: #0369a1;
    border: 1px solid #bae6fd;
    border-radius: 999px;
    padding: 2px 8px;
    font-size: 11px;
    font-weight: 600;
    margin: 2px 3px;
    cursor: pointer;
  }
  .cmd-example-chip:hover {
    background: #bae6fd;
  }
  .suggestion-chip {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    background: #f5f3ff;
    border: 1px solid #c4b5fd;
    color: #4c1d95;
    padding: 4px 9px;
    border-radius: 7px;
    font-size: 11.5px;
    font-weight: 600;
  }
  .inline-select {
    padding: 3px 6px;
    border-radius: 5px;
    font-size: 11.5px;
    font-weight: 700;
    border: 1px solid #cbd5e1;
    cursor: pointer;
  }
  .inline-date {
    padding: 3px 6px;
    border-radius: 5px;
    font-size: 11.5px;
    border: 1px solid #cbd5e1;
    background: #fff;
    color: var(--text-main);
  }
  @keyframes rowFlash {
    0% { background-color: #bbf7d0; }
    100% { background-color: transparent; }
  }
  .flash-updated {
    animation: rowFlash 2.2s ease-out;
  }

  /* TABLE STYLING */
  .table-card {
    background: var(--card-bg);
    border: 1px solid var(--border-col);
    border-radius: 10px;
    overflow: hidden;
    box-shadow: 0 1px 3px rgba(15, 23, 42, 0.04);
  }
  .table-scroll {
    overflow-x: auto;
    max-height: 72vh;
    overflow-y: auto;
  }
  table.comp-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 12.5px;
  }
  table.comp-table thead th {
    position: sticky;
    top: 0;
    z-index: 10;
    background: #f8fafc;
    color: var(--ucsf-navy);
    font-weight: 700;
    text-align: left;
    padding: 10px 10px;
    border-bottom: 2px solid #cbd5e1;
    white-space: nowrap;
    cursor: pointer;
    user-select: none;
  }
  table.comp-table thead th:hover {
    background: #eef2ff;
  }
  table.comp-table tbody tr {
    border-bottom: 1px solid #e2e8f0;
    transition: background 0.1s;
  }
  table.comp-table tbody tr:hover {
    background: #f1f5f9;
  }
  table.comp-table tbody tr.custom-row {
    background: #f0fdf4;
  }
  table.comp-table tbody tr.custom-row:hover {
    background: #dcfce7;
  }
  table.comp-table tbody tr.checkedout-row {
    background: #faf5ff;
  }
  table.comp-table tbody tr.co-row-gillian {
    background: #f5ecfe !important;
  }
  table.comp-table tbody tr.co-row-gillian:hover {
    background: #ebd9fd !important;
  }
  table.comp-table tbody tr.co-row-lauren {
    background: #e6fcf5 !important;
  }
  table.comp-table tbody tr.co-row-lauren:hover {
    background: #c3fae8 !important;
  }
  table.comp-table tbody tr.co-row-michael {
    background: #e8f1ff !important;
  }
  table.comp-table tbody tr.co-row-michael:hover {
    background: #d0e2ff !important;
  }
  table.comp-table tbody tr.co-row-emilio {
    background: #e0f2fe !important;
  }
  table.comp-table tbody tr.co-row-emilio:hover {
    background: #bae6fd !important;
  }
  table.comp-table tbody tr.co-row-other {
    background: #fef3c7 !important;
  }
  .checkout-wrap {
    display: inline-flex;
    align-items: center;
    gap: 7px;
    cursor: pointer;
    user-select: none;
    white-space: nowrap;
  }
  .checkout-cb {
    width: 16px;
    height: 16px;
    cursor: pointer;
    margin: 0;
    border-radius: 4px;
  }
  .co-badge-gillian {
    background: #e9d5ff;
    color: #581c87;
    border: 1.5px solid #a855f7;
  }
  .co-badge-lauren {
    background: #99f6e4;
    color: #115e59;
    border: 1.5px solid #14b8a6;
  }
  .co-badge-michael {
    background: #bfdbfe;
    color: #1e3a8a;
    border: 1.5px solid #3b82f6;
  }
  .co-badge-emilio {
    background: #bae6fd;
    color: #0c4a6e;
    border: 1.5px solid #0ea5e9;
  }
  .co-badge-other {
    background: #fde68a;
    color: #78350f;
    border: 1.5px solid #f59e0b;
  }
  table.comp-table td {
    padding: 8px 10px;
    vertical-align: middle;
  }
  .mono {
    font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;
    font-size: 12px;
  }

  /* STATUS BADGES */
  .badge {
    display: inline-flex;
    align-items: center;
    padding: 2.5px 8px;
    border-radius: 999px;
    font-size: 11.5px;
    font-weight: 700;
    white-space: nowrap;
  }
  .badge-active { background: var(--green-bg); color: var(--green-fg); }
  .badge-pending { background: var(--amber-bg); color: var(--amber-fg); }
  .badge-removed { background: var(--red-bg); color: var(--red-fg); }
  .badge-custom { background: var(--purple-bg); color: var(--purple-fg); }
  .badge-blue { background: var(--blue-bg); color: var(--blue-fg); }
  .badge-mix {
    background: #e0f2fe;
    color: #0369a1;
    border: 1px solid #bae6fd;
  }

  /* MODAL DIALOGS */
  .modal-backdrop {
    position: fixed;
    inset: 0;
    background: rgba(15, 23, 42, 0.55);
    display: none;
    align-items: center;
    justify-content: center;
    z-index: 9000;
    padding: 20px;
  }
  .modal-backdrop.open { display: flex; }
  .modal-box {
    background: #fff;
    width: 100%;
    max-width: 860px;
    max-height: 90vh;
    border-radius: 12px;
    box-shadow: 0 20px 45px rgba(0, 0, 0, 0.28);
    display: flex;
    flex-direction: column;
    overflow: hidden;
  }
  .modal-header {
    background: var(--ucsf-navy);
    color: #fff;
    padding: 15px 20px;
    display: flex;
    align-items: center;
    justify-content: space-between;
  }
  .modal-header h3 {
    margin: 0;
    font-size: 16px;
    font-weight: 800;
  }
  .modal-close {
    background: transparent;
    border: none;
    color: #cbd5e1;
    font-size: 20px;
    cursor: pointer;
  }
  .modal-body {
    padding: 20px;
    overflow-y: auto;
  }
  .form-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
    gap: 14px;
  }
  .form-field {
    display: flex;
    flex-direction: column;
    gap: 5px;
  }
  .form-field.full-width {
    grid-column: 1 / -1;
  }
  .form-field label {
    font-size: 12px;
    font-weight: 700;
    color: #334155;
  }
  .form-field input, .form-field select, .form-field textarea {
    padding: 8px 10px;
    border: 1px solid #cbd5e1;
    border-radius: 7px;
    font-size: 13px;
  }
  .modal-footer {
    padding: 14px 20px;
    background: #f8fafc;
    border-top: 1px solid #e2e8f0;
    display: flex;
    align-items: center;
    justify-content: flex-end;
    gap: 10px;
  }

  /* TOAST NOTIFICATION */
  #toast-box {
    position: fixed;
    bottom: 22px;
    right: 22px;
    background: var(--ucsf-navy);
    color: #fff;
    padding: 12px 18px;
    border-radius: 9px;
    font-size: 13px;
    font-weight: 600;
    box-shadow: 0 8px 24px rgba(0, 0, 0, 0.25);
    display: none;
    z-index: 99999;
  }
  .action-link {
    background: none;
    border: 1px solid #cbd5e1;
    border-radius: 5px;
    padding: 3px 7px;
    font-size: 11.5px;
    font-weight: 600;
    color: var(--ucsf-navy);
    cursor: pointer;
    margin-right: 4px;
  }
  .action-link:hover { background: #e2e8f0; }
</style>
<style id="dynamic-palette-css"></style>
</head>
<body>

<!-- PASSWORD PROTECTION OVERLAY -->
<div id="auth-overlay">
  <div class="auth-card">
    <div class="auth-logo">UCSF</div>
    <h2 class="auth-title">Lynch Lab Method & Inventory Portal</h2>
    <p class="auth-subtitle">
      Sign in with your Lynch Lab user account and password to view and manage MS Method Compounds (X500R CDHR) and the 2026 Chemical Standards Inventory.
    </p>
    <form onsubmit="handleUnlock(event)">
      <div style="text-align:left;margin-bottom:6px;font-size:12px;font-weight:700;color:#334155;">Lab User</div>
      <select id="auth-user-input" class="auth-input" style="background:#fff;cursor:pointer;" required>
        <option value="Gillian">Gillian</option>
        <option value="Lauren">Lauren</option>
        <option value="Michael">Michael</option>
      </select>
      <div style="text-align:left;margin-bottom:6px;font-size:12px;font-weight:700;color:#334155;">Password</div>
      <input type="password" id="auth-pw-input" class="auth-input" placeholder="Enter your password..." autocomplete="off" autofocus required>
      <button type="submit" class="auth-btn" style="margin-top:4px;">Sign In to Dashboard</button>
      <div id="auth-err" class="auth-err"></div>
    </form>
  </div>
</div>

<!-- TOP HEADER -->
<header class="top-header">
  <div class="header-row">
    <div class="brand-group">
      <div class="brand-badge">LYNCH LAB · UCSF</div>
      <div>
        <h1 class="brand-title">Method Compound Registry & Standards Inventory Dashboard</h1>
        <p class="brand-sub" id="header-meta-sub">Integrated: CDHR_X500R_drug_list.xlsx (234 Method Compounds) + CURRENT STANDARDS INVENTORY 2026 (994 Vials)</p>
      </div>
    </div>
    <div class="header-actions">
      <span id="active-user-pill" style="display:inline-flex;align-items:center;gap:6px;background:rgba(255,255,255,0.16);border:1px solid rgba(255,255,255,0.3);color:#fff;padding:6px 12px;border-radius:999px;font-size:12.5px;font-weight:700;">👤 Not Signed In</span>
      <!-- MoMA Color Palette Picker -->
      <div style="display:inline-flex;align-items:center;gap:6px;background:rgba(255,255,255,0.14);border:1px solid rgba(255,255,255,0.28);border-radius:8px;padding:4px 8px;">
        <label for="palette-select" style="font-size:11px;font-weight:700;color:#e2e8f0;white-space:nowrap;">🎨 Palette:</label>
        <select id="palette-select" onchange="applyColorPalette(this.value)" style="padding:3px 6px;border-radius:5px;font-size:11.5px;font-weight:700;background:#ffffff;color:#0f172a;border:1px solid #cbd5e1;cursor:pointer;max-width:155px;">
          <option value="ucsf_original">UCSF Original (Light)</option>
          <option value="original">Original (Dark)</option>
          <optgroup label="Featured 5 Re-Skins">
            <option value="VanGogh">VanGogh (Dark)</option>
            <option value="Doughton">Doughton (Light)</option>
            <option value="Ernst">Ernst (Dark)</option>
            <option value="OKeeffe">OKeeffe (Dark)</option>
            <option value="ustwo">ustwo (Dark)</option>
          </optgroup>
          <optgroup label="MoMA Palettes">
            <option value="Abbott">Abbott (Light)</option>
            <option value="Alkalay1">Alkalay1 (Light)</option>
            <option value="Alkalay2">Alkalay2 (Dark)</option>
            <option value="Althoff">Althoff (Light)</option>
            <option value="Andri">Andri (Light)</option>
            <option value="Avedon">Avedon (Dark)</option>
            <option value="Budnitz">Budnitz (Dark)</option>
            <option value="Clay">Clay (Light)</option>
            <option value="Connors">Connors (Dark)</option>
            <option value="Dali">Dali (Dark)</option>
            <option value="Exter">Exter (Light)</option>
            <option value="Flash">Flash (Light)</option>
            <option value="Fritsch">Fritsch (Light)</option>
            <option value="Kippenberger">Kippenberger (Light)</option>
            <option value="Klein">Klein (Dark)</option>
            <option value="Koons">Koons (Dark)</option>
            <option value="Levine1">Levine1 (Light)</option>
            <option value="Levine2">Levine2 (Dark)</option>
            <option value="Liu">Liu (Dark)</option>
            <option value="Lupi">Lupi (Dark)</option>
            <option value="Ohchi">Ohchi (Light)</option>
            <option value="Palermo">Palermo (Light)</option>
            <option value="Panton">Panton (Light)</option>
            <option value="Picabia">Picabia (Light)</option>
            <option value="Picasso">Picasso (Light)</option>
            <option value="Rattner">Rattner (Dark)</option>
            <option value="Sidhu">Sidhu (Light)</option>
            <option value="Smith">Smith (Dark)</option>
            <option value="vonHeyl">vonHeyl (Dark)</option>
            <option value="Warhol">Warhol (Dark)</option>
          </optgroup>
        </select>
        <div id="palette-swatch-bar" style="display:inline-flex;align-items:center;height:14px;width:68px;border-radius:4px;overflow:hidden;border:1px solid rgba(255,255,255,0.4);" title="Active MoMA Palette + Background &amp; Card Swatches"></div>
        <span id="palette-bg-mode-badge" style="padding:2px 6px;border-radius:4px;font-size:10px;font-weight:700;background:rgba(15,23,42,0.45);color:#fff;white-space:nowrap;">☀ Light</span>
      </div>
      <button class="btn btn-primary" onclick="openCompoundModal()">➕ Submit Method Compound</button>
      <button class="btn btn-teal" onclick="openExcelMethodModal()">📊 Upload Excel Method (New Tab)</button>
      <button class="btn btn-teal" style="background:#6d28d9;" onclick="openLibraryUploadModal()">🧬 Upload Library (.lbp / .sdf)</button>
      <button class="btn btn-teal" onclick="openVialModal()">📦 Add Standard Vial</button>
      <button class="btn btn-glass" onclick="copyActiveTabToCSV()">📋 Copy CSV</button>
      <button class="btn btn-glass" onclick="downloadActiveTabCSV()">⬇️ Download CSV</button>
      <button class="btn btn-glass" onclick="exportStandaloneHTML()" title="Downloads an updated HTML file containing all newly submitted compounds & inventory updates so you can save it to OneDrive for all lab users">💾 Save Updated HTML for OneDrive</button>
      <button class="btn btn-glass" onclick="openPasswordModal()">🔑 Change Password</button>
      <button class="btn btn-glass" onclick="lockDashboard()">🔒 Log Out</button>
    </div>
  </div>

  <!-- TOP-LEVEL METHOD & INVENTORY TABS -->
  <div class="method-tabs" id="method-tabs-bar"></div>
</header>

<!-- MAIN CONTENT -->
<main class="container">
  <!-- KPI SUMMARY CARDS -->
  <div class="kpi-grid" id="kpi-grid"></div>

  <!-- MODE A: METHOD COMPOUNDS WORKSPACE -->
  <div id="workspace-methods">
    <div class="subnav-bar">
      <div class="subnav-group">
        <button class="subnav-btn active" id="subbtn-table" onclick="switchSubView('table')">📋 Method Compound Master Table</button>
        <button class="subnav-btn" id="subbtn-validation" onclick="switchSubView('validation')">🧪 LOD Validation Queue & Mix 8 Recipe</button>
        <button class="subnav-btn" id="subbtn-audit" onclick="switchSubView('audit')">🕒 Lab User Submissions & Audit Log (<span id="audit-count-badge">0</span>)</button>
      </div>
      <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">
        <button class="btn btn-outline" id="delete-method-tab-btn" style="display:none;color:#dc2626;border-color:#fca5a5;" onclick="deleteActiveMethodTab()">🗑️ Delete Current Tab</button>
      </div>
    </div>

    <!-- SUBVIEW 1: MASTER COMPOUND TABLE -->
    <section id="subview-table">
      <div class="toolbar-card">
        <!-- REAL-TIME NATURAL-LANGUAGE CHANGE / SUGGESTION INPUT BOX -->
        <div class="realtime-bar" id="live-command-bar">
          <div class="realtime-bar-row">
            <span class="realtime-badge">⚡ Real-Time Change / Suggestion Box</span>
            <input type="text" id="live-cmd-input" class="realtime-input"
              placeholder='Type a change or suggestion (e.g., "Set 11-hydroxy-THC RT to 8.80", "Mark dronabinol In 7600_CID library? Yes", "Set 4-ANPP first W93 analysis to 2026-05-15", "Check out fentanyl to Gillian")...'
              oninput="previewLiveCommand()"
              onkeydown="if(event.key==='Enter'){event.preventDefault();executeLiveCommand(false);}">
            <button type="button" class="btn btn-primary" style="padding:5px 10px;font-size:11.5px;" onclick="executeLiveCommand(false)" title="Immediately implement this change in real-time on the active method tab">⚡ Apply Live Now</button>
            <button type="button" class="btn btn-outline" style="padding:5px 10px;font-size:11.5px;color:#6d28d9;border-color:#c4b5fd;background:#f5f3ff;" onclick="executeLiveCommand(true)" title="Queue as a real-time suggestion chip so any user can review & 1-click apply">💡 Suggest for Review</button>
            <button type="button" class="btn btn-outline" style="padding:5px 8px;font-size:11px;" onclick="toggleLiveHelp()" title="Show clickable example commands">? Examples</button>
          </div>
          <div id="live-cmd-preview" style="display:none;font-size:11.5px;padding:4px 8px;border-radius:6px;background:#f0fdf4;border:1px solid #bbf7d0;color:#166534;font-weight:600;"></div>
          <div id="live-cmd-help" style="display:none;font-size:11.5px;padding:6px 10px;border-radius:6px;background:#ffffff;border:1px solid #cbd5e1;color:#334155;">
            <strong>Click any example to test real-time updates on the current tab:</strong>
            <span class="cmd-example-chip" onclick="fillLiveCmdExample('Set 11-hydroxy-THC RT to 8.80')">Set 11-hydroxy-THC RT to 8.80</span>
            <span class="cmd-example-chip" onclick="fillLiveCmdExample('Mark dronabinol In 7600_CID library? Yes')">Mark dronabinol In 7600_CID library? Yes</span>
            <span class="cmd-example-chip" onclick="fillLiveCmdExample('Set 4-ANPP first W93 analysis to 2026-05-15')">Set 4-ANPP first W93 analysis to 2026-05-15</span>
            <span class="cmd-example-chip" onclick="fillLiveCmdExample('Set atenolol Active in CDS W93 method to Active')">Set atenolol Active in CDS W93 method to Active</span>
            <span class="cmd-example-chip" onclick="fillLiveCmdExample('Check out fentanyl to Gillian')">Check out fentanyl to Gillian</span>
            <span class="cmd-example-chip" onclick="fillLiveCmdExample('Return fentanyl')">Return fentanyl</span>
            <span class="cmd-example-chip" onclick="fillLiveCmdExample('Set xylazine box to CAY2-41')">Set xylazine box to CAY2-41</span>
            <span class="cmd-example-chip" onclick="fillLiveCmdExample('Add compound Protonitazepyne, formula C25H32N4O3, RT 4.62')">Add compound Protonitazepyne, formula C25H32N4O3, RT 4.62</span>
          </div>
          <div id="live-suggestions-tray" style="display:none;flex-wrap:wrap;gap:6px;align-items:center;padding-top:2px;"></div>
        </div>

        <div class="toolbar-row">
          <div class="filter-group">
            <input type="text" id="search-input" class="search-box" placeholder="Search compound, formula, box, CAS..." oninput="renderTable()">
            <span class="ctrl-label" id="filter-status-label">Status:</span>
            <select id="filter-status" class="select-ctrl" style="max-width:125px;" onchange="renderTable()">
              <option value="ALL">All Statuses</option>
              <option value="Active">Active in Method (1)</option>
              <option value="Inactive / Pending">Inactive / Pending (0)</option>
              <option value="Removed">Removed from Method</option>
              <option value="LINKED_INV">Matched to Inventory Vial</option>
              <option value="MISSING_INV">No Linked Inventory Vial</option>
              <option value="CUSTOM">Newly Added / Edited by Users</option>
            </select>
            <span class="ctrl-label">Checked Out:</span>
            <select id="filter-checkout" class="select-ctrl" style="max-width:140px;" onchange="renderTable()">
              <option value="ALL">All (Out & In Box)</option>
              <option value="ANY_OUT">Checked Out (Any User)</option>
              <option value="Gillian">🟣 Checked Out: Gillian</option>
              <option value="Lauren">🟢 Checked Out: Lauren</option>
              <option value="Michael">🔵 Checked Out: Michael</option>
              <option value="Emilio">🔷 Checked Out: Emilio</option>
              <option value="IN_BOX">In Box (Not Checked Out)</option>
            </select>
            <button class="btn btn-outline" style="padding:5px 8px;font-size:11.5px;font-weight:700;color:#4c1d95;border-color:#c4b5fd;background:#f5f3ff;" onclick="sortByCheckedOut()" title="Sort table so Checked-Out compounds appear at the top, grouped by user">📌 Sort Checked-Out</button>
            <span id="filter-mix-wrap" style="display:inline-flex;align-items:center;gap:6px;">
              <span class="ctrl-label">Mix:</span>
              <select id="filter-mix" class="select-ctrl" style="max-width:85px;" onchange="renderTable()">
                <option value="ALL">All Mixes</option>
                <option value="1">Mix 1</option>
                <option value="2">Mix 2</option>
                <option value="3">Mix 3</option>
                <option value="4">Mix 4</option>
                <option value="5">Mix 5</option>
                <option value="6">Mix 6</option>
                <option value="7">Mix 7</option>
                <option value="8">Mix 8</option>
              </select>
            </span>
            <span class="ctrl-label">Date:</span>
            <select id="filter-date-preset" class="select-ctrl" style="max-width:105px;" onchange="handleDatePresetChange()">
              <option value="ALL">All Dates</option>
              <option value="2026-04-14">2026-04-14 (Initial X500R Go-Live)</option>
              <option value="2026-05-01">Early May 2026 (2026-05-01)</option>
              <option value="NONE">No Date (Pending Addition)</option>
              <option value="RANGE">Custom Date Range...</option>
            </select>
            <span id="date-range-wrap" style="display:none;align-items:center;gap:4px;">
              <input type="date" id="filter-date-from" class="date-ctrl" onchange="renderTable()">
              <span>to</span>
              <input type="date" id="filter-date-to" class="date-ctrl" onchange="renderTable()">
            </span>
            <span style="margin-left:auto;display:inline-flex;align-items:center;gap:6px;flex-shrink:0;">
              <span class="ctrl-label">Columns:</span>
              <select id="col-preset" class="select-ctrl" style="max-width:165px;" onchange="renderTable()">
                <option value="standard">Standard + Inventory View</option>
                <option value="validation">Full Validation & LODs</option>
                <option value="inventory">Chemical & Stock Inventory</option>
                <option value="all">All 44 Columns (Full Sheet)</option>
              </select>
              <button class="btn btn-outline" style="padding:5px 8px;font-size:11.5px;" onclick="resetFilters()">Reset Filters</button>
            </span>
          </div>
        </div>

        <div class="toolbar-row" style="border-top:1px solid #e2e8f0;padding-top:8px;">
          <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;min-width:0;">
            <div id="showing-summary" style="font-size:12px;font-weight:700;color:var(--ucsf-navy);white-space:nowrap;"></div>
            <div id="checkout-legend-bar" style="display:inline-flex;align-items:center;gap:5px;flex-wrap:wrap;font-size:11.5px;"></div>
          </div>
          <div style="display:flex;align-items:center;gap:6px;flex-wrap:nowrap;flex-shrink:0;">
            <span class="ctrl-label">Data Export:</span>
            <select id="method-export-menu" class="select-ctrl" style="font-size:11.5px;font-weight:700;padding:5px 10px;background:var(--ucsf-blue);color:#fff;border-color:var(--ucsf-blue);cursor:pointer;" onchange="handleMethodExportMenu(this)">
              <option value="" style="background:#fff;color:#0f172a;">📤 Export / Copy Data...</option>
              <option value="copy_filtered" style="background:#fff;color:#0f172a;">📋 Copy Filtered CSV (Current View)</option>
              <option value="copy_all" style="background:#fff;color:#0f172a;">📋 Copy All Columns CSV</option>
              <option value="download_filtered" style="background:#fff;color:#0f172a;">⬇️ Download Filtered CSV</option>
              <option value="download_full" style="background:#fff;color:#0f172a;">⬇️ Download Full Master CSV</option>
            </select>
          </div>
        </div>
      </div>

      <div class="table-card">
        <div class="table-scroll">
          <table class="comp-table" id="comp-table">
            <thead id="comp-thead"></thead>
            <tbody id="comp-tbody"></tbody>
          </table>
        </div>
      </div>
    </section>

    <!-- SUBVIEW 2: LOD VALIDATION QUEUE & MIX 8 -->
    <section id="subview-validation" style="display:none;">
      <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(460px,1fr));gap:16px;">
        <div class="table-card" style="padding:16px;">
          <h3 style="margin:0 0 6px 0;color:var(--ucsf-navy);">Mix 8 & High-LOD Preparation Recipe (Sheet 4)</h3>
          <p style="margin:0 0 12px 0;font-size:12.5px;color:var(--text-muted);">
            Dilution table for Mix 8 and high-LOD compounds from <code>CDHR_X500R_drug_list.xlsx</code>.
          </p>
          <table class="comp-table">
            <thead>
              <tr>
                <th>Compound</th>
                <th>Stock (µg/mL)</th>
                <th>Final (µg/mL)</th>
                <th>Dilution for 1 mL Mix (µL)</th>
                <th>Validated in Phase 1?</th>
              </tr>
            </thead>
            <tbody id="mix8-tbody"></tbody>
          </table>
        </div>

        <div class="table-card" style="padding:16px;">
          <h3 style="margin:0 0 6px 0;color:var(--ucsf-navy);">Outstanding Urine & Serum LOD Validation Queue (Sheets 2 & 3)</h3>
          <p style="margin:0 0 12px 0;font-size:12.5px;color:var(--text-muted);">
            Compounds flagged as <code>OUTSTANDING</code> or <code>REVISIT with extended range</code> in the Urine and Serum LOD sheets.
          </p>
          <table class="comp-table">
            <thead>
              <tr>
                <th>Matrix</th>
                <th>Compound</th>
                <th>Current / Target LOD (ng/mL)</th>
                <th>Validation Note</th>
              </tr>
            </thead>
            <tbody id="outstanding-tbody"></tbody>
          </table>
        </div>
      </div>
    </section>

    <!-- SUBVIEW 3: USER SUBMISSIONS & AUDIT LOG -->
    <section id="subview-audit" style="display:none;">
      <div class="table-card" style="padding:16px;">
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:12px;flex-wrap:wrap;gap:10px;">
          <div>
            <h3 style="margin:0 0 4px 0;color:var(--ucsf-navy);">Lab User Submissions & Inventory Change Log</h3>
            <p style="margin:0;font-size:12.5px;color:var(--text-muted);">
              Tracks all new compounds submitted, method edits, and standard vial checkouts/additions. Click <strong>💾 Save Updated HTML for OneDrive</strong> in the header to permanently bake these changes into the shared OneDrive file.
            </p>
          </div>
          <div style="display:flex;gap:8px;">
            <button class="btn btn-outline" onclick="exportAuditLogCSV()">⬇️ Export Change Log CSV</button>
            <button class="btn btn-outline" style="color:#dc2626;border-color:#fca5a5;" onclick="clearLocalCustomChanges()">🗑️ Reset Local Unsaved Changes</button>
          </div>
        </div>
        <table class="comp-table">
          <thead>
            <tr>
              <th>Timestamp</th>
              <th>Action</th>
              <th>Target / Method</th>
              <th>Compound / Standard</th>
              <th>Date / Box</th>
              <th>Status</th>
              <th>User</th>
              <th>Notes / Summary</th>
            </tr>
          </thead>
          <tbody id="audit-tbody"></tbody>
        </table>
      </div>
    </section>
  </div>

  <!-- MODE B: STANDARDS INVENTORY 2026 WORKSPACE -->
  <div id="workspace-inventory" style="display:none;">
    <div class="toolbar-card">
      <div class="toolbar-row">
        <div class="filter-group">
          <input type="text" id="inv-search-input" class="search-box" placeholder="Search 996 vials by standard name, box (A-01, CAY2, HJL), manufacturer, catalog #, lot #, CAS #, or notes..." oninput="renderInventoryTable()">
          <span class="ctrl-label">Sheet:</span>
          <select id="inv-filter-sheet" class="select-ctrl" onchange="renderInventoryTable()">
            <option value="Current Stock">Current Stock (Active Inventory)</option>
            <option value="Old Orders">Old Orders (Archived / Empty)</option>
            <option value="ALL">All Vials (Current + Old Orders)</option>
          </select>
          <span class="ctrl-label">Vial Status:</span>
          <select id="inv-filter-status" class="select-ctrl" onchange="renderInventoryTable()">
            <option value="ALL">All Vial Statuses</option>
            <option value="OPEN">OPEN / OPENED</option>
            <option value="UNOPENED">UNOPENED</option>
            <option value="CHECKED_OUT">Checked Out by Lab Member</option>
            <option value="IN_METHOD">Used in X500R / Lab Method</option>
            <option value="NOT_IN_METHOD">Not Yet in X500R Method</option>
            <option value="TAGGED">DEA / Tagged (Has Tag = Y)</option>
            <option value="EXPIRED">Past Expiration Date</option>
          </select>
          <span class="ctrl-label">Manufacturer:</span>
          <select id="inv-filter-mfg" class="select-ctrl" onchange="renderInventoryTable()">
            <option value="ALL">All Manufacturers</option>
            <option value="Cerilliant">Cerilliant</option>
            <option value="Cayman">Cayman</option>
            <option value="ChemFaces">ChemFaces</option>
            <option value="Sigma">Sigma</option>
            <option value="TCI">TCI</option>
            <option value="OTHER">Other Manufacturers</option>
          </select>
          <button class="btn btn-outline" onclick="resetInventoryFilters()">Reset Filters</button>
        </div>
      </div>

      <div class="toolbar-row" style="border-top:1px solid #e2e8f0;padding-top:10px;">
        <div id="inv-showing-summary" style="font-size:12.5px;font-weight:700;color:var(--ucsf-navy);"></div>
        <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">
          <span style="font-size:12px;color:var(--text-muted);margin-right:6px;">
            Legend: <span class="badge badge-custom">Checked Out (Gillian / Emilio / etc.)</span>
            <span class="badge badge-blue">Linked to X500R Method</span>
          </span>
          <button class="btn btn-teal" onclick="openVialModal()">📦 + Add New Standard Vial</button>
          <button class="btn btn-outline" onclick="copyInventoryToCSV()">📋 Copy Inventory CSV</button>
          <button class="btn btn-blue" onclick="downloadInventoryCSV()">⬇️ Download Inventory CSV</button>
        </div>
      </div>
    </div>

    <div class="table-card">
      <div class="table-scroll">
        <table class="comp-table" id="inv-table">
          <thead id="inv-thead"></thead>
          <tbody id="inv-tbody"></tbody>
        </table>
      </div>
    </div>
  </div>
</main>

<!-- ADD / EDIT METHOD COMPOUND MODAL -->
<div class="modal-backdrop" id="compound-modal">
  <div class="modal-box">
    <div class="modal-header">
      <h3 id="compound-modal-title">Submit New Compound to Lab Method</h3>
      <button class="modal-close" onclick="closeCompoundModal()">✕</button>
    </div>
    <form onsubmit="saveCompoundForm(event)">
      <div class="modal-body">
        <input type="hidden" id="edit-comp-id">
        <div style="background:#f0f9ff;border:1px solid #bae6fd;border-radius:8px;padding:10px 14px;margin-bottom:14px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px;">
          <div style="font-size:12.5px;color:#0369a1;">
            <strong>💡 Auto-Fill from Standards Inventory (891 Current Stock Vials):</strong> Pick a vial below to auto-populate Box Location, Concentration, Solvent, Manufacturer, Catalog #, Lot #, and CAS #.
          </div>
          <input type="text" list="inv-datalist" id="f-inv-lookup" placeholder="Type standard name or box to auto-fill..." style="min-width:280px;padding:6px 10px;border:1px solid #7dd3fc;border-radius:6px;font-size:12.5px;" onchange="applyInventoryAutofill()">
          <datalist id="inv-datalist"></datalist>
        </div>

        <div class="form-grid">
          <div class="form-field">
            <label>Lab Method / Instrument *</label>
            <select id="f-method" required onchange="syncCompoundModalForMethod()"></select>
          </div>
          <div class="form-field">
            <label>Compound Name *</label>
            <input type="text" id="f-compound" placeholder="e.g., N-pyrrolidino etonitazene" required>
          </div>
          <div class="form-field">
            <label>Date Added to Method *</label>
            <input type="date" id="f-date-added" required>
          </div>
          <div class="form-field" id="field-wrap-w93-date" style="display:none;">
            <label>Date of first W93 analysis</label>
            <input type="date" id="f-date-first-w93">
          </div>
          <div class="form-field">
            <label id="label-f-status">Method Status *</label>
            <select id="f-status">
              <option value="Active">Active in Method (1)</option>
              <option value="Inactive / Pending">Inactive / Pending Validation (0)</option>
              <option value="Removed">Removed from Method</option>
            </select>
          </div>
          <div class="form-field">
            <label id="label-f-in-library">Added to Instrument Library?</label>
            <select id="f-in-library">
              <option value="Yes">Yes</option>
              <option value="No">No</option>
              <option value="Needs Review" id="opt-lib-needs-review">Needs Review</option>
            </select>
          </div>
          <div class="form-field">
            <label>Submitted / Updated By (Initials) *</label>
            <input type="text" id="f-submitted-by" placeholder="e.g., MM, KL, GR, PY" required>
          </div>

          <div class="form-field">
            <label>Molecular Formula (auto-calculates mass)</label>
            <input type="text" id="f-formula" placeholder="e.g., C22H28N2O" oninput="autoCalcMassFromFormula()">
          </div>
          <div class="form-field">
            <label>Monoisotopic Mass (Da)</label>
            <input type="text" id="f-mass" placeholder="e.g., 336.2202" oninput="autoCalcMHFromMass()">
          </div>
          <div class="form-field">
            <label>Protonated [M+H]+ (m/z)</label>
            <input type="text" id="f-mh" placeholder="e.g., 337.2274">
          </div>

          <div class="form-field" id="field-wrap-rt-x500r">
            <label>X500R RT (min, Current)</label>
            <input type="text" id="f-rt-x500r" placeholder="e.g., 5.12">
          </div>
          <div class="form-field" id="field-wrap-rt-dec25">
            <label>X500R Dec 2025 RT (min)</label>
            <input type="text" id="f-rt-dec25" placeholder="e.g., 5.08">
          </div>
          <div class="form-field" id="field-wrap-rt-7600">
            <label id="label-f-rt-7600">7600 RT (min)</label>
            <input type="text" id="f-rt-7600" placeholder="e.g., 4.95">
          </div>
          <div class="form-field" id="field-wrap-last-rt-date">
            <label>Date of last RT update</label>
            <input type="date" id="f-date-last-rt">
          </div>

          <div class="form-field" id="field-wrap-mix1">
            <label>Primary Mix</label>
            <input type="text" id="f-mix1" placeholder="e.g., 1 - 8">
          </div>
          <div class="form-field" id="field-wrap-mix2">
            <label>Secondary Mix</label>
            <input type="text" id="f-mix2" placeholder="Optional">
          </div>
          <div class="form-field" id="field-wrap-checked-7600">
            <label>Checked in 7600 Library?</label>
            <input type="text" id="f-checked-7600" placeholder="yes / no">
          </div>

          <div class="form-field">
            <label>Stock Box Location</label>
            <input type="text" id="f-box" placeholder="e.g., CAY2-37 or D-11">
          </div>
          <div class="form-field">
            <label>Box Position / Well</label>
            <input type="text" id="f-pos" placeholder="e.g., A11">
          </div>
          <div class="form-field">
            <label>Stock Concentration</label>
            <input type="text" id="f-conc" placeholder="e.g., 1.0 mg/mL">
          </div>
          <div class="form-field">
            <label>Stock Solvent</label>
            <input type="text" id="f-solvent" placeholder="e.g., Methanol">
          </div>
          <div class="form-field">
            <label>Manufacturer</label>
            <input type="text" id="f-mfg" placeholder="e.g., Cerilliant / Cayman">
          </div>
          <div class="form-field">
            <label>Catalog # (COA)</label>
            <input type="text" id="f-cat" placeholder="e.g., F-013">
          </div>
          <div class="form-field">
            <label>Lot / Batch #</label>
            <input type="text" id="f-lot" placeholder="e.g., FE05072413">
          </div>
          <div class="form-field">
            <label>CAS #</label>
            <input type="text" id="f-cas" placeholder="e.g., 437-38-7">
          </div>

          <div class="form-field">
            <label>Consensus Urine LOD (ng/mL)</label>
            <input type="text" id="f-lod-u" placeholder="e.g., 5">
          </div>
          <div class="form-field">
            <label>Consensus Serum LOD (ng/mL)</label>
            <input type="text" id="f-lod-s" placeholder="e.g., 1">
          </div>
          <div class="form-field">
            <label>InChIKey</label>
            <input type="text" id="f-inchikey" placeholder="e.g., YCBKSSAWEUDACY-IAGOWNOFSA-N">
          </div>
          <div class="form-field">
            <label>Possible Isomers</label>
            <input type="text" id="f-isomers" placeholder="e.g., ortho/meta/para isomers">
          </div>
          <div class="form-field">
            <label>Synonyms</label>
            <input type="text" id="f-synonyms" placeholder="e.g., common/street names">
          </div>

          <div class="form-field full-width">
            <label>Comments / Validation Notes</label>
            <textarea id="f-comments" rows="2" placeholder="Any notes on MS/MS transitions, matrix interference, false positives, or validation status..."></textarea>
          </div>
        </div>
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-outline" onclick="closeCompoundModal()">Cancel</button>
        <button type="submit" class="btn btn-primary">💾 Save Method Compound</button>
      </div>
    </form>
  </div>
</div>

<!-- ADD / EDIT INVENTORY VIAL MODAL -->
<div class="modal-backdrop" id="vial-modal">
  <div class="modal-box">
    <div class="modal-header" style="background:#0f766e;">
      <h3 id="vial-modal-title">Add / Edit Chemical Standard Vial</h3>
      <button class="modal-close" onclick="closeVialModal()">✕</button>
    </div>
    <form onsubmit="saveVialForm(event)">
      <div class="modal-body">
        <input type="hidden" id="edit-vial-id">
        <div class="form-grid">
          <div class="form-field">
            <label>Standard Name (include salt / alias) *</label>
            <input type="text" id="v-name" required placeholder="e.g., Fentanyl HCl (1 mg/mL)">
          </div>
          <div class="form-field">
            <label>Inventory Sheet</label>
            <select id="v-sheet">
              <option value="Current Stock">Current Stock (Active)</option>
              <option value="Old Orders">Old Orders (Archived / Empty)</option>
            </select>
          </div>
          <div class="form-field">
            <label>Vial Status *</label>
            <select id="v-status">
              <option value="UNOPENED">UNOPENED</option>
              <option value="OPEN">OPEN / OPENED</option>
              <option value="EMPTY">EMPTY</option>
            </select>
          </div>
          <div class="form-field">
            <label>Box / Freezer Location *</label>
            <input type="text" id="v-loc" required placeholder="e.g., A-23, CAY3-02, -80 / 01">
          </div>
          <div class="form-field">
            <label>Checked Out By (In Possession)</label>
            <select id="v-checked-out">
              <option value="">In Box (Not Checked Out)</option>
              <option value="Gillian">Gillian has it out</option>
              <option value="Lauren">Lauren has it out</option>
              <option value="Emilio">Emilio has it out</option>
              <option value="Hannah">Hannah has it out</option>
              <option value="Michael">Michael has it out</option>
              <option value="Other Lab User">Other Lab User (see notes)</option>
            </select>
          </div>
          <div class="form-field">
            <label>Concentration or Amount (Powder)</label>
            <input type="text" id="v-conc" placeholder="e.g., 1.0 mg/mL or 5 mg">
          </div>
          <div class="form-field">
            <label>Solvent</label>
            <input type="text" id="v-solvent" placeholder="e.g., Methanol, Acetonitrile, powder">
          </div>
          <div class="form-field">
            <label>Manufacturer</label>
            <input type="text" id="v-mfg" placeholder="e.g., Cerilliant, Cayman, Sigma">
          </div>
          <div class="form-field">
            <label>Catalog Number (COA)</label>
            <input type="text" id="v-cat" placeholder="e.g., F-013">
          </div>
          <div class="form-field">
            <label>Lot / Batch No.</label>
            <input type="text" id="v-lot" placeholder="e.g., FE05072413">
          </div>
          <div class="form-field">
            <label>CAS No.</label>
            <input type="text" id="v-cas" placeholder="e.g., 437-38-7">
          </div>
          <div class="form-field">
            <label>Expiration Date</label>
            <input type="date" id="v-exp">
          </div>
          <div class="form-field">
            <label>Receive Date</label>
            <input type="date" id="v-rec">
          </div>
          <div class="form-field">
            <label>Open Date</label>
            <input type="date" id="v-open">
          </div>
          <div class="form-field">
            <label>Initial Volume or Mass</label>
            <input type="text" id="v-vol" placeholder="e.g., 1000 uL or 1 mL">
          </div>
          <div class="form-field">
            <label>DEA / Tag Required (Y/N)</label>
            <select id="v-tag">
              <option value="">No</option>
              <option value="Y">Yes (DEA / Tagged)</option>
              <option value="N">N</option>
            </select>
          </div>
          <div class="form-field">
            <label>Updated By (Initials) *</label>
            <input type="text" id="v-user" required placeholder="e.g., MM, GR, KL">
          </div>
          <div class="form-field full-width">
            <label>Additional Notes</label>
            <textarea id="v-notes" rows="2" placeholder="Any notes on COA, storage, or ordering status..."></textarea>
          </div>
        </div>
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-outline" onclick="closeVialModal()">Cancel</button>
        <button type="submit" class="btn btn-teal">💾 Save Inventory Vial</button>
      </div>
    </form>
  </div>
</div>

<!-- LINKED VIALS INSPECTOR MODAL -->
<div class="modal-backdrop" id="linked-vials-modal">
  <div class="modal-box" style="max-width:920px;">
    <div class="modal-header">
      <h3 id="linked-vials-title">Matching Standards Inventory Vials</h3>
      <button class="modal-close" onclick="document.getElementById('linked-vials-modal').classList.remove('open')">✕</button>
    </div>
    <div class="modal-body" id="linked-vials-body"></div>
    <div class="modal-footer">
      <button type="button" class="btn btn-outline" onclick="document.getElementById('linked-vials-modal').classList.remove('open')">Close</button>
    </div>
  </div>
</div>

<!-- CHANGE PASSWORD MODAL -->
<div class="modal-backdrop" id="pw-modal">
  <div class="modal-box" style="max-width:420px;">
    <div class="modal-header">
      <h3>Manage Lab User Password</h3>
      <button class="modal-close" onclick="closePasswordModal()">✕</button>
    </div>
    <form onsubmit="saveNewPassword(event)">
      <div class="modal-body">
        <div class="form-field" style="margin-bottom:12px;">
          <label>User Account *</label>
          <select id="pw-user-select" required>
            <option value="Gillian">Gillian</option>
            <option value="Lauren">Lauren</option>
            <option value="Michael">Michael</option>
          </select>
        </div>
        <div class="form-field" style="margin-bottom:12px;">
          <label>New Password *</label>
          <input type="password" id="new-pw-1" required placeholder="Enter new password...">
        </div>
        <div class="form-field" style="margin-bottom:8px;">
          <label>Confirm New Password *</label>
          <input type="password" id="new-pw-2" required placeholder="Re-enter new password...">
        </div>
        <p style="font-size:12px;color:var(--text-muted);margin:8px 0 0 0;">
          Note: After changing a user's password, click <strong>💾 Save Updated HTML for OneDrive</strong> if you want the updated password baked into the shared OneDrive HTML file.
        </p>
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-outline" onclick="closePasswordModal()">Cancel</button>
        <button type="submit" class="btn btn-primary">Update Password</button>
      </div>
    </form>
  </div>
</div>

<!-- NEW METHOD TAB MODAL -->
<div class="modal-backdrop" id="method-modal">
  <div class="modal-box" style="max-width:440px;">
    <div class="modal-header">
      <h3>Add Lab Method / Instrument Tab</h3>
      <button class="modal-close" onclick="closeNewMethodModal()">✕</button>
    </div>
    <form onsubmit="saveNewMethodTab(event)">
      <div class="modal-body">
        <div class="form-field" style="margin-bottom:12px;">
          <label>Method / Instrument Name *</label>
          <input type="text" id="new-method-name" required placeholder="e.g., ZenoTOF 7600 (IDA/SWATH) or TripleTOF 5600">
        </div>
        <div class="form-field">
          <label>Pre-populate compounds?</label>
          <select id="new-method-seed">
            <option value="empty">Start with an empty method tab (add compounds manually)</option>
            <option value="clone_7600">Seed from X500R sheet compounds that have 7600 RT / Library data</option>
            <option value="clone_5600">Seed from X500R sheet compounds that have 5600 Urine LOD data</option>
          </select>
        </div>
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-outline" onclick="closeNewMethodModal()">Cancel</button>
        <button type="submit" class="btn btn-primary">Create Method Tab</button>
      </div>
    </form>
  </div>
</div>

<!-- UPLOAD FULL METHOD BY EXCEL MODAL -->
<div class="modal-backdrop" id="excel-method-modal">
  <div class="modal-box" style="max-width:940px;">
    <div class="modal-header" style="background:linear-gradient(90deg, #052049 0%, #007CBE 100%);">
      <h3>📊 Upload Full Method by Excel (.xlsx / .csv) → Generate New Method Tab</h3>
      <button class="modal-close" onclick="closeExcelMethodModal()">✕</button>
    </div>
    <form onsubmit="confirmExcelMethodImport(event)">
      <div class="modal-body">
        <!-- STEP 1: FILE PICKER -->
        <div style="border:2px dashed #93c5fd;background:#f0f9ff;border-radius:10px;padding:18px;text-align:center;margin-bottom:16px;">
          <div style="font-size:14px;font-weight:800;color:var(--ucsf-navy);margin-bottom:6px;">
            Select an Excel Workbook (<code>.xlsx</code>) or CSV (<code>.csv</code>) Method File
          </div>
          <div style="font-size:12.5px;color:#475569;margin-bottom:12px;">
            Parses multi-sheet Excel files directly in your browser, auto-detects columns (Compound, RT, Formula, Exact Mass, [M+H]+, Mix, Box, LODs), and cross-links every compound to the 2026 Standards Inventory.
          </div>
          <input type="file" id="excel-method-file-input" accept=".xlsx,.csv,.tsv,.json" onchange="handleExcelMethodFileChosen(event)" style="font-size:13px;padding:6px;background:#fff;border:1px solid #cbd5e1;border-radius:6px;cursor:pointer;">
          <div id="excel-parse-status" style="margin-top:8px;font-size:12.5px;font-weight:700;color:#0f766e;"></div>
        </div>

        <!-- STEP 2: TAB NAME & WORKSHEET CONFIG -->
        <div id="excel-config-section" style="display:none;">
          <div class="form-grid" style="margin-bottom:14px;">
            <div class="form-field">
              <label>New Method Tab Name *</label>
              <input type="text" id="excel-new-tab-name" required placeholder="e.g., ZenoTOF 7600 Method, QTRAP 6500, TripleTOF 5600">
            </div>
            <div class="form-field">
              <label>Excel Worksheet to Import</label>
              <select id="excel-sheet-select" onchange="onExcelSheetOrHeaderChanged()"></select>
            </div>
            <div class="form-field">
              <label>Header Row in Sheet</label>
              <select id="excel-header-row-select" onchange="onExcelSheetOrHeaderChanged()">
                <option value="0">Row 1 (First Row)</option>
                <option value="1">Row 2</option>
                <option value="2">Row 3</option>
                <option value="3">Row 4</option>
              </select>
            </div>
            <div class="form-field">
              <label>Default Date Added (if blank in sheet)</label>
              <input type="date" id="excel-default-date">
            </div>
            <div class="form-field">
              <label>If Tab Name Already Exists</label>
              <select id="excel-existing-mode">
                <option value="append">Append / Merge into Tab</option>
                <option value="replace">Replace Existing Compounds in Tab</option>
              </select>
            </div>
            <div class="form-field">
              <label>Uploaded By (User) *</label>
              <input type="text" id="excel-uploaded-by" required placeholder="e.g., Gillian, Lauren, Michael">
            </div>
          </div>

          <!-- STEP 3: COLUMN MAPPING -->
          <div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:12px 14px;margin-bottom:14px;">
            <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;flex-wrap:wrap;gap:8px;">
              <div style="font-size:12.5px;font-weight:800;color:var(--ucsf-navy);">
                🔗 Auto-Detected Column Mapping (Adjust any column dropdown if needed)
              </div>
              <label style="font-size:12px;color:#475569;display:inline-flex;align-items:center;gap:6px;cursor:pointer;">
                <input type="checkbox" id="excel-preserve-extra" checked onchange="renderExcelPreview()">
                Include extra unmapped Excel columns in Comments / Notes
              </label>
            </div>
            <div class="form-grid" id="excel-col-map-grid" style="grid-template-columns:repeat(auto-fit, minmax(195px, 1fr));gap:10px;"></div>
          </div>

          <!-- STEP 4: LIVE PREVIEW -->
          <div>
            <div id="excel-preview-summary" style="font-size:12.5px;font-weight:700;color:var(--ucsf-navy);margin-bottom:6px;"></div>
            <div class="table-scroll" style="max-height:210px;border:1px solid #e2e8f0;border-radius:8px;">
              <table class="comp-table" id="excel-preview-table">
                <thead id="excel-preview-thead"></thead>
                <tbody id="excel-preview-tbody"></tbody>
              </table>
            </div>
          </div>
        </div>
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-outline" onclick="closeExcelMethodModal()">Cancel</button>
        <button type="submit" class="btn btn-primary" id="excel-import-submit-btn" disabled>🚀 Generate Method Tab & Import Compounds</button>
      </div>
    </form>
  </div>
</div>

<!-- UPLOAD MS/MS SPECTRAL LIBRARY (.LBP / .SDF) MODAL -->
<div class="modal-backdrop" id="lib-upload-modal">
  <div class="modal-box" style="max-width:960px;">
    <div class="modal-header" style="background:linear-gradient(90deg, #4c1d95 0%, #6d28d9 100%);">
      <h3>🧬 Upload Spectral Library File (.lbp / .sdf)</h3>
      <button class="modal-close" onclick="closeLibraryUploadModal()">✕</button>
    </div>
    <form onsubmit="confirmLibraryUpload(event)">
      <div class="modal-body">
        <!-- STEP 1: FILE PICKER -->
        <div style="border:2px dashed #c4b5fd;background:#f5f3ff;border-radius:10px;padding:18px;text-align:center;margin-bottom:16px;">
          <div style="font-size:14px;font-weight:800;color:#4c1d95;margin-bottom:6px;">
            Select a SCIEX LibraryView / SCIEX OS Binary Library (<code>.lbp</code>) or Structure-Data File (<code>.sdf</code>)
          </div>
          <div style="font-size:12.5px;color:#475569;margin-bottom:12px;">
            Parses binary SQL CE 3.5 <code>.lbp</code> packages and <code>.sdf</code> spectral libraries directly in your browser — extracting Compound Names, Formulas, Exact Masses, Precursor <em>m/z</em>, Library RTs, Collision Energies, and High-Resolution MS/MS Centroid Spectra.
          </div>
          <input type="file" id="lib-file-input" accept=".lbp,.sdf" onchange="handleLibraryFileChosen(event)" style="font-size:13px;padding:6px;background:#fff;border:1px solid #cbd5e1;border-radius:6px;cursor:pointer;">
          <div id="lib-parse-status" style="margin-top:8px;font-size:12.5px;font-weight:700;color:#6d28d9;"></div>
        </div>

        <!-- STEP 2: IMPORT MODE & TARGET TAB -->
        <div id="lib-config-section" style="display:none;">
          <div class="form-grid" style="margin-bottom:14px;">
            <div class="form-field">
              <label>How to Apply Uploaded Library *</label>
              <select id="lib-import-mode" onchange="onLibImportModeChanged()">
                <option value="sync_and_append">Sync Existing Method Tab + Append New Library Compounds</option>
                <option value="sync_only">Sync / Enrich Existing Method Tab Only (Mark In Library = Yes & Attach MS/MS)</option>
                <option value="new_tab">Create a Brand-New Method / Library Tab from File</option>
              </select>
            </div>
            <div class="form-field" id="lib-target-tab-wrap">
              <label>Target Method Tab *</label>
              <select id="lib-target-method-select" onchange="renderLibraryPreview()"></select>
            </div>
            <div class="form-field" id="lib-new-tab-wrap" style="display:none;">
              <label>New Method / Library Tab Name *</label>
              <input type="text" id="lib-new-tab-name" placeholder="e.g., X500R Library (.lbp) or 7600 CID Library" oninput="renderLibraryPreview()">
            </div>
            <div class="form-field">
              <label>Default Date Added (for new entries)</label>
              <input type="date" id="lib-default-date" onchange="renderLibraryPreview()">
            </div>
            <div class="form-field">
              <label>Uploaded By (User) *</label>
              <input type="text" id="lib-uploaded-by" required placeholder="e.g., Gillian, Lauren, Michael">
            </div>
          </div>

          <!-- STEP 3: PREVIEW TABLE -->
          <div>
            <div id="lib-preview-summary" style="font-size:12.5px;font-weight:700;color:#4c1d95;margin-bottom:6px;"></div>
            <div class="table-scroll" style="max-height:240px;border:1px solid #e2e8f0;border-radius:8px;">
              <table class="comp-table" id="lib-preview-table">
                <thead>
                  <tr>
                    <th>Compound Name</th>
                    <th>Match in Target Tab</th>
                    <th>Standards Inventory</th>
                    <th>Formula</th>
                    <th>Exact Mass (Da)</th>
                    <th>Precursor [M+H]+</th>
                    <th>Library RT (min)</th>
                    <th>MS/MS Peaks</th>
                    <th>Top Fragment Ions (m/z)</th>
                  </tr>
                </thead>
                <tbody id="lib-preview-tbody"></tbody>
              </table>
            </div>
          </div>
        </div>
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-outline" onclick="closeLibraryUploadModal()">Cancel</button>
        <button type="submit" class="btn btn-primary" id="lib-import-submit-btn" style="background:#6d28d9;" disabled>🧬 Apply Library (.lbp / .sdf) to Dashboard</button>
      </div>
    </form>
  </div>
</div>

<!-- MS/MS SPECTRUM VIEWER MODAL -->
<div class="modal-backdrop" id="msms-modal">
  <div class="modal-box" style="max-width:900px;">
    <div class="modal-header" style="background:linear-gradient(90deg, #052049 0%, #4c1d95 100%);">
      <h3 id="msms-modal-title">MS/MS Reference Spectrum</h3>
      <button class="modal-close" onclick="document.getElementById('msms-modal').classList.remove('open')">✕</button>
    </div>
    <div class="modal-body" id="msms-modal-body"></div>
    <div class="modal-footer">
      <button type="button" class="btn btn-outline" onclick="document.getElementById('msms-modal').classList.remove('open')">Close</button>
    </div>
  </div>
</div>

<div id="toast-box"></div>

<script id="embedded-lab-data" type="application/json">__EMBEDDED_JSON_DATA__</script>

<script>
const STORAGE_KEY_DATA = "lynch_lab_dashboard_custom_state_v2";
const STORAGE_KEY_USERS = "lynch_lab_dashboard_users_v1";

let EMBEDDED_STATE = JSON.parse(document.getElementById("embedded-lab-data").textContent);
let DEFAULT_USERS = EMBEDDED_STATE.users || {
  "Gillian": "G1ll1an",
  "Lauren": "Laur3n",
  "Michael": "M1chael"
};

let state = {
  currentUser: null,
  methods: EMBEDDED_STATE.methods || ["X500R (CDHR)", "7600 (CDS W93)"],
  activeMethod: "X500R (CDHR)",
  compounds: [],
  inventory: [],
  auditLog: EMBEDDED_STATE.audit_log || [],
  suggestions: [],
  lastFlashedId: null,
  sortCol: "compound",
  sortAsc: true,
  invSortCol: "location",
  invSortAsc: true
};

const REALTIME_CHANNEL = (typeof BroadcastChannel !== "undefined")
  ? new BroadcastChannel("lynch_lab_realtime_channel_v1")
  : null;
if (REALTIME_CHANNEL) {
  REALTIME_CHANNEL.onmessage = (ev) => {
    if (ev.data && ev.data.type === "STATE_UPDATED") {
      reloadSavedStateFromStorage();
    }
  };
}
window.addEventListener("storage", (ev) => {
  if (ev.key === STORAGE_KEY_DATA) {
    reloadSavedStateFromStorage();
  }
});

const ATOMIC_MASSES = {
  H: 1.007825032, C: 12.0000000, N: 14.003074004, O: 15.994914620,
  F: 18.99840316, Na: 22.98976928, P: 30.973761998, S: 31.972071174,
  Cl: 34.96885268, K: 38.96370649, Br: 78.9183371, I: 126.904473
};

const COLUMN_DEFS = [
  { key: "status", label: "Status", presets: ["standard", "validation", "inventory", "all"] },
  { key: "compound", label: "Compound", presets: ["standard", "validation", "inventory", "all"] },
  { key: "checked_out_by", label: "Check Out / Holder", presets: ["standard", "validation", "inventory", "all"] },
  { key: "box_location", label: "Box Location", presets: ["standard", "inventory", "all"] },
  { key: "method", label: "Method", presets: ["standard", "all"] },
  { key: "date_added", label: "Date Added", presets: ["standard", "validation", "inventory", "all"] },
  { key: "date_first_w93", label: "Date of first W93 analysis", presets: ["all"] },
  { key: "in_library", label: "In Library?", presets: ["standard", "all"] },
  { key: "rt_x500r_jul26", label: "X500R RT (min)", presets: ["standard", "validation", "all"], numeric: true },
  { key: "rt_x500r_dec25", label: "X500R Dec '25 RT", presets: ["standard", "all"], numeric: true },
  { key: "rt_7600_jul26", label: "7600 RT (min)", presets: ["standard", "validation", "all"], numeric: true },
  { key: "date_last_rt_update", label: "Date of last RT update", presets: ["standard", "validation", "all"] },
  { key: "formula", label: "Formula", presets: ["standard", "inventory", "all"] },
  { key: "monoisotopic_mass", label: "Exact Mass (Da)", presets: ["standard", "inventory", "all"], numeric: true },
  { key: "mh_plus_mz", label: "[M+H]+ (m/z)", presets: ["standard", "inventory", "all"], numeric: true },
  { key: "primary_mix", label: "Primary Mix", presets: ["standard", "inventory", "all"], numeric: true },
  { key: "secondary_mix", label: "Sec. Mix", presets: ["standard", "inventory", "all"] },
  { key: "box_position", label: "Pos/Well", presets: ["inventory", "all"] },
  { key: "stock_conc", label: "Stock Conc.", presets: ["standard", "inventory", "all"] },
  { key: "stock_solvent", label: "Solvent", presets: ["standard", "inventory", "all"] },
  { key: "manufacturer", label: "Manufacturer", presets: ["standard", "inventory", "all"] },
  { key: "catalog_no", label: "Catalog # (COA)", presets: ["inventory", "all"] },
  { key: "lot_no", label: "Lot / Batch #", presets: ["inventory", "all"] },
  { key: "cas_no", label: "CAS #", presets: ["inventory", "all"] },
  { key: "vial_status", label: "Vial Status", presets: ["standard", "inventory", "all"] },
  { key: "lod_x500r_u_consensus", label: "X500R Urine LOD (ng/mL)", presets: ["standard", "validation", "all"], numeric: true },
  { key: "lod_x500r_s_consensus", label: "X500R Serum LOD (ng/mL)", presets: ["standard", "validation", "all"], numeric: true },
  { key: "lod_x500r_u_p1", label: "X500R Urine P1", presets: ["validation", "all"], numeric: true },
  { key: "lod_x500r_u_p2", label: "X500R Urine P2", presets: ["validation", "all"], numeric: true },
  { key: "x500r_u_agree", label: "Urine Agree", presets: ["validation", "all"] },
  { key: "lod_x500r_s_p1", label: "X500R Serum P1", presets: ["validation", "all"], numeric: true },
  { key: "lod_x500r_s_p2", label: "X500R Serum P2", presets: ["validation", "all"], numeric: true },
  { key: "x500r_s_agree", label: "Serum Agree", presets: ["validation", "all"] },
  { key: "lod_7600_u_swath", label: "7600 Urine SWATH LOD", presets: ["validation", "all"], numeric: true },
  { key: "lod_7600_u_ida", label: "7600 Urine IDA LOD", presets: ["validation", "all"], numeric: true },
  { key: "lod_7600_s_ida", label: "7600 Serum IDA LOD", presets: ["validation", "all"], numeric: true },
  { key: "lod_5600_u_ida", label: "5600 Urine IDA LOD", presets: ["validation", "all"], numeric: true },
  { key: "lod_brown_u", label: "Brown Urine LOD", presets: ["validation", "all"], numeric: true },
  { key: "lod_brown_s", label: "Brown Serum LOD", presets: ["validation", "all"], numeric: true },
  { key: "me_7600_i", label: "7600_I_ME (%)", presets: ["validation", "all"], numeric: true },
  { key: "me_7600_s", label: "7600_S_ME (%)", presets: ["validation", "all"], numeric: true },
  { key: "me_5600_u", label: "5600_U_ME (%)", presets: ["validation", "all"], numeric: true },
  { key: "checked_7600", label: "7600 Checked", presets: ["validation", "all"] },
  { key: "isomers", label: "Possible Isomers", presets: ["inventory", "all"] },
  { key: "synonyms", label: "Synonyms", presets: ["inventory", "all"] },
  { key: "inchikey", label: "InChIKey", presets: ["inventory", "all"] },
  { key: "comments", label: "Comments / Notes", presets: ["standard", "validation", "inventory", "all"] },
  { key: "submitted_by", label: "Submitted By", presets: ["standard", "all"] }
];

const INV_COLUMNS = [
  { key: "vial_status", label: "Vial Status" },
  { key: "standard_name", label: "Standard Name" },
  { key: "location", label: "Box / Location" },
  { key: "checked_out_by", label: "Possession / Checked Out" },
  { key: "linked_methods", label: "Used in Lab Method?" },
  { key: "concentration", label: "Concentration / Amount" },
  { key: "solvent", label: "Solvent" },
  { key: "manufacturer", label: "Manufacturer" },
  { key: "catalog_no", label: "Catalog # (COA)" },
  { key: "lot_no", label: "Lot / Batch #" },
  { key: "cas_no", label: "CAS #" },
  { key: "exp_date", label: "Expiration Date" },
  { key: "receive_date", label: "Receive Date" },
  { key: "initial_vol_mass", label: "Initial Vol/Mass" },
  { key: "open_date", label: "Open Date" },
  { key: "has_tag", label: "DEA / Tag" },
  { key: "notes", label: "Notes" },
  { key: "sheet", label: "Inventory Sheet" }
];

function getActiveUsers() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY_USERS);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (parsed && typeof parsed === "object") {
        return { ...DEFAULT_USERS, ...parsed };
      }
    }
  } catch (e) {}
  return { ...DEFAULT_USERS };
}

function updateActiveUserUI() {
  const pill = document.getElementById("active-user-pill");
  if (pill) {
    pill.textContent = state.currentUser ? `👤 Signed in: ${state.currentUser}` : "👤 Not Signed In";
  }
}

function handleUnlock(e) {
  e.preventDefault();
  const users = getActiveUsers();
  const selUserEl = document.getElementById("auth-user-input");
  const selectedUser = selUserEl ? selUserEl.value : "Gillian";
  const val = document.getElementById("auth-pw-input").value;

  let matchedUser = null;
  if (users[selectedUser] && val === users[selectedUser]) {
    matchedUser = selectedUser;
  } else {
    for (const [uname, upw] of Object.entries(users)) {
      if (val === upw) {
        matchedUser = uname;
        if (selUserEl) selUserEl.value = uname;
        break;
      }
    }
  }

  if (matchedUser) {
    state.currentUser = matchedUser;
    document.getElementById("auth-overlay").style.display = "none";
    document.getElementById("auth-err").textContent = "";
    document.getElementById("auth-pw-input").value = "";
    updateActiveUserUI();
    showToast(`Welcome, ${matchedUser}!`);
  } else {
    document.getElementById("auth-err").textContent = `Incorrect password for ${selectedUser}. Please try again.`;
  }
}

function lockDashboard() {
  state.currentUser = null;
  updateActiveUserUI();
  document.getElementById("auth-pw-input").value = "";
  document.getElementById("auth-err").textContent = "";
  document.getElementById("auth-overlay").style.display = "flex";
  document.getElementById("auth-pw-input").focus();
}

function reloadSavedStateFromStorage() {
  const baseCompounds = JSON.parse(JSON.stringify(EMBEDDED_STATE.compounds || []));
  const baseInventory = JSON.parse(JSON.stringify(EMBEDDED_STATE.inventory || []));
  const savedRaw = localStorage.getItem(STORAGE_KEY_DATA);

  if (savedRaw) {
    try {
      const saved = JSON.parse(savedRaw);
      const defaultMethods = EMBEDDED_STATE.methods || ["X500R (CDHR)", "7600 (CDS W93)"];
      if (Array.isArray(saved.methods) && saved.methods.length) {
        const merged = [...defaultMethods];
        saved.methods.forEach(m => {
          if (!merged.includes(m)) merged.push(m);
        });
        state.methods = merged;
      } else {
        state.methods = [...defaultMethods];
      }
      if (Array.isArray(saved.auditLog)) {
        state.auditLog = saved.auditLog;
      }
      if (Array.isArray(saved.suggestions)) {
        state.suggestions = saved.suggestions;
      }
      const mapComp = new Map();
      baseCompounds.forEach(c => mapComp.set(c.id, c));
      if (Array.isArray(saved.customCompounds)) {
        saved.customCompounds.forEach(c => mapComp.set(c.id, c));
      }
      state.compounds = Array.from(mapComp.values());

      const mapInv = new Map();
      baseInventory.forEach(v => mapInv.set(v.vial_id, v));
      if (Array.isArray(saved.customInventory)) {
        saved.customInventory.forEach(v => mapInv.set(v.vial_id, v));
      }
      state.inventory = Array.from(mapInv.values());
    } catch (err) {
      state.compounds = baseCompounds;
      state.inventory = baseInventory;
    }
  } else {
    state.compounds = baseCompounds;
    state.inventory = baseInventory;
  }
  populateInventoryDatalist();
  renderMethodTabs();
  renderValidationSubView();
  renderAuditLog();
  renderSuggestionsTray();
  renderActiveWorkspace();
}

function initState() {
  // Always require login with password when the page is loaded
  state.currentUser = null;
  document.getElementById("auth-overlay").style.display = "flex";
  const pwInput = document.getElementById("auth-pw-input");
  if (pwInput) {
    pwInput.value = "";
    setTimeout(() => pwInput.focus(), 50);
  }
  updateActiveUserUI();
  reloadSavedStateFromStorage();
  const savedPal = localStorage.getItem("lynch_lab_moma_palette_v1") || "ucsf_original";
  applyColorPalette(savedPal);
}

function persistCustomState() {
  const customCompounds = state.compounds.filter(c => c.is_custom);
  const customInventory = state.inventory.filter(v => v.is_custom);
  localStorage.setItem(STORAGE_KEY_DATA, JSON.stringify({
    methods: state.methods,
    customCompounds: customCompounds,
    customInventory: customInventory,
    auditLog: state.auditLog,
    suggestions: state.suggestions
  }));
  if (REALTIME_CHANNEL) {
    try {
      REALTIME_CHANNEL.postMessage({ type: "STATE_UPDATED", ts: Date.now() });
    } catch (e) {}
  }
}

function populateInventoryDatalist() {
  const dl = document.getElementById("inv-datalist");
  if (!dl) return;
  const curVials = state.inventory.filter(v => v.sheet === "Current Stock");
  dl.innerHTML = curVials.map(v => {
    const label = `${v.standard_name} | Box: ${v.location || 'N/A'} | ${v.concentration || ''} (${v.manufacturer || ''}) [${v.vial_id}]`;
    return `<option value="${escapeHtml(label)}"></option>`;
  }).join("");
}

function applyInventoryAutofill() {
  const val = document.getElementById("f-inv-lookup").value || "";
  const m = val.match(/\[(INV-\d+|OLD-\d+|USR-INV-\d+)\]/);
  if (!m) return;
  const vid = m[1];
  const v = state.inventory.find(x => x.vial_id === vid);
  if (!v) return;
  if (!document.getElementById("f-compound").value) {
    document.getElementById("f-compound").value = v.standard_name;
  }
  document.getElementById("f-box").value = v.location || "";
  document.getElementById("f-conc").value = v.concentration || "";
  document.getElementById("f-solvent").value = v.solvent || "";
  document.getElementById("f-mfg").value = v.manufacturer || "";
  document.getElementById("f-cat").value = v.catalog_no || "";
  document.getElementById("f-lot").value = v.lot_no || "";
  document.getElementById("f-cas").value = v.cas_no || "";
  showToast(`Auto-filled stock details from ${v.standard_name} (${v.location})`);
}

function renderMethodTabs() {
  const bar = document.getElementById("method-tabs-bar");
  const curStockCount = state.inventory.filter(v => v.sheet === "Current Stock").length;
  const methodButtons = state.methods.map(m => {
    const count = state.compounds.filter(c => c.method === m).length;
    const activeCls = (state.activeMethod === m) ? "active" : "";
    return `<button class="method-tab ${activeCls}" onclick="selectMethodTab('${m.replace(/'/g, "\\'")}')">
      <span>🔬 ${escapeHtml(m)}</span>
      <span class="tab-pill">${count}</span>
    </button>`;
  });

  if (state.methods.length > 1) {
    const activeAll = (state.activeMethod === "ALL_METHODS") ? "active" : "";
    methodButtons.push(`<button class="method-tab ${activeAll}" onclick="selectMethodTab('ALL_METHODS')">
      <span>🔬 All Lab Methods</span>
      <span class="tab-pill">${state.compounds.length}</span>
    </button>`);
  }

  const activeInv = (state.activeMethod === "STANDARDS_INVENTORY") ? "active" : "";
  methodButtons.push(`<button class="method-tab ${activeInv}" onclick="selectMethodTab('STANDARDS_INVENTORY')">
    <span>📦 Standards Inventory 2026</span>
    <span class="tab-pill">${curStockCount} Stock / ${state.inventory.length} Total</span>
  </button>`);

  bar.innerHTML = methodButtons.join("");

  const delBtn = document.getElementById("delete-method-tab-btn");
  if (delBtn) {
    const canDelete = state.activeMethod !== "X500R (CDHR)" && state.activeMethod !== "7600 (CDS W93)" && state.activeMethod !== "ALL_METHODS" && state.activeMethod !== "STANDARDS_INVENTORY";
    delBtn.style.display = canDelete ? "inline-flex" : "none";
  }

  const fMethod = document.getElementById("f-method");
  if (fMethod) {
    fMethod.innerHTML = state.methods.map(m => `<option value="${escapeHtml(m)}">${escapeHtml(m)}</option>`).join("");
    if (state.activeMethod !== "ALL_METHODS" && state.activeMethod !== "STANDARDS_INVENTORY") {
      fMethod.value = state.activeMethod;
    }
  }
}

function selectMethodTab(m) {
  state.activeMethod = m;
  renderMethodTabs();
  renderActiveWorkspace();
  previewLiveCommand();
}

function renderActiveWorkspace() {
  const isInv = (state.activeMethod === "STANDARDS_INVENTORY");
  document.getElementById("workspace-methods").style.display = isInv ? "none" : "block";
  document.getElementById("workspace-inventory").style.display = isInv ? "block" : "none";
  if (isInv) {
    renderInventoryTable();
  } else {
    renderTable();
  }
}

function switchSubView(view) {
  ["table", "validation", "audit"].forEach(v => {
    document.getElementById("subview-" + v).style.display = (v === view) ? "block" : "none";
    document.getElementById("subbtn-" + v).classList.toggle("active", v === view);
  });
}

function getUserCheckoutStyle(userName) {
  const u = String(userName || "").trim().toLowerCase();
  if (!u) return { rowCls: "", badgeCls: "", accent: "#64748b", dot: "⚪" };
  if (u.includes("gillian")) return { rowCls: "co-row-gillian", badgeCls: "co-badge-gillian", accent: "#7e22ce", dot: "🟣" };
  if (u.includes("lauren")) return { rowCls: "co-row-lauren", badgeCls: "co-badge-lauren", accent: "#0d9488", dot: "🟢" };
  if (u.includes("michael")) return { rowCls: "co-row-michael", badgeCls: "co-badge-michael", accent: "#2563eb", dot: "🔵" };
  if (u.includes("emilio")) return { rowCls: "co-row-emilio", badgeCls: "co-badge-emilio", accent: "#0284c7", dot: "🔷" };
  return { rowCls: "co-row-other", badgeCls: "co-badge-other", accent: "#d97706", dot: "🟠" };
}

function toggleCompoundCheckout(compId, isChecked, overrideUser = null) {
  const c = state.compounds.find(x => x.id === compId);
  if (!c) return;
  const activeUser = overrideUser || state.currentUser || "Gillian";
  const newHolder = isChecked ? activeUser : "";
  c.checked_out_by = newHolder;
  c.is_custom = true;

  // Sync across all method tabs sharing the same compound name or linked vials
  const linkedIds = c.linked_vial_ids || [];
  state.compounds.forEach(other => {
    if (other.id !== c.id) {
      const sameName = String(other.compound || "").toLowerCase() === String(c.compound || "").toLowerCase();
      const sharesVial = linkedIds.length > 0 && Array.isArray(other.linked_vial_ids) && other.linked_vial_ids.some(vid => linkedIds.includes(vid));
      if (sameName || sharesVial) {
        other.checked_out_by = newHolder;
        other.is_custom = true;
      }
    }
  });

  // Sync with linked physical vials in Standards Inventory
  if (linkedIds.length > 0) {
    const primaryVid = linkedIds[0];
    const v = state.inventory.find(x => x.vial_id === primaryVid);
    if (v) {
      v.checked_out_by = newHolder;
      v.is_custom = true;
    }
  }

  state.auditLog.push({
    timestamp: new Date().toISOString().replace("T", " ").slice(0, 16),
    action: newHolder ? `Checked Out (${newHolder})` : `Returned to Box (by ${activeUser})`,
    method: c.method,
    compound: c.compound,
    date_added: c.box_location || "",
    status: newHolder ? `Checked Out: ${newHolder}` : "In Box",
    submitted_by: activeUser,
    notes: newHolder ? `${activeUser} checked out ${c.compound}` : `${activeUser} returned ${c.compound}`
  });

  persistCustomState();
  renderAuditLog();
  renderTable();
  showToast(newHolder ? `Checked out "${c.compound}" to ${newHolder}` : `Returned "${c.compound}" to box`);
}

function toggleVialCheckout(vialId, isChecked) {
  const v = state.inventory.find(x => x.vial_id === vialId);
  if (!v) return;
  const activeUser = state.currentUser || "Gillian";
  const newHolder = isChecked ? activeUser : "";
  v.checked_out_by = newHolder;
  v.is_custom = true;

  // Also sync to any method compounds linked to this vial
  state.compounds.forEach(c => {
    if (Array.isArray(c.linked_vial_ids) && c.linked_vial_ids.includes(vialId)) {
      c.checked_out_by = newHolder;
      c.is_custom = true;
    }
  });

  state.auditLog.push({
    timestamp: new Date().toISOString().replace("T", " ").slice(0, 16),
    action: newHolder ? `Checked Out Vial (${newHolder})` : `Returned Vial (by ${activeUser})`,
    method: "Standards Inventory",
    compound: v.standard_name,
    date_added: v.location || "",
    status: newHolder ? `Checked Out: ${newHolder}` : "In Box",
    submitted_by: activeUser,
    notes: newHolder ? `${activeUser} checked out ${v.standard_name} [${v.location}]` : `${activeUser} returned ${v.standard_name} [${v.location}]`
  });

  persistCustomState();
  renderAuditLog();
  renderInventoryTable();
  showToast(newHolder ? `Checked out vial "${v.standard_name}" to ${newHolder}` : `Returned vial "${v.standard_name}" to box`);
}

function sortByCheckedOut() {
  if (state.sortCol === "checked_out_by") {
    state.sortAsc = !state.sortAsc;
  } else {
    state.sortCol = "checked_out_by";
    state.sortAsc = true;
  }
  renderTable();
  showToast(state.sortAsc ? "Sorted: Checked-Out compounds at top (grouped by user)" : "Sorted: In-Box compounds first");
}

function handleDatePresetChange() {
  const preset = document.getElementById("filter-date-preset").value;
  document.getElementById("date-range-wrap").style.display = (preset === "RANGE") ? "inline-flex" : "none";
  renderTable();
}

function resetFilters() {
  document.getElementById("search-input").value = "";
  document.getElementById("filter-status").value = "ALL";
  const coEl = document.getElementById("filter-checkout");
  if (coEl) coEl.value = "ALL";
  document.getElementById("filter-mix").value = "ALL";
  document.getElementById("filter-date-preset").value = "ALL";
  document.getElementById("filter-date-from").value = "";
  document.getElementById("filter-date-to").value = "";
  document.getElementById("date-range-wrap").style.display = "none";
  renderTable();
}

function getFilteredCompounds() {
  const q = (document.getElementById("search-input").value || "").trim().toLowerCase();
  const statusFilter = document.getElementById("filter-status").value;
  const coFilterEl = document.getElementById("filter-checkout");
  const coFilter = coFilterEl ? coFilterEl.value : "ALL";
  const mixFilter = document.getElementById("filter-mix").value;
  const datePreset = document.getElementById("filter-date-preset").value;
  const dtFrom = document.getElementById("filter-date-from").value;
  const dtTo = document.getElementById("filter-date-to").value;

  let list = state.compounds.filter(c => {
    if (state.activeMethod !== "ALL_METHODS" && state.activeMethod !== "STANDARDS_INVENTORY" && c.method !== state.activeMethod) return false;
    if (statusFilter === "CUSTOM") {
      if (!c.is_custom) return false;
    } else if (statusFilter === "LINKED_INV") {
      if (!c.linked_vial_ids || !c.linked_vial_ids.length) return false;
    } else if (statusFilter === "MISSING_INV") {
      if (c.linked_vial_ids && c.linked_vial_ids.length) return false;
    } else if (statusFilter !== "ALL" && c.status !== statusFilter) {
      return false;
    }

    if (coFilter === "ANY_OUT") {
      if (!c.checked_out_by) return false;
    } else if (coFilter === "IN_BOX") {
      if (c.checked_out_by) return false;
    } else if (coFilter !== "ALL") {
      if (String(c.checked_out_by || "").toLowerCase() !== coFilter.toLowerCase()) return false;
    }

    if (mixFilter !== "ALL" && String(c.primary_mix) !== mixFilter) return false;

    if (datePreset === "NONE") {
      if (c.date_added) return false;
    } else if (datePreset === "RANGE") {
      if (dtFrom && (!c.date_added || c.date_added < dtFrom)) return false;
      if (dtTo && (!c.date_added || c.date_added > dtTo)) return false;
    } else if (datePreset !== "ALL") {
      if (c.date_added !== datePreset) return false;
    }

    if (q) {
      const hay = [
        c.compound, c.formula, c.inchikey, c.synonyms, c.isomers,
        c.box_location, c.box_position, c.manufacturer, c.catalog_no, c.lot_no, c.cas_no,
        c.checked_out_by, c.comments, c.method, c.date_added, c.submitted_by
      ].join(" ").toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  });

  const colDef = COLUMN_DEFS.find(d => d.key === state.sortCol);
  const isNum = colDef && colDef.numeric;
  list.sort((a, b) => {
    let va = a[state.sortCol] ?? "";
    let vb = b[state.sortCol] ?? "";
    if (state.sortCol === "checked_out_by") {
      const aOut = Boolean(String(va).trim());
      const bOut = Boolean(String(vb).trim());
      if (aOut !== bOut) {
        return state.sortAsc ? (aOut ? -1 : 1) : (aOut ? 1 : -1);
      }
      if (aOut && bOut) {
        const cmpUser = String(va).toLowerCase().localeCompare(String(vb).toLowerCase());
        if (cmpUser !== 0) return state.sortAsc ? cmpUser : -cmpUser;
      }
      return String(a.compound || "").toLowerCase().localeCompare(String(b.compound || "").toLowerCase());
    }
    if (isNum) {
      const na = parseFloat(String(va).replace(/[^0-9.-]/g, ""));
      const nb = parseFloat(String(vb).replace(/[^0-9.-]/g, ""));
      const aValid = !isNaN(na);
      const bValid = !isNaN(nb);
      if (aValid && bValid) return state.sortAsc ? (na - nb) : (nb - na);
      if (aValid) return -1;
      if (bValid) return 1;
    }
    va = String(va).toLowerCase();
    vb = String(vb).toLowerCase();
    if (va < vb) return state.sortAsc ? -1 : 1;
    if (va > vb) return state.sortAsc ? 1 : -1;
    return 0;
  });

  return list;
}

function toggleSort(colKey) {
  if (state.sortCol === colKey) {
    state.sortAsc = !state.sortAsc;
  } else {
    state.sortCol = colKey;
    state.sortAsc = true;
  }
  renderTable();
}

function renderMethodKPIs(methodSubset, filteredList) {
  const is7600 = (state.activeMethod === "7600 (CDS W93)");
  const total = methodSubset.length;
  const activeCount = methodSubset.filter(c => c.status === "Active").length;
  const pendingCount = methodSubset.filter(c => c.status === "Inactive / Pending").length;
  const removedCount = methodSubset.filter(c => c.status === "Removed").length;
  const invLinkedCount = methodSubset.filter(c => (c.linked_vial_ids && c.linked_vial_ids.length) || c.box_location).length;

  let card4Html = "";
  if (is7600) {
    const inCidCount = methodSubset.filter(c => String(c.in_library || "").toLowerCase() === "yes").length;
    const uSwathCount = methodSubset.filter(c => c.lod_7600_u_swath && !c.lod_7600_u_swath.toLowerCase().includes("not")).length;
    const uIdaCount = methodSubset.filter(c => c.lod_7600_u_ida && !c.lod_7600_u_ida.toLowerCase().includes("not")).length;
    card4Html = `
      <div class="kpi-card">
        <div class="kpi-label">In 7600_CID Library & LODs</div>
        <div class="kpi-val" style="color:#007CBE;">${inCidCount} / ${total}</div>
        <div class="kpi-sub">${uSwathCount} Urine SWATH · ${uIdaCount} Urine IDA LODs</div>
      </div>`;
  } else {
    const uLodCount = methodSubset.filter(c => c.lod_x500r_u_consensus && !c.lod_x500r_u_consensus.toLowerCase().includes("not")).length;
    const sLodCount = methodSubset.filter(c => c.lod_x500r_s_consensus && !c.lod_x500r_s_consensus.toLowerCase().includes("not")).length;
    card4Html = `
      <div class="kpi-card">
        <div class="kpi-label">Consensus LOD Coverage</div>
        <div class="kpi-val" style="color:#007CBE;">${uLodCount} U / ${sLodCount} S</div>
        <div class="kpi-sub">Validated Urine & Serum IDA LODs (ng/mL)</div>
      </div>`;
  }

  document.getElementById("kpi-grid").innerHTML = `
    <div class="kpi-card">
      <div class="kpi-label">Total Method Compounds</div>
      <div class="kpi-val">${total}</div>
      <div class="kpi-sub">Showing ${filteredList.length} after active filters</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">${is7600 ? "Active in CDS W93 Method" : "Active in Method"}</div>
      <div class="kpi-val" style="color:#166534;">${activeCount}</div>
      <div class="kpi-sub">${total ? ((100 * activeCount / total).toFixed(1)) : 0}% active (${pendingCount} pending, ${removedCount} removed)</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">Linked Standards Inventory</div>
      <div class="kpi-val" style="color:#0f766e;">${invLinkedCount} / ${total}</div>
      <div class="kpi-sub">Cross-linked to ${state.inventory.filter(v => v.sheet === 'Current Stock').length} Current Stock vials</div>
    </div>
    ${card4Html}
    <div class="kpi-card">
      <div class="kpi-label">Lab Activity & Audit Log</div>
      <div class="kpi-val" style="color:#6b21a8;">${state.auditLog.length}</div>
      <div class="kpi-sub">User compound submissions & vial updates</div>
    </div>
  `;
}

function getVisibleColumns(forceAll = false) {
  const preset = forceAll ? "all" : document.getElementById("col-preset").value;
  const is7600 = (state.activeMethod === "7600 (CDS W93)");
  const isX500R = (state.activeMethod === "X500R (CDHR)");

  return COLUMN_DEFS.filter(c => {
    if (is7600) {
      // Remove Mix columns and X500R-only RT columns on the 7600 tab (just one column for current RT)
      if (["primary_mix", "secondary_mix", "rt_x500r_jul26", "rt_x500r_dec25", "checked_7600"].includes(c.key)) {
        return false;
      }
      if (preset === "standard") {
        if (["lod_x500r_u_consensus", "lod_x500r_s_consensus"].includes(c.key)) return false;
        if (["date_first_w93", "lod_7600_u_swath", "lod_7600_u_ida", "lod_7600_s_ida"].includes(c.key)) return true;
      } else if (preset === "validation" || preset === "inventory") {
        if (c.key === "date_first_w93") return true;
      }
    } else if (isX500R) {
      if (c.key === "date_first_w93" && !forceAll) return false;
    }
    return c.presets.includes(preset);
  }).map(c => {
    if (is7600) {
      if (c.key === "status") return { ...c, label: "Active in CDS W93 method" };
      if (c.key === "in_library") return { ...c, label: "In 7600_CID library?" };
      if (c.key === "rt_7600_jul26") return { ...c, label: "Current RT (min)" };
    }
    return c;
  });
}

function renderTable() {
  const is7600 = (state.activeMethod === "7600 (CDS W93)");
  const mixWrap = document.getElementById("filter-mix-wrap");
  if (mixWrap) {
    mixWrap.style.display = is7600 ? "none" : "inline-flex";
    if (is7600 && document.getElementById("filter-mix").value !== "ALL") {
      document.getElementById("filter-mix").value = "ALL";
    }
  }
  const statusLbl = document.getElementById("filter-status-label");
  if (statusLbl) {
    statusLbl.textContent = is7600 ? "W93 Status:" : "Status:";
  }

  const methodSubset = (state.activeMethod === "ALL_METHODS")
    ? state.compounds
    : state.compounds.filter(c => c.method === state.activeMethod);
  const filtered = getFilteredCompounds();
  renderMethodKPIs(methodSubset, filtered);

  const cols = getVisibleColumns(false);
  const thead = document.getElementById("comp-thead");
  thead.innerHTML = `<tr>
    <th>Actions / Vials</th>
    ${cols.map(c => {
      const arrow = (state.sortCol === c.key) ? (state.sortAsc ? " ▲" : " ▼") : "";
      return `<th onclick="toggleSort('${c.key}')">${escapeHtml(c.label)}${arrow}</th>`;
    }).join("")}
  </tr>`;

  const tbody = document.getElementById("comp-tbody");
  tbody.innerHTML = filtered.map(row => {
    const coStyle = getUserCheckoutStyle(row.checked_out_by);
    const flashCls = (state.lastFlashedId === row.id) ? "flash-updated" : "";
    const rowCls = [row.is_custom ? "custom-row" : "", coStyle.rowCls, flashCls].filter(Boolean).join(" ");
    const nVials = (row.linked_vial_ids || []).length;
    const vialBtn = nVials > 0
      ? `<button class="action-link" style="color:#0f766e;border-color:#99f6e4;background:#f0fdfa;" onclick="openLinkedVialsModal('${row.id}')">📦 Vials (${nVials})</button>`
      : "";
    const nPeaks = (row.msms_peaks || []).length;
    const msmsBtn = nPeaks > 0
      ? `<button class="action-link" style="color:#6d28d9;border-color:#ddd6fe;background:#f5f3ff;" onclick="openMsmsModal('${row.id}')">📈 MS/MS (${nPeaks})</button>`
      : "";
    return `<tr class="${rowCls}" id="row-${row.id}">
      <td style="white-space:nowrap;">
        <button class="action-link" onclick="openCompoundModal('${row.id}')">✏️ Edit</button>
        ${vialBtn}
        ${msmsBtn}
      </td>
      ${cols.map(c => {
        const val = row[c.key] ?? "";
        if (c.key === "status") {
          let bCls = "badge-pending";
          if (val === "Active") bCls = "badge-active";
          else if (val === "Removed") bCls = "badge-removed";
          const customTag = row.is_custom ? ` <span class="badge badge-custom">Updated</span>` : "";
          return `<td><span class="badge ${bCls}">${escapeHtml(val)}</span>${customTag}</td>`;
        }
        if (c.key === "compound") {
          return `<td style="font-weight:700;color:var(--ucsf-navy);white-space:nowrap;">${escapeHtml(val)}</td>`;
        }
        if (c.key === "in_library" && (is7600 || row.method === "7600 (CDS W93)")) {
          const isYes = String(val).toLowerCase() !== "no";
          const selStyle = isYes
            ? "background:#dcfce7;color:#166534;border-color:#86efac;"
            : "background:#fee2e2;color:#991b1b;border-color:#fca5a5;";
          return `<td>
            <select class="inline-select" style="${selStyle}" onchange="updateCompoundFieldInline('${row.id}', 'in_library', this.value)" title="In 7600_CID library? (Yes/No)">
              <option value="Yes" ${isYes ? "selected" : ""}>Yes</option>
              <option value="No" ${!isYes ? "selected" : ""}>No</option>
            </select>
          </td>`;
        }
        if (c.key === "date_first_w93") {
          return `<td>
            <input type="date" class="inline-date mono" value="${escapeHtml(val)}" onchange="updateCompoundFieldInline('${row.id}', 'date_first_w93', this.value)" title="Set Date of first W93 analysis">
          </td>`;
        }
        if (c.key === "date_last_rt_update") {
          return `<td>
            <input type="date" class="inline-date mono" value="${escapeHtml(val)}" onchange="updateCompoundFieldInline('${row.id}', 'date_last_rt_update', this.value)" title="Set Date of last RT update">
          </td>`;
        }
        if (c.key === "primary_mix" && val) {
          return `<td><span class="badge badge-mix">Mix ${escapeHtml(val)}</span></td>`;
        }
        if (c.key === "checked_out_by") {
          const isChecked = Boolean(String(val).trim());
          const activeUser = state.currentUser || "Gillian";
          const tip = isChecked
            ? `Checked out by ${val}. Click checkbox to return to box.`
            : `Click checkbox to check out to ${activeUser}.`;
          const badgeHtml = isChecked
            ? `<span class="badge ${coStyle.badgeCls}">${coStyle.dot} ${escapeHtml(val)}</span>`
            : `<span style="font-size:11.5px;color:#94a3b8;">Check out</span>`;
          return `<td>
            <label class="checkout-wrap" title="${escapeHtml(tip)}">
              <input type="checkbox" class="checkout-cb" style="accent-color:${coStyle.accent};" ${isChecked ? "checked" : ""} onchange="toggleCompoundCheckout('${row.id}', this.checked)">
              ${badgeHtml}
            </label>
          </td>`;
        }
        if (c.key === "vial_status" && val) {
          const vsCls = (val === "UNOPENED") ? "badge-blue" : ((val === "OPEN") ? "badge-active" : "badge-pending");
          return `<td><span class="badge ${vsCls}">${escapeHtml(val)}</span></td>`;
        }
        if (c.key === "formula" || c.key === "monoisotopic_mass" || c.key === "mh_plus_mz" || c.key === "inchikey" || c.key === "date_added" || c.key === "cas_no") {
          return `<td class="mono" style="white-space:nowrap;">${escapeHtml(val)}</td>`;
        }
        return `<td>${escapeHtml(val)}</td>`;
      }).join("")}
    </tr>`;
  }).join("");

  document.getElementById("showing-summary").textContent =
    `Displaying ${filtered.length} of ${methodSubset.length} compounds in ${state.activeMethod === "ALL_METHODS" ? "All Lab Methods" : state.activeMethod}`;

  const legendBar = document.getElementById("checkout-legend-bar");
  if (legendBar) {
    const counts = { Gillian: 0, Lauren: 0, Michael: 0, Emilio: 0 };
    methodSubset.forEach(c => {
      const h = String(c.checked_out_by || "").trim();
      if (h) counts[h] = (counts[h] || 0) + 1;
    });
    legendBar.innerHTML = `<span style="color:var(--text-muted);font-weight:700;">Checkout Colors:</span>` +
      Object.entries(counts).map(([uName, cnt]) => {
        const st = getUserCheckoutStyle(uName);
        return `<span class="badge ${st.badgeCls}" style="cursor:pointer;" onclick="document.getElementById('filter-checkout').value='${uName}';renderTable();" title="Click to filter compounds checked out by ${uName}">${st.dot} ${escapeHtml(uName)} (${cnt})</span>`;
      }).join("") +
      `<span class="badge" style="background:#f1f5f9;color:#475569;border:1px solid #cbd5e1;cursor:pointer;" onclick="document.getElementById('filter-checkout').value='ALL';renderTable();" title="Show all compounds">Show All</span>`;
  }
}

/* ==================== STANDARDS INVENTORY 2026 VIEW ==================== */

function resetInventoryFilters() {
  document.getElementById("inv-search-input").value = "";
  document.getElementById("inv-filter-sheet").value = "Current Stock";
  document.getElementById("inv-filter-status").value = "ALL";
  document.getElementById("inv-filter-mfg").value = "ALL";
  renderInventoryTable();
}

function getFilteredInventory() {
  const q = (document.getElementById("inv-search-input").value || "").trim().toLowerCase();
  const sheetFilter = document.getElementById("inv-filter-sheet").value;
  const statusFilter = document.getElementById("inv-filter-status").value;
  const mfgFilter = document.getElementById("inv-filter-mfg").value;

  let list = state.inventory.filter(v => {
    if (sheetFilter !== "ALL" && v.sheet !== sheetFilter) return false;
    if (statusFilter === "OPEN" && v.vial_status !== "OPEN") return false;
    if (statusFilter === "UNOPENED" && v.vial_status !== "UNOPENED") return false;
    if (statusFilter === "CHECKED_OUT" && !v.checked_out_by) return false;
    if (statusFilter === "IN_METHOD" && (!v.linked_methods || !v.linked_methods.length)) return false;
    if (statusFilter === "NOT_IN_METHOD" && v.linked_methods && v.linked_methods.length) return false;
    if (statusFilter === "TAGGED" && (!v.has_tag || v.has_tag.toUpperCase() === "N")) return false;
    if (statusFilter === "EXPIRED" && !v.is_expired) return false;

    if (mfgFilter !== "ALL") {
      const mLow = (v.manufacturer || "").toLowerCase();
      if (mfgFilter === "OTHER") {
        if (["cerilliant", "cayman", "chemfaces", "sigma", "tci"].some(k => mLow.includes(k))) return false;
      } else if (!mLow.includes(mfgFilter.toLowerCase())) {
        return false;
      }
    }

    if (q) {
      const hay = [
        v.standard_name, v.location, v.manufacturer, v.catalog_no, v.lot_no,
        v.cas_no, v.concentration, v.solvent, v.checked_out_by, v.notes,
        (v.linked_methods || []).join(" ")
      ].join(" ").toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  });

  list.sort((a, b) => {
    let va = String(a[state.invSortCol] ?? "");
    let vb = String(b[state.invSortCol] ?? "");
    if (state.invSortCol === "checked_out_by") {
      const aOut = Boolean(va.trim());
      const bOut = Boolean(vb.trim());
      if (aOut !== bOut) {
        return state.invSortAsc ? (aOut ? -1 : 1) : (aOut ? 1 : -1);
      }
      if (aOut && bOut) {
        const cmpUser = va.toLowerCase().localeCompare(vb.toLowerCase());
        if (cmpUser !== 0) return state.invSortAsc ? cmpUser : -cmpUser;
      }
      return String(a.standard_name || "").toLowerCase().localeCompare(String(b.standard_name || "").toLowerCase());
    }
    va = va.toLowerCase();
    vb = vb.toLowerCase();
    if (va < vb) return state.invSortAsc ? -1 : 1;
    if (va > vb) return state.invSortAsc ? 1 : -1;
    return 0;
  });
  return list;
}

function toggleInvSort(colKey) {
  if (state.invSortCol === colKey) {
    state.invSortAsc = !state.invSortAsc;
  } else {
    state.invSortCol = colKey;
    state.invSortAsc = true;
  }
  renderInventoryTable();
}

function renderInventoryTable() {
  const curStock = state.inventory.filter(v => v.sheet === "Current Stock");
  const oldOrders = state.inventory.filter(v => v.sheet === "Old Orders");
  const openCount = curStock.filter(v => v.vial_status === "OPEN").length;
  const unopenCount = curStock.filter(v => v.vial_status === "UNOPENED").length;
  const inMethodCount = curStock.filter(v => v.linked_methods && v.linked_methods.length).length;
  const checkedOutList = curStock.filter(v => v.checked_out_by);
  const taggedCount = curStock.filter(v => v.has_tag && v.has_tag.toUpperCase() !== "N").length;

  const filtered = getFilteredInventory();

  document.getElementById("kpi-grid").innerHTML = `
    <div class="kpi-card">
      <div class="kpi-label">Current Stock Vials (2026)</div>
      <div class="kpi-val">${curStock.length}</div>
      <div class="kpi-sub">${openCount} Open · ${unopenCount} Unopened (${oldOrders.length} in Old Orders)</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">Used in Active MS Methods</div>
      <div class="kpi-val" style="color:#007CBE;">${inMethodCount}</div>
      <div class="kpi-sub">Vials cross-linked to X500R (CDHR) & lab methods</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">Checked Out (In Possession)</div>
      <div class="kpi-val" style="color:#6b21a8;">${checkedOutList.length}</div>
      <div class="kpi-sub">${checkedOutList.map(v => `${v.standard_name} (${v.checked_out_by})`).slice(0, 3).join(", ") || "All vials in boxes"}</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">DEA / Tagged Standards</div>
      <div class="kpi-val" style="color:#92400e;">${taggedCount}</div>
      <div class="kpi-sub">Additional DEA paperwork / notification required</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">Filtered Vials Shown</div>
      <div class="kpi-val" style="color:#166534;">${filtered.length}</div>
      <div class="kpi-sub">Click "+ Method" on any vial to add to an MS method</div>
    </div>
  `;

  const thead = document.getElementById("inv-thead");
  thead.innerHTML = `<tr>
    <th>Actions</th>
    ${INV_COLUMNS.map(c => {
      const arrow = (state.invSortCol === c.key) ? (state.invSortAsc ? " ▲" : " ▼") : "";
      return `<th onclick="toggleInvSort('${c.key}')">${escapeHtml(c.label)}${arrow}</th>`;
    }).join("")}
  </tr>`;

  const tbody = document.getElementById("inv-tbody");
  tbody.innerHTML = filtered.map(v => {
    const coStyle = getUserCheckoutStyle(v.checked_out_by);
    const rowCls = [v.is_custom ? "custom-row" : "", coStyle.rowCls].filter(Boolean).join(" ");
    return `<tr class="${rowCls}">
      <td style="white-space:nowrap;">
        <button class="action-link" onclick="openVialModal('${v.vial_id}')">✏️ Edit / Checkout</button>
        <button class="action-link" style="color:#0369a1;border-color:#bae6fd;background:#f0f9ff;" onclick="promoteVialToMethod('${v.vial_id}')" title="Add this standard to X500R or another lab MS method">+ Method</button>
      </td>
      ${INV_COLUMNS.map(c => {
        const val = v[c.key] ?? "";
        if (c.key === "vial_status") {
          let bCls = "badge-pending";
          if (val === "OPEN") bCls = "badge-active";
          else if (val === "UNOPENED") bCls = "badge-blue";
          else if (val === "EMPTY") bCls = "badge-removed";
          return `<td><span class="badge ${bCls}">${escapeHtml(val)}</span></td>`;
        }
        if (c.key === "standard_name") {
          return `<td style="font-weight:700;color:var(--ucsf-navy);white-space:nowrap;">${escapeHtml(val)}</td>`;
        }
        if (c.key === "location") {
          return `<td class="mono" style="font-weight:700;color:#0f766e;white-space:nowrap;">${escapeHtml(val)}</td>`;
        }
        if (c.key === "checked_out_by") {
          const isChecked = Boolean(String(val).trim());
          const activeUser = state.currentUser || "Gillian";
          const tip = isChecked
            ? `Checked out by ${val}. Click checkbox to return to box.`
            : `Click checkbox to check out to ${activeUser}.`;
          const badgeHtml = isChecked
            ? `<span class="badge ${coStyle.badgeCls}">${coStyle.dot} ${escapeHtml(val)}</span>`
            : `<span style="color:#94a3b8;font-size:11.5px;">In Box</span>`;
          return `<td>
            <label class="checkout-wrap" title="${escapeHtml(tip)}">
              <input type="checkbox" class="checkout-cb" style="accent-color:${coStyle.accent};" ${isChecked ? "checked" : ""} onchange="toggleVialCheckout('${v.vial_id}', this.checked)">
              ${badgeHtml}
            </label>
          </td>`;
        }
        if (c.key === "linked_methods") {
          const mList = v.linked_methods || [];
          if (!mList.length) return `<td style="color:#94a3b8;font-size:11.5px;">—</td>`;
          return `<td>${mList.map(m => `<span class="badge badge-mix" style="margin-right:3px;">🔬 ${escapeHtml(m)}</span>`).join("")}</td>`;
        }
        if (c.key === "exp_date") {
          if (!val) return `<td></td>`;
          const expTag = v.is_expired ? ` <span class="badge badge-removed" style="font-size:10px;">Past Exp</span>` : "";
          return `<td class="mono" style="white-space:nowrap;">${escapeHtml(val)}${expTag}</td>`;
        }
        if (c.key === "has_tag" && val && val.toUpperCase() !== "N") {
          return `<td><span class="badge badge-pending">DEA / Tag (${escapeHtml(val)})</span></td>`;
        }
        if (c.key === "catalog_no" || c.key === "lot_no" || c.key === "cas_no" || c.key === "receive_date" || c.key === "open_date") {
          return `<td class="mono" style="white-space:nowrap;">${escapeHtml(val)}</td>`;
        }
        return `<td>${escapeHtml(val)}</td>`;
      }).join("")}
    </tr>`;
  }).join("");

  document.getElementById("inv-showing-summary").textContent =
    `Displaying ${filtered.length} of ${state.inventory.length} standard vials`;
}

function openLinkedVialsModal(compId) {
  const c = state.compounds.find(x => x.id === compId);
  if (!c) return;
  const vials = (c.linked_vial_ids || []).map(vid => state.inventory.find(v => v.vial_id === vid)).filter(Boolean);
  document.getElementById("linked-vials-title").textContent =
    `Standards Inventory Vials Linked to "${c.compound}" (${vials.length} found)`;
  document.getElementById("linked-vials-body").innerHTML = `
    <table class="comp-table">
      <thead>
        <tr>
          <th>Sheet</th>
          <th>Vial Status</th>
          <th>Standard Name</th>
          <th>Box / Location</th>
          <th>Checked Out By</th>
          <th>Conc.</th>
          <th>Solvent</th>
          <th>Manufacturer</th>
          <th>Catalog #</th>
          <th>Lot #</th>
          <th>Exp. Date</th>
          <th>Action</th>
        </tr>
      </thead>
      <tbody>
        ${vials.map(v => `
          <tr>
            <td>${escapeHtml(v.sheet)}</td>
            <td><span class="badge ${v.vial_status === 'UNOPENED' ? 'badge-blue' : 'badge-active'}">${escapeHtml(v.vial_status)}</span></td>
            <td style="font-weight:700;color:var(--ucsf-navy);">${escapeHtml(v.standard_name)}</td>
            <td class="mono" style="font-weight:700;color:#0f766e;">${escapeHtml(v.location)}</td>
            <td>${v.checked_out_by ? `<span class="badge badge-custom">👋 ${escapeHtml(v.checked_out_by)}</span>` : "In Box"}</td>
            <td>${escapeHtml(v.concentration)}</td>
            <td>${escapeHtml(v.solvent)}</td>
            <td>${escapeHtml(v.manufacturer)}</td>
            <td class="mono">${escapeHtml(v.catalog_no)}</td>
            <td class="mono">${escapeHtml(v.lot_no)}</td>
            <td class="mono">${escapeHtml(v.exp_date)}</td>
            <td><button class="action-link" onclick="document.getElementById('linked-vials-modal').classList.remove('open');openVialModal('${v.vial_id}')">✏️ Edit Vial</button></td>
          </tr>
        `).join("")}
      </tbody>
    </table>
  `;
  document.getElementById("linked-vials-modal").classList.add("open");
}

function openVialModal(vialId = null) {
  const modal = document.getElementById("vial-modal");
  const today = new Date().toISOString().slice(0, 10);
  if (vialId) {
    const v = state.inventory.find(x => x.vial_id === vialId);
    if (!v) return;
    document.getElementById("vial-modal-title").textContent = `Edit Standard Vial: ${v.standard_name} (${v.location})`;
    document.getElementById("edit-vial-id").value = v.vial_id;
    document.getElementById("v-name").value = v.standard_name || "";
    document.getElementById("v-sheet").value = v.sheet || "Current Stock";
    document.getElementById("v-status").value = v.vial_status || "OPEN";
    document.getElementById("v-loc").value = v.location || "";
    document.getElementById("v-checked-out").value = v.checked_out_by || "";
    document.getElementById("v-conc").value = v.concentration || "";
    document.getElementById("v-solvent").value = v.solvent || "";
    document.getElementById("v-mfg").value = v.manufacturer || "";
    document.getElementById("v-cat").value = v.catalog_no || "";
    document.getElementById("v-lot").value = v.lot_no || "";
    document.getElementById("v-cas").value = v.cas_no || "";
    document.getElementById("v-exp").value = /^\d{4}-\d{2}-\d{2}$/.test(v.exp_date || "") ? v.exp_date : "";
    document.getElementById("v-rec").value = /^\d{4}-\d{2}-\d{2}$/.test(v.receive_date || "") ? v.receive_date : "";
    document.getElementById("v-open").value = /^\d{4}-\d{2}-\d{2}$/.test(v.open_date || "") ? v.open_date : "";
    document.getElementById("v-vol").value = v.initial_vol_mass || "";
    document.getElementById("v-tag").value = v.has_tag || "";
    document.getElementById("v-user").value = state.currentUser || "";
    document.getElementById("v-notes").value = v.notes || "";
  } else {
    document.getElementById("vial-modal-title").textContent = "Add New Chemical Standard Vial to Inventory";
    document.getElementById("edit-vial-id").value = "";
    document.getElementById("v-name").value = "";
    document.getElementById("v-sheet").value = "Current Stock";
    document.getElementById("v-status").value = "UNOPENED";
    document.getElementById("v-loc").value = "";
    document.getElementById("v-checked-out").value = "";
    document.getElementById("v-conc").value = "1.0 mg/mL";
    document.getElementById("v-solvent").value = "Methanol";
    document.getElementById("v-mfg").value = "Cerilliant";
    document.getElementById("v-cat").value = "";
    document.getElementById("v-lot").value = "";
    document.getElementById("v-cas").value = "";
    document.getElementById("v-exp").value = "";
    document.getElementById("v-rec").value = today;
    document.getElementById("v-open").value = "";
    document.getElementById("v-vol").value = "1000";
    document.getElementById("v-tag").value = "";
    document.getElementById("v-user").value = state.currentUser || "";
    document.getElementById("v-notes").value = "";
  }
  modal.classList.add("open");
}

function closeVialModal() {
  document.getElementById("vial-modal").classList.remove("open");
}

function saveVialForm(e) {
  e.preventDefault();
  const editId = document.getElementById("edit-vial-id").value;
  const nowStr = new Date().toISOString().replace("T", " ").slice(0, 16);
  const userInit = document.getElementById("v-user").value.trim();
  const sname = document.getElementById("v-name").value.trim();
  const loc = document.getElementById("v-loc").value.trim();
  const statusVal = document.getElementById("v-status").value;
  const checkedOut = document.getElementById("v-checked-out").value;

  if (editId) {
    const idx = state.inventory.findIndex(v => v.vial_id === editId);
    if (idx !== -1) {
      state.inventory[idx] = {
        ...state.inventory[idx],
        standard_name: sname,
        sheet: document.getElementById("v-sheet").value,
        vial_status: statusVal,
        location: loc,
        checked_out_by: checkedOut,
        concentration: document.getElementById("v-conc").value.trim(),
        solvent: document.getElementById("v-solvent").value.trim(),
        manufacturer: document.getElementById("v-mfg").value.trim(),
        catalog_no: document.getElementById("v-cat").value.trim(),
        lot_no: document.getElementById("v-lot").value.trim(),
        cas_no: document.getElementById("v-cas").value.trim(),
        exp_date: document.getElementById("v-exp").value,
        receive_date: document.getElementById("v-rec").value,
        open_date: document.getElementById("v-open").value,
        initial_vol_mass: document.getElementById("v-vol").value.trim(),
        has_tag: document.getElementById("v-tag").value,
        notes: document.getElementById("v-notes").value.trim(),
        is_custom: true
      };
      state.auditLog.push({
        timestamp: nowStr,
        action: checkedOut ? `Checked Out Vial (${checkedOut})` : "Updated Inventory Vial",
        method: "Standards Inventory",
        compound: sname,
        date_added: loc,
        status: statusVal,
        submitted_by: userInit,
        notes: state.inventory[idx].notes
      });
      showToast(`Updated standard vial "${sname}" (${loc})`);
    }
  } else {
    const newVid = "USR-INV-" + Date.now();
    const newVial = {
      vial_id: newVid,
      sheet: document.getElementById("v-sheet").value,
      standard_name: sname,
      vial_status: statusVal,
      checked_out_by: checkedOut,
      location: loc,
      concentration: document.getElementById("v-conc").value.trim(),
      solvent: document.getElementById("v-solvent").value.trim(),
      manufacturer: document.getElementById("v-mfg").value.trim(),
      catalog_no: document.getElementById("v-cat").value.trim(),
      lot_no: document.getElementById("v-lot").value.trim(),
      cas_no: document.getElementById("v-cas").value.trim(),
      exp_date: document.getElementById("v-exp").value,
      is_expired: false,
      receive_date: document.getElementById("v-rec").value,
      initial_vol_mass: document.getElementById("v-vol").value.trim(),
      open_date: document.getElementById("v-open").value,
      has_tag: document.getElementById("v-tag").value,
      notes: document.getElementById("v-notes").value.trim(),
      linked_methods: [],
      linked_compound_ids: [],
      is_custom: true
    };
    state.inventory.unshift(newVial);
    state.auditLog.push({
      timestamp: nowStr,
      action: "Added New Inventory Vial",
      method: "Standards Inventory",
      compound: sname,
      date_added: loc,
      status: statusVal,
      submitted_by: userInit,
      notes: newVial.notes
    });
    showToast(`Added "${sname}" (${loc}) to Standards Inventory`);
  }

  persistCustomState();
  populateInventoryDatalist();
  closeVialModal();
  renderMethodTabs();
  renderAuditLog();
  renderActiveWorkspace();
}

function promoteVialToMethod(vialId) {
  const v = state.inventory.find(x => x.vial_id === vialId);
  if (!v) return;
  openCompoundModal(null);
  document.getElementById("f-compound").value = v.standard_name;
  document.getElementById("f-box").value = v.location || "";
  document.getElementById("f-conc").value = v.concentration || "";
  document.getElementById("f-solvent").value = v.solvent || "";
  document.getElementById("f-mfg").value = v.manufacturer || "";
  document.getElementById("f-cat").value = v.catalog_no || "";
  document.getElementById("f-lot").value = v.lot_no || "";
  document.getElementById("f-cas").value = v.cas_no || "";
  document.getElementById("f-comments").value = v.notes || "";
}

function renderValidationSubView() {
  const mix8Tbody = document.getElementById("mix8-tbody");
  mix8Tbody.innerHTML = (EMBEDDED_STATE.mix8_recipe || []).map(r => `
    <tr>
      <td style="font-weight:700;color:var(--ucsf-navy);">${escapeHtml(r.compound)}</td>
      <td class="mono">${escapeHtml(r.stock_ug_ml)}</td>
      <td class="mono">${escapeHtml(r.final_ug_ml)}</td>
      <td class="mono" style="font-weight:700;">${escapeHtml(r.vol_ul_per_ml)} µL</td>
      <td>${escapeHtml(r.validated_phase1)}</td>
    </tr>
  `).join("");

  const outRows = [
    ...(EMBEDDED_STATE.urine_outstanding || []).map(r => ({ matrix: "Urine", ...r })),
    ...(EMBEDDED_STATE.serum_outstanding || []).map(r => ({ matrix: "Serum", lod_note: "TBD", ...r }))
  ];
  document.getElementById("outstanding-tbody").innerHTML = outRows.map(r => `
    <tr>
      <td><span class="badge ${r.matrix === 'Urine' ? 'badge-pending' : 'badge-mix'}">${escapeHtml(r.matrix)}</span></td>
      <td style="font-weight:700;color:var(--ucsf-navy);">${escapeHtml(r.compound)}</td>
      <td class="mono">${escapeHtml(r.lod_note || "")}</td>
      <td>${escapeHtml(r.comment || "")}</td>
    </tr>
  `).join("");
}

function renderAuditLog() {
  document.getElementById("audit-count-badge").textContent = state.auditLog.length;
  const tbody = document.getElementById("audit-tbody");
  if (!state.auditLog.length) {
    tbody.innerHTML = `<tr><td colspan="8" style="text-align:center;color:var(--text-muted);padding:20px;">No user compound submissions or inventory updates recorded yet.</td></tr>`;
    return;
  }
  tbody.innerHTML = state.auditLog.slice().reverse().map(item => `
    <tr>
      <td class="mono" style="white-space:nowrap;">${escapeHtml(item.timestamp)}</td>
      <td><span class="badge badge-custom">${escapeHtml(item.action)}</span></td>
      <td>${escapeHtml(item.method)}</td>
      <td style="font-weight:700;color:var(--ucsf-navy);">${escapeHtml(item.compound)}</td>
      <td class="mono">${escapeHtml(item.date_added)}</td>
      <td>${escapeHtml(item.status)}</td>
      <td style="font-weight:700;">${escapeHtml(item.submitted_by)}</td>
      <td>${escapeHtml(item.notes || "")}</td>
    </tr>
  `).join("");
}

function autoCalcMassFromFormula() {
  const formula = (document.getElementById("f-formula").value || "").trim();
  if (!formula) return;
  const regex = /([A-Z][a-z]?)(\d*)/g;
  let match;
  let totalMass = 0;
  let matchedLen = 0;
  while ((match = regex.exec(formula)) !== null) {
    const el = match[1];
    const count = match[2] ? parseInt(match[2], 10) : 1;
    if (!(el in ATOMIC_MASSES)) return;
    totalMass += ATOMIC_MASSES[el] * count;
    matchedLen += match[0].length;
  }
  if (matchedLen === formula.length && totalMass > 0) {
    document.getElementById("f-mass").value = totalMass.toFixed(4);
    document.getElementById("f-mh").value = (totalMass + 1.007276).toFixed(4);
  }
}

function autoCalcMHFromMass() {
  const m = parseFloat(document.getElementById("f-mass").value);
  if (!isNaN(m) && m > 0) {
    document.getElementById("f-mh").value = (m + 1.007276).toFixed(4);
  }
}

function syncCompoundModalForMethod() {
  const m = document.getElementById("f-method").value;
  const is7600 = (m === "7600 (CDS W93)");
  const lblStatus = document.getElementById("label-f-status");
  if (lblStatus) lblStatus.textContent = is7600 ? "Active in CDS W93 method *" : "Method Status *";

  const lblLib = document.getElementById("label-f-in-library");
  if (lblLib) lblLib.textContent = is7600 ? "In 7600_CID library? *" : "Added to Instrument Library?";

  const optNeedsReview = document.getElementById("opt-lib-needs-review");
  if (optNeedsReview) {
    optNeedsReview.style.display = is7600 ? "none" : "block";
    if (is7600 && document.getElementById("f-in-library").value === "Needs Review") {
      document.getElementById("f-in-library").value = "Yes";
    }
  }

  const w93Wrap = document.getElementById("field-wrap-w93-date");
  if (w93Wrap) w93Wrap.style.display = is7600 ? "flex" : "none";

  const rtX500Wrap = document.getElementById("field-wrap-rt-x500r");
  if (rtX500Wrap) rtX500Wrap.style.display = is7600 ? "none" : "flex";

  const rtDecWrap = document.getElementById("field-wrap-rt-dec25");
  if (rtDecWrap) rtDecWrap.style.display = is7600 ? "none" : "flex";

  const lblRt7600 = document.getElementById("label-f-rt-7600");
  if (lblRt7600) lblRt7600.textContent = is7600 ? "Current RT (min)" : "7600 RT (min)";

  const mix1Wrap = document.getElementById("field-wrap-mix1");
  if (mix1Wrap) mix1Wrap.style.display = is7600 ? "none" : "flex";

  const mix2Wrap = document.getElementById("field-wrap-mix2");
  if (mix2Wrap) mix2Wrap.style.display = is7600 ? "none" : "flex";

  const chk7600Wrap = document.getElementById("field-wrap-checked-7600");
  if (chk7600Wrap) chk7600Wrap.style.display = is7600 ? "none" : "flex";
}

function openCompoundModal(compId = null) {
  renderMethodTabs();
  const modal = document.getElementById("compound-modal");
  const today = new Date().toISOString().slice(0, 10);
  document.getElementById("f-inv-lookup").value = "";

  if (compId) {
    const c = state.compounds.find(x => x.id === compId);
    if (!c) return;
    document.getElementById("compound-modal-title").textContent = `Edit Method Compound: ${c.compound}`;
    document.getElementById("edit-comp-id").value = c.id;
    document.getElementById("f-method").value = c.method;
    document.getElementById("f-compound").value = c.compound || "";
    document.getElementById("f-date-added").value = c.date_added || today;
    document.getElementById("f-date-first-w93").value = c.date_first_w93 || "";
    document.getElementById("f-date-last-rt").value = c.date_last_rt_update || "";
    document.getElementById("f-status").value = c.status || "Active";
    document.getElementById("f-in-library").value = c.in_library || "Yes";
    document.getElementById("f-submitted-by").value = state.currentUser || ((c.submitted_by && !c.submitted_by.startsWith("Initial")) ? c.submitted_by : "");
    document.getElementById("f-formula").value = c.formula || "";
    document.getElementById("f-mass").value = c.monoisotopic_mass || "";
    document.getElementById("f-mh").value = c.mh_plus_mz || "";
    document.getElementById("f-rt-x500r").value = c.rt_x500r_jul26 || "";
    document.getElementById("f-rt-dec25").value = c.rt_x500r_dec25 || "";
    document.getElementById("f-rt-7600").value = c.rt_7600_jul26 || "";
    document.getElementById("f-mix1").value = c.primary_mix || "";
    document.getElementById("f-mix2").value = c.secondary_mix || "";
    document.getElementById("f-checked-7600").value = c.checked_7600 || "";
    document.getElementById("f-box").value = c.box_location || "";
    document.getElementById("f-pos").value = c.box_position || "";
    document.getElementById("f-conc").value = c.stock_conc || "";
    document.getElementById("f-solvent").value = c.stock_solvent || "";
    document.getElementById("f-mfg").value = c.manufacturer || "";
    document.getElementById("f-cat").value = c.catalog_no || "";
    document.getElementById("f-lot").value = c.lot_no || "";
    document.getElementById("f-cas").value = c.cas_no || "";
    document.getElementById("f-lod-u").value = c.lod_x500r_u_consensus || "";
    document.getElementById("f-lod-s").value = c.lod_x500r_s_consensus || "";
    document.getElementById("f-inchikey").value = c.inchikey || "";
    document.getElementById("f-isomers").value = c.isomers || "";
    document.getElementById("f-synonyms").value = c.synonyms || "";
    document.getElementById("f-comments").value = c.comments || "";
  } else {
    document.getElementById("compound-modal-title").textContent = "Submit New Compound to Lab Method";
    document.getElementById("edit-comp-id").value = "";
    document.getElementById("f-method").value = (state.activeMethod === "ALL_METHODS" || state.activeMethod === "STANDARDS_INVENTORY") ? state.methods[0] : state.activeMethod;
    document.getElementById("f-compound").value = "";
    document.getElementById("f-date-added").value = today;
    document.getElementById("f-date-first-w93").value = "";
    document.getElementById("f-date-last-rt").value = today;
    document.getElementById("f-status").value = "Active";
    document.getElementById("f-in-library").value = "Yes";
    document.getElementById("f-submitted-by").value = state.currentUser || "";
    document.getElementById("f-formula").value = "";
    document.getElementById("f-mass").value = "";
    document.getElementById("f-mh").value = "";
    document.getElementById("f-rt-x500r").value = "";
    document.getElementById("f-rt-dec25").value = "";
    document.getElementById("f-rt-7600").value = "";
    document.getElementById("f-mix1").value = "";
    document.getElementById("f-mix2").value = "";
    document.getElementById("f-checked-7600").value = "";
    document.getElementById("f-box").value = "";
    document.getElementById("f-pos").value = "";
    document.getElementById("f-conc").value = "1.0 mg/mL";
    document.getElementById("f-solvent").value = "Methanol";
    document.getElementById("f-mfg").value = "";
    document.getElementById("f-cat").value = "";
    document.getElementById("f-lot").value = "";
    document.getElementById("f-cas").value = "";
    document.getElementById("f-lod-u").value = "";
    document.getElementById("f-lod-s").value = "";
    document.getElementById("f-inchikey").value = "";
    document.getElementById("f-isomers").value = "";
    document.getElementById("f-synonyms").value = "";
    document.getElementById("f-comments").value = "";
  }
  syncCompoundModalForMethod();
  modal.classList.add("open");
}

function closeCompoundModal() {
  document.getElementById("compound-modal").classList.remove("open");
}

function saveCompoundForm(e) {
  e.preventDefault();
  const editId = document.getElementById("edit-comp-id").value;
  const nowStr = new Date().toISOString().replace("T", " ").slice(0, 16);
  const methodVal = document.getElementById("f-method").value;
  const compName = document.getElementById("f-compound").value.trim();
  const dateAdded = document.getElementById("f-date-added").value;
  const dateFirstW93 = document.getElementById("f-date-first-w93").value;
  let dateLastRt = document.getElementById("f-date-last-rt").value;
  const statusVal = document.getElementById("f-status").value;
  const subBy = document.getElementById("f-submitted-by").value.trim();

  if (editId) {
    const idx = state.compounds.findIndex(x => x.id === editId);
    if (idx !== -1) {
      const prev = state.compounds[idx];
      const newRtX500 = document.getElementById("f-rt-x500r").value.trim();
      const newRt7600 = document.getElementById("f-rt-7600").value.trim();
      if ((newRtX500 !== (prev.rt_x500r_jul26 || "") || newRt7600 !== (prev.rt_7600_jul26 || "")) && dateLastRt === (prev.date_last_rt_update || "")) {
        dateLastRt = nowStr.slice(0, 10);
      }
      const updated = {
        ...prev,
        method: methodVal,
        compound: compName,
        date_added: dateAdded,
        date_first_w93: dateFirstW93,
        date_last_rt_update: dateLastRt,
        status: statusVal,
        in_library: document.getElementById("f-in-library").value,
        submitted_by: subBy,
        formula: document.getElementById("f-formula").value.trim(),
        monoisotopic_mass: document.getElementById("f-mass").value.trim(),
        mh_plus_mz: document.getElementById("f-mh").value.trim(),
        rt_x500r_jul26: document.getElementById("f-rt-x500r").value.trim(),
        rt_x500r_dec25: document.getElementById("f-rt-dec25").value.trim(),
        rt_7600_jul26: document.getElementById("f-rt-7600").value.trim(),
        primary_mix: document.getElementById("f-mix1").value.trim(),
        secondary_mix: document.getElementById("f-mix2").value.trim(),
        checked_7600: document.getElementById("f-checked-7600").value.trim(),
        box_location: document.getElementById("f-box").value.trim(),
        box_position: document.getElementById("f-pos").value.trim(),
        stock_conc: document.getElementById("f-conc").value.trim(),
        stock_solvent: document.getElementById("f-solvent").value.trim(),
        manufacturer: document.getElementById("f-mfg").value.trim(),
        catalog_no: document.getElementById("f-cat").value.trim(),
        lot_no: document.getElementById("f-lot").value.trim(),
        cas_no: document.getElementById("f-cas").value.trim(),
        lod_x500r_u_consensus: document.getElementById("f-lod-u").value.trim(),
        lod_x500r_s_consensus: document.getElementById("f-lod-s").value.trim(),
        inchikey: document.getElementById("f-inchikey").value.trim(),
        isomers: document.getElementById("f-isomers").value.trim(),
        synonyms: document.getElementById("f-synonyms").value.trim(),
        comments: document.getElementById("f-comments").value.trim(),
        updated_at: nowStr,
        is_custom: true
      };
      state.compounds[idx] = updated;
      state.auditLog.push({
        timestamp: nowStr,
        action: "Edited Method Compound",
        method: methodVal,
        compound: compName,
        date_added: dateAdded,
        status: statusVal,
        submitted_by: subBy,
        notes: updated.comments
      });
      showToast(`Updated "${compName}" in ${methodVal}`);
    }
  } else {
    const newId = "USR-" + Date.now();
    const newEntry = {
      id: newId,
      method: methodVal,
      compound: compName,
      status: statusVal,
      in_library: document.getElementById("f-in-library").value,
      date_added: dateAdded,
      date_first_w93: dateFirstW93,
      date_last_rt_update: dateLastRt || dateAdded,
      checked_7600: document.getElementById("f-checked-7600").value.trim(),
      rt_x500r_jul26: document.getElementById("f-rt-x500r").value.trim(),
      rt_x500r_dec25: document.getElementById("f-rt-dec25").value.trim(),
      rt_7600_jul26: document.getElementById("f-rt-7600").value.trim(),
      primary_mix: document.getElementById("f-mix1").value.trim(),
      secondary_mix: document.getElementById("f-mix2").value.trim(),
      monoisotopic_mass: document.getElementById("f-mass").value.trim(),
      mh_plus_mz: document.getElementById("f-mh").value.trim(),
      formula: document.getElementById("f-formula").value.trim(),
      isomers: document.getElementById("f-isomers").value.trim(),
      synonyms: document.getElementById("f-synonyms").value.trim(),
      box_location: document.getElementById("f-box").value.trim(),
      box_position: document.getElementById("f-pos").value.trim(),
      stock_conc: document.getElementById("f-conc").value.trim(),
      stock_solvent: document.getElementById("f-solvent").value.trim(),
      manufacturer: document.getElementById("f-mfg").value.trim(),
      catalog_no: document.getElementById("f-cat").value.trim(),
      lot_no: document.getElementById("f-lot").value.trim(),
      cas_no: document.getElementById("f-cas").value.trim(),
      vial_status: "OPEN",
      checked_out_by: "",
      linked_vial_ids: [],
      inchikey: document.getElementById("f-inchikey").value.trim(),
      lod_7600_u_swath: "",
      lod_7600_u_ida: "",
      lod_7600_s_ida: "",
      lod_5600_u_ida: "",
      lod_brown_u: "",
      lod_x500r_u_p1: "",
      lod_x500r_u_p2: "",
      x500r_u_agree: "",
      lod_x500r_u_consensus: document.getElementById("f-lod-u").value.trim(),
      lod_brown_s: "",
      lod_x500r_s_p1: "",
      lod_x500r_s_p2: "",
      x500r_s_agree: "",
      lod_x500r_s_consensus: document.getElementById("f-lod-s").value.trim(),
      me_7600_i: "",
      me_7600_s: "",
      me_5600_u: "",
      include_flag: "1",
      comments: document.getElementById("f-comments").value.trim(),
      submitted_by: subBy,
      updated_at: nowStr,
      is_custom: true
    };
    state.compounds.unshift(newEntry);
    state.auditLog.push({
      timestamp: nowStr,
      action: "Submitted New Compound",
      method: methodVal,
      compound: compName,
      date_added: dateAdded,
      status: statusVal,
      submitted_by: subBy,
      notes: newEntry.comments
    });
    if (state.activeMethod === "STANDARDS_INVENTORY") {
      state.activeMethod = methodVal;
    }
    showToast(`Submitted "${compName}" (${dateAdded}) to ${methodVal}`);
  }

  persistCustomState();
  closeCompoundModal();
  renderMethodTabs();
  renderAuditLog();
  renderActiveWorkspace();
}

// =========================================================================
// INLINE FIELD UPDATES & REAL-TIME NATURAL-LANGUAGE CHANGE / SUGGESTION ENGINE
// =========================================================================

function updateCompoundFieldInline(compId, fieldKey, newVal) {
  const c = state.compounds.find(x => x.id === compId);
  if (!c) return;
  const activeUser = state.currentUser || "Gillian";
  const nowStr = new Date().toISOString().replace("T", " ").slice(0, 16);
  const oldVal = c[fieldKey] || "";
  if (String(oldVal) === String(newVal)) return;

  c[fieldKey] = newVal;
  c.updated_at = nowStr;
  c.submitted_by = activeUser;
  c.is_custom = true;
  state.lastFlashedId = c.id;

  const friendlyField = (fieldKey === "in_library")
    ? (c.method === "7600 (CDS W93)" ? "In 7600_CID library?" : "In Library?")
    : (fieldKey === "date_first_w93"
        ? "Date of first W93 analysis"
        : (fieldKey === "date_last_rt_update" ? "Date of last RT update" : fieldKey));

  state.auditLog.push({
    timestamp: nowStr,
    action: `Updated ${friendlyField}`,
    method: c.method,
    compound: c.compound,
    date_added: newVal || "—",
    status: c.status,
    submitted_by: activeUser,
    notes: `${friendlyField}: "${oldVal || '—'}" → "${newVal || '—'}"`
  });

  persistCustomState();
  renderAuditLog();
  renderTable();
  showToast(`Updated ${c.compound}: ${friendlyField} = ${newVal || "—"}`);
}

function toggleLiveHelp() {
  const el = document.getElementById("live-cmd-help");
  if (!el) return;
  el.style.display = (el.style.display === "none" || !el.style.display) ? "block" : "none";
}

function fillLiveCmdExample(txt) {
  const inp = document.getElementById("live-cmd-input");
  if (!inp) return;
  inp.value = txt;
  inp.focus();
  previewLiveCommand();
}

function _calcMassFromFormulaStr(formula) {
  const f = (formula || "").trim();
  if (!f) return null;
  const regex = /([A-Z][a-z]?)(\d*)/g;
  let match, totalMass = 0, matchedLen = 0;
  while ((match = regex.exec(f)) !== null) {
    const el = match[1];
    const count = match[2] ? parseInt(match[2], 10) : 1;
    if (!(el in ATOMIC_MASSES)) return null;
    totalMass += ATOMIC_MASSES[el] * count;
    matchedLen += match[0].length;
  }
  if (matchedLen === f.length && totalMass > 0) {
    return {
      mass: totalMass.toFixed(4),
      mh: (totalMass + 1.007276).toFixed(4)
    };
  }
  return null;
}

function _normalizeDateInput(raw) {
  const s = (raw || "").trim();
  if (!s) return "";
  if (s.toLowerCase() === "today" || s.toLowerCase() === "now") {
    return new Date().toISOString().slice(0, 10);
  }
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m) {
    const mm = m[1].padStart(2, "0");
    const dd = m[2].padStart(2, "0");
    const yyyy = m[3].length === 2 ? ("20" + m[3]) : m[3];
    return `${yyyy}-${mm}-${dd}`;
  }
  return s;
}

function parseLiveCommand(rawText) {
  const text = (rawText || "").trim();
  if (!text) return null;

  const targetMethod = (state.activeMethod === "ALL_METHODS" || state.activeMethod === "STANDARDS_INVENTORY")
    ? "X500R (CDHR)"
    : state.activeMethod;
  const is7600 = (targetMethod === "7600 (CDS W93)");

  // 1. Check if adding a new compound
  const addMatch = text.match(/^(?:add|create|new|submit)\s+(?:compound\s+)?([^,;]+)(.*)$/i);
  if (addMatch) {
    const compName = addMatch[1].replace(/^(?:named|called)\s+/i, "").trim();
    const rest = addMatch[2] || "";
    const fMatch = rest.match(/formula\s*[:=]?\s*([A-Za-z0-9]+)/i);
    const rtMatch = rest.match(/\brt\s*[:=to]*\s*([0-9]+(?:\.[0-9]+)?)/i);
    const boxMatch = rest.match(/\bbox\s*[:=to]*\s*([A-Za-z0-9\-_]+)/i);
    const w93Match = rest.match(/w93(?:\s+date|\s+analysis)?\s*[:=to]*\s*([0-9\/\-]+|today)/i);
    const formula = fMatch ? fMatch[1] : "";
    const rt = rtMatch ? rtMatch[1] : "";
    const box = boxMatch ? boxMatch[1] : "";
    const w93Date = w93Match ? _normalizeDateInput(w93Match[1]) : "";
    return {
      kind: "ADD_COMPOUND",
      method: targetMethod,
      compound: compName,
      formula,
      rt,
      box,
      w93Date,
      summary: `Add new compound "${compName}" to ${targetMethod}` +
        (formula ? ` (Formula: ${formula})` : "") +
        (rt ? ` (RT: ${rt} min)` : "") +
        (box ? ` (Box: ${box})` : "")
    };
  }

  // 2. Match an existing compound in the active method tab (or across all methods)
  const subset = state.compounds.filter(c => c.method === targetMethod);
  const pool = subset.length ? subset : state.compounds;
  const lowerText = text.toLowerCase();

  let bestComp = null;
  let bestLen = 0;
  for (const c of pool) {
    const cname = (c.compound || "").trim();
    if (!cname) continue;
    const clow = cname.toLowerCase();
    if (lowerText.includes(clow) && clow.length > bestLen) {
      bestComp = c;
      bestLen = clow.length;
    }
    if (c.synonyms) {
      for (const syn of String(c.synonyms).split(/[,;/]/)) {
        const slow = syn.trim().toLowerCase();
        if (slow.length >= 4 && lowerText.includes(slow) && slow.length > bestLen) {
          bestComp = c;
          bestLen = slow.length;
        }
      }
    }
  }

  if (!bestComp) {
    return {
      kind: "GENERAL_SUGGESTION",
      method: targetMethod,
      rawText: text,
      summary: `Log general method suggestion on ${targetMethod}: "${text}"`
    };
  }

  // Remove the matched compound name from the text to inspect the requested field & value
  const remainder = text.replace(new RegExp(bestComp.compound.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "i"), " ").trim();

  // Check-out / Return patterns
  if (/\b(return|check\s*in|uncheck|in\s+box)\b/i.test(text)) {
    return {
      kind: "CHECKOUT",
      compId: bestComp.id,
      compound: bestComp.compound,
      method: bestComp.method,
      isChecked: false,
      holder: "",
      summary: `[${bestComp.compound}] Return to Box (clear check-out)`
    };
  }
  if (/\b(check\s*out|checkout|assign\s+to)\b/i.test(text)) {
    let holder = state.currentUser || "Gillian";
    if (/\bgillian\b/i.test(remainder)) holder = "Gillian";
    else if (/\blauren\b/i.test(remainder)) holder = "Lauren";
    else if (/\bmichael\b/i.test(remainder)) holder = "Michael";
    else if (/\bemilio\b/i.test(remainder)) holder = "Emilio";
    return {
      kind: "CHECKOUT",
      compId: bestComp.id,
      compound: bestComp.compound,
      method: bestComp.method,
      isChecked: true,
      holder,
      summary: `[${bestComp.compound}] Check out to ${holder}`
    };
  }

  // In Library / In 7600_CID library?
  if (/\b(library|7600_cid|cid\s+library|in_library)\b/i.test(remainder)) {
    const isNo = /\b(no|false|0|missing|remove)\b/i.test(remainder);
    const newVal = isNo ? "No" : "Yes";
    const colLabel = (bestComp.method === "7600 (CDS W93)") ? "In 7600_CID library?" : "In Library?";
    return {
      kind: "UPDATE_FIELD",
      compId: bestComp.id,
      compound: bestComp.compound,
      method: bestComp.method,
      fieldKey: "in_library",
      fieldLabel: colLabel,
      newVal,
      summary: `[${bestComp.compound}] Set ${colLabel} → "${newVal}"`
    };
  }

  // Date of last RT update
  if (/\b(last\s*rt\s*update|rt\s*update\s*date|rt\s*date)\b/i.test(remainder)) {
    const dtToken = remainder.match(/(\d{4}-\d{2}-\d{2}|\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}|today|now)/i);
    const newVal = dtToken ? _normalizeDateInput(dtToken[1]) : new Date().toISOString().slice(0, 10);
    return {
      kind: "UPDATE_FIELD",
      compId: bestComp.id,
      compound: bestComp.compound,
      method: bestComp.method,
      fieldKey: "date_last_rt_update",
      fieldLabel: "Date of last RT update",
      newVal,
      summary: `[${bestComp.compound}] Set Date of last RT update → "${newVal}"`
    };
  }

  // Date of first W93 analysis
  if (/\b(w93\s*analysis|first\s*w93|w93\s*date|first\s*analysis)\b/i.test(remainder)) {
    const dtToken = remainder.match(/(\d{4}-\d{2}-\d{2}|\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}|today|now)/i);
    const newVal = dtToken ? _normalizeDateInput(dtToken[1]) : new Date().toISOString().slice(0, 10);
    return {
      kind: "UPDATE_FIELD",
      compId: bestComp.id,
      compound: bestComp.compound,
      method: bestComp.method,
      fieldKey: "date_first_w93",
      fieldLabel: "Date of first W93 analysis",
      newVal,
      summary: `[${bestComp.compound}] Set Date of first W93 analysis → "${newVal}"`
    };
  }

  // Date Added
  if (/\b(date\s*added|added\s*date|\bdate\b)\b/i.test(remainder)) {
    const dtToken = remainder.match(/(\d{4}-\d{2}-\d{2}|\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}|today|now)/i);
    const newVal = dtToken ? _normalizeDateInput(dtToken[1]) : new Date().toISOString().slice(0, 10);
    return {
      kind: "UPDATE_FIELD",
      compId: bestComp.id,
      compound: bestComp.compound,
      method: bestComp.method,
      fieldKey: "date_added",
      fieldLabel: "Date Added",
      newVal,
      summary: `[${bestComp.compound}] Set Date Added → "${newVal}"`
    };
  }

  // Status / Active in CDS W93 method
  if (/\b(status|active\s+in\s+cds|w93\s+method|active|inactive|pending|removed|remove|deactivate|activate)\b/i.test(remainder)) {
    let newVal = "Active";
    if (/\b(removed|remove)\b/i.test(remainder)) newVal = "Removed";
    else if (/\b(inactive|pending|no|0|false|deactivate)\b/i.test(remainder)) newVal = "Inactive / Pending";
    const colLabel = (bestComp.method === "7600 (CDS W93)") ? "Active in CDS W93 method" : "Status";
    return {
      kind: "UPDATE_FIELD",
      compId: bestComp.id,
      compound: bestComp.compound,
      method: bestComp.method,
      fieldKey: "status",
      fieldLabel: colLabel,
      newVal,
      summary: `[${bestComp.compound}] Set ${colLabel} → "${newVal}"`
    };
  }

  // Retention Time (RT / Current RT)
  if (/\b(rt|retention\s*time|current\s*rt)\b/i.test(remainder)) {
    const numMatch = remainder.replace(/7600|500r|x500r|dec\s*25/gi, "").match(/([0-9]+(?:\.[0-9]+)?)/);
    if (numMatch) {
      const use7600 = (bestComp.method === "7600 (CDS W93)") || /\b7600\b/i.test(remainder);
      const fieldKey = use7600 ? "rt_7600_jul26" : "rt_x500r_jul26";
      const fieldLabel = (bestComp.method === "7600 (CDS W93)") ? "Current RT (min)" : (use7600 ? "7600 RT (min)" : "X500R RT (min)");
      const newVal = parseFloat(numMatch[1]).toFixed(2);
      return {
        kind: "UPDATE_FIELD",
        compId: bestComp.id,
        compound: bestComp.compound,
        method: bestComp.method,
        fieldKey,
        fieldLabel,
        newVal,
        summary: `[${bestComp.compound}] Set ${fieldLabel} → "${newVal}"`
      };
    }
  }

  // Box Location
  if (/\b(box|location|freezer)\b/i.test(remainder)) {
    const m = remainder.match(/(?:box|location|freezer|to|=|:)\s*([A-Za-z0-9\-_\/ ]+)$/i);
    const newVal = m ? m[1].trim() : remainder.replace(/^(set|change|update|box|location|to|=|:|\s)+/gi, "").trim();
    if (newVal) {
      return {
        kind: "UPDATE_FIELD",
        compId: bestComp.id,
        compound: bestComp.compound,
        method: bestComp.method,
        fieldKey: "box_location",
        fieldLabel: "Box Location",
        newVal,
        summary: `[${bestComp.compound}] Set Box Location → "${newVal}"`
      };
    }
  }

  // Formula
  if (/\bformula\b/i.test(remainder)) {
    const m = remainder.match(/formula\s*(?:to|=|:)?\s*([A-Za-z0-9]+)/i);
    if (m) {
      const newVal = m[1].trim();
      return {
        kind: "UPDATE_FIELD",
        compId: bestComp.id,
        compound: bestComp.compound,
        method: bestComp.method,
        fieldKey: "formula",
        fieldLabel: "Formula",
        newVal,
        summary: `[${bestComp.compound}] Set Formula → "${newVal}" (and auto-recalculate Exact Mass & [M+H]+)`
      };
    }
  }

  // Urine or Serum LOD
  if (/\b(lod|swath|ida)\b/i.test(remainder)) {
    const numMatch = remainder.replace(/7600|500r|x500r|5600/gi, "").match(/([0-9]+(?:\.[0-9]+)?)/);
    if (numMatch) {
      const newVal = numMatch[1];
      let fieldKey = is7600 ? "lod_7600_u_ida" : "lod_x500r_u_consensus";
      let fieldLabel = is7600 ? "7600 Urine IDA LOD" : "X500R Urine LOD (ng/mL)";
      if (/\bswath\b/i.test(remainder)) {
        fieldKey = "lod_7600_u_swath";
        fieldLabel = "7600 Urine SWATH LOD";
      } else if (/\bserum\b/i.test(remainder)) {
        fieldKey = is7600 ? "lod_7600_s_ida" : "lod_x500r_s_consensus";
        fieldLabel = is7600 ? "7600 Serum IDA LOD" : "X500R Serum LOD (ng/mL)";
      }
      return {
        kind: "UPDATE_FIELD",
        compId: bestComp.id,
        compound: bestComp.compound,
        method: bestComp.method,
        fieldKey,
        fieldLabel,
        newVal,
        summary: `[${bestComp.compound}] Set ${fieldLabel} → "${newVal} ng/mL"`
      };
    }
  }

  // Default for matched compound: append/update Comments / Notes
  const cleanNote = remainder.replace(/^(?:comment|note|set|to|=|:|-)+\s*/i, "").trim() || text;
  return {
    kind: "UPDATE_FIELD",
    compId: bestComp.id,
    compound: bestComp.compound,
    method: bestComp.method,
    fieldKey: "comments",
    fieldLabel: "Comments / Notes",
    newVal: bestComp.comments ? `${bestComp.comments} | ${cleanNote}` : cleanNote,
    summary: `[${bestComp.compound}] Append Note → "${cleanNote}"`
  };
}

function previewLiveCommand() {
  const inp = document.getElementById("live-cmd-input");
  const prev = document.getElementById("live-cmd-preview");
  if (!inp || !prev) return;
  const parsed = parseLiveCommand(inp.value);
  if (!parsed) {
    prev.style.display = "none";
    prev.textContent = "";
    return;
  }
  prev.style.display = "block";
  prev.innerHTML = `🟢 <strong>Live Action Preview (${escapeHtml(parsed.method || state.activeMethod)}):</strong> ${escapeHtml(parsed.summary)} <span style="color:#64748b;font-weight:400;">(Press Enter or click ⚡ Apply Live Now)</span>`;
}

function executeLiveCommand(asSuggestion = false) {
  const inp = document.getElementById("live-cmd-input");
  if (!inp) return;
  const parsed = parseLiveCommand(inp.value);
  if (!parsed) {
    showToast("Type a compound change or suggestion in the input box first.");
    inp.focus();
    return;
  }

  const activeUser = state.currentUser || "Gillian";
  const nowStr = new Date().toISOString().replace("T", " ").slice(0, 16);

  if (asSuggestion) {
    const sug = {
      id: "SUG-" + Date.now(),
      user: activeUser,
      timestamp: nowStr,
      parsed,
      rawText: inp.value.trim()
    };
    state.suggestions.push(sug);
    state.auditLog.push({
      timestamp: nowStr,
      action: "Posted Real-Time Suggestion",
      method: parsed.method || state.activeMethod,
      compound: parsed.compound || "(General)",
      date_added: nowStr.slice(0, 10),
      status: "Pending Review",
      submitted_by: activeUser,
      notes: parsed.summary
    });
    inp.value = "";
    previewLiveCommand();
    persistCustomState();
    renderSuggestionsTray();
    renderAuditLog();
    showToast(`Posted suggestion for review: ${parsed.summary}`);
    return;
  }

  applyParsedLiveCommand(parsed, activeUser);
  inp.value = "";
  previewLiveCommand();
}

function applyParsedLiveCommand(parsed, actingUser) {
  const nowStr = new Date().toISOString().replace("T", " ").slice(0, 16);
  const today = nowStr.slice(0, 10);

  if (parsed.kind === "ADD_COMPOUND") {
    const massInfo = _calcMassFromFormulaStr(parsed.formula);
    const is7600 = (parsed.method === "7600 (CDS W93)");
    const newId = "USR-" + Date.now();
    const newEntry = {
      id: newId,
      method: parsed.method,
      compound: parsed.compound,
      status: "Active",
      in_library: "Yes",
      date_added: today,
      date_first_w93: parsed.w93Date || (is7600 ? today : ""),
      date_last_rt_update: parsed.rt ? today : "",
      checked_7600: "yes",
      rt_x500r_jul26: is7600 ? "" : (parsed.rt || ""),
      rt_x500r_dec25: "",
      rt_7600_jul26: is7600 ? (parsed.rt || "") : "",
      primary_mix: "",
      secondary_mix: "",
      monoisotopic_mass: massInfo ? massInfo.mass : "",
      mh_plus_mz: massInfo ? massInfo.mh : "",
      formula: parsed.formula || "",
      isomers: "",
      synonyms: "",
      box_location: parsed.box || "",
      box_position: "",
      stock_conc: "1.0 mg/mL",
      stock_solvent: "Methanol",
      manufacturer: "",
      catalog_no: "",
      lot_no: "",
      cas_no: "",
      vial_status: "OPEN",
      checked_out_by: "",
      linked_vial_ids: [],
      inchikey: "",
      lod_7600_u_swath: "",
      lod_7600_u_ida: "",
      lod_7600_s_ida: "",
      lod_5600_u_ida: "",
      lod_brown_u: "",
      lod_x500r_u_p1: "",
      lod_x500r_u_p2: "",
      x500r_u_agree: "",
      lod_x500r_u_consensus: "",
      lod_brown_s: "",
      lod_x500r_s_p1: "",
      lod_x500r_s_p2: "",
      x500r_s_agree: "",
      lod_x500r_s_consensus: "",
      me_7600_i: "",
      me_7600_s: "",
      me_5600_u: "",
      include_flag: "1",
      comments: "Added via Real-Time Command Box",
      submitted_by: actingUser,
      updated_at: nowStr,
      is_custom: true
    };
    state.compounds.unshift(newEntry);
    state.lastFlashedId = newId;
    state.auditLog.push({
      timestamp: nowStr,
      action: "Live Added Compound",
      method: parsed.method,
      compound: parsed.compound,
      date_added: today,
      status: "Active",
      submitted_by: actingUser,
      notes: parsed.summary
    });
    persistCustomState();
    renderMethodTabs();
    renderAuditLog();
    renderTable();
    showToast(`⚡ Live Added: ${parsed.compound} (${parsed.method})`);
    return;
  }

  if (parsed.kind === "CHECKOUT") {
    toggleCompoundCheckout(parsed.compId, parsed.isChecked, parsed.holder || actingUser);
    return;
  }

  if (parsed.kind === "UPDATE_FIELD") {
    const c = state.compounds.find(x => x.id === parsed.compId);
    if (!c) return;
    const oldVal = c[parsed.fieldKey] || "";
    c[parsed.fieldKey] = parsed.newVal;
    if (parsed.fieldKey === "rt_7600_jul26" || parsed.fieldKey === "rt_x500r_jul26") {
      c.date_last_rt_update = today;
    }
    if (parsed.fieldKey === "formula") {
      const mInfo = _calcMassFromFormulaStr(parsed.newVal);
      if (mInfo) {
        c.monoisotopic_mass = mInfo.mass;
        c.mh_plus_mz = mInfo.mh;
      }
    }
    c.updated_at = nowStr;
    c.submitted_by = actingUser;
    c.is_custom = true;
    state.lastFlashedId = c.id;

    state.auditLog.push({
      timestamp: nowStr,
      action: `Live Updated ${parsed.fieldLabel}`,
      method: c.method,
      compound: c.compound,
      date_added: parsed.newVal || "—",
      status: c.status,
      submitted_by: actingUser,
      notes: `${parsed.fieldLabel}: "${oldVal || '—'}" → "${parsed.newVal || '—'}"`
    });
    persistCustomState();
    renderAuditLog();
    renderTable();
    const rowEl = document.getElementById("row-" + c.id);
    if (rowEl) {
      rowEl.scrollIntoView({ behavior: "smooth", block: "center" });
    }
    showToast(`⚡ Live Applied: ${parsed.summary}`);
    return;
  }

  if (parsed.kind === "GENERAL_SUGGESTION") {
    state.auditLog.push({
      timestamp: nowStr,
      action: "Logged Method Suggestion",
      method: parsed.method,
      compound: "(General)",
      date_added: today,
      status: "Logged",
      submitted_by: actingUser,
      notes: parsed.rawText
    });
    persistCustomState();
    renderAuditLog();
    renderTable();
    showToast(`⚡ Logged suggestion in Audit Log: "${parsed.rawText}"`);
  }
}

function renderSuggestionsTray() {
  const tray = document.getElementById("live-suggestions-tray");
  if (!tray) return;
  if (!state.suggestions || !state.suggestions.length) {
    tray.style.display = "none";
    tray.innerHTML = "";
    return;
  }
  tray.style.display = "flex";
  tray.innerHTML = `<span style="font-size:11px;font-weight:800;color:#4c1d95;">💡 Pending Lab Suggestions (${state.suggestions.length}):</span>` +
    state.suggestions.map(s => `
      <span class="suggestion-chip">
        <span><strong>${escapeHtml(s.user)}:</strong> ${escapeHtml(s.parsed.summary)}</span>
        <button type="button" class="action-link" style="background:#10b981;color:#fff;border-color:#059669;padding:2px 7px;" onclick="applySuggestionById('${s.id}')">✅ Implement</button>
        <button type="button" class="action-link" style="background:#fff;color:#dc2626;border-color:#fca5a5;padding:2px 6px;" onclick="dismissSuggestionById('${s.id}')">✕</button>
      </span>
    `).join("");
}

function applySuggestionById(sugId) {
  const idx = state.suggestions.findIndex(s => s.id === sugId);
  if (idx === -1) return;
  const sug = state.suggestions[idx];
  state.suggestions.splice(idx, 1);
  applyParsedLiveCommand(sug.parsed, state.currentUser || sug.user || "Gillian");
  persistCustomState();
  renderSuggestionsTray();
}

function dismissSuggestionById(sugId) {
  const idx = state.suggestions.findIndex(s => s.id === sugId);
  if (idx === -1) return;
  state.suggestions.splice(idx, 1);
  persistCustomState();
  renderSuggestionsTray();
  showToast("Suggestion dismissed.");
}

function buildCSVString(forceAllCols = false) {
  const rows = getFilteredCompounds();
  const cols = getVisibleColumns(forceAllCols);
  const esc = v => {
    const s = String(v ?? "");
    if (s.includes(",") || s.includes('"') || s.includes("\n")) {
      return '"' + s.replace(/"/g, '""') + '"';
    }
    return s;
  };
  const header = cols.map(c => esc(c.label)).join(",");
  const lines = rows.map(r => cols.map(c => esc(r[c.key] ?? "")).join(","));
  return [header, ...lines].join("\r\n");
}

function buildInventoryCSVString() {
  const rows = getFilteredInventory();
  const esc = v => {
    const s = Array.isArray(v) ? v.join("; ") : String(v ?? "");
    if (s.includes(",") || s.includes('"') || s.includes("\n")) {
      return '"' + s.replace(/"/g, '""') + '"';
    }
    return s;
  };
  const header = INV_COLUMNS.map(c => esc(c.label)).join(",");
  const lines = rows.map(r => INV_COLUMNS.map(c => esc(r[c.key] ?? "")).join(","));
  return [header, ...lines].join("\r\n");
}

function handleMethodExportMenu(selectEl) {
  const action = selectEl.value;
  selectEl.value = "";
  if (!action) return;
  if (action === "copy_filtered") {
    copyFilteredToCSV(false);
  } else if (action === "copy_all") {
    copyFilteredToCSV(true);
  } else if (action === "download_filtered") {
    downloadFilteredCSV(false);
  } else if (action === "download_full") {
    downloadFilteredCSV(true);
  }
}

function copyActiveTabToCSV() {
  if (state.activeMethod === "STANDARDS_INVENTORY") {
    copyInventoryToCSV();
  } else {
    copyFilteredToCSV(false);
  }
}

function downloadActiveTabCSV() {
  if (state.activeMethod === "STANDARDS_INVENTORY") {
    downloadInventoryCSV();
  } else {
    downloadFilteredCSV(false);
  }
}

function copyFilteredToCSV(forceAllCols = false) {
  const csv = buildCSVString(forceAllCols);
  const rowsCount = getFilteredCompounds().length;
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(csv).then(() => {
      showToast(`Copied ${rowsCount} method compounds (${forceAllCols ? 'All Columns' : 'Current View'}) to clipboard as CSV!`);
    }).catch(() => fallbackCopy(csv, rowsCount));
  } else {
    fallbackCopy(csv, rowsCount);
  }
}

function copyInventoryToCSV() {
  const csv = buildInventoryCSVString();
  const rowsCount = getFilteredInventory().length;
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(csv).then(() => {
      showToast(`Copied ${rowsCount} inventory vials to clipboard as CSV!`);
    }).catch(() => fallbackCopy(csv, rowsCount));
  } else {
    fallbackCopy(csv, rowsCount);
  }
}

function fallbackCopy(text, rowsCount) {
  const ta = document.createElement("textarea");
  ta.value = text;
  document.body.appendChild(ta);
  ta.select();
  document.execCommand("copy");
  document.body.removeChild(ta);
  showToast(`Copied ${rowsCount} rows to clipboard as CSV!`);
}

function downloadFilteredCSV(forceAllCols = false) {
  const csv = buildCSVString(forceAllCols);
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  const dt = new Date().toISOString().slice(0, 10);
  const mSlug = state.activeMethod.replace(/[^a-zA-Z0-9]+/g, "_");
  a.href = url;
  a.download = `Lynch_Lab_${mSlug}_${forceAllCols ? "Full" : "Filtered"}_${dt}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  showToast(`Downloaded CSV (${getFilteredCompounds().length} compounds)`);
}

function downloadInventoryCSV() {
  const csv = buildInventoryCSVString();
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  const dt = new Date().toISOString().slice(0, 10);
  a.href = url;
  a.download = `Lynch_Lab_Standards_Inventory_${dt}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  showToast(`Downloaded Inventory CSV (${getFilteredInventory().length} vials)`);
}

function exportAuditLogCSV() {
  if (!state.auditLog.length) {
    showToast("No audit log entries to export yet.");
    return;
  }
  const cols = ["timestamp", "action", "method", "compound", "date_added", "status", "submitted_by", "notes"];
  const esc = v => '"' + String(v ?? "").replace(/"/g, '""') + '"';
  const csv = [
    cols.join(","),
    ...state.auditLog.map(r => cols.map(k => esc(r[k])).join(","))
  ].join("\r\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `Lynch_Lab_Audit_Log_${new Date().toISOString().slice(0, 10)}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}

function exportStandaloneHTML() {
  const updatedBundle = {
    ...EMBEDDED_STATE,
    users: getActiveUsers(),
    methods: state.methods,
    compounds: state.compounds,
    inventory: state.inventory,
    audit_log: state.auditLog,
    generated_at: new Date().toISOString().replace("T", " ").slice(0, 16)
  };
  const currentDoc = document.documentElement.outerHTML;
  const jsonText = JSON.stringify(updatedBundle);
  let replaced = currentDoc.replace(
    /<script id="embedded-lab-data" type="application\/json">[\s\S]*?<\/script>/,
    `<script id="embedded-lab-data" type="application/json">${jsonText}<\/script>`
  );
  // Ensure the auth overlay is always visible on initial page load in the exported HTML
  replaced = replaced.replace(
    /<div id="auth-overlay"[^>]*>/,
    '<div id="auth-overlay">'
  );
  const blob = new Blob(["<!DOCTYPE html>\n" + replaced], { type: "text/html;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "Lynch_Lab_X500R_Method_Dashboard.html";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  showToast("Downloaded updated HTML dashboard with all method & inventory updates embedded!");
}

function openPasswordModal() {
  const sel = document.getElementById("pw-user-select");
  if (sel && state.currentUser) {
    sel.value = state.currentUser;
  }
  document.getElementById("new-pw-1").value = "";
  document.getElementById("new-pw-2").value = "";
  document.getElementById("pw-modal").classList.add("open");
}

function closePasswordModal() {
  document.getElementById("pw-modal").classList.remove("open");
}

function saveNewPassword(e) {
  e.preventDefault();
  const targetUser = document.getElementById("pw-user-select").value;
  const p1 = document.getElementById("new-pw-1").value;
  const p2 = document.getElementById("new-pw-2").value;
  if (p1 !== p2) {
    alert("Passwords do not match.");
    return;
  }
  const users = getActiveUsers();
  users[targetUser] = p1;
  localStorage.setItem(STORAGE_KEY_USERS, JSON.stringify(users));
  closePasswordModal();
  showToast(`Password updated for ${targetUser}!`);
}

function openNewMethodModal() {
  document.getElementById("new-method-name").value = "";
  document.getElementById("method-modal").classList.add("open");
}

function closeNewMethodModal() {
  document.getElementById("method-modal").classList.remove("open");
}

function saveNewMethodTab(e) {
  e.preventDefault();
  const name = document.getElementById("new-method-name").value.trim();
  const seed = document.getElementById("new-method-seed").value;
  if (!name) return;
  if (!state.methods.includes(name)) {
    state.methods.push(name);
  }
  if (seed === "clone_7600") {
    const base = state.compounds.filter(c => c.method === "X500R (CDHR)" && (c.rt_7600_jul26 || c.checked_7600 === "yes"));
    base.forEach((c, idx) => {
      state.compounds.push({
        ...c,
        id: `${name.slice(0,4).toUpperCase()}-${Date.now()}-${idx}`,
        method: name,
        is_custom: true
      });
    });
  } else if (seed === "clone_5600") {
    const base = state.compounds.filter(c => c.method === "X500R (CDHR)" && c.lod_5600_u_ida);
    base.forEach((c, idx) => {
      state.compounds.push({
        ...c,
        id: `${name.slice(0,4).toUpperCase()}-${Date.now()}-${idx}`,
        method: name,
        is_custom: true
      });
    });
  }
  state.activeMethod = name;
  persistCustomState();
  closeNewMethodModal();
  renderMethodTabs();
  renderActiveWorkspace();
  showToast(`Created method tab "${name}"`);
}

function deleteActiveMethodTab() {
  const m = state.activeMethod;
  if (!m || m === "X500R (CDHR)" || m === "ALL_METHODS" || m === "STANDARDS_INVENTORY") return;
  const count = state.compounds.filter(c => c.method === m).length;
  if (!confirm(`Delete method tab "${m}" and its ${count} compounds?`)) return;
  state.methods = state.methods.filter(x => x !== m);
  state.compounds = state.compounds.filter(c => c.method !== m);
  state.activeMethod = state.methods[0] || "X500R (CDHR)";
  state.auditLog.push({
    timestamp: new Date().toISOString().replace("T", " ").slice(0, 16),
    action: "Deleted Method Tab",
    method: m,
    compound: `(${count} compounds removed)`,
    date_added: "",
    status: "Removed",
    submitted_by: state.currentUser || "Lab User",
    notes: `Deleted tab ${m}`
  });
  persistCustomState();
  renderMethodTabs();
  renderAuditLog();
  renderActiveWorkspace();
  showToast(`Deleted method tab "${m}"`);
}

/* ==================== EXCEL (.XLSX / .CSV) FULL METHOD UPLOAD -> NEW TAB ==================== */

let excelUploadState = {
  fileName: "",
  sheets: {}, // { sheetName: [ [cell, cell, ...], ... ] }
  sheetNames: []
};

const EXCEL_MAP_FIELDS = [
  { key: "compound", label: "Compound Name *", required: true, patterns: [/^compound$/i, /^drug/i, /^analyte/i, /^standard name/i, /^chemical/i, /^name$/i, /compound.*name/i] },
  { key: "status", label: "Method Status (1/0/Active)", patterns: [/in\s*x500r\s*method/i, /in\s*method/i, /^status$/i, /^active$/i, /^include$/i] },
  { key: "in_library", label: "In Library? (1/0/Yes)", patterns: [/in.*library/i, /^library$/i] },
  { key: "date_added", label: "Date Added", patterns: [/date\s*added/i, /^added$/i, /^date$/i] },
  { key: "rt_x500r_jul26", label: "Primary RT (min)", patterns: [/x500r.*rt/i, /^rt$/i, /retention\s*time/i, /rt\s*\(min\)/i, /^rt_/i] },
  { key: "rt_x500r_dec25", label: "Secondary / Old RT (min)", patterns: [/dec.*2025.*rt/i, /old\s*rt/i, /prev.*rt/i] },
  { key: "rt_7600_jul26", label: "7600 / Alt Instrument RT", patterns: [/7600.*rt/i, /alt.*rt/i] },
  { key: "formula", label: "Molecular Formula", patterns: [/formula/i, /mol.*form/i] },
  { key: "monoisotopic_mass", label: "Exact / Monoisotopic Mass", patterns: [/monoisotopic/i, /exact\s*mass/i, /^mass$/i, /mol.*weight/i, /^mw$/i] },
  { key: "mh_plus_mz", label: "[M+H]+ / Precursor m/z", patterns: [/m\+h/i, /precursor/i, /^m\/z$/i, /^q1/i, /target\s*mass/i] },
  { key: "primary_mix", label: "Primary Mix", patterns: [/primary\s*mix/i, /^mix$/i, /mix\s*#/i] },
  { key: "secondary_mix", label: "Secondary Mix", patterns: [/secondary\s*mix/i, /sec.*mix/i] },
  { key: "box_location", label: "Box / Freezer Location", patterns: [/location\s*\(box\)/i, /^location$/i, /^box$/i, /box\s*loc/i, /freezer/i] },
  { key: "box_position", label: "Box Position / Well", patterns: [/position/i, /^pos$/i, /^well$/i] },
  { key: "stock_conc", label: "Stock Concentration", patterns: [/concentration/i, /^conc/i, /stock\s*conc/i] },
  { key: "stock_solvent", label: "Solvent", patterns: [/solvent/i, /diluent/i] },
  { key: "manufacturer", label: "Manufacturer", patterns: [/manufacturer/i, /vendor/i, /supplier/i] },
  { key: "catalog_no", label: "Catalog # (COA)", patterns: [/catalog/i, /cat\s*#/i, /part\s*#/i] },
  { key: "lot_no", label: "Lot / Batch #", patterns: [/lot/i, /batch/i] },
  { key: "cas_no", label: "CAS #", patterns: [/^cas/i] },
  { key: "lod_x500r_u_consensus", label: "Urine LOD (ng/mL)", patterns: [/consensus_u_lod/i, /urine.*lod/i, /u_lod/i, /^lod$/i] },
  { key: "lod_x500r_s_consensus", label: "Serum LOD (ng/mL)", patterns: [/consensus_s_lod/i, /serum.*lod/i, /s_lod/i, /plasma.*lod/i] },
  { key: "inchikey", label: "InChIKey", patterns: [/inchikey/i, /inchi/i] },
  { key: "isomers", label: "Possible Isomers", patterns: [/isomer/i] },
  { key: "synonyms", label: "Synonyms", patterns: [/synonym/i, /alias/i] },
  { key: "comments", label: "Comments / Notes", patterns: [/comment/i, /note/i, /reason/i, /remark/i] }
];

function openExcelMethodModal() {
  excelUploadState = { fileName: "", sheets: {}, sheetNames: [] };
  document.getElementById("excel-method-file-input").value = "";
  document.getElementById("excel-parse-status").textContent = "";
  document.getElementById("excel-config-section").style.display = "none";
  document.getElementById("excel-import-submit-btn").disabled = true;
  document.getElementById("excel-default-date").value = new Date().toISOString().slice(0, 10);
  document.getElementById("excel-uploaded-by").value = state.currentUser || "";
  document.getElementById("excel-method-modal").classList.add("open");
}

function closeExcelMethodModal() {
  document.getElementById("excel-method-modal").classList.remove("open");
}

function colLettersToIndex(letters) {
  let idx = 0;
  for (let i = 0; i < letters.length; i++) {
    idx = idx * 26 + (letters.charCodeAt(i) - 64);
  }
  return idx - 1;
}

async function decompressZipEntry(uint8Data, compressionMethod) {
  if (compressionMethod === 0) {
    return new TextDecoder("utf-8").decode(uint8Data);
  }
  if (compressionMethod === 8 && typeof DecompressionStream !== "undefined") {
    const ds = new DecompressionStream("deflate-raw");
    const writer = ds.writable.getWriter();
    writer.write(uint8Data);
    writer.close();
    const resp = new Response(ds.readable);
    const ab = await resp.arrayBuffer();
    return new TextDecoder("utf-8").decode(ab);
  }
  throw new Error("Unsupported ZIP compression method: " + compressionMethod);
}

async function parseXlsxArrayBuffer(arrayBuffer) {
  const bytes = new Uint8Array(arrayBuffer);
  const view = new DataView(arrayBuffer);

  // Locate End of Central Directory (EOCD) signature 0x06054b50
  let eocdPos = -1;
  const minSearch = Math.max(0, bytes.length - 65557);
  for (let i = bytes.length - 22; i >= minSearch; i--) {
    if (view.getUint32(i, true) === 0x06054b50) {
      eocdPos = i;
      break;
    }
  }
  if (eocdPos === -1) {
    throw new Error("Not a valid .xlsx (ZIP) archive.");
  }

  const totalEntries = view.getUint16(eocdPos + 10, true);
  const cdOffset = view.getUint32(eocdPos + 16, true);

  const files = {};
  let ptr = cdOffset;
  const dec = new TextDecoder("utf-8");

  for (let e = 0; e < totalEntries; e++) {
    if (view.getUint32(ptr, true) !== 0x02014b50) break;
    const compMethod = view.getUint16(ptr + 10, true);
    const compSize = view.getUint32(ptr + 20, true);
    const fileNameLen = view.getUint16(ptr + 28, true);
    const extraLen = view.getUint16(ptr + 30, true);
    const commentLen = view.getUint16(ptr + 32, true);
    const localHeaderOffset = view.getUint32(ptr + 42, true);

    const fileName = dec.decode(bytes.subarray(ptr + 46, ptr + 46 + fileNameLen));
    ptr += 46 + fileNameLen + extraLen + commentLen;

    if (
      fileName === "xl/sharedStrings.xml" ||
      fileName === "xl/workbook.xml" ||
      fileName === "xl/_rels/workbook.xml.rels" ||
      fileName.startsWith("xl/worksheets/sheet")
    ) {
      const localFnameLen = view.getUint16(localHeaderOffset + 26, true);
      const localExtraLen = view.getUint16(localHeaderOffset + 28, true);
      const dataStart = localHeaderOffset + 30 + localFnameLen + localExtraLen;
      const slice = bytes.subarray(dataStart, dataStart + compSize);
      files[fileName] = await decompressZipEntry(slice, compMethod);
    }
  }

  const domParser = new DOMParser();

  // 1. Shared Strings
  const sharedStrings = [];
  if (files["xl/sharedStrings.xml"]) {
    const ssDoc = domParser.parseFromString(files["xl/sharedStrings.xml"], "application/xml");
    const siNodes = ssDoc.getElementsByTagName("si");
    for (let i = 0; i < siNodes.length; i++) {
      const tNodes = siNodes[i].getElementsByTagName("t");
      let parts = [];
      for (let j = 0; j < tNodes.length; j++) {
        parts.push(tNodes[j].textContent || "");
      }
      sharedStrings.push(parts.join(""));
    }
  }

  // 2. Workbook Relationships
  const relMap = {};
  if (files["xl/_rels/workbook.xml.rels"]) {
    const relDoc = domParser.parseFromString(files["xl/_rels/workbook.xml.rels"], "application/xml");
    const relNodes = relDoc.getElementsByTagName("Relationship");
    for (let i = 0; i < relNodes.length; i++) {
      const id = relNodes[i].getAttribute("Id");
      let target = (relNodes[i].getAttribute("Target") || "").replace(/^\/+/, "");
      if (!target.startsWith("xl/")) target = "xl/" + target;
      relMap[id] = target;
    }
  }

  // 3. Workbook Sheet List
  const sheetList = [];
  if (files["xl/workbook.xml"]) {
    const wbDoc = domParser.parseFromString(files["xl/workbook.xml"], "application/xml");
    const sNodes = wbDoc.getElementsByTagName("sheet");
    for (let i = 0; i < sNodes.length; i++) {
      const name = sNodes[i].getAttribute("name") || `Sheet${i + 1}`;
      const rId = sNodes[i].getAttribute("r:id") || sNodes[i].getAttribute("id");
      const target = relMap[rId] || `xl/worksheets/sheet${i + 1}.xml`;
      sheetList.push({ name, target });
    }
  }

  // 4. Parse Each Worksheet into a 2D array of rows
  const sheetsOut = {};
  for (const sInfo of sheetList) {
    const xmlStr = files[sInfo.target];
    if (!xmlStr) continue;
    const shDoc = domParser.parseFromString(xmlStr, "application/xml");
    const rowNodes = shDoc.getElementsByTagName("row");
    const rows2D = [];
    let maxCol = 0;

    for (let r = 0; r < rowNodes.length; r++) {
      const cNodes = rowNodes[r].getElementsByTagName("c");
      const rowArr = [];
      for (let c = 0; c < cNodes.length; c++) {
        const cell = cNodes[c];
        const ref = cell.getAttribute("r") || "";
        const m = ref.match(/^([A-Z]+)/);
        const colIdx = m ? colLettersToIndex(m[1]) : c;
        if (colIdx > maxCol) maxCol = colIdx;

        const cType = cell.getAttribute("t") || "";
        let val = "";
        if (cType === "inlineStr") {
          const tNodes = cell.getElementsByTagName("t");
          let p = [];
          for (let k = 0; k < tNodes.length; k++) p.push(tNodes[k].textContent || "");
          val = p.join("");
        } else {
          const vNode = cell.getElementsByTagName("v")[0];
          const rawV = vNode ? (vNode.textContent || "") : "";
          if (cType === "s") {
            const sIdx = parseInt(rawV, 10);
            val = !isNaN(sIdx) && sIdx < sharedStrings.length ? sharedStrings[sIdx] : rawV;
          } else {
            val = rawV;
          }
        }
        rowArr[colIdx] = String(val ?? "").trim();
      }
      if (rowArr.some(x => x !== undefined && x !== "")) {
        rows2D.push(rowArr);
      }
    }
    // Normalize row lengths
    for (let r = 0; r < rows2D.length; r++) {
      for (let c = 0; c <= maxCol; c++) {
        if (rows2D[r][c] === undefined) rows2D[r][c] = "";
      }
    }
    if (rows2D.length > 0) {
      sheetsOut[sInfo.name] = rows2D;
    }
  }
  return sheetsOut;
}

function parseDelimitedTextTo2D(text, delimiter = ",") {
  const rows = [];
  let curRow = [];
  let curVal = "";
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    const next = text[i + 1];
    if (inQuotes) {
      if (ch === '"' && next === '"') {
        curVal += '"';
        i++;
      } else if (ch === '"') {
        inQuotes = false;
      } else {
        curVal += ch;
      }
    } else {
      if (ch === '"') {
        inQuotes = true;
      } else if (ch === delimiter) {
        curRow.push(curVal.trim());
        curVal = "";
      } else if (ch === "\r" && next === "\n") {
        curRow.push(curVal.trim());
        if (curRow.some(x => x !== "")) rows.push(curRow);
        curRow = [];
        curVal = "";
        i++;
      } else if (ch === "\n" || ch === "\r") {
        curRow.push(curVal.trim());
        if (curRow.some(x => x !== "")) rows.push(curRow);
        curRow = [];
        curVal = "";
      } else {
        curVal += ch;
      }
    }
  }
  if (curVal !== "" || curRow.length > 0) {
    curRow.push(curVal.trim());
    if (curRow.some(x => x !== "")) rows.push(curRow);
  }
  return rows;
}

async function handleExcelMethodFileChosen(e) {
  const file = e.target.files[0];
  if (!file) return;
  const statusEl = document.getElementById("excel-parse-status");
  statusEl.style.color = "#0369a1";
  statusEl.textContent = `⏳ Reading "${file.name}"...`;

  try {
    const baseName = file.name.replace(/\.(xlsx|csv|tsv|json)$/i, "").replace(/[_-]+/g, " ").trim();
    excelUploadState.fileName = file.name;

    if (file.name.toLowerCase().endsWith(".xlsx")) {
      const ab = await file.arrayBuffer();
      const sheets = await parseXlsxArrayBuffer(ab);
      excelUploadState.sheets = sheets;
      excelUploadState.sheetNames = Object.keys(sheets);
    } else {
      const text = await file.text();
      const delim = file.name.toLowerCase().endsWith(".tsv") ? "\t" : ",";
      const rows2D = parseDelimitedTextTo2D(text, delim);
      const sName = baseName || "Uploaded_Sheet";
      excelUploadState.sheets = { [sName]: rows2D };
      excelUploadState.sheetNames = [sName];
    }

    if (!excelUploadState.sheetNames.length) {
      throw new Error("No non-empty worksheets found in file.");
    }

    // Populate sheet selector
    const sheetSel = document.getElementById("excel-sheet-select");
    sheetSel.innerHTML = excelUploadState.sheetNames.map(sn => {
      const nRows = Math.max(0, (excelUploadState.sheets[sn] || []).length - 1);
      return `<option value="${escapeHtml(sn)}">${escapeHtml(sn)} (~${nRows} rows)</option>`;
    }).join("");

    // Suggest a clean method tab name
    const tabInput = document.getElementById("excel-new-tab-name");
    let suggestedTab = baseName
      .replace(/\bdrug\s*list\b/i, "")
      .replace(/\bmethod\b/i, "")
      .trim() || baseName;
    if (suggestedTab.length > 36) suggestedTab = suggestedTab.slice(0, 36).trim();
    tabInput.value = suggestedTab;

    // Auto-detect header row (Row 0 vs Row 1 if Row 0 has very few non-empty cells)
    const firstSheet = excelUploadState.sheets[excelUploadState.sheetNames[0]];
    let bestHeaderIdx = 0;
    if (firstSheet.length > 1) {
      const nonEmpty0 = firstSheet[0].filter(Boolean).length;
      const nonEmpty1 = firstSheet[1].filter(Boolean).length;
      if (nonEmpty1 > nonEmpty0 * 1.5 && nonEmpty0 <= 4) {
        bestHeaderIdx = 1;
      }
    }
    document.getElementById("excel-header-row-select").value = String(bestHeaderIdx);

    statusEl.style.color = "#0f766e";
    statusEl.textContent = `✅ Loaded "${file.name}" (${excelUploadState.sheetNames.length} worksheet${excelUploadState.sheetNames.length > 1 ? "s" : ""}). Review settings below and click Generate Method Tab.`;
    document.getElementById("excel-config-section").style.display = "block";
    document.getElementById("excel-import-submit-btn").disabled = false;

    onExcelSheetOrHeaderChanged();
  } catch (err) {
    statusEl.style.color = "#dc2626";
    statusEl.textContent = `❌ Error reading file: ${err.message || err}`;
    document.getElementById("excel-config-section").style.display = "none";
    document.getElementById("excel-import-submit-btn").disabled = true;
  }
}

function onExcelSheetOrHeaderChanged() {
  const sheetName = document.getElementById("excel-sheet-select").value;
  const headerIdx = parseInt(document.getElementById("excel-header-row-select").value, 10) || 0;
  const rows = excelUploadState.sheets[sheetName] || [];
  const headerRow = rows[headerIdx] || [];

  const grid = document.getElementById("excel-col-map-grid");
  const usedCols = new Set();

  grid.innerHTML = EXCEL_MAP_FIELDS.map(field => {
    let matchedCol = -1;
    for (const pat of field.patterns) {
      const idx = headerRow.findIndex((h, cIdx) => !usedCols.has(cIdx) && pat.test(String(h || "").trim()));
      if (idx !== -1) {
        matchedCol = idx;
        usedCols.add(idx);
        break;
      }
    }
    // If compound didn't match by header name, default to first non-empty text column
    if (field.key === "compound" && matchedCol === -1 && headerRow.length > 0) {
      matchedCol = headerRow.findIndex(h => String(h || "").trim() !== "");
      if (matchedCol === -1) matchedCol = 0;
    }

    const opts = [`<option value="-1">— Not Mapped —</option>`];
    headerRow.forEach((h, cIdx) => {
      const colLabel = h ? `${h} (Col ${cIdx + 1})` : `Column ${cIdx + 1}`;
      const sel = (cIdx === matchedCol) ? "selected" : "";
      opts.push(`<option value="${cIdx}" ${sel}>${escapeHtml(colLabel)}</option>`);
    });

    return `<div class="form-field">
      <label style="font-size:11.5px;">${escapeHtml(field.label)}</label>
      <select id="excel-map-${field.key}" class="select-ctrl" onchange="renderExcelPreview()">
        ${opts.join("")}
      </select>
    </div>`;
  }).join("");

  renderExcelPreview();
}

function formatExcelSerialDateJS(val) {
  if (!val) return "";
  const s = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  const num = parseFloat(s);
  if (!isNaN(num) && num >= 35000 && num <= 60000 && /^\d+(\.\d+)?$/.test(s)) {
    const utcDays = Math.floor(num - 25569);
    const d = new Date(utcDays * 86400 * 1000);
    if (!isNaN(d.getTime())) {
      return d.toISOString().slice(0, 10);
    }
  }
  return s;
}

function normalizeStatusJS(rawStatus) {
  const s = String(rawStatus ?? "").trim();
  const low = s.toLowerCase();
  if (s === "1" || low === "yes" || low === "true" || low === "active" || low === "y") return "Active";
  if (s === "0" || low === "no" || low === "false" || low === "inactive" || low === "pending" || low === "n") return "Inactive / Pending";
  if (low.includes("remove") || low.includes("drop")) return "Removed";
  return s ? "Active" : "Active";
}

function normalizeChemKeyJS(name) {
  if (!name) return "";
  let s = String(name).toLowerCase();
  s = s.replace(/\([^)]*\)/g, " ");
  s = s.replace(/\b(hcl|hydrochloride|hemifumarate|fumarate|citrate|tartrate|sulfate|maleate|mesylate|bromide|sodium|potassium|hydrate|dihydrate|monohydrate|solution|powder|salt|d\d+)\b/g, " ");
  s = s.replace(/[^a-z0-9]/g, "");
  return s;
}

function calcExactMassFromFormulaJS(formulaStr) {
  const f = String(formulaStr || "").trim();
  if (!f) return { mass: "", mh: "" };
  const regex = /([A-Z][a-z]?)(\d*)/g;
  let match;
  let totalMass = 0;
  let matchedLen = 0;
  while ((match = regex.exec(f)) !== null) {
    const el = match[1];
    const count = match[2] ? parseInt(match[2], 10) : 1;
    if (!(el in ATOMIC_MASSES)) return { mass: "", mh: "" };
    totalMass += ATOMIC_MASSES[el] * count;
    matchedLen += match[0].length;
  }
  if (matchedLen === f.length && totalMass > 0) {
    return {
      mass: totalMass.toFixed(4),
      mh: (totalMass + 1.007276).toFixed(4)
    };
  }
  return { mass: "", mh: "" };
}

function buildMappedCompoundsFromSheet() {
  const sheetName = document.getElementById("excel-sheet-select").value;
  const headerIdx = parseInt(document.getElementById("excel-header-row-select").value, 10) || 0;
  const defaultDate = document.getElementById("excel-default-date").value || new Date().toISOString().slice(0, 10);
  const targetMethod = (document.getElementById("excel-new-tab-name").value || sheetName || "Uploaded Method").trim();
  const uploadedBy = (document.getElementById("excel-uploaded-by").value || state.currentUser || "Excel Upload").trim();
  const preserveExtra = document.getElementById("excel-preserve-extra").checked;

  const rows = excelUploadState.sheets[sheetName] || [];
  const headerRow = rows[headerIdx] || [];
  const colMap = {};
  const mappedIndices = new Set();

  EXCEL_MAP_FIELDS.forEach(f => {
    const el = document.getElementById("excel-map-" + f.key);
    const cIdx = el ? parseInt(el.value, 10) : -1;
    colMap[f.key] = isNaN(cIdx) ? -1 : cIdx;
    if (cIdx >= 0) mappedIndices.add(cIdx);
  });

  // Build lookup index for Standards Inventory
  const invByKey = new Map();
  const invByLoc = new Map();
  state.inventory.forEach(v => {
    const k = normalizeChemKeyJS(v.standard_name);
    if (k && k.length >= 3) {
      if (!invByKey.has(k)) invByKey.set(k, []);
      invByKey.get(k).push(v);
    }
    const loc = String(v.location || "").trim().toUpperCase();
    if (loc) {
      if (!invByLoc.has(loc)) invByLoc.set(loc, []);
      invByLoc.get(loc).push(v);
    }
  });

  const parsedCompounds = [];
  const nowStr = new Date().toISOString().replace("T", " ").slice(0, 16);

  for (let r = headerIdx + 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row || !row.length) continue;
    const getCell = key => {
      const idx = colMap[key];
      return (idx >= 0 && idx < row.length) ? String(row[idx] ?? "").trim() : "";
    };

    const cname = getCell("compound");
    if (!cname || cname.toLowerCase() === "compound" || cname.toLowerCase() === "standard name") continue;

    const rawStatus = getCell("status");
    const status = normalizeStatusJS(rawStatus);
    const rawLib = getCell("in_library");
    const inLib = (rawLib === "0" || rawLib.toLowerCase() === "no") ? "No" : "Yes";
    const dateAdded = formatExcelSerialDateJS(getCell("date_added")) || defaultDate;

    const formula = getCell("formula");
    let monoMass = getCell("monoisotopic_mass");
    let mhMz = getCell("mh_plus_mz");
    if (monoMass && !isNaN(parseFloat(monoMass))) {
      monoMass = parseFloat(monoMass).toFixed(4);
    }
    if (mhMz && !isNaN(parseFloat(mhMz))) {
      mhMz = parseFloat(mhMz).toFixed(4);
    }
    if (formula && (!monoMass || !mhMz)) {
      const calc = calcExactMassFromFormulaJS(formula);
      if (!monoMass && calc.mass) monoMass = calc.mass;
      if (!mhMz && calc.mh) mhMz = calc.mh;
    } else if (monoMass && !mhMz && !isNaN(parseFloat(monoMass))) {
      mhMz = (parseFloat(monoMass) + 1.007276).toFixed(4);
    }

    let boxLoc = getCell("box_location");
    let stockConc = getCell("stock_conc");
    let stockSolv = getCell("stock_solvent");
    let mfg = getCell("manufacturer");
    let catNo = getCell("catalog_no");
    let lotNo = getCell("lot_no");
    let casNo = getCell("cas_no");
    const synonyms = getCell("synonyms");

    // Cross-link to Standards Inventory
    const matchedVials = [];
    const seenVids = new Set();
    const cKey = normalizeChemKeyJS(cname);
    if (cKey && invByKey.has(cKey)) {
      invByKey.get(cKey).forEach(v => {
        if (!seenVids.has(v.vial_id)) {
          matchedVials.push(v);
          seenVids.add(v.vial_id);
        }
      });
    }
    if (boxLoc && invByLoc.has(boxLoc.toUpperCase())) {
      invByLoc.get(boxLoc.toUpperCase()).forEach(v => {
        if (!seenVids.has(v.vial_id)) {
          matchedVials.push(v);
          seenVids.add(v.vial_id);
        }
      });
    }

    const primaryVial = matchedVials[0] || null;
    if (primaryVial) {
      if (!boxLoc && primaryVial.location) boxLoc = `${primaryVial.location} (Inventory)`;
      if (!stockConc && primaryVial.concentration) stockConc = primaryVial.concentration;
      if (!stockSolv && primaryVial.solvent) stockSolv = primaryVial.solvent;
      if (!mfg && primaryVial.manufacturer) mfg = primaryVial.manufacturer;
      if (!catNo && primaryVial.catalog_no) catNo = primaryVial.catalog_no;
      if (!lotNo && primaryVial.lot_no) lotNo = primaryVial.lot_no;
      if (!casNo && primaryVial.cas_no) casNo = primaryVial.cas_no;
    }

    const notesParts = [];
    if (getCell("comments")) notesParts.push(getCell("comments"));
    if (preserveExtra) {
      headerRow.forEach((hName, cIdx) => {
        if (!mappedIndices.has(cIdx) && hName && row[cIdx]) {
          notesParts.push(`${hName}: ${row[cIdx]}`);
        }
      });
    }

    parsedCompounds.push({
      id: `XLS-${Date.now()}-${r}`,
      method: targetMethod,
      compound: cname,
      status: status,
      in_library: inLib,
      date_added: dateAdded,
      checked_7600: "",
      rt_x500r_jul26: getCell("rt_x500r_jul26"),
      rt_x500r_dec25: getCell("rt_x500r_dec25"),
      rt_7600_jul26: getCell("rt_7600_jul26"),
      primary_mix: getCell("primary_mix"),
      secondary_mix: getCell("secondary_mix"),
      monoisotopic_mass: monoMass,
      mh_plus_mz: mhMz,
      formula: formula,
      isomers: getCell("isomers"),
      synonyms: synonyms,
      box_location: boxLoc,
      box_position: getCell("box_position"),
      stock_conc: stockConc,
      stock_solvent: stockSolv,
      manufacturer: mfg,
      catalog_no: catNo,
      lot_no: lotNo,
      cas_no: casNo,
      vial_status: primaryVial ? primaryVial.vial_status : "",
      vial_exp_date: primaryVial ? primaryVial.exp_date : "",
      checked_out_by: primaryVial ? primaryVial.checked_out_by : "",
      linked_vial_ids: matchedVials.map(v => v.vial_id),
      inchikey: getCell("inchikey"),
      lod_7600_u_swath: "",
      lod_7600_u_ida: "",
      lod_5600_u_ida: "",
      lod_brown_u: "",
      lod_x500r_u_p1: "",
      lod_x500r_u_p2: "",
      x500r_u_agree: "",
      lod_x500r_u_consensus: getCell("lod_x500r_u_consensus"),
      lod_brown_s: "",
      lod_x500r_s_p1: "",
      lod_x500r_s_p2: "",
      x500r_s_agree: "",
      lod_x500r_s_consensus: getCell("lod_x500r_s_consensus"),
      me_7600_i: "",
      me_7600_s: "",
      me_5600_u: "",
      include_flag: "1",
      comments: notesParts.slice(0, 6).join(" | "),
      submitted_by: `${uploadedBy} (${excelUploadState.fileName})`,
      updated_at: nowStr,
      is_custom: true
    });
  }

  return parsedCompounds;
}

function renderExcelPreview() {
  const list = buildMappedCompoundsFromSheet();
  const linkedCount = list.filter(c => c.linked_vial_ids && c.linked_vial_ids.length > 0).length;
  const summaryEl = document.getElementById("excel-preview-summary");
  summaryEl.innerHTML = `Previewing first 8 of <strong>${list.length}</strong> parsed compounds — <span style="color:#0f766e;"><strong>${linkedCount}</strong> automatically matched to 2026 Standards Inventory vials</span>:`;

  document.getElementById("excel-preview-thead").innerHTML = `<tr>
    <th>Compound</th>
    <th>Status</th>
    <th>Date Added</th>
    <th>RT (min)</th>
    <th>Formula</th>
    <th>Exact Mass</th>
    <th>[M+H]+</th>
    <th>Mix</th>
    <th>Box Location</th>
    <th>Matched Vials</th>
  </tr>`;

  document.getElementById("excel-preview-tbody").innerHTML = list.slice(0, 8).map(c => `<tr>
    <td style="font-weight:700;color:var(--ucsf-navy);">${escapeHtml(c.compound)}</td>
    <td><span class="badge ${c.status === 'Active' ? 'badge-active' : 'badge-pending'}">${escapeHtml(c.status)}</span></td>
    <td class="mono">${escapeHtml(c.date_added)}</td>
    <td class="mono">${escapeHtml(c.rt_x500r_jul26)}</td>
    <td class="mono">${escapeHtml(c.formula)}</td>
    <td class="mono">${escapeHtml(c.monoisotopic_mass)}</td>
    <td class="mono">${escapeHtml(c.mh_plus_mz)}</td>
    <td>${escapeHtml(c.primary_mix)}</td>
    <td class="mono">${escapeHtml(c.box_location)}</td>
    <td>${c.linked_vial_ids.length ? `<span class="badge badge-blue">📦 ${c.linked_vial_ids.length} vial(s)</span>` : "—"}</td>
  </tr>`).join("");

  document.getElementById("excel-import-submit-btn").disabled = (list.length === 0);
}

function confirmExcelMethodImport(e) {
  e.preventDefault();
  const tabName = (document.getElementById("excel-new-tab-name").value || "").trim();
  if (!tabName) {
    alert("Please enter a New Method Tab Name.");
    return;
  }
  const existingMode = document.getElementById("excel-existing-mode").value;
  const uploadedBy = (document.getElementById("excel-uploaded-by").value || state.currentUser || "Lab User").trim();
  const parsed = buildMappedCompoundsFromSheet();
  if (!parsed.length) {
    alert("No valid compounds found with current column mapping.");
    return;
  }

  if (!state.methods.includes(tabName)) {
    state.methods.push(tabName);
  } else if (existingMode === "replace") {
    state.compounds = state.compounds.filter(c => c.method !== tabName);
  }

  // Also link the newly imported method name onto matched inventory vials
  const vialMap = new Map();
  state.inventory.forEach(v => vialMap.set(v.vial_id, v));

  parsed.forEach(c => {
    c.method = tabName;
    (c.linked_vial_ids || []).forEach(vid => {
      const v = vialMap.get(vid);
      if (v) {
        if (!Array.isArray(v.linked_methods)) v.linked_methods = [];
        if (!v.linked_methods.includes(tabName)) {
          v.linked_methods.push(tabName);
          v.is_custom = true;
        }
      }
    });
    state.compounds.push(c);
  });

  state.auditLog.push({
    timestamp: new Date().toISOString().replace("T", " ").slice(0, 16),
    action: "Uploaded Method Excel Tab",
    method: tabName,
    compound: `${parsed.length} compounds imported from ${excelUploadState.fileName}`,
    date_added: document.getElementById("excel-default-date").value || "",
    status: "Active",
    submitted_by: uploadedBy,
    notes: `Worksheet: ${document.getElementById("excel-sheet-select").value}`
  });

  state.activeMethod = tabName;
  persistCustomState();
  closeExcelMethodModal();
  renderMethodTabs();
  renderAuditLog();
  renderActiveWorkspace();
  showToast(`Created tab "${tabName}" with ${parsed.length} compounds from "${excelUploadState.fileName}"!`);
}

/* ==================== SPECTRAL LIBRARY (.LBP / .SDF) UPLOAD & MS/MS VIEWER ==================== */

let libUploadState = {
  fileName: "",
  fileType: "",
  entries: []
};

function openLibraryUploadModal() {
  libUploadState = { fileName: "", fileType: "", entries: [] };
  document.getElementById("lib-file-input").value = "";
  document.getElementById("lib-parse-status").textContent = "";
  document.getElementById("lib-config-section").style.display = "none";
  document.getElementById("lib-import-submit-btn").disabled = true;
  document.getElementById("lib-default-date").value = new Date().toISOString().slice(0, 10);
  document.getElementById("lib-uploaded-by").value = state.currentUser || "";

  const targetSel = document.getElementById("lib-target-method-select");
  targetSel.innerHTML = state.methods.map(m => `<option value="${escapeHtml(m)}">${escapeHtml(m)}</option>`).join("");
  if (state.activeMethod !== "ALL_METHODS" && state.activeMethod !== "STANDARDS_INVENTORY") {
    targetSel.value = state.activeMethod;
  }
  onLibImportModeChanged();
  document.getElementById("lib-upload-modal").classList.add("open");
}

function closeLibraryUploadModal() {
  document.getElementById("lib-upload-modal").classList.remove("open");
}

function onLibImportModeChanged() {
  const mode = document.getElementById("lib-import-mode").value;
  document.getElementById("lib-target-tab-wrap").style.display = (mode === "new_tab") ? "none" : "flex";
  document.getElementById("lib-new-tab-wrap").style.display = (mode === "new_tab") ? "flex" : "none";
  if (libUploadState.entries.length > 0) {
    renderLibraryPreview();
  }
}

function condenseTopPeaks(peaks, maxKeep = 35) {
  if (!Array.isArray(peaks) || !peaks.length) return { topPeaks: [], top5Summary: "" };
  let maxInt = 0;
  for (let i = 0; i < peaks.length; i++) {
    if (peaks[i][1] > maxInt) maxInt = peaks[i][1];
  }
  if (maxInt <= 0) maxInt = 1;
  const normalized = peaks.map(p => [
    Number(p[0].toFixed(4)),
    Number(((p[1] / maxInt) * 100.0).toFixed(1))
  ]);
  const byInt = normalized.slice().sort((a, b) => b[1] - a[1]);
  const top5 = byInt.slice(0, 5).map(p => `${p[0].toFixed(4)} (${p[1].toFixed(0)}%)`).join(", ");
  const kept = byInt.slice(0, maxKeep).sort((a, b) => a[0] - b[0]);
  return { topPeaks: kept, top5Summary: top5 };
}

function parseLbpArrayBuffer(arrayBuffer) {
  const buf = new Uint8Array(arrayBuffer);
  const view = new DataView(arrayBuffer);
  const numPages = Math.floor(buf.length / 4096);
  if (numPages < 2) {
    throw new Error("File is too small to be a valid SCIEX .lbp SQL CE 3.5 library.");
  }
  const pidToPno = new Map();

  for (let pno = 0; pno < numPages; pno++) {
    const base = pno * 4096;
    const pid = view.getUint16(base + 4, true);
    if (pid !== 0) pidToPno.set(pid, pno);
  }

  function getPageSlots(pno) {
    const base = pno * 4096;
    const nslots = buf[base + 20];
    const slots = [];
    for (let i = 0; i < nslots; i++) {
      const slotPos = base + 4096 - 4 * (i + 1);
      if (slotPos < base + 24) break;
      const w0 = view.getUint16(slotPos, true);
      const w1 = view.getUint16(slotPos + 2, true);
      const off = 24 + (w0 & 0x0fff);
      const ln = ((w1 & 0x00ff) << 4) | (w0 >> 12);
      const flags = w1 >> 8;
      if (off + ln <= 4096 && ln > 0) {
        slots.push({ i, flags, rec: buf.subarray(base + off, base + off + ln) });
      }
    }
    return slots;
  }

  const pagesByGrp = new Map();
  for (let pno = 0; pno < numPages; pno++) {
    const base = pno * 4096;
    if (buf[base + 6] === 0x40) {
      const grp = view.getUint16(base + 16, true);
      const pid = view.getUint16(base + 4, true);
      const nextPid = view.getUint16(base + 8, true);
      const prevPid = view.getUint16(base + 10, true);
      if (!pagesByGrp.has(grp)) pagesByGrp.set(grp, []);
      pagesByGrp.get(grp).push({ pno, pid, prevPid, nextPid });
    }
  }

  function orderChain(pageList) {
    const byPid = new Map();
    pageList.forEach(x => byPid.set(x.pid, x));
    const heads = pageList.filter(x => x.prevPid === 0 || !byPid.has(x.prevPid)).sort((a, b) => a.pno - b.pno);
    const ordered = [];
    const visited = new Set();
    for (const h of heads) {
      let cur = h;
      while (cur && !visited.has(cur.pid)) {
        visited.add(cur.pid);
        ordered.push(cur.pno);
        cur = byPid.get(cur.nextPid);
      }
    }
    for (const x of pageList) {
      if (!visited.has(x.pid)) ordered.push(x.pno);
    }
    return ordered;
  }

  function concatUint8(a, b) {
    const out = new Uint8Array(a.length + b.length);
    out.set(a, 0);
    out.set(b, a.length);
    return out;
  }

  const allRecords = [];
  for (const [grp, plist] of pagesByGrp.entries()) {
    const orderedPnos = orderChain(plist);
    let pending = null;
    for (const pno of orderedPnos) {
      const slots = getPageSlots(pno);
      for (const s of slots) {
        if (s.flags === 0) {
          if (pending !== null && s.rec.length > 4) {
            const combined = concatUint8(pending, s.rec.subarray(4));
            allRecords.push(combined);
            pending = null;
          }
        } else if (s.flags === 2) {
          pending = s.rec;
          allRecords.push(s.rec);
        } else {
          allRecords.push(s.rec);
        }
      }
    }
  }

  function bytesToHex(u8) {
    let s = "";
    for (let i = 0; i < u8.length; i++) s += u8[i].toString(16).padStart(2, "0");
    return s;
  }

  function extractLobPids(desc) {
    const pids = [];
    const dv = new DataView(desc.buffer, desc.byteOffset, desc.byteLength);
    for (let k = 8; k < desc.length; k += 8) {
      const rem = desc.length - k;
      if (rem < 2) break;
      const p1 = dv.getUint16(k, true);
      if (p1) pids.push(p1);
      if (rem >= 5) {
        const p2 = (desc[k + 2] >> 4) | (desc[k + 3] << 4) | ((desc[k + 4] & 0x0f) << 12);
        if (p2) pids.push(p2);
      }
      if (rem >= 7) {
        const p3 = dv.getUint16(k + 5, true);
        if (p3) pids.push(p3);
      }
    }
    return pids;
  }

  function findPattern(u8, pat) {
    outer: for (let i = 0; i <= u8.length - pat.length; i++) {
      for (let j = 0; j < pat.length; j++) {
        if (u8[i + j] !== pat[j]) continue outer;
      }
      return i;
    }
    return -1;
  }

  function parseTransactXY(blob) {
    const headStr = new TextDecoder("latin1").decode(blob.subarray(0, Math.min(800, blob.length)));
    const mTitle = headStr.match(/(Spectrum from [^\r\n\x00-\x1f]+)/);
    const title = mTitle ? mTitle[1] : "";
    let rt = null;
    const mRt = title.match(/from\s+([0-9.]+)(?:\s+to\s+([0-9.]+))?\s+min/);
    if (mRt) {
      rt = parseFloat(mRt[1]);
      if (mRt[2]) rt = 0.5 * (rt + parseFloat(mRt[2]));
    }
    const ix = findPattern(blob, [0x0f, 0x03, 0x00, 0x00, 0x00]);
    const iy = findPattern(blob, [0x0f, 0x04, 0x00, 0x00, 0x00]);
    const peaks = [];
    if (ix !== -1 && iy !== -1 && iy + 10 <= blob.length) {
      const dv = new DataView(blob.buffer, blob.byteOffset, blob.byteLength);
      const nx = dv.getUint32(ix + 5, true);
      const tx = blob[ix + 9];
      const ny = dv.getUint32(iy + 5, true);
      const ty = blob[iy + 9];
      if (tx === 6 && ty === 6 && nx === ny && nx > 0 && nx <= 10000 && iy + 10 + 8 * ny <= blob.length) {
        for (let k = 0; k < nx; k++) {
          const mz = dv.getFloat64(ix + 10 + k * 8, true);
          const inten = dv.getFloat64(iy + 10 + k * 8, true);
          if (!isNaN(mz) && !isNaN(inten) && inten > 0) {
            peaks.push([mz, inten]);
          }
        }
      }
    }
    return { title, rt, peaks };
  }

  const compounds = new Map();
  const cnames = new Map();
  const msByCguid = new Map();
  const utf8Dec = new TextDecoder("utf-8");

  for (const rec of allRecords) {
    if (rec.length < 42) continue;
    const dv = new DataView(rec.buffer, rec.byteOffset, rec.byteLength);
    const ncols = dv.getUint32(4, true);
    if (ncols === 14 && rec.length >= 83) {
      const cguid = bytesToHex(rec.subarray(11, 27));
      const mw = dv.getFloat64(27, true);
      const mim = dv.getFloat64(35, true);
      const tailStr = new TextDecoder("latin1").decode(rec.subarray(80));
      const mForm = tailStr.match(/([A-Z][A-Za-z0-9]{1,30})/);
      const formula = mForm ? mForm[1] : "";
      compounds.set(cguid, { mw, mim, formula });
    } else if (ncols === 6 && rec.length >= 68 && rec[64] === 0 && rec[65] === 0 && rec[66] === 0 && rec[67] === 0x80) {
      const cguid = bytesToHex(rec.subarray(26, 42));
      const name = utf8Dec.decode(rec.subarray(68)).replace(/\x00+$/g, "").trim();
      if (name && (!cnames.has(cguid) || name.length >= cnames.get(cguid).length)) {
        cnames.set(cguid, name);
      }
    } else if (ncols === 41 && rec.length >= 300) {
      const cguid = bytesToHex(rec.subarray(31, 47));
      const startRt = dv.getFloat64(59, true);
      const endRt = dv.getFloat64(67, true);
      const precMz = dv.getFloat64(123, true);
      const ce = dv.getFloat64(155, true);
      const ces = dv.getFloat64(171, true);
      const wRaw = dv.getUint16(272, true) & 0x7fff;
      const wCent = dv.getUint16(274, true) & 0x7fff;
      if (292 + wCent <= rec.length && wRaw >= 16 && (wCent - wRaw) >= 16) {
        const rawDesc = rec.subarray(292, 292 + wRaw);
        const centDesc = rec.subarray(292 + wRaw, 292 + wCent);
        const centPids = extractLobPids(centDesc);
        const rawPids = extractLobPids(rawDesc);
        const pidsToUse = centPids.length ? centPids : rawPids;
        const chunks = [];
        let totalLen = 0;
        for (const pid of pidsToUse) {
          if (pidToPno.has(pid)) {
            const lpno = pidToPno.get(pid);
            const slice = buf.subarray(lpno * 4096 + 16, (lpno + 1) * 4096);
            chunks.push(slice);
            totalLen += slice.length;
          }
        }
        if (totalLen > 0) {
          const blob = new Uint8Array(totalLen);
          let off = 0;
          for (const ch of chunks) { blob.set(ch, off); off += ch.length; }
          const parsed = parseTransactXY(blob);
          if (parsed.peaks.length > 0) {
            if (!msByCguid.has(cguid)) msByCguid.set(cguid, []);
            msByCguid.get(cguid).push({
              startRt,
              endRt,
              titleRt: parsed.rt,
              precMz,
              ce,
              ces,
              title: parsed.title,
              peaks: parsed.peaks
            });
          }
        }
      }
    }
  }

  const outEntries = [];
  for (const [cguid, rawName] of cnames.entries()) {
    const cInfo = compounds.get(cguid) || {};
    const specs = msByCguid.get(cguid) || [];
    // Pick spectrum with most peaks or first valid spectrum
    specs.sort((a, b) => (b.peaks ? b.peaks.length : 0) - (a.peaks ? a.peaks.length : 0));
    const bestSpec = specs[0] || {};
    const mimVal = (cInfo.mim && cInfo.mim > 10 && cInfo.mim < 5000) ? cInfo.mim.toFixed(4) : "";
    let precVal = (bestSpec.precMz && bestSpec.precMz > 10 && bestSpec.precMz < 5000)
      ? bestSpec.precMz.toFixed(4)
      : (mimVal ? (parseFloat(mimVal) + 1.007276).toFixed(4) : "");
    let rtVal = "";
    if (bestSpec.titleRt && bestSpec.titleRt > 0 && bestSpec.titleRt < 60) {
      rtVal = bestSpec.titleRt.toFixed(2);
    } else if (bestSpec.startRt && bestSpec.startRt > 0 && bestSpec.startRt < 60) {
      rtVal = bestSpec.startRt.toFixed(2);
    }

    const condensed = condenseTopPeaks(bestSpec.peaks || [], 35);
    outEntries.push({
      name: rawName,
      formula: cInfo.formula || "",
      monoisotopic_mass: mimVal,
      mh_plus_mz: precVal,
      rt: rtVal,
      cas_no: "",
      inchikey: "",
      ce: bestSpec.ce ? String(bestSpec.ce) : "",
      ces: bestSpec.ces ? String(bestSpec.ces) : "",
      instrument: "SCIEX (.lbp)",
      created_date: "",
      total_peaks: bestSpec.peaks ? bestSpec.peaks.length : 0,
      msms_peaks: condensed.topPeaks,
      top5_summary: condensed.top5Summary
    });
  }
  return outEntries;
}

function parseSdfText(sdfText) {
  const rawRecords = sdfText.split(/\$\$\$\$\r?\n?/);
  const outEntries = [];

  for (const raw of rawRecords) {
    const trimmed = raw.trim();
    if (!trimmed) continue;
    const lines = raw.split(/\r?\n/);
    const headerName = (lines[0] || "").trim();
    const tags = {};
    const tagRegex = />\s*<([^>]+)>\s*\r?\n([\s\S]*?)(?=(?:\r?\n>\s*<)|\r?\n\s*$|$)/g;
    let m;
    while ((m = tagRegex.exec(raw)) !== null) {
      tags[m[1].trim().toUpperCase()] = (m[2] || "").trim();
    }
    const name = tags["NAME"] || tags["COMPOUND_NAME"] || tags["GENERIC_NAME"] || headerName;
    if (!name) continue;

    const formula = tags["FORMULA"] || tags["MOLECULAR_FORMULA"] || "";
    let mim = tags["MONOISOTOPIC MASS"] || tags["EXACT_MASS"] || tags["MONOISOTOPIC_MASS"] || "";
    let precMz = tags["PRECURSOR M/Z"] || tags["PRECURSOR_MZ"] || "";
    if (mim && !isNaN(parseFloat(mim))) mim = parseFloat(mim).toFixed(4);
    if (precMz && !isNaN(parseFloat(precMz))) precMz = parseFloat(precMz).toFixed(4);
    if (!mim && formula) {
      const calc = calcExactMassFromFormulaJS(formula);
      if (calc.mass) mim = calc.mass;
      if (!precMz && calc.mh) precMz = calc.mh;
    } else if (mim && !precMz && !isNaN(parseFloat(mim))) {
      precMz = (parseFloat(mim) + 1.007276).toFixed(4);
    }

    let rt = tags["RETENTION TIME"] || tags["RT"] || tags["PREDICTED_RT_MIN"] || "";
    if (!rt) {
      const contrib = tags["CONTRIBUTOR"] || tags["COMMENT"] || "";
      const mRt = contrib.match(/(?:Pred_RT|RT)\s*=\s*([0-9.]+)/i);
      if (mRt) rt = parseFloat(mRt[1]).toFixed(2);
    } else if (!isNaN(parseFloat(rt))) {
      rt = parseFloat(rt).toFixed(2);
    }

    const casNo = (tags["CASNO"] || tags["CAS"] || "").replace(/^CAS\s*rn\s*=\s*/i, "").trim();
    const inchikey = tags["INCHIKEY"] || "";
    const ce = tags["COLLISION ENERGY"] || "";
    const ces = tags["COLLISION ENERGY SPREAD"] || "";
    const instrument = tags["INSTRUMENT"] || "SCIEX (.sdf)";
    const createdDate = /^\d{4}-\d{2}-\d{2}/.test(tags["SPECTRUM CREATED DATE"] || "")
      ? tags["SPECTRUM CREATED DATE"].slice(0, 10)
      : "";

    const peaksRaw = tags["MASS SPECTRAL PEAKS"] || "";
    const peaks = [];
    if (peaksRaw) {
      for (const pl of peaksRaw.split(/\r?\n/)) {
        const parts = pl.trim().split(/\s+/);
        if (parts.length >= 2) {
          const mz = parseFloat(parts[0]);
          const inten = parseFloat(parts[1]);
          if (!isNaN(mz) && !isNaN(inten) && inten > 0) {
            peaks.push([mz, inten]);
          }
        }
      }
    }

    const condensed = condenseTopPeaks(peaks, 35);
    outEntries.push({
      name,
      formula,
      monoisotopic_mass: mim,
      mh_plus_mz: precMz,
      rt,
      cas_no: casNo,
      inchikey,
      ce,
      ces,
      instrument,
      created_date: createdDate,
      total_peaks: peaks.length,
      msms_peaks: condensed.topPeaks,
      top5_summary: condensed.top5Summary
    });
  }
  return outEntries;
}

async function handleLibraryFileChosen(e) {
  const file = e.target.files[0];
  if (!file) return;
  const statusEl = document.getElementById("lib-parse-status");
  statusEl.style.color = "#4c1d95";
  statusEl.textContent = `⏳ Parsing "${file.name}"...`;

  try {
    const ext = file.name.toLowerCase().endsWith(".lbp") ? "LBP" : "SDF";
    let entries = [];
    if (ext === "LBP") {
      const ab = await file.arrayBuffer();
      entries = parseLbpArrayBuffer(ab);
    } else {
      const text = await file.text();
      entries = parseSdfText(text);
    }

    if (!entries.length) {
      throw new Error(`No valid compound/spectral records found in ${file.name}.`);
    }

    libUploadState = {
      fileName: file.name,
      fileType: ext,
      entries
    };

    const baseName = file.name.replace(/\.(lbp|sdf)$/i, "").replace(/[_-]+/g, " ").trim();
    document.getElementById("lib-new-tab-name").value = `${baseName} (${ext})`;

    const withPeaks = entries.filter(x => x.total_peaks > 0).length;
    statusEl.style.color = "#0f766e";
    statusEl.textContent = `✅ Parsed ${entries.length} library records (${withPeaks} with MS/MS spectra) from "${file.name}".`;
    document.getElementById("lib-config-section").style.display = "block";
    document.getElementById("lib-import-submit-btn").disabled = false;

    renderLibraryPreview();
  } catch (err) {
    statusEl.style.color = "#dc2626";
    statusEl.textContent = `❌ Error parsing library file: ${err.message || err}`;
    document.getElementById("lib-config-section").style.display = "none";
    document.getElementById("lib-import-submit-btn").disabled = true;
  }
}

function renderLibraryPreview() {
  const mode = document.getElementById("lib-import-mode").value;
  const targetMethod = (mode === "new_tab")
    ? (document.getElementById("lib-new-tab-name").value || "New Library Tab").trim()
    : document.getElementById("lib-target-method-select").value;

  const methodCompByKey = new Map();
  state.compounds.filter(c => c.method === targetMethod).forEach(c => {
    const k = normalizeChemKeyJS(c.compound);
    if (k) methodCompByKey.set(k, c);
  });

  const invByKey = new Map();
  state.inventory.forEach(v => {
    const k = normalizeChemKeyJS(v.standard_name);
    if (k && k.length >= 3) {
      if (!invByKey.has(k)) invByKey.set(k, []);
      invByKey.get(k).push(v);
    }
  });

  let matchedInMethod = 0;
  let matchedInInv = 0;
  libUploadState.entries.forEach(ent => {
    const k = normalizeChemKeyJS(ent.name);
    if (k && methodCompByKey.has(k)) matchedInMethod++;
    if (k && invByKey.has(k)) matchedInInv++;
  });

  const newCount = libUploadState.entries.length - matchedInMethod;
  document.getElementById("lib-preview-summary").innerHTML =
    `Total Library Entries: <strong>${libUploadState.entries.length}</strong> · ` +
    `<span style="color:#166534;">Matched in "${escapeHtml(targetMethod)}": <strong>${matchedInMethod}</strong></span> · ` +
    `<span style="color:#0369a1;">New Compounds not in "${escapeHtml(targetMethod)}": <strong>${newCount}</strong></span> · ` +
    `<span style="color:#0f766e;">Matched to 2026 Standards Inventory: <strong>${matchedInInv}</strong></span> (Previewing first 12):`;

  document.getElementById("lib-preview-tbody").innerHTML = libUploadState.entries.slice(0, 12).map(ent => {
    const k = normalizeChemKeyJS(ent.name);
    const inM = k && methodCompByKey.has(k);
    const invVials = (k && invByKey.get(k)) || [];
    return `<tr>
      <td style="font-weight:700;color:var(--ucsf-navy);">${escapeHtml(ent.name)}</td>
      <td>${inM ? `<span class="badge badge-active">✓ In ${escapeHtml(targetMethod)}</span>` : `<span class="badge badge-blue">+ New to Tab</span>`}</td>
      <td>${invVials.length ? `<span class="badge badge-active">📦 ${escapeHtml(invVials[0].location || "In Stock")}</span>` : `<span style="color:#94a3b8;">—</span>`}</td>
      <td class="mono">${escapeHtml(ent.formula)}</td>
      <td class="mono">${escapeHtml(ent.monoisotopic_mass)}</td>
      <td class="mono">${escapeHtml(ent.mh_plus_mz)}</td>
      <td class="mono">${escapeHtml(ent.rt)}</td>
      <td class="mono">${ent.total_peaks}</td>
      <td class="mono" style="font-size:11px;color:#475569;">${escapeHtml(ent.top5_summary || "—")}</td>
    </tr>`;
  }).join("");
}

function confirmLibraryUpload(e) {
  e.preventDefault();
  if (!libUploadState.entries.length) return;

  const mode = document.getElementById("lib-import-mode").value;
  let targetMethod = document.getElementById("lib-target-method-select").value;
  if (mode === "new_tab") {
    targetMethod = (document.getElementById("lib-new-tab-name").value || "").trim();
    if (!targetMethod) {
      alert("Please enter a New Method / Library Tab Name.");
      return;
    }
    if (!state.methods.includes(targetMethod)) {
      state.methods.push(targetMethod);
    }
  }

  const defaultDate = document.getElementById("lib-default-date").value || new Date().toISOString().slice(0, 10);
  const uploadedBy = (document.getElementById("lib-uploaded-by").value || state.currentUser || "Lab User").trim();
  const nowStr = new Date().toISOString().replace("T", " ").slice(0, 16);

  // Build lookup of existing compounds in targetMethod
  const compIndexByKey = new Map();
  state.compounds.forEach((c, idx) => {
    if (c.method === targetMethod) {
      const k = normalizeChemKeyJS(c.compound);
      if (k && !compIndexByKey.has(k)) compIndexByKey.set(k, idx);
    }
  });

  // Build lookup of Standards Inventory vials
  const invByKey = new Map();
  state.inventory.forEach(v => {
    const k = normalizeChemKeyJS(v.standard_name);
    if (k && k.length >= 3) {
      if (!invByKey.has(k)) invByKey.set(k, []);
      invByKey.get(k).push(v);
    }
  });

  let updatedCount = 0;
  let addedCount = 0;

  libUploadState.entries.forEach((ent, i) => {
    const k = normalizeChemKeyJS(ent.name);
    const existingIdx = (k && compIndexByKey.has(k)) ? compIndexByKey.get(k) : -1;

    if (existingIdx !== -1 && mode !== "new_tab") {
      const cur = state.compounds[existingIdx];
      cur.in_library = "Yes";
      if (!cur.formula && ent.formula) cur.formula = ent.formula;
      if (!cur.monoisotopic_mass && ent.monoisotopic_mass) cur.monoisotopic_mass = ent.monoisotopic_mass;
      if (!cur.mh_plus_mz && ent.mh_plus_mz) cur.mh_plus_mz = ent.mh_plus_mz;
      if (!cur.rt_x500r_jul26 && ent.rt) cur.rt_x500r_jul26 = ent.rt;
      if (!cur.cas_no && ent.cas_no) cur.cas_no = ent.cas_no;
      if (!cur.inchikey && ent.inchikey) cur.inchikey = ent.inchikey;
      if (ent.msms_peaks && ent.msms_peaks.length) {
        cur.msms_peaks = ent.msms_peaks;
        cur.msms_total_peaks = ent.total_peaks;
        cur.msms_ce = ent.ce;
        cur.msms_ces = ent.ces;
        cur.msms_instrument = ent.instrument;
      }
      const libNote = `Synced from ${libUploadState.fileName} (${ent.total_peaks} MS/MS peaks${ent.top5_summary ? "; Top ions: " + ent.top5_summary : ""})`;
      if (!cur.comments || !cur.comments.includes(libUploadState.fileName)) {
        cur.comments = cur.comments ? `${cur.comments} | ${libNote}` : libNote;
      }
      cur.is_custom = true;
      updatedCount++;
    } else if (mode === "sync_and_append" || mode === "new_tab") {
      const matchedVials = (k && invByKey.get(k)) ? invByKey.get(k) : [];
      const primaryVial = matchedVials[0] || null;
      matchedVials.forEach(v => {
        if (!Array.isArray(v.linked_methods)) v.linked_methods = [];
        if (!v.linked_methods.includes(targetMethod)) {
          v.linked_methods.push(targetMethod);
          v.is_custom = true;
        }
      });

      state.compounds.push({
        id: `LIB-${Date.now()}-${i}`,
        method: targetMethod,
        compound: ent.name,
        status: "Active",
        in_library: "Yes",
        date_added: ent.created_date || defaultDate,
        checked_7600: libUploadState.fileName.includes("7600") ? "yes" : "",
        rt_x500r_jul26: ent.rt || "",
        rt_x500r_dec25: "",
        rt_7600_jul26: libUploadState.fileName.includes("7600") ? (ent.rt || "") : "",
        primary_mix: "",
        secondary_mix: "",
        monoisotopic_mass: ent.monoisotopic_mass || "",
        mh_plus_mz: ent.mh_plus_mz || "",
        formula: ent.formula || "",
        isomers: "",
        synonyms: "",
        box_location: primaryVial ? `${primaryVial.location} (Inventory)` : "",
        box_position: "",
        stock_conc: primaryVial ? primaryVial.concentration : "",
        stock_solvent: primaryVial ? primaryVial.solvent : "",
        manufacturer: primaryVial ? primaryVial.manufacturer : "",
        catalog_no: primaryVial ? primaryVial.catalog_no : "",
        lot_no: primaryVial ? primaryVial.lot_no : "",
        cas_no: ent.cas_no || (primaryVial ? primaryVial.cas_no : ""),
        vial_status: primaryVial ? primaryVial.vial_status : "",
        vial_exp_date: primaryVial ? primaryVial.exp_date : "",
        checked_out_by: primaryVial ? primaryVial.checked_out_by : "",
        linked_vial_ids: matchedVials.map(v => v.vial_id),
        inchikey: ent.inchikey || "",
        lod_7600_u_swath: "",
        lod_7600_u_ida: "",
        lod_5600_u_ida: "",
        lod_brown_u: "",
        lod_x500r_u_p1: "",
        lod_x500r_u_p2: "",
        x500r_u_agree: "",
        lod_x500r_u_consensus: "",
        lod_brown_s: "",
        lod_x500r_s_p1: "",
        lod_x500r_s_p2: "",
        x500r_s_agree: "",
        lod_x500r_s_consensus: "",
        me_7600_i: "",
        me_7600_s: "",
        me_5600_u: "",
        include_flag: "1",
        comments: `Imported from ${libUploadState.fileName} (${ent.total_peaks} MS/MS peaks${ent.top5_summary ? "; Top ions: " + ent.top5_summary : ""})`,
        submitted_by: `${uploadedBy} (${libUploadState.fileName})`,
        updated_at: nowStr,
        msms_peaks: ent.msms_peaks || [],
        msms_total_peaks: ent.total_peaks || 0,
        msms_ce: ent.ce || "",
        msms_ces: ent.ces || "",
        msms_instrument: ent.instrument || "",
        is_custom: true
      });
      addedCount++;
    }
  });

  state.auditLog.push({
    timestamp: nowStr,
    action: `Uploaded ${libUploadState.fileType} Library`,
    method: targetMethod,
    compound: `${libUploadState.fileName} (${updatedCount} synced, ${addedCount} added)`,
    date_added: defaultDate,
    status: "In Library = Yes",
    submitted_by: uploadedBy,
    notes: `Mode: ${mode} | Total library records: ${libUploadState.entries.length}`
  });

  state.activeMethod = targetMethod;
  persistCustomState();
  closeLibraryUploadModal();
  renderMethodTabs();
  renderAuditLog();
  renderActiveWorkspace();
  showToast(`Applied "${libUploadState.fileName}" to ${targetMethod}: ${updatedCount} synced, ${addedCount} added!`);
}

function openMsmsModal(compId) {
  const c = state.compounds.find(x => x.id === compId);
  if (!c || !c.msms_peaks || !c.msms_peaks.length) return;
  const peaks = c.msms_peaks;
  const minMz = Math.max(30, Math.floor(Math.min(...peaks.map(p => p[0])) - 15));
  const maxMz = Math.ceil(Math.max(...peaks.map(p => p[0]), parseFloat(c.mh_plus_mz) || 100) + 20);
  const spanMz = Math.max(10, maxMz - minMz);

  // Identify top 8 peaks for m/z text labels above sticks
  const top8Set = new Set(
    peaks.slice().sort((a, b) => b[1] - a[1]).slice(0, 8).map(p => p[0])
  );

  const W = 820, H = 260, padL = 50, padR = 25, padT = 28, padB = 34;
  const plotW = W - padL - padR;
  const plotH = H - padT - padB;

  const sticksSvg = peaks.map(p => {
    const mz = p[0];
    const relInt = p[1];
    const x = padL + ((mz - minMz) / spanMz) * plotW;
    const y = padT + plotH - (relInt / 100.0) * plotH;
    const isTop = top8Set.has(mz);
    const col = isTop ? "#4c1d95" : "#007CBE";
    const sw = isTop ? 2.2 : 1.4;
    const label = isTop
      ? `<text x="${x.toFixed(1)}" y="${Math.max(12, y - 5).toFixed(1)}" text-anchor="middle" font-size="10" font-family="monospace" font-weight="700" fill="#4c1d95">${mz.toFixed(4)}</text>`
      : "";
    return `<line x1="${x.toFixed(1)}" y1="${padT + plotH}" x2="${x.toFixed(1)}" y2="${y.toFixed(1)}" stroke="${col}" stroke-width="${sw}" />${label}`;
  }).join("");

  const xTicks = [];
  const step = Math.ceil(spanMz / 8 / 10) * 10 || 20;
  for (let mVal = Math.ceil(minMz / step) * step; mVal <= maxMz; mVal += step) {
    const x = padL + ((mVal - minMz) / spanMz) * plotW;
    xTicks.push(`<line x1="${x.toFixed(1)}" y1="${padT + plotH}" x2="${x.toFixed(1)}" y2="${padT + plotH + 5}" stroke="#64748b" />
      <text x="${x.toFixed(1)}" y="${padT + plotH + 18}" text-anchor="middle" font-size="10.5" fill="#475569">${mVal}</text>`);
  }

  const sortedByInt = peaks.slice().sort((a, b) => b[1] - a[1]);

  document.getElementById("msms-modal-title").textContent =
    `MS/MS Reference Spectrum: ${c.compound} (${c.formula || "No Formula"})`;
  document.getElementById("msms-modal-body").innerHTML = `
    <div style="display:flex;gap:14px;flex-wrap:wrap;margin-bottom:12px;background:#f8fafc;border:1px solid #e2e8f0;padding:10px 14px;border-radius:8px;font-size:12.5px;">
      <div><strong>Compound:</strong> ${escapeHtml(c.compound)}</div>
      <div><strong>Formula:</strong> <span class="mono">${escapeHtml(c.formula || "—")}</span></div>
      <div><strong>Exact Mass:</strong> <span class="mono">${escapeHtml(c.monoisotopic_mass || "—")} Da</span></div>
      <div><strong>Precursor [M+H]+:</strong> <span class="mono">${escapeHtml(c.mh_plus_mz || "—")} m/z</span></div>
      <div><strong>Library RT:</strong> <span class="mono">${escapeHtml(c.rt_x500r_jul26 || "—")} min</span></div>
      <div><strong>CE / CES:</strong> <span class="mono">${escapeHtml(c.msms_ce || "35")} / ${escapeHtml(c.msms_ces || "15")} eV</span></div>
      <div><strong>Raw Centroid Peaks:</strong> <span class="mono">${c.msms_total_peaks || peaks.length}</span> (Top ${peaks.length} shown)</div>
    </div>
    <div style="border:1px solid #cbd5e1;border-radius:8px;background:#fff;padding:8px;margin-bottom:14px;overflow-x:auto;">
      <svg viewBox="0 0 ${W} ${H}" style="width:100%;height:auto;display:block;">
        <line x1="${padL}" y1="${padT}" x2="${padL}" y2="${padT + plotH}" stroke="#94a3b8" />
        <line x1="${padL}" y1="${padT + plotH}" x2="${padL + plotW}" y2="${padT + plotH}" stroke="#94a3b8" />
        <text x="14" y="${padT + plotH / 2}" transform="rotate(-90 14 ${padT + plotH / 2})" text-anchor="middle" font-size="11" fill="#475569" font-weight="700">Relative Intensity (%)</text>
        <text x="${padL + plotW / 2}" y="${H - 3}" text-anchor="middle" font-size="11" fill="#475569" font-weight="700">m/z (Da)</text>
        ${xTicks.join("")}
        ${sticksSvg}
      </svg>
    </div>
    <div style="font-size:12.5px;font-weight:700;color:var(--ucsf-navy);margin-bottom:6px;">Top Diagnostic Fragment Ions (Base-Peak Normalized):</div>
    <div class="table-scroll" style="max-height:190px;border:1px solid #e2e8f0;border-radius:8px;">
      <table class="comp-table">
        <thead>
          <tr>
            <th>Rank</th>
            <th>Fragment m/z</th>
            <th>Relative Intensity (%)</th>
            <th>Neutral Loss from [M+H]+ (Da)</th>
          </tr>
        </thead>
        <tbody>
          ${sortedByInt.map((p, idx) => {
            const prec = parseFloat(c.mh_plus_mz);
            const nl = (!isNaN(prec) && prec > p[0]) ? (prec - p[0]).toFixed(4) : "—";
            return `<tr>
              <td class="mono">#${idx + 1}</td>
              <td class="mono" style="font-weight:700;color:#4c1d95;">${p[0].toFixed(4)}</td>
              <td class="mono">${p[1].toFixed(1)}%</td>
              <td class="mono">${nl}</td>
            </tr>`;
          }).join("")}
        </tbody>
      </table>
    </div>
  `;
  document.getElementById("msms-modal").classList.add("open");
}

function clearLocalCustomChanges() {
  if (!confirm("Reset local browser additions/edits back to the embedded spreadsheet baseline?")) return;
  localStorage.removeItem(STORAGE_KEY_DATA);
  state.auditLog = [];
  initState();
  showToast("Reset to baseline dataset.");
}

function showToast(msg) {
  const box = document.getElementById("toast-box");
  box.textContent = msg;
  box.style.display = "block";
  clearTimeout(window._toastTimer);
  window._toastTimer = setTimeout(() => {
    box.style.display = "none";
  }, 3600);
}

function escapeHtml(str) {
  return String(str ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// =========================================================================
// MOMA COLOR PALETTE PICKER ENGINE (BlakeRMills/MoMAColors)
// =========================================================================
const MOMA_PALETTES = {
  "ucsf_original": null,
  "original":  ["#3b82f6", "#ef4444", "#f97316", "#eab308", "#10b981", "#06b6d4", "#a855f7", "#ec4899"],
  "VanGogh":   ["#41395E", "#6F6941", "#937141", "#A39960", "#B5B59C", "#A1B7B9", "#6996A6", "#4B687F"],
  "Doughton":  ["#5928ED", "#1B94F4", "#64C1EC", "#98E0EB", "#F9D0B6", "#F38343", "#F0409A", "#AC167C"],
  "Ernst":     ["#91323A", "#B7484B", "#D29B52", "#E6CA63", "#B9C5B9", "#8AA6A3", "#4F6F76", "#31485A"],
  "OKeeffe":   ["#B24422", "#D26344", "#E48F66", "#F0B98D", "#F4DEC4", "#B8C7C7", "#6B909E", "#2D5266"],
  "ustwo":     ["#FA5E5B", "#F9A153", "#F7D969", "#9FD08A", "#00B6B5", "#0082B4", "#5F5296", "#35284E"],
  "Abbott":    ["#99281A", "#D34A28", "#E98A33", "#F0C565", "#76A08A", "#3F7879", "#285260", "#1A3240"],
  "Alkalay1":  ["#8C2B32", "#C44B39", "#DF7E38", "#E8B752", "#9AA889", "#5C8385", "#355D68", "#1F3642"],
  "Alkalay2":  ["#E63946", "#F4A261", "#E9C46A", "#2A9D8F", "#264653", "#457B9D", "#A8DADC", "#6D597A"],
  "Althoff":   ["#6E2B3B", "#A64942", "#D97F43", "#E8B85F", "#B6C2A5", "#7B9E9A", "#486E75", "#2B4450"],
  "Andri":     ["#C83E2B", "#E67F33", "#F3C354", "#D8DEC3", "#7AA695", "#3A787E", "#224B5E", "#152B38"],
  "Avedon":    ["#D9383A", "#F27935", "#F5B841", "#9EC488", "#3B9A8B", "#266D78", "#2B4560", "#1E2A38"],
  "Budnitz":   ["#B33951", "#E3655B", "#F49D6E", "#F7D08A", "#A3C9A8", "#69A297", "#50808E", "#3B5260"],
  "Clay":      ["#8F3B2B", "#C7633B", "#E29B52", "#EDD082", "#B5C4A7", "#739A92", "#436B70", "#26424A"],
  "Connors":   ["#C7362F", "#E66C37", "#F2A949", "#F5D97E", "#8FB89C", "#4B8B8B", "#2E5E6E", "#1B3847"],
  "Dali":      ["#B8332A", "#D96B27", "#E6A632", "#EBD471", "#92B088", "#50847C", "#2F5C66", "#1C3642"],
  "Exter":     ["#B8292F", "#E05A33", "#F09436", "#F5CA5C", "#78A889", "#3B7D7D", "#255264", "#173142"],
  "Flash":     ["#D92B4B", "#F26430", "#F2A922", "#F2E205", "#65A694", "#2B7270", "#1E4854", "#122A36"],
  "Fritsch":   ["#C92C3B", "#E86838", "#F4A641", "#F7D86B", "#94BD97", "#4F8F8B", "#31606E", "#1D3B4A"],
  "Kippenberger": ["#8A2B36", "#C24E3A", "#E3843E", "#EFC05B", "#A3B899", "#648E8C", "#3B626B", "#223C47"],
  "Klein":     ["#002FA7", "#1E56C6", "#4282DE", "#6FAEF0", "#F2C94C", "#F2994A", "#EB5757", "#9B51E0"],
  "Koons":     ["#D92B6B", "#F25E3D", "#F2A03D", "#F2D43D", "#45B39D", "#2980B9", "#8E44AD", "#2C3E50"],
  "Levine1":   ["#9E2A2B", "#D15834", "#E89242", "#F2C86B", "#A7BM9E".replace("M","8"), "#69918E", "#3D656E", "#233F4A"],
  "Levine2":   ["#7A2E3A", "#B35240", "#D9844A", "#E8BA6B", "#B8C4A9", "#7C9E99", "#4A6F75", "#2B464F"],
  "Liu":       ["#B52B38", "#DE5B39", "#ED9342", "#F5C963", "#8EB396", "#4C8787", "#2E5B6B", "#1B3745"],
  "Lupi":      ["#C93B5B", "#E66C64", "#F2A272", "#F7D394", "#A6C9B2", "#689F9E", "#427380", "#294957"],
  "Ohchi":     ["#D93829", "#F27435", "#F5B041", "#F7DC6F", "#7FB39A", "#3D8484", "#265969", "#173645"],
  "Palermo":   ["#B83227", "#E0632E", "#F09B38", "#F5CE62", "#85AD8E", "#45807D", "#295561", "#183340"],
  "Panton":    ["#D9232E", "#F25C27", "#F2911B", "#F2C12E", "#008C72", "#005A70", "#3B2854", "#1F142E"],
  "Picabia":   ["#8F2D38", "#C4533D", "#E08745", "#EEC068", "#ABBAA0", "#6D9491", "#416770", "#26404A"],
  "Picasso":   ["#A62B2B", "#D65A31", "#E8963E", "#F2CA68", "#9BB89D", "#5A8C8B", "#36606B", "#1F3B45"],
  "Rattner":   ["#992638", "#CC4E3C", "#E68342", "#F2BC61", "#9CB599", "#5B8A88", "#365E69", "#1F3942"],
  "Sidhu":     ["#B82D42", "#E05D44", "#F0944D", "#F5C96B", "#91B89E", "#508A8A", "#325F6E", "#1D3B47"],
  "Smith":     ["#A82E2E", "#D95D34", "#EB9640", "#F2CA66", "#96B598", "#558987", "#335D69", "#1D3942"],
  "vonHeyl":   ["#8C2D46", "#C2544D", "#E08855", "#EFC073", "#B2C2AA", "#759996", "#476D73", "#2A444D"],
  "Warhol":    ["#F22973", "#F26329", "#F2A629", "#F2D929", "#29D9A6", "#299BF2", "#8C29F2", "#291E40"]
};

function _hexToRgb(hex) {
  const h = (hex || "#000000").replace("#", "").trim();
  const n = parseInt(h.length === 3 ? h.split("").map(c => c + c).join("") : h, 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}
function _rgbToHex(r, g, b) {
  const c = v => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, "0");
  return "#" + c(r) + c(g) + c(b);
}
function _mixHex(h1, h2, t) {
  const [r1, g1, b1] = _hexToRgb(h1);
  const [r2, g2, b2] = _hexToRgb(h2);
  return _rgbToHex(r1 + (r2 - r1) * t, g1 + (g2 - g1) * t, b1 + (b2 - b1) * t);
}
function _relLum(hex) {
  const [r, g, b] = _hexToRgb(hex).map(v => {
    const s = v / 255;
    return s <= 0.03928 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4);
  });
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

function computePaletteSurfaceTheme(key, p) {
  if (!key || key === "ucsf_original") {
    return {
      isLight: true,
      bg: "#f4f6f9",
      card: "#ffffff",
      secondary: "#f8fafc",
      border: "#dde3ec",
      fg: "#1e293b",
      mutedFg: "#64748b",
      leadSwatches: ["#052049", "#007CBE"]
    };
  }
  if (key === "original") {
    return {
      isLight: false,
      bg: "#090d16",
      card: "#111827",
      secondary: "#1e293b",
      border: "#1f2937",
      fg: "#f8fafc",
      mutedFg: "#94a3b8",
      leadSwatches: ["#090d16", "#111827"]
    };
  }
  const scored = p.map((hex, idx) => ({ hex, idx, lum: _relLum(hex) }));
  const sortedByLum = [...scored].sort((a, b) => a.lum - b.lum);
  const darkest = sortedByLum[0].hex;
  const secondDarkest = sortedByLum[1].hex;
  const lightest = sortedByLum[sortedByLum.length - 1].hex;
  const secondLightest = sortedByLum[sortedByLum.length - 2].hex;

  const avgLum = scored.reduce((s, x) => s + x.lum, 0) / scored.length;
  const lightCount = scored.filter(x => x.lum > 0.36).length;
  const isLight = (avgLum >= 0.30) || (lightCount >= 3 && sortedByLum[sortedByLum.length - 1].lum > 0.52);

  if (isLight) {
    const bg = _mixHex(lightest, "#ffffff", 0.62);
    const card = _mixHex(secondLightest, "#ffffff", 0.82);
    const secondary = _mixHex(lightest, secondLightest, 0.45);
    const border = _mixHex(darkest, bg, 0.76);
    const fg = _mixHex(darkest, "#090d16", 0.45);
    const mutedFg = _mixHex(darkest, bg, 0.38);
    return { isLight: true, bg, card, secondary, border, fg, mutedFg, leadSwatches: [lightest, secondLightest] };
  } else {
    const bg = _mixHex(darkest, "#06080e", 0.62);
    const card = _mixHex(secondDarkest, "#0b0f19", 0.56);
    const secondary = _mixHex(secondDarkest, darkest, 0.35);
    const border = _mixHex(lightest, card, 0.78);
    const fg = _mixHex(lightest, "#ffffff", 0.65);
    const mutedFg = _mixHex(lightest, card, 0.35);
    return { isLight: false, bg, card, secondary, border, fg, mutedFg, leadSwatches: [darkest, secondDarkest] };
  }
}

function applyColorPalette(key) {
  const sel = document.getElementById("palette-select");
  if (sel && sel.value !== key && (key === "ucsf_original" || MOMA_PALETTES[key])) {
    sel.value = key;
  }
  const resolvedKey = (key === "ucsf_original" || MOMA_PALETTES[key]) ? key : "ucsf_original";
  localStorage.setItem("lynch_lab_moma_palette_v1", resolvedKey);

  const p = MOMA_PALETTES[resolvedKey] || ["#052049", "#007CBE", "#008080", "#10b981", "#f59e0b", "#6366f1", "#ec4899", "#06b6d4"];
  const surf = computePaletteSurfaceTheme(resolvedKey, p);

  const bar = document.getElementById("palette-swatch-bar");
  if (bar) {
    const leadHtml = surf.leadSwatches.map((c, i) =>
      `<span title="Surface ${i === 0 ? 'Background' : 'Card'} Lead (${c})" style="display:inline-block;width:11px;height:11px;border-radius:3px;background:${c};border:1.5px solid ${surf.isLight ? '#334155' : '#f8fafc'};margin-right:${i === 1 ? '4px' : '1px'};"></span>`
    ).join("");
    const palHtml = p.map(c =>
      `<span title="Accent Swatch (${c})" style="display:inline-block;width:10px;height:10px;border-radius:2px;background:${c};border:1px solid rgba(128,128,128,0.35);"></span>`
    ).join("");
    bar.innerHTML = leadHtml + palHtml;
  }

  const badge = document.getElementById("palette-bg-mode-badge");
  if (badge) {
    badge.textContent = surf.isLight ? "☀ Light" : "☾ Dark";
    badge.style.background = surf.isLight ? "rgba(255,255,255,0.88)" : "rgba(15,23,42,0.75)";
    badge.style.color = surf.isLight ? "#0f172a" : "#f8fafc";
    badge.style.border = `1px solid ${surf.isLight ? 'rgba(203,213,225,0.6)' : 'rgba(148,163,184,0.4)'}`;
  }

  const styleEl = document.getElementById("dynamic-palette-css");
  if (!styleEl) return;

  if (resolvedKey === "ucsf_original") {
    styleEl.textContent = "";
    return;
  }

  const headerGradStart = _mixHex(p[0], "#050811", 0.35);
  const headerGradMid   = _mixHex(p[6] || p[0], "#091526", 0.25);
  const headerGradEnd   = p[5] || p[1];

  styleEl.textContent = `
    :root {
      --ucsf-navy: ${surf.isLight ? _mixHex(p[0], "#090d16", 0.25) : p[5]};
      --ucsf-blue: ${p[0]};
      --ucsf-teal: ${p[4]};
      --bg-main: ${surf.bg};
      --card-bg: ${surf.card};
      --border-col: ${surf.border};
      --text-main: ${surf.fg};
      --text-muted: ${surf.mutedFg};
    }
    body {
      background: ${surf.bg} !important;
      color: ${surf.fg} !important;
    }
    .top-header {
      background: linear-gradient(90deg, ${headerGradStart} 0%, ${headerGradMid} 62%, ${headerGradEnd} 100%) !important;
    }
    .btn-primary {
      background: ${p[4]} !important;
      border-color: ${_mixHex(p[4], "#000000", 0.2)} !important;
      color: #ffffff !important;
    }
    .btn-teal {
      background: ${p[5]} !important;
      border-color: ${_mixHex(p[5], "#000000", 0.2)} !important;
      color: #ffffff !important;
    }
    .btn-blue {
      background: ${p[0]} !important;
      color: #ffffff !important;
    }
    .btn-outline {
      background: ${surf.card} !important;
      color: ${surf.fg} !important;
      border-color: ${surf.border} !important;
    }
    .btn-outline:hover {
      background: ${surf.secondary} !important;
    }
    .kpi-card, .toolbar-card, .table-card, .val-card, .modal-card, .auth-card {
      background: ${surf.card} !important;
      border-color: ${surf.border} !important;
      color: ${surf.fg} !important;
    }
    .kpi-val {
      color: ${surf.isLight ? _mixHex(p[0], "#090d16", 0.2) : p[3]} !important;
    }
    .subnav-group {
      background: ${surf.secondary} !important;
      border: 1px solid ${surf.border} !important;
    }
    .subnav-btn {
      color: ${surf.mutedFg} !important;
    }
    .subnav-btn.active {
      background: ${surf.card} !important;
      color: ${surf.fg} !important;
    }
    .comp-table th {
      background: ${surf.secondary} !important;
      color: ${surf.mutedFg} !important;
      border-bottom-color: ${surf.border} !important;
    }
    .comp-table td {
      border-bottom-color: ${surf.border} !important;
      color: ${surf.fg} !important;
    }
    .comp-table tbody tr:hover {
      background: ${_mixHex(p[0], surf.card, 0.88)} !important;
    }
    .search-box, .select-ctrl, .date-ctrl, .realtime-input, .inline-date, .form-field input, .form-field select, .form-field textarea, .auth-input {
      background: ${surf.isLight ? "#ffffff" : surf.secondary} !important;
      color: ${surf.fg} !important;
      border-color: ${surf.border} !important;
    }
    .realtime-bar {
      background: ${surf.secondary} !important;
      border-color: ${surf.border} !important;
    }
    .table-footer, .modal-header, .modal-footer {
      background: ${surf.secondary} !important;
      border-color: ${surf.border} !important;
      color: ${surf.fg} !important;
    }
    .method-tab.active {
      background: ${surf.bg} !important;
      color: ${surf.fg} !important;
    }
    .method-tab.active .tab-pill {
      background: ${_mixHex(p[0], surf.card, 0.75)} !important;
      color: ${surf.fg} !important;
    }
  `;
}

window.addEventListener("DOMContentLoaded", initState);
</script>
</body>
</html>
"""


def main():
    dataset = build_dataset()
    dataset["users"] = {
        "Gillian": "G1ll1an",
        "Lauren": "Laur3n",
        "Michael": "M1chael",
    }
    dataset["methods"] = ["X500R (CDHR)", "7600 (CDS W93)"]
    dataset["audit_log"] = []

    json_str = json.dumps(dataset, separators=(",", ":"))
    html_out = HTML_TEMPLATE.replace("__EMBEDDED_JSON_DATA__", json_str)

    with open(OUT_HTML, "w", encoding="utf-8") as f:
        f.write(html_out)

    n_linked = sum(1 for c in dataset["compounds"] if c.get("linked_vial_ids"))
    n_enriched = sum(1 for c in dataset["compounds"] if c.get("inv_enriched"))
    print(
        f"Built {OUT_HTML} ({len(dataset['compounds'])} method compounds [{n_linked} linked to inventory, {n_enriched} missing box locations enriched], "
        f"{len(dataset['inventory'])} inventory vials, {os.path.getsize(OUT_HTML):,} bytes)"
    )


if __name__ == "__main__":
    main()
