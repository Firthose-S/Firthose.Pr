"""
Unified Water Quality Calculator
--------------------------------
Calculates/records:
- Total solids (TS), total dissolved solids (TDS), suspended solids (TSS)
- Chloride concentration
- Total hardness + automatic hardness classification
- pH + automatic pH classification
- Metals and other chemicals
- Nitrogen and nitrogen compounds
- Dissolved gases

The program exports a report to Excel and PDF.

Dependencies:
    pip install openpyxl reportlab
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Optional
import math

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer


# ----------------------------
# Calculation functions
# ----------------------------

def total_solids_from_evaporation(
    dish_plus_residue_g: float,
    empty_dish_g: float,
    sample_volume_ml: float,
) -> float:
    """Total solids, mg/L, from gravimetric evaporation."""
    if sample_volume_ml <= 0:
        raise ValueError("Sample volume must be > 0 mL.")
    return (dish_plus_residue_g - empty_dish_g) * 1_000_000 / sample_volume_ml


def tds_from_evaporation(
    dish_plus_residue_g: float,
    empty_dish_g: float,
    sample_volume_ml: float,
) -> float:
    """TDS, mg/L, from gravimetric evaporation of a filtered sample."""
    return total_solids_from_evaporation(
        dish_plus_residue_g, empty_dish_g, sample_volume_ml
    )


def tss_from_ts_tds(ts_mg_l: float, tds_mg_l: float) -> float:
    """Approximate TSS = TS - TDS."""
    return max(0.0, ts_mg_l - tds_mg_l)


def chloride_from_agno3(
    agno3_molarity: float,
    sample_volume_ml: float,
    titrant_volume_ml: float,
    blank_volume_ml: float = 0.0,
) -> float:
    """
    Chloride as mg/L Cl- using AgNO3 titration.
    Ag+ reacts 1:1 with Cl-.
    """
    if sample_volume_ml <= 0:
        raise ValueError("Sample volume must be > 0 mL.")
    net_titrant_ml = titrant_volume_ml - blank_volume_ml
    if net_titrant_ml < 0:
        raise ValueError("Titrant volume cannot be less than blank volume.")
    return (
        net_titrant_ml / 1000.0
        * agno3_molarity
        * 35.453
        * 1000.0
        / (sample_volume_ml / 1000.0)
    )


def hardness_from_edta(
    edta_molarity: float,
    sample_volume_ml: float,
    edta_volume_ml: float,
    blank_volume_ml: float = 0.0,
) -> float:
    """
    Total hardness as mg/L CaCO3 using EDTA titration.
    Assumes 1 mol EDTA equivalent to 1 mol CaCO3 for this calculation.
    """
    if sample_volume_ml <= 0:
        raise ValueError("Sample volume must be > 0 mL.")
    net_ml = edta_volume_ml - blank_volume_ml
    if net_ml < 0:
        raise ValueError("EDTA volume cannot be less than blank volume.")
    return (
        net_ml / 1000.0
        * edta_molarity
        * 100.0869
        * 1000.0
        / (sample_volume_ml / 1000.0)
    )


def classify_hardness(mg_l_as_caco3: float) -> str:
    """
    Common classification:
      0-60     Soft
      >60-120  Moderately hard
      >120-180 Hard
      >180     Very hard
    """
    if mg_l_as_caco3 < 0:
        return "Invalid"
    if mg_l_as_caco3 <= 60:
        return "Soft"
    if mg_l_as_caco3 <= 120:
        return "Moderately hard"
    if mg_l_as_caco3 <= 180:
        return "Hard"
    return "Very hard"


def classify_ph(ph: float) -> str:
    """Simple pH-range classification."""
    if not 0 <= ph <= 14:
        return "Invalid pH"
    if ph < 6.5:
        return "Acidic"
    if ph <= 8.5:
        return "Near-neutral / acceptable range"
    return "Alkaline"


def ph_category_detail(ph: float) -> str:
    if ph < 0 or ph > 14:
        return "Outside normal pH scale"
    if ph < 3:
        return "Strongly acidic"
    if ph < 6.5:
        return "Acidic"
    if ph < 7:
        return "Slightly acidic"
    if ph == 7:
        return "Neutral"
    if ph <= 8.5:
        return "Slightly alkaline"
    if ph <= 11:
        return "Alkaline"
    return "Strongly alkaline"


# ----------------------------
# Report data helpers
# ----------------------------

def make_parameter(name, value, unit, method, category=""):
    return {
        "Parameter": name,
        "Value": value,
        "Unit": unit,
        "Method": method,
        "Classification": category,
    }


def evaluate_water_sample(data: dict) -> list[dict]:
    """Calculate all derived parameters from a single input dictionary."""

    results = []

    ts = total_solids_from_evaporation(
        data["ts_dish_plus_residue_g"],
        data["ts_empty_dish_g"],
        data["ts_sample_volume_ml"],
    )
    tds = tds_from_evaporation(
        data["tds_dish_plus_residue_g"],
        data["tds_empty_dish_g"],
        data["tds_sample_volume_ml"],
    )
    tss = tss_from_ts_tds(ts, tds)

    chloride = chloride_from_agno3(
        data["agno3_molarity"],
        data["chloride_sample_volume_ml"],
        data["agno3_volume_ml"],
        data["chloride_blank_ml"],
    )

    hardness = hardness_from_edta(
        data["edta_molarity"],
        data["hardness_sample_volume_ml"],
        data["edta_volume_ml"],
        data["hardness_blank_ml"],
    )

    results.extend([
        make_parameter("Total Solids (TS)", round(ts, 2), "mg/L", "Gravimetric evaporation"),
        make_parameter("Total Dissolved Solids (TDS)", round(tds, 2), "mg/L", "Gravimetric evaporation"),
        make_parameter("Total Suspended Solids (TSS)", round(tss, 2), "mg/L", "TS - TDS"),
        make_parameter("Chloride (as Cl-)", round(chloride, 2), "mg/L", "AgNO3 titration"),
        make_parameter(
            "Total Hardness (as CaCO3)",
            round(hardness, 2),
            "mg/L",
            "EDTA titration",
            classify_hardness(hardness),
        ),
        make_parameter(
            "pH",
            round(data["ph"], 2),
            "pH units",
            "pH meter",
            ph_category_detail(data["ph"]),
        ),
    ])

    # User-entered laboratory measurements for metals/chemicals
    for name, value in data.get("metals_chemicals", {}).items():
        results.append(make_parameter(name, value, "mg/L", "Laboratory measurement"))

    # Nitrogen species
    for name, value in data.get("nitrogen_compounds", {}).items():
        results.append(make_parameter(name, value, "mg/L", "Laboratory measurement"))

    # Dissolved gases
    for name, value in data.get("dissolved_gases", {}).items():
        results.append(make_parameter(name, value, "mg/L", "Laboratory measurement"))

    return results


# ----------------------------
# Excel export
# ----------------------------

def export_excel(results: list[dict], sample_name: str, output_path: str | Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Water Quality Report"

    title = f"Unified Water Quality Report - {sample_name}"
    ws["A1"] = title
    ws["A1"].font = Font(bold=True, size=16)
    ws.merge_cells("A1:E1")

    ws["A2"] = "Generated"
    ws["B2"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    headers = ["Parameter", "Value", "Unit", "Method", "Classification"]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col, value=header)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9EAF7")
        cell.alignment = Alignment(horizontal="center")

    for row_idx, item in enumerate(results, 5):
        for col_idx, key in enumerate(headers, 1):
            ws.cell(row=row_idx, column=col_idx, value=item.get(key, ""))

    widths = [34, 16, 15, 30, 35]
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:E{4 + len(results)}"

    notes = wb.create_sheet("Notes")
    notes["A1"] = "Important notes"
    notes["A1"].font = Font(bold=True, size=14)
    notes["A3"] = (
        "This calculator performs calculations and categorization; it does not "
        "replace a laboratory test or a regulatory compliance determination."
    )
    notes["A4"] = (
        "Metals, chemicals, nitrogen compounds and dissolved gases are entered "
        "as laboratory-measured concentrations because different analytical "
        "methods require different equations."
    )
    notes["A5"] = (
        "Hardness categories used: Soft <=60; Moderately hard 60-120; "
        "Hard 120-180; Very hard >180 mg/L as CaCO3."
    )
    notes["A6"] = (
        "pH categories: acidic <6.5; near-neutral/acceptable 6.5-8.5; "
        "alkaline >8.5. Actual drinking-water limits depend on the applicable standard."
    )
    notes.column_dimensions["A"].width = 120
    notes["A3"].alignment = Alignment(wrap_text=True)
    notes["A4"].alignment = Alignment(wrap_text=True)
    notes["A5"].alignment = Alignment(wrap_text=True)
    notes["A6"].alignment = Alignment(wrap_text=True)

    wb.save(output_path)


# ----------------------------
# PDF export
# ----------------------------

def export_pdf(results: list[dict], sample_name: str, output_path: str | Path):
    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        rightMargin=12 * mm,
        leftMargin=12 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
    )
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="Small", parent=styles["BodyText"], fontSize=7.5, leading=9))

    story = [
        Paragraph(f"Unified Water Quality Report", styles["Title"]),
        Paragraph(f"Sample: {sample_name}", styles["Heading2"]),
        Paragraph(f"Generated: {datetime.now():%Y-%m-%d %H:%M:%S}", styles["BodyText"]),
        Spacer(1, 8),
    ]

    table_data = [["Parameter", "Value", "Unit", "Method", "Classification"]]
    for r in results:
        table_data.append([
            Paragraph(str(r["Parameter"]), styles["Small"]),
            Paragraph(str(r["Value"]), styles["Small"]),
            Paragraph(str(r["Unit"]), styles["Small"]),
            Paragraph(str(r["Method"]), styles["Small"]),
            Paragraph(str(r["Classification"]), styles["Small"]),
        ])

    table = Table(table_data, colWidths=[48*mm, 22*mm, 22*mm, 48*mm, 38*mm], repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(table)
    story.append(Spacer(1, 10))
    story.append(Paragraph(
        "<b>Disclaimer:</b> Results are calculations/classifications based on the "
        "input data and stated methods. They should be interpreted against the "
        "applicable laboratory QA/QC procedures and regulatory standard.",
        styles["Small"],
    ))
    doc.build(story)


# ----------------------------
# Example / command-line use
# ----------------------------

def example_data():
    return {
        "ts_empty_dish_g": 50.000,
        "ts_dish_plus_residue_g": 50.125,
        "ts_sample_volume_ml": 100.0,

        "tds_empty_dish_g": 45.000,
        "tds_dish_plus_residue_g": 45.090,
        "tds_sample_volume_ml": 100.0,

        "agno3_molarity": 0.0141,
        "chloride_sample_volume_ml": 100.0,
        "agno3_volume_ml": 12.5,
        "chloride_blank_ml": 0.2,

        "edta_molarity": 0.0100,
        "hardness_sample_volume_ml": 50.0,
        "edta_volume_ml": 8.0,
        "hardness_blank_ml": 0.1,

        "ph": 7.4,

        "metals_chemicals": {
            "Calcium (Ca)": 32.0,
            "Magnesium (Mg)": 8.0,
            "Iron (Fe)": 0.10,
            "Manganese (Mn)": 0.03,
            "Lead (Pb)": 0.002,
            "Fluoride (F-)": 0.70,
            "Sulphate (SO4 2-)": 45.0,
        },

        "nitrogen_compounds": {
            "Ammonia-N (NH3-N)": 0.08,
            "Nitrite-N (NO2-N)": 0.01,
            "Nitrate-N (NO3-N)": 4.2,
        },

        "dissolved_gases": {
            "Dissolved Oxygen (DO)": 7.1,
            "Carbon Dioxide (CO2)": 4.0,
            "Hydrogen Sulfide (H2S)": 0.00,
        },
    }


def main():
    sample_name = "Example Water Sample"
    data = example_data()
    results = evaluate_water_sample(data)

    out_dir = Path(".")
    excel_file = out_dir / "water_quality_report.xlsx"
    pdf_file = out_dir / "water_quality_report.pdf"

    export_excel(results, sample_name, excel_file)
    export_pdf(results, sample_name, pdf_file)

    print("Unified Water Quality Calculator")
    print("--------------------------------")
    for r in results:
        classification = f" | {r['Classification']}" if r["Classification"] else ""
        print(f"{r['Parameter']}: {r['Value']} {r['Unit']}{classification}")

    print(f"\nExcel report: {excel_file.resolve()}")
    print(f"PDF report:   {pdf_file.resolve()}")


if __name__ == "__main__":
    main()
