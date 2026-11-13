import json
import os
import re
import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_TAB_ALIGNMENT


def sanitize_filename(text: str) -> str:
    """Replaces characters that are invalid in file names (like / or :) with hyphens."""
    return re.sub(r'[\\/*?:"<>|]', "-", text.strip())


def fill_evaluation_grid(
    template_path: str,
    nom: str,
    date_str: str,
    sujet: str,
    output_dir: str = "."
) -> str:
    """Fills a single Word document from the template and saves it as {Date}_{Nom}.docx."""
    doc = docx.Document(template_path)

    for p in doc.paragraphs:
        # Fill Paragraph containing "Nom :" and "Date :"
        if "Nom" in p.text and "Date" in p.text:
            p.clear()
            p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.LEFT
            
            # Set fixed tab stops so "Date :" and "Note :" stay aligned regardless of name length
            tab_stops = p.paragraph_format.tab_stops
            tab_stops.clear_all()
            tab_stops.add_tab_stop(Cm(9.5), WD_TAB_ALIGNMENT.LEFT)
            tab_stops.add_tab_stop(Cm(15.0), WD_TAB_ALIGNMENT.LEFT)

            r_nom_lbl = p.add_run("Nom\xa0: ")
            r_nom_lbl.bold = True
            r_nom_lbl.font.size = Pt(10)

            r_nom_val = p.add_run(f"{nom}\t")
            r_nom_val.bold = False
            r_nom_val.font.size = Pt(10)

            r_date_lbl = p.add_run("Date\xa0: ")
            r_date_lbl.bold = True
            r_date_lbl.font.size = Pt(10)

            r_date_val = p.add_run(f"{date_str}\t")
            r_date_val.bold = False
            r_date_val.font.size = Pt(10)

            r_note_lbl = p.add_run("Note\xa0: ")
            r_note_lbl.bold = True
            r_note_lbl.font.size = Pt(10)

        # Fill Paragraph containing "Sujet :"
        elif p.text.strip().startswith("Sujet"):
            p.clear()
            r_sujet_lbl = p.add_run("Sujet\xa0: ")
            r_sujet_lbl.bold = True
            r_sujet_lbl.font.size = Pt(10)

            r_sujet_val = p.add_run(sujet)
            r_sujet_val.bold = False
            r_sujet_val.font.size = Pt(10)

    os.makedirs(output_dir, exist_ok=True)
    safe_date = sanitize_filename(date_str)
    safe_nom = sanitize_filename(nom)
    output_filename = os.path.join(output_dir, f"{safe_date}_{safe_nom}.docx")
    doc.save(output_filename)
    return output_filename


def main(
    json_path: str = "colles_data.json",
    template_path: str = "Grille evaluation colles eco.docx",
    output_dir: str = "."
):
    with open(json_path, "r", encoding="utf-8") as f:
        config = json.load(f)

    # Support both a top-level dict with defaults or a flat list of records
    if isinstance(config, dict):
        default_date = config.get("default_date", "")
        default_sujet = config.get("default_sujet", "")
        records = config.get("etudiants", [])
    elif isinstance(config, list):
        default_date = ""
        default_sujet = ""
        records = config
    else:
        raise ValueError("Invalid JSON format.")

    for entry in records:
        nom = entry.get("nom") or entry.get("Nom", "")
        date_str = entry.get("date") or entry.get("Date") or default_date
        sujet = entry.get("sujet") or entry.get("Sujet") or default_sujet

        out_path = fill_evaluation_grid(template_path, nom, date_str, sujet, output_dir)
        print(f"Generated: {out_path}")


if __name__ == "__main__":
    main()
