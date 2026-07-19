import os
import json
import re
import qrcode
import shutil
from pathlib import Path
from pptx import Presentation
from pptx.util import Cm
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor

# --- AUTOMATISCHE PFAD-ERMITTLUNG ---
DOWNLOADS_DIR = str(os.path.join(Path.home(), "Downloads"))[cite: 7]
BASE_DIR = DOWNLOADS_DIR[cite: 7]

# --- KONFIGURATION ---
OUTPUT_ROOT = os.path.join(BASE_DIR, "lve_online")[cite: 7]
JSON_INPUT = os.path.join(BASE_DIR, "kombiniert_mit_losungen.json")[cite: 7]
UNI_LOGO_PATH = os.path.join(BASE_DIR, "11_Logo_uulm_Vorlage_100mm_schwarz.png")[cite: 7]

TEMPLATE_HTML_DEFAULT = "Leere_RLB.html"[cite: 7]
TEMPLATE_HTML_FSR = "Leere_FSR_RLB.html"[cite: 7]

TEILBEREICHE = [
    "Biologie", "Chemie", "Humboldt Studienzentrum", "Informatik", 
    "Ingenieurwissenschaften", "Mathematik", "Physik", "Psychologie", 
    "Wirtschaftswissenschaften", "Zentrum für Sprachen und Philologie"
][cite: 7]

ART_MAPPING = {
    "Vorlesung": "1", "Seminar": "2", "Übung": "4", "Praktikum": "5",
    "Vorlesung/Seminar": "8", "Vorlesung/Übung": "12", "Tutorium": "31",
    "Online-Seminar": "41", "Vorlesung/Übung/Tutorium": "43", "Klinisches Praktikum": "44"
}[cite: 7]

SEMESTER = "Sommersemester 2026"[cite: 7]

INFO_SHEETS = {
    "Mathematik": os.path.join(DOWNLOADS_DIR, "Infoblatt_Mathe_SoSe26.pdf"),[cite: 7]
    "Wirtschaftswissenschaften": os.path.join(DOWNLOADS_DIR, "Infoblatt_WiWi_SoSe26.pdf"),[cite: 7]
    "Allgemein": os.path.join(DOWNLOADS_DIR, "Infoblatt_allgemein_SoSe26.pdf")[cite: 7]
}

def clean_filename(name):
    if not name: return "Unbekannt"[cite: 7]
    name = str(name).replace('ä', 'ae').replace('ö', 'oe').replace('ü', 'ue').replace('ß', 'ss')[cite: 7]
    cleaned = re.sub(r'[\\/:*?"<>|&#~,;+"!§$^°\[\]{}=]', "_", name)[cite: 7]
    cleaned = re.sub(r'_+', "_", cleaned)[cite: 7]
    return cleaned.strip("_ ")[cite: 7]

def get_last_name(full_name):
    parts = str(full_name).split()[cite: 7]
    return parts[-1] if parts else "Unbekannt"[cite: 7]

def should_add_suffix(titel, art_str):
    if not art_str:[cite: 7]
        return False[cite: 7]
    titel_clean = titel.lower()[cite: 7]
    art_lower = art_str.lower()[cite: 7]
    if art_lower in titel_clean:[cite: 7]
        return False[cite: 7]
    return True[cite: 7]

def process_html_template(template_name, output_path, data, v_art_str):
    try:
        with open(os.path.join(BASE_DIR, template_name), 'r', encoding='utf-8') as f:[cite: 7]
            content = f.read()[cite: 7]

        titel_raw = data.get("lve_titel", "")[cite: 7]
        if should_add_suffix(titel_raw, v_art_str):[cite: 7]
            titel_anzeige = f"{titel_raw} ({v_art_str})"[cite: 7]
        else:
            titel_anzeige = titel_raw[cite: 7]

        replacements = [
            ("semester", SEMESTER),[cite: 7]
            ("titel", titel_anzeige),[cite: 7]
            ("kennung", data.get("lve_kennung", "")),[cite: 7]
            ("dozent", data.get("prof_name", "")),[cite: 7]
            ("dateiname_RLB", os.path.basename(output_path))[cite: 7]
        ]
        
        for field_id, value in replacements:
            pattern = rf'(<[^>]*id="{field_id}"[^>]*value=")([^"]*)(")'[cite: 7]
            content = re.sub(pattern, rf'\1{value}\3', content)[cite: 7]

        with open(output_path, 'w', encoding='utf-8') as f:[cite: 7]
            f.write(content)[cite: 7]
    except Exception as e:
        print(f"HTML Fehler: {e}")[cite: 7]

def create_complex_ppt(output_path, qr_code_path, data, prof_name, suffix_clean):
    veranstaltung_pure = data.get("lve_titel", "Unbekannt")[cite: 7]
    
    if suffix_clean and suffix_clean.lower() in veranstaltung_pure.lower():[cite: 7]
        veranstaltung_mit_suffix = veranstaltung_pure[cite: 7]
    elif suffix_clean:
        veranstaltung_mit_suffix = f"{veranstaltung_pure} ({suffix_clean})"[cite: 7]
    else:
        veranstaltung_mit_suffix = veranstaltung_pure[cite: 7]

    lehrperson = prof_name[cite: 7]
    login_losung = data.get("losung", "")[cite: 7]
    qr_code_url = f"https://evaluation.uni-ulm.de/evasys/online.php?p={login_losung}"[cite: 7]

    prs = Presentation()[cite: 7]
    slide_layout = prs.slide_layouts[6][cite: 7]
    
    # FOLIE 1
    slide = prs.slides.add_slide(slide_layout)[cite: 7]
    text_box = slide.shapes.add_textbox(Cm(0.1), Cm(0), Cm(8), Cm(1))[cite: 7]
    text_frame = text_box.text_frame[cite: 7]
    text_frame.text = SEMESTER[cite: 7]
    text_frame.paragraphs[0].runs[0].font.bold = True[cite: 7]
    
    if os.path.exists(UNI_LOGO_PATH):[cite: 7]
        slide.shapes.add_picture(UNI_LOGO_PATH, Cm(18), Cm(0.2), width=Cm(7))[cite: 7]

    x, y, cx, cy = Cm(1.25), Cm(3), Cm(23), Cm(5)[cite: 7]
    shape = slide.shapes.add_table(4, 2, x, y, cx, cy)[cite: 7]
    table = shape.table[cite: 7]
    table.columns[0].width = Cm(7)[cite: 7]
    table.columns[1].width = Cm(16)[cite: 7]
    
    try:
        tbl = shape._element.graphic.graphicData.tbl[cite: 7]
        style_id = '{7E9639D4-E3E2-4D34-9284-5A2195B3D0D7}'[cite: 7]
        tbl[0][-1].text = style_id[cite: 7]
    except:
        pass
    
    border_color = RGBColor(125, 154, 170)[cite: 7]
    headers = ["Lehrveranstaltung:", "Lehrperson:", "URL:", "TAN/Losung:"][cite: 7]
    contents = [veranstaltung_mit_suffix, prof_name, "https://evaluation.uni-ulm.de/evasys/online", data.get("losung", "")][cite: 7]

    for row in range(4):
        cell_label = table.cell(row, 0)[cite: 7]
        cell_label.fill.solid()[cite: 7]
        cell_label.fill.fore_color.rgb = border_color[cite: 7]
        cell_label.text = headers[row][cite: 7]
        run_label = cell_label.text_frame.paragraphs[0].runs[0][cite: 7]
        run_label.font.bold = True[cite: 7]
        run_label.font.color.rgb = RGBColor(255, 255, 255)[cite: 7]
        cell_label.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT[cite: 7]
        cell_label.vertical_anchor = MSO_ANCHOR.MIDDLE[cite: 7]
        
        cell_val = table.cell(row, 1)[cite: 7]
        p = cell_val.text_frame.paragraphs[0][cite: 7]
        p.alignment = PP_ALIGN.CENTER[cite: 7]
        cell_val.vertical_anchor = MSO_ANCHOR.MIDDLE[cite: 7]
        r = p.add_run()[cite: 7]
        r.text = contents[row][cite: 7]
        if row >= 2:
            r.hyperlink.address = qr_code_url[cite: 7]

    text_box_qr = slide.shapes.add_textbox(Cm(3.5), Cm(14), Cm(8), Cm(1))[cite: 7]
    text_box_qr.text_frame.text = "QR-Code zum Scannen:"[cite: 7]
    slide.shapes.add_picture(qr_code_path, Cm(13.32), Cm(11.5), width=Cm(6), height=Cm(6))[cite: 7]
    
    # FOLIE 2
    slide2 = prs.slides.add_slide(slide_layout)[cite: 7]
    tb2 = slide2.shapes.add_textbox(Cm(0.1), Cm(0), Cm(8), Cm(1))[cite: 7]
    tb2.text_frame.text = SEMESTER[cite: 7]
    tb2.text_frame.paragraphs[0].runs[0].font.bold = True[cite: 7]
    
    tb3 = slide2.shapes.add_textbox(Cm(0.1), Cm(0.8), Cm(17), Cm(6.8))[cite: 7]
    tf3 = tb3.text_frame[cite: 7]
    tf3.word_wrap = True[cite: 7]
    p1 = tf3.add_paragraph()[cite: 7]
    p1.text = veranstaltung_mit_suffix[cite: 7]
    p2 = tf3.add_paragraph()[cite: 7]
    p2.text = f"Lehrperson: {lehrperson}"[cite: 7]
    p2.runs[0].font.bold = True[cite: 7]
    
    if os.path.exists(UNI_LOGO_PATH):[cite: 7]
        slide2.shapes.add_picture(UNI_LOGO_PATH, Cm(18), Cm(0.2), width=Cm(7))[cite: 7]
    slide2.shapes.add_picture(qr_code_path, Cm(6.3), Cm(5.6), width=Cm(12.78), height=Cm(12.78))[cite: 7]

    prs.save(output_path)[cite: 7]

def convert_to_pdf():
    try:
        import comtypes.client[cite: 7]
        powerpoint = comtypes.client.CreateObject("PowerPoint.Application")[cite: 7]
        powerpoint.Visible = 1[cite: 7]
        for root, _, files in os.walk(OUTPUT_ROOT):[cite: 7]
            for f in files:[cite: 7]
                if f.endswith(".pptx") and not f.startswith("~$"):[cite: 7]
                    path = os.path.normpath(os.path.abspath(os.path.join(root, f)))[cite: 7]
                    pdf_path = os.path.splitext(path)[0] + ".pdf"[cite: 7]
                    if not os.path.exists(pdf_path):[cite: 7]
                        pres = powerpoint.Presentations.Open(path, WithWindow=0, ReadOnly=1)[cite: 7]
                        pres.ExportAsFixedFormat(pdf_path, 2, 1, 0, 1, 1, 0, None, 1, "", True, True, True, False)[cite: 7]
                        pres.Close()[cite: 7]
        powerpoint.Quit()[cite: 7]
    except Exception as e: 
        print(f"PDF Fehler: {e}")[cite: 7]

def copy_info_sheets():
    for tb_name in os.listdir(OUTPUT_ROOT):[cite: 7]
        tb_path = os.path.join(OUTPUT_ROOT, tb_name)[cite: 7]
        if not os.path.isdir(tb_path): continue[cite: 7]
        src = INFO_SHEETS.get(tb_name, INFO_SHEETS["Allgemein"])[cite: 7]
        if os.path.exists(src):[cite: 7]
            for prof in os.listdir(tb_path):[cite: 7]
                dest = os.path.join(tb_path, prof, os.path.basename(src))[cite: 7]
                if os.path.isdir(os.path.join(tb_path, prof)):[cite: 7]
                    shutil.copy2(src, dest)[cite: 7]

def main():
    os.makedirs(OUTPUT_ROOT, exist_ok=True)[cite: 7]
    if not os.path.exists(JSON_INPUT): return[cite: 7]
    with open(JSON_INPUT, 'r', encoding='utf-8') as f:[cite: 7]
        data_struct = json.load(f)[cite: 7]

    for tb in TEILBEREICHE:[cite: 7]
        if tb not in data_struct: continue[cite: 7]
        for prof in data_struct[tb]:[cite: 7]
            prof_full = prof["prof_name"][cite: 7]
            nachname_clean = clean_filename(get_last_name(prof_full))[cite: 7]
            prof_folder = os.path.join(OUTPUT_ROOT, tb, nachname_clean)[cite: 7]
            os.makedirs(prof_folder, exist_ok=True)[cite: 7]
            
            for match in prof["matches"]:[cite: 7]
                if not match.get("losung"): continue[cite: 7]

                kennung_clean = clean_filename(match["lve_kennung"])[cite: 7]
                titel_raw = match["lve_titel"][cite: 7]
                titel_clean = clean_filename(titel_raw)[cite: 7]
                
                v_art_id = str(match["codes"]["veranstaltungsart_id"])[cite: 7]
                v_art_str = next((k for k, v in ART_MAPPING.items() if v == v_art_id), "")[cite: 7]
                v_art_clean = clean_filename(v_art_str)[cite: 7]
                
                if v_art_clean and not titel_clean.endswith(v_art_clean):[cite: 7]
                    v_art_suffix = f"_{v_art_clean}"[cite: 7]
                else:
                    v_art_suffix = ""[cite: 7]
                
                titel_gekürzt = titel_clean[:80].strip("_ ") if len(titel_clean) > 80 else titel_clean[cite: 7]
                if len(titel_clean) > 80 and v_art_clean:[cite: 7]
                    v_art_suffix = f"_{v_art_clean}"[cite: 7]

                file_base = f"{nachname_clean}_{kennung_clean}_{titel_gekürzt}{v_art_suffix}"[cite: 7]
                
                # HTML template erzeugen
                html_path = os.path.join(prof_folder, f"{nachname_clean}_RLB_{kennung_clean}_{titel_gekürzt}{v_art_suffix}.html")[cite: 7]
                process_html_template(TEMPLATE_HTML_FSR if tb in ["Mathematik", "Wirtschaftswissenschaften"] else TEMPLATE_HTML_DEFAULT, html_path, match, v_art_str)[cite: 7]
                
                # QR-Code generieren
                qr_temp = os.path.join(prof_folder, f"temp_{kennung_clean}.png")[cite: 7]
                qr = qrcode.QRCode(version=None, border=1)[cite: 7]
                qr.add_data(f"https://evaluation.uni-ulm.de/evasys/online.php?p={match['losung']}")[cite: 7]
                qr.make(fit=True)[cite: 7]
                qr.make_image().save(qr_temp)[cite: 7]
                
                # PPT erstellen
                ppt_path = os.path.join(prof_folder, f"{file_base}.pptx")[cite: 7]
                create_complex_ppt(ppt_path, qr_temp, match, prof_full, v_art_str)[cite: 7]
                if os.path.exists(qr_temp): os.remove(qr_temp)[cite: 7]

    convert_to_pdf()[cite: 7]
    copy_info_sheets()[cite: 7]

if __name__ == "__main__":
    main()[cite: 7]
