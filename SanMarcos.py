import pdfplumber
import openpyxl
from openpyxl.styles import Border, Side
from openpyxl.utils import get_column_letter
import unicodedata
import streamlit as st
import re
import io

# --- HELPER FUNCTIONS ---
def normalize_text(text):
    if not text: return ""
    nfd = unicodedata.normalize('NFD', str(text))
    return ''.join(char for char in nfd if unicodedata.category(char) != 'Mn').lower()

def squish_text(text):
    """Aggressively removes ALL spaces, punctuation, hyphens, and hidden characters for a 100% reliable match."""
    if not text: return ""
    t = normalize_text(text)
    return re.sub(r'[^a-z0-9]', '', t)

def safe_float(val):
    if val is None: return 0.0
    s = str(val).strip()
    if not s or s == '-': return 0.0
    s = s.replace(',', '')
    s = re.sub(r'[^\d\.\-]', '', s)
    if s.count('.') > 1:
        parts = s.rsplit('.', 1)
        s = parts[0].replace('.', '') + '.' + parts[1]
    try: return float(s)
    except ValueError: return 0.0

def clean_currency(value):
    if not value: return 0.0
    raw = str(value).strip().replace(' ', '')
    raw = re.sub(r'[^\d\.,]', '', raw)
    if not raw: return 0.0

    if re.search(r',\d{1,2}$', raw):
        parts = raw.rsplit(',', 1)
        raw = parts[0].replace('.', '').replace(',', '') + '.' + parts[1]
    else:
        raw = raw.replace(',', '')

    if raw.count('.') > 1:
        parts = raw.rsplit('.', 1)
        raw = parts[0].replace('.', '') + '.' + parts[1]

    try: return float(raw)
    except ValueError: return 0.0

def extract_value_from_row(row_list, total_idx):
    if total_idx != -1 and len(row_list) > total_idx:
        val = clean_currency(row_list[total_idx])
        if val > 0: return val
    for item in reversed(row_list):
        val = clean_currency(item)
        if val > 0: return val
    return 0.0

def get_master_cell(ws, r_idx, c_idx):
    cell = ws.cell(row=r_idx, column=c_idx)
    if type(cell).__name__ == 'MergedCell':
        for m_range in ws.merged_cells.ranges:
            if cell.coordinate in m_range:
                return ws.cell(row=m_range.min_row, column=m_range.min_col)
    return cell

# --- TRUCO CSS PARA TRADUCIR LA INTERFAZ A ESPAÑOL ---
st.markdown("""
    <style>
        div[data-testid="stFileUploader"] label p {
            font-size: 40px !important;
        }
    </style>
""", unsafe_allow_html=True)

# --- WEB UI ---
st.title("🇬🇹 MAGA: Procesador de Facturas por la LAE: San Marcos")
uploaded_pdfs = st.file_uploader(label='1. Seleccione sus Facturas (PDFs)', type='pdf', accept_multiple_files=True)
uploaded_xlsx = st.file_uploader(label='2. Seleccione su Archivo de Excel', type='xlsx')

if st.button("INICIAR PROCESO") and uploaded_pdfs and uploaded_xlsx:
    try:
        input_buffer = io.BytesIO(uploaded_xlsx.read())
        wb = openpyxl.load_workbook(input_buffer)
        ws = wb.active

        if "Extra Detalles" not in wb.sheetnames:
            ws_det = wb.create_sheet("Extra Detalles")
            ws_det.append(['Nombre Emisor', 'NIT Emisor', 'NIT Receptor',
                           'Nombre Receptor', 'Num. DTE', 'Municipio'])
        else:
            ws_det = wb["Extra Detalles"]

        # Collect DTEs already recorded in "Extra Detalles" (Num. DTE is column 5)
        # so we can skip invoices that were already processed in a prior run.
        existing_dtes = set()
        for row in ws_det.iter_rows(min_row=2, min_col=5, max_col=5, values_only=True):
            if row[0] is not None:
                existing_dtes.add(str(row[0]).strip())

        # 1. Map Excel Columns dynamically
        col_map = {}
        for row in ws.iter_rows(min_row=1, max_row=15):
            for cell in row:
                if type(cell).__name__ == 'MergedCell': continue
                if not cell.value: continue
                val = normalize_text(str(cell.value))

                if 'agricultura' in val: col_map['agri'] = cell.column
                if 'escuela' in val or 'establecimiento' in val: col_map['escuelas'] = cell.column
                if 'proveedor' in val or 'productor' in val:
                    base_col, base_row, found_total = cell.column, cell.row, False
                    for r_offset in range(1, 4):
                        for c_offset in range(3):
                            sub_cell = ws.cell(row=base_row + r_offset, column=base_col + c_offset)
                            if sub_cell.value and 'total' in normalize_text(str(sub_cell.value)):
                                col_map['productores'] = sub_cell.column
                                found_total = True
                                break
                        if found_total: break
                    if 'productores' not in col_map: col_map['productores'] = base_col

        if 'agri' not in col_map:
            st.error("No encontré la columna de Agricultura en el Excel.")
            st.stop()

        department_name = 'san marcos'
        # 2. MASTER MUNICIPALITY DICTIONARY (San Marcos department)
        MUNICIPIOS = {
            1:  {"nombre_oficial": "Ayutla",                        "alias_pdf": ["ayutla"]},
            2:  {"nombre_oficial": "Catarina",                      "alias_pdf": ["catarina"]},
            3:  {"nombre_oficial": "Comitancillo",                  "alias_pdf": ["comitancillo"]},
            4:  {"nombre_oficial": "Concepción Tutuapa",            "alias_pdf": ["concepcion tutuapa"]},
            5:  {"nombre_oficial": "El Quetzal",                    "alias_pdf": ["el quetzal"]},
            6:  {"nombre_oficial": "El Tumbador",                   "alias_pdf": ["el tumbador"]},
            7:  {"nombre_oficial": "Esquipulas Palo Gordo",         "alias_pdf": ["esquipulas palo gordo"]},
            8:  {"nombre_oficial": "Ixchiguán",                     "alias_pdf": ["ixchiguan"]},
            9:  {"nombre_oficial": "La Blanca",                     "alias_pdf": ["la blanca"]},
            10: {"nombre_oficial": "La Reforma",                    "alias_pdf": ["la reforma"]},
            11: {"nombre_oficial": "Malacatán",                     "alias_pdf": ["malacatan"]},
            12: {"nombre_oficial": "Nuevo Progreso",                "alias_pdf": ["nuevo progreso"]},
            13: {"nombre_oficial": "Ocós",                          "alias_pdf": ["ocos"]},
            14: {"nombre_oficial": "Pajapita",                      "alias_pdf": ["pajapita"]},
            15: {"nombre_oficial": "Río Blanco",                    "alias_pdf": ["rio blanco"]},
            16: {"nombre_oficial": "San Antonio Sacatepéquez",      "alias_pdf": ["san antonio sacatepequez"]},
            17: {"nombre_oficial": "San Cristóbal Cucho",           "alias_pdf": ["san cristobal cucho"]},
            18: {"nombre_oficial": "San José El Rodeo",             "alias_pdf": ["san jose el rodeo", "el rodeo"]},
            19: {"nombre_oficial": "San José Ojetenam",             "alias_pdf": ["san jose ojetenam"]},
            20: {"nombre_oficial": "San Lorenzo",                   "alias_pdf": ["san lorenzo"]},
            21: {"nombre_oficial": "San Marcos",                    "alias_pdf": ["san marcos, san marcos", "san marcos san marcos"]},
            22: {"nombre_oficial": "San Miguel Ixtahuacán",         "alias_pdf": ["san miguel ixtahuacan"]},
            23: {"nombre_oficial": "San Pablo",                     "alias_pdf": ["san pablo"]},
            24: {"nombre_oficial": "San Pedro Sacatepéquez",        "alias_pdf": ["san pedro sacatepequez"]},
            25: {"nombre_oficial": "San Rafael Pie De La Cuesta",   "alias_pdf": ["san rafael pie de la cuesta", "san rafael"]},
            26: {"nombre_oficial": "Sibinal",                       "alias_pdf": ["sibinal"]},
            27: {"nombre_oficial": "Sipacapa",                      "alias_pdf": ["sipacapa"]},
            28: {"nombre_oficial": "Tacaná",                        "alias_pdf": ["tacana"]},
            29: {"nombre_oficial": "Tajumulco",                     "alias_pdf": ["tajumulco"]},
            30: {"nombre_oficial": "Tejutla",                       "alias_pdf": ["tejutla"]},
        }

        search_list = []
        for m_id, data in MUNICIPIOS.items():
            aliases = data.get("alias_pdf", [data["nombre_oficial"]])
            for alias in aliases:
                if alias:  # skip empties defensively
                    search_list.append((alias, m_id, data["nombre_oficial"]))

        # Sort so the department capital (ambiguous "San Marcos") is evaluated last;
        # within the rest, longer aliases first so specific names win.
        search_list.sort(key=lambda x: (
            squish_text(x[2]) == squish_text(department_name),
            -len(x[0])
        ))

        EXCEL_MAPPINGS = {
            1: "ayutla", 2: "catarina", 3: "comitancillo", 4: "concepcion tutuapa",
            5: "el quetzal", 6: "el tumbador", 7: "esquipulas palo gordo", 8: "ixchiguan",
            9: "la blanca", 10: "la reforma", 11: "malacatan", 12: "nuevo progreso",
            13: "ocos", 14: "pajapita", 15: "rio blanco", 16: "san antonio sacatepequez",
            17: "san cristobal cucho", 18: "san jose el rodeo", 19: "san jose ojetenam",
            20: "san lorenzo", 21: "san marcos", 22: "san miguel ixtahuacan", 23: "san pablo",
            24: "san pedro sacatepequez", 25: "san rafael pie de la cuesta", 26: "sibinal",
            27: "sipacapa", 28: "tacana", 29: "tajumulco", 30: "tejutla",
        }

        # 3. Map Excel Rows to Municipalities
        row_map = {}
        for row_ex in ws.iter_rows(min_row=5, max_row=150):
            row_text = " ".join([str(c.value) for c in row_ex if c.value and type(c).__name__ != 'MergedCell'])
            row_squished = squish_text(row_text)
            for m_id, search_key in EXCEL_MAPPINGS.items():
                if m_id in row_map: continue
                key_squished = squish_text(search_key)
                if key_squished in row_squished:
                    row_map[m_id] = row_ex[0].row

        batch_totals = {m_id: {'agri': 0.0, 'emisores': set(), 'receptores': set()} for m_id in MUNICIPIOS.keys()}
        new_count = 0
        progress_bar = st.progress(0)

        # 4. Process each PDF
        for i, pdf_file in enumerate(uploaded_pdfs):
            with pdfplumber.open(pdf_file) as pdf:
                text = "".join([p.extract_text() or "" for p in pdf.pages])
                tables = []
                for p in pdf.pages:
                    t = p.extract_table()
                    if t: tables.extend(t)

                dte_m = re.search(r'N[úu]mero\s*de\s*DTE:\s*(\d+)', text, re.IGNORECASE)
                dte_val = dte_m.group(1) if dte_m else pdf_file.name

                # Duplicate-DTE guard: skip invoices whose DTE already exists in "Extra Detalles"
                # (either from a prior run or from an earlier PDF in this same batch).
                if str(dte_val).strip() in existing_dtes:
                    st.warning(f"Esta factura (Num. DTE: {dte_val}) no ha sido agregado al archivo de Excel: ya ha sido procesado")
                    progress_bar.progress((i + 1) / len(uploaded_pdfs))
                    continue

                text_squished = squish_text(text)
                m_id, m_name = None, "N/A"

                for alias, mun_id, official_name in search_list:
                    alias_squished = squish_text(alias)
                    if alias_squished and alias_squished in text_squished:
                        m_id = mun_id
                        m_name = official_name
                        break

                if m_id:
                    agri_sum = 0.0

                    # Find the Total (Q) column index
                    total_col_idx = -1
                    for row_tbl in tables:
                        if not row_tbl: continue
                        for idx, cell in enumerate(row_tbl):
                            if not cell: continue
                            cell_norm = normalize_text(str(cell))
                            if 'total' in cell_norm and 'descuento' not in cell_norm and '(q)' in cell_norm:
                                total_col_idx = idx
                                break
                        if total_col_idx != -1:
                            break

                    # Sum every line-item row's total
                    for row_tbl in tables:
                        if not row_tbl: continue

                        row_text = " ".join([str(x) for x in row_tbl if x])
                        row_text_normalized = normalize_text(row_text)

                        skip_keywords = ['totales', 'superintendencia', 'datos del certificador',
                                         'contribuyendo', 'sujeto a pagos', 'no genera derecho',
                                         'descripcion', 'cantidad', 'unitario', 'descuentos', 'impuestos']
                        if any(keyword in row_text_normalized for keyword in skip_keywords):
                            continue

                        # Line-item rows start with an item number like "1", "2", "1."
                        if not (row_tbl and row_tbl[0] and re.match(r'^\s*\d+\.?\s*$', str(row_tbl[0]).strip())):
                            continue

                        val = extract_value_from_row(row_tbl, total_col_idx)
                        if val <= 0:
                            continue

                        agri_sum += val

                    nit_e_match = re.search(r'Emisor:\s*([0-9Kk\-]+)', text, re.I)
                    nit_r_match = re.search(r'Receptor:\s*([0-9Kk\-]+)', text, re.I)
                    name_e_match = re.search(r'(?:Factura(?:\s*Pequeño\s*Contribuyente)?)\s*\n+(.*?)\n+Nit\s*Emisor', text, re.IGNORECASE | re.DOTALL)

                    nit_e = nit_e_match.group(1).strip() if nit_e_match else "N/A"
                    nit_r = nit_r_match.group(1).strip() if nit_r_match else "N/A"
                    raw_name = re.sub(r'\s+', ' ', name_e_match.group(1).strip() if name_e_match else "N/A")
                    name_e = re.split(r'(?i)n[úu]mero\s*de\s*autorizaci[óo]n', raw_name)[0]
                    name_e = re.split(r'(?i)\bserie\b', name_e)[0].strip()

                    # Receptor (school) name: captured from "Nombre Receptor:" up to "Dirección comprador:",
                    # with the interleaved "Fecha y hora de certificación: DD-MMM-YYYY HH:MM:SS" stripped out.
                    name_r_match = re.search(
                        r'Nombre\s*Receptor:\s*(.*?)(?=Direcci[oó]n\s*comprador:)',
                        text, re.IGNORECASE | re.DOTALL)
                    if name_r_match:
                        raw_r = re.sub(
                            r'\s*Fecha\s*y\s*hora\s*de\s*certificaci[oó]n:\s*\d{1,2}-[A-Za-zÁÉÍÓÚáéíóúñ]+-\d{4}\s+\d{1,2}:\d{2}:\d{2}\s*',
                            ' ', name_r_match.group(1), flags=re.IGNORECASE)
                        name_r = re.sub(r'\s+', ' ', raw_r).strip().rstrip(',').strip()
                    else:
                        name_r = "N/A"

                    batch_totals[m_id]['agri'] += agri_sum
                    if nit_e != "N/A": batch_totals[m_id]['emisores'].add(nit_e)
                    if nit_r != "N/A": batch_totals[m_id]['receptores'].add(nit_r)

                    ws_det.append([name_e, nit_e, nit_r, name_r, dte_val, m_name])
                    existing_dtes.add(str(dte_val).strip())
                    new_count += 1
                else:
                    st.warning(f"No se pudo identificar el municipio en la factura: {pdf_file.name}")

            progress_bar.progress((i + 1) / len(uploaded_pdfs))

        # 5. Write totals to the main sheet
        for target_m_id, r_idx in row_map.items():
            data = batch_totals.get(target_m_id)
            if not data: continue

            if 'agri' in col_map and data['agri'] > 0:
                target_cell = get_master_cell(ws, r_idx, col_map['agri'])
                target_cell.value = safe_float(target_cell.value) + data['agri']

            if 'escuelas' in col_map and len(data['receptores']) > 0:
                target_cell = get_master_cell(ws, r_idx, col_map['escuelas'])
                target_cell.value = int(safe_float(target_cell.value)) + len(data['receptores'])

            if 'productores' in col_map and len(data['emisores']) > 0:
                target_cell = get_master_cell(ws, r_idx, col_map['productores'])
                target_cell.value = int(safe_float(target_cell.value)) + len(data['emisores'])

        # 6. Format "Extra Detalles"
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        for col in ws_det.columns:
            max_length = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                cell.border = thin_border
                try: max_length = max(max_length, len(str(cell.value)))
                except: pass
            ws_det.column_dimensions[col_letter].width = max_length + 2

        # 7. Final Export
        output = io.BytesIO()
        wb.save(output)

        st.success(f"¡Proceso completado! {new_count} facturas procesadas y agregadas al Excel con éxito.")
        output.seek(0)
        st.download_button("Descargar Reporte Final", data=output.getvalue(),
                           file_name="Reporte_MAGA_Actualizado.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Error crítico detectado: {e}")
