from datetime import date
from io import BytesIO
from fpdf import FPDF
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
import streamlit as st

st.set_page_config(
    page_title="Medición de Vivienda", page_icon="📐", layout="centered"
)


def texto_pdf_seguro(texto: str) -> str:
    """Evita errores de codificación en FPDF con tildes, ñ o símbolos."""
    return str(texto).encode("latin-1", "replace").decode("latin-1")


def parsear_decimal(valor_str: str) -> float:
    """Convierte texto con coma o punto decimal a número flotante."""
    if not valor_str or not valor_str.strip():
        return 0.0
    try:
        return float(valor_str.strip().replace(",", "."))
    except ValueError:
        return -1.0


def generar_pdf_bytes(
    estancias: list,
    tot_paredes: float,
    tot_suelos: float,
    tot_techos: float,
    cliente: str,
    direccion: str,
    fecha_str: str,
) -> bytes:
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # --- Cabecera personalizada ---
    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(
        0,
        10,
        texto_pdf_seguro("Informe de Mediciones y Superficies de Vivienda"),
        ln=True,
    )

    pdf.set_font("Helvetica", "", 10)
    if cliente.strip():
        pdf.cell(0, 6, texto_pdf_seguro(f"Cliente: {cliente.strip()}"), ln=True)
    if direccion.strip():
        pdf.cell(
            0, 6, texto_pdf_seguro(f"Obra / Dirección: {direccion.strip()}"), ln=True
        )
    pdf.cell(0, 6, texto_pdf_seguro(f"Fecha: {fecha_str}"), ln=True)
    pdf.ln(5)

    # --- Tabla ---
    pdf.set_font("Helvetica", "B", 9)
    anchos = (42, 22, 22, 22, 28, 27, 27)
    headers = (
        "Estancia",
        "Largo",
        "Ancho",
        "Altura",
        "Paredes",
        "Suelo",
        "Techo",
    )
    for h, w in zip(headers, anchos):
        pdf.cell(w, 8, texto_pdf_seguro(h), border=1, align="C")
    pdf.ln()

    pdf.set_font("Helvetica", "", 9)
    for e in estancias:
        fila = (
            texto_pdf_seguro(e["Estancia"]),
            f"{e['Largo (m)']:.2f} m".replace(".", ","),
            f"{e['Ancho (m)']:.2f} m".replace(".", ","),
            f"{e['Altura (m)']:.2f} m".replace(".", ","),
            f"{e['Paredes (m²)']:.2f} m2".replace(".", ","),
            f"{e['Suelo (m²)']:.2f} m2".replace(".", ","),
            f"{e['Techo (m²)']:.2f} m2".replace(".", ","),
        )
        for idx, (val, w) in enumerate(zip(fila, anchos)):
            pdf.cell(w, 7, val, border=1, align="L" if idx == 0 else "R")
        pdf.ln()

    pdf.set_font("Helvetica", "B", 9)
    pdf.cell(sum(anchos[:4]), 8, "TOTALES GENERALES", border=1, align="R")
    pdf.cell(
        anchos[4], 8, f"{tot_paredes:.2f} m2".replace(".", ","), border=1, align="R"
    )
    pdf.cell(
        anchos[5], 8, f"{tot_suelos:.2f} m2".replace(".", ","), border=1, align="R"
    )
    pdf.cell(
        anchos[6], 8, f"{tot_techos:.2f} m2".replace(".", ","), border=1, align="R"
    )

    return bytes(pdf.output())


def generar_excel_bytes(
    estancias: list, cliente: str, direccion: str, fecha_str: str
) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Mediciones"

    # --- Cabecera personalizada en Excel ---
    ws["A1"] = "INFORME DE MEDICIONES Y SUPERFICIES"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = f"Cliente: {cliente.strip() or 'No especificado'}"
    ws["A3"] = f"Dirección: {direccion.strip() or 'No especificada'}"
    ws["A4"] = f"Fecha: {fecha_str}"

    fila_inicio_tabla = 6
    headers = [
        "Estancia",
        "Largo (m)",
        "Ancho (m)",
        "Altura (m)",
        "Paredes (m²)",
        "Suelo (m²)",
        "Techo (m²)",
    ]

    fill_header = PatternFill(
        start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
    )
    for col_idx, text in enumerate(headers, start=1):
        cell = ws.cell(row=fila_inicio_tabla, column=col_idx, value=text)
        cell.font = Font(bold=True)
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center")

    # --- Filas de datos con fórmulas dinámicas ---
    for i, e in enumerate(estancias, start=fila_inicio_tabla + 1):
        ws.cell(row=i, column=1, value=e["Estancia"])
        ws.cell(row=i, column=2, value=e["Largo (m)"]).number_format = "0.00"
        ws.cell(row=i, column=3, value=e["Ancho (m)"]).number_format = "0.00"
        ws.cell(row=i, column=4, value=e["Altura (m)"]).number_format = "0.00"

        # Fórmulas dinámicas por fila: Paredes = (Largo + Ancho)*2*Altura | Suelo/Techo = Largo*Ancho
        c_paredes = ws.cell(row=i, column=5, value=f"=(B{i}+C{i})*2*D{i}")
        c_suelo = ws.cell(row=i, column=6, value=f"=B{i}*C{i}")
        c_techo = ws.cell(row=i, column=7, value=f"=B{i}*C{i}")

        for c in (c_paredes, c_suelo, c_techo):
            c.number_format = "0.00"

    # --- Fila de Totales con =SUM(...) ---
    fila_primera_dato = fila_inicio_tabla + 1
    fila_ultima_dato = fila_inicio_tabla + len(estancias)
    fila_total = fila_ultima_dato + 1

    ws.cell(row=fila_total, column=1, value="TOTALES GENERALES").font = Font(
        bold=True
    )
    for col_idx, col_letra in [(5, "E"), (6, "F"), (7, "G")]:
        cell = ws.cell(
            row=fila_total,
            column=col_idx,
            value=f"=SUM({col_letra}{fila_primera_dato}:{col_letra}{fila_ultima_dato})",
        )
        cell.font = Font(bold=True)
        cell.number_format = "0.00"

    # Ajuste automático de ancho de columnas
    for col in ws.columns:
        max_len = max(
            len(str(cell.value))
            if cell.value is not None and not str(cell.value).startswith("=")
            else 12
            for cell in col
        )
        ws.column_dimensions[col.column_letter].width = max(max_len + 4, 14)

    output = BytesIO()
    wb.save(output)
    return output.getvalue()


# --- INTERFAZ ---
st.title("📐 Medición de Vivienda")

if "estancias" not in st.session_state:
    st.session_state.estancias = []

# Bloque de datos de cabecera
with st.expander("📋 Datos del proyecto / cabecera (Opcional)", expanded=False):
    c_cli, c_fec = st.columns(2)
    with c_cli:
        cliente = st.text_input("Nombre del cliente", placeholder="Ej: Juan Pérez")
    with c_fec:
        fecha_proyecto = st.date_input("Fecha", value=date.today())
    direccion = st.text_input(
        "Dirección de la vivienda", placeholder="Ej: Calle Mayor 12, 3ºB"
    )

fecha_formateada = fecha_proyecto.strftime("%d/%m/%Y")

with st.form("form_estancia", clear_on_submit=True):
    st.subheader("Añadir nueva estancia")

    opciones = (
        "Hall",
        "Pasillo 1",
        "Pasillo 2",
        "Cocina",
        "Baño",
        "Aseo",
        "Despensa",
        "Salón",
        "Dormitorio 1",
        "Dormitorio 2",
        "Dormitorio 3",
        "Terraza",
        "Galería",
        "Otra (escribir abajo)",
    )
    seleccion = st.selectbox("Estancia", opciones)
    nombre_custom = st.text_input("Nombre personalizado (si elegiste 'Otra')")

    col1, col2, col3 = st.columns(3)
    with col1:
        largo_txt = st.text_input("Largo (m)", value="", placeholder="Ej: 4,25")
    with col2:
        ancho_txt = st.text_input("Ancho (m)", value="", placeholder="Ej: 3,10")
    with col3:
        altura_txt = st.text_input("Altura (m)", value="2,60", placeholder="Ej: 2,60")

    enviado = st.form_submit_button("➕ Añadir estancia", use_container_width=True)

    if enviado:
        largo = parsear_decimal(largo_txt)
        ancho = parsear_decimal(ancho_txt)
        altura = parsear_decimal(altura_txt)

        if seleccion == "Otra (escribir abajo)" and not nombre_custom.strip():
            st.error("Escribe un nombre personalizado para la estancia.")
        elif largo < 0 or ancho < 0 or altura < 0:
            st.error("Introduce números válidos (usa coma o punto para decimales).")
        elif largo == 0 or ancho == 0 or altura == 0:
            st.error("Introduce medidas mayores que 0 en largo, ancho y altura.")
        else:
            nombre_final = (
                nombre_custom.strip()
                if seleccion == "Otra (escribir abajo)"
                else seleccion
            )
            paredes = (largo + ancho) * 2 * altura
            area_planta = largo * ancho

            st.session_state.estancias.append({
                "Estancia": nombre_final,
                "Largo (m)": largo,
                "Ancho (m)": ancho,
                "Altura (m)": altura,
                "Paredes (m²)": round(paredes, 2),
                "Suelo (m²)": round(area_planta, 2),
                "Techo (m²)": round(area_planta, 2),
            })
            st.success(f"Añadida: {nombre_final}")

if st.session_state.estancias:
    st.subheader("Resumen de estancias medidas")
    st.dataframe(st.session_state.estancias, use_container_width=True)

    tot_paredes = sum(e["Paredes (m²)"] for e in st.session_state.estancias)
    tot_suelos = sum(e["Suelo (m²)"] for e in st.session_state.estancias)
    tot_techos = sum(e["Techo (m²)"] for e in st.session_state.estancias)

    c1, c2, c3 = st.columns(3)
    c1.metric("Total Paredes", f"{tot_paredes:.2f} m²".replace(".", ","))
    c2.metric("Total Suelos", f"{tot_suelos:.2f} m²".replace(".", ","))
    c3.metric("Total Techos", f"{tot_techos:.2f} m²".replace(".", ","))

    st.divider()

    col_pdf, col_xlsx = st.columns(2)
    with col_pdf:
        st.download_button(
            label="📄 Descargar PDF (DIN A4)",
            data=generar_pdf_bytes(
                st.session_state.estancias,
                tot_paredes,
                tot_suelos,
                tot_techos,
                cliente,
                direccion,
                fecha_formateada,
            ),
            file_name="informe_calculo_vivienda.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
    with col_xlsx:
        st.download_button(
            label="📊 Descargar Excel (.xlsx)",
            data=generar_excel_bytes(
                st.session_state.estancias, cliente, direccion, fecha_formateada
            ),
            file_name="informe_calculo_vivienda.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    col_del1, col_del2 = st.columns(2)
    with col_del1:
        if st.button("↩️ Borrar última estancia", use_container_width=True):
            st.session_state.estancias.pop()
            st.rerun()
    with col_del2:
        if st.button("🗑️ Vaciar todas las estancias", use_container_width=True):
            st.session_state.estancias.clear()
            st.rerun()
