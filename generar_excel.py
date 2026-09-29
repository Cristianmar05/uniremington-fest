import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

wb = openpyxl.Workbook()
wb.remove(wb.active)  # Eliminar hoja por defecto vacía

# Lista oficial de eventos con ÚNICAMENTE los 6 campos requeridos
events = [
    {
        "tab_name": "01_Apertura",
        "title": "Apertura Uniremington Fest",
        "fecha": "21 de septiembre",
        "hora": "2:00 PM",
        "lugar": "Auditorio",
        "descripcion": "Acto inaugural de la Semana de Aniversario, bienvenida institucional y muestras culturales de apertura.",
        "participacion": "Obligatoria"
    },
    {
        "tab_name": "02_Talentos",
        "title": "Talentos Uniremington",
        "fecha": "22 de septiembre",
        "hora": "2:00 PM",
        "lugar": "Auditorio",
        "descripcion": "Presentaciones artísticas, musicales y dancísticas de la comunidad universitaria.",
        "participacion": "Obligatoria"
    },
    {
        "tab_name": "03_Juegos",
        "title": "Juegos – Cultura General",
        "fecha": "23 de septiembre",
        "hora": "2:00 PM",
        "lugar": "Auditorio",
        "descripcion": "Jornada lúdica de trivias de conocimiento general, competencias académicas y retos en equipo.",
        "participacion": "Opcional"
    },
    {
        "tab_name": "04_Declamadores",
        "title": "Declamadores Invitados",
        "fecha": "24 de septiembre",
        "hora": "2:00 PM",
        "lugar": "Auditorio",
        "descripcion": "Recital poético y encuentro literario con declamadores invitados especiales del ámbito cultural.",
        "participacion": "Opcional"
    },
    {
        "tab_name": "05_Distinciones",
        "title": "Entrega de Distinciones",
        "fecha": "25 de septiembre",
        "hora": "2:00 PM",
        "lugar": "Auditorio",
        "descripcion": "Ceremonia oficial de reconocimiento a la excelencia académica, mérito investigativo y egresados destacados.",
        "participacion": "Opcional"
    },
    {
        "tab_name": "06_Familia",
        "title": "Día de la Familia Uniremington y Feria Empresarial",
        "fecha": "26 de septiembre",
        "hora": "2:00 PM",
        "lugar": "Auditorio",
        "descripcion": "Feria de emprendimientos estudiantiles, muestras comerciales y actividades recreativas para la familia.",
        "participacion": "Opcional"
    }
]

# Estilos visuales institucionales
font_title = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
font_subtitle = Font(name="Calibri", size=10, italic=True, color="E0E8F5")
fill_navy = PatternFill(start_color="0B2F64",
                        end_color="0B2F64", fill_type="solid")
fill_header = PatternFill(start_color="1A4A8B",
                          end_color="1A4A8B", fill_type="solid")
font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

fill_obligatoria = PatternFill(
    start_color="FDE8E8", end_color="FDE8E8", fill_type="solid")
font_obligatoria = Font(name="Calibri", size=10, bold=True, color="9C1C1C")
fill_opcional = PatternFill(
    start_color="E6F4EA", end_color="E6F4EA", fill_type="solid")
font_opcional = Font(name="Calibri", size=10, bold=True, color="137333")

fill_even = PatternFill(start_color="F8FAFD",
                        end_color="F8FAFD", fill_type="solid")
border_thin = Border(
    left=Side(style='thin', color='D3D3D3'),
    right=Side(style='thin', color='D3D3D3'),
    top=Side(style='thin', color='D3D3D3'),
    bottom=Side(style='thin', color='D3D3D3')
)

# -------------------------------------------------------------
# 1. PESTAÑA: Programación General (Tabla Resumen)
# -------------------------------------------------------------
ws_gen = wb.create_sheet(title="Programacion General")
ws_gen.views.sheetView[0].showGridLines = True

# Banner superior
ws_gen.merge_cells("A1:F1")
ws_gen["A1"] = "SEMANA DE ANIVERSARIO UNIREMINGTON - PROGRAMACION DE EVENTOS 2026"
ws_gen["A1"].font = font_title
ws_gen["A1"].fill = fill_navy
ws_gen["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws_gen.row_dimensions[1].height = 30

ws_gen.merge_cells("A2:F2")
ws_gen["A2"] = "Haga clic en el nombre del evento para ir a su ficha tecnica detallada"
ws_gen["A2"].font = font_subtitle
ws_gen["A2"].fill = fill_navy
ws_gen["A2"].alignment = Alignment(horizontal="center", vertical="center")
ws_gen.row_dimensions[2].height = 20

# Encabezados de tabla
headers = ["Evento", "Fecha", "Hora", "Lugar", "Descripción", "Participación"]
ws_gen.row_dimensions[3].height = 25
for col_idx, h in enumerate(headers, 1):
    c = ws_gen.cell(row=3, column=col_idx, value=h)
    c.font = font_header
    c.fill = fill_header
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = border_thin

# Filas de datos
for idx, ev in enumerate(events, 4):
    ws_gen.row_dimensions[idx].height = 35
    row_fill = fill_even if idx % 2 == 0 else PatternFill(fill_type=None)

    # Evento con hipervínculo a su pestaña
    c_ev = ws_gen.cell(row=idx, column=1, value=ev["title"])
    c_ev.hyperlink = f"#'{ev['tab_name']}'!A1"
    c_ev.font = Font(name="Calibri", size=10, bold=True,
                     color="0B2F64", underline="single")
    c_ev.alignment = Alignment(horizontal="left", vertical="center")
    c_ev.border = border_thin
    c_ev.fill = row_fill

    # Fecha
    c_f = ws_gen.cell(row=idx, column=2, value=ev["fecha"])
    c_f.font = Font(name="Calibri", size=10)
    c_f.alignment = Alignment(horizontal="center", vertical="center")
    c_f.border = border_thin
    c_f.fill = row_fill

    # Hora
    c_h = ws_gen.cell(row=idx, column=3, value=ev["hora"])
    c_h.font = Font(name="Calibri", size=10)
    c_h.alignment = Alignment(horizontal="center", vertical="center")
    c_h.border = border_thin
    c_h.fill = row_fill

    # Lugar
    c_l = ws_gen.cell(row=idx, column=4, value=ev["lugar"])
    c_l.font = Font(name="Calibri", size=10)
    c_l.alignment = Alignment(horizontal="center", vertical="center")
    c_l.border = border_thin
    c_l.fill = row_fill

    # Descripción
    c_d = ws_gen.cell(row=idx, column=5, value=ev["descripcion"])
    c_d.font = Font(name="Calibri", size=9)
    c_d.alignment = Alignment(
        horizontal="left", vertical="center", wrap_text=True)
    c_d.border = border_thin
    c_d.fill = row_fill

    # Participación
    c_p = ws_gen.cell(row=idx, column=6, value=ev["participacion"])
    c_p.alignment = Alignment(horizontal="center", vertical="center")
    c_p.border = border_thin
    if ev["participacion"] == "Obligatoria":
        c_p.fill = fill_obligatoria
        c_p.font = font_obligatoria
    else:
        c_p.fill = fill_opcional
        c_p.font = font_opcional

ws_gen.column_dimensions["A"].width = 30
ws_gen.column_dimensions["B"].width = 20
ws_gen.column_dimensions["C"].width = 14
ws_gen.column_dimensions["D"].width = 18
ws_gen.column_dimensions["E"].width = 55
ws_gen.column_dimensions["F"].width = 16

# -------------------------------------------------------------
# 2. PESTAÑAS INDIVIDUALES: Ficha con ÚNICAMENTE los 6 campos
# -------------------------------------------------------------
for ev in events:
    ws = wb.create_sheet(title=ev["tab_name"])
    ws.views.sheetView[0].showGridLines = True

    # Encabezado institucional
    ws.merge_cells("A1:D1")
    ws["A1"] = "FICHA TECNICA DEL EVENTO"
    ws["A1"].font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    ws["A1"].fill = fill_navy
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 24

    ws.merge_cells("A2:D2")
    ws["A2"] = ev["title"]
    ws["A2"].font = Font(name="Calibri", size=14, bold=True, color="0B2F64")
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 28

    ws.merge_cells("A3:D3")
    ws["A3"] = "Semana de Aniversario Uniremington"
    ws["A3"].font = Font(name="Calibri", size=10, italic=True, color="666666")
    ws["A3"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[3].height = 18

    ws.row_dimensions[4].height = 10  # Separador

    # ÚNICAMENTE los 6 campos solicitados
    ficha_rows = [
        ("Evento", ev["title"]),
        ("Fecha", ev["fecha"]),
        ("Hora", ev["hora"]),
        ("Lugar", ev["lugar"]),
        ("Descripción", ev["descripcion"]),
        ("Participación", ev["participacion"])
    ]

    current_row = 5
    for label, val in ficha_rows:
        # Columna A: Etiqueta del campo
        c_lbl = ws.cell(row=current_row, column=1, value=label)
        c_lbl.font = Font(name="Calibri", size=10, bold=True, color="1A4A8B")
        c_lbl.fill = PatternFill(start_color="F0F4FA",
                                 end_color="F0F4FA", fill_type="solid")
        c_lbl.alignment = Alignment(horizontal="left", vertical="center")
        c_lbl.border = border_thin

        # Columnas B a D combinadas para el contenido
        ws.merge_cells(start_row=current_row, start_column=2,
                       end_row=current_row, end_column=4)
        c_val = ws.cell(row=current_row, column=2, value=val)
        c_val.font = Font(name="Calibri", size=10)
        c_val.alignment = Alignment(
            horizontal="left", vertical="center", wrap_text=True)
        c_val.border = border_thin

        for c in range(2, 5):
            ws.cell(row=current_row, column=c).border = border_thin

        if label == "Participación":
            if val == "Obligatoria":
                c_val.fill = fill_obligatoria
                c_val.font = font_obligatoria
            else:
                c_val.fill = fill_opcional
                c_val.font = font_opcional

        if label == "Descripción":
            ws.row_dimensions[current_row].height = 45
        else:
            ws.row_dimensions[current_row].height = 25

        current_row += 1

    ws.row_dimensions[current_row].height = 12
    current_row += 1

    # Botón interactivo de retorno
    ws.merge_cells(start_row=current_row, start_column=1,
                   end_row=current_row, end_column=4)
    btn = ws.cell(row=current_row, column=1,
                  value="⬅ Volver a Programacion General")
    btn.hyperlink = "#'Programacion General'!A1"
    btn.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    btn.fill = PatternFill(start_color="0B2F64",
                           end_color="0B2F64", fill_type="solid")
    btn.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[current_row].height = 26

    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 25
    ws.column_dimensions["C"].width = 25
    ws.column_dimensions["D"].width = 30

wb.save("eventos_uniremington.xlsx")
print("✅ Archivo 'eventos_uniremington.xlsx' generado con éxito con ÚNICAMENTE los 6 campos.")
