import flet as ft
from datetime import datetime, date
import os
import sys

# ── Localizar base_dir correctamente en Android / Desktop ──
if getattr(sys, 'frozen', False):
    base_dir = os.path.dirname(sys.executable)
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

# Importar lógica de negocio
from logic import (
    extract_vendors_and_defaults,
    generate_planilla_excel,
    load_dnis,
    BANCOS_VALIDOS,
)

# Colores de la app
COLOR_PRIMARY   = "#1565C0"
COLOR_SECONDARY = "#1E88E5"
COLOR_BG        = "#F0F4F8"
COLOR_CARD      = "#FFFFFF"
COLOR_DANGER    = "#E53935"
COLOR_SUCCESS   = "#2E7D32"
COLOR_TEXT      = "#1A237E"
COLOR_SUBTEXT   = "#546E7A"

LUGARES = ["SULLANA", "PAITA", "TALARA", "PIURA", "PERIFERIA", "AUTOVENTA", "OTRO"]
DIAS_SEMANA = ["LUNES", "MARTES", "MIERCOLES", "JUEVES", "VIERNES", "SABADO", "DOMINGO"]


def main(page: ft.Page):
    page.title = "Planilla de Movilidad"
    page.bgcolor = COLOR_BG
    page.padding = 0
    page.scroll = ft.ScrollMode.HIDDEN
    page.theme_mode = ft.ThemeMode.LIGHT

    # ── ESTADO GLOBAL ──
    vendors_data = extract_vendors_and_defaults()
    dnis_data    = load_dnis()

    # Filas de desplazamiento (cada elemento es un dict con controles Flet)
    trip_rows: list = []

    # ── REFERENCIAS A CONTROLES PRINCIPALES ──
    # Vendedor
    vendedor_dd    = ft.Dropdown(label="Vendedor / Trabajador", width=280,
                                  border_color=COLOR_SECONDARY,
                                  focused_border_color=COLOR_PRIMARY)
    nuevo_vendor_tf = ft.TextField(label="Nombre completo del nuevo vendedor",
                                   width=280, visible=False,
                                   border_color=COLOR_SECONDARY)
    dni_tf         = ft.TextField(label="D.N.I.", width=140, max_length=8,
                                  keyboard_type=ft.KeyboardType.NUMBER,
                                  border_color=COLOR_SECONDARY)
    banco_dd       = ft.Dropdown(label="Banco", width=160,
                                  border_color=COLOR_SECONDARY,
                                  options=[ft.dropdown.Option(b) for b in BANCOS_VALIDOS])
    periodo_tf     = ft.TextField(label="Periodo (YYYY-MM)", width=150,
                                   value=datetime.now().strftime("%Y-%m"),
                                   border_color=COLOR_SECONDARY)
    fecha_em_tf    = ft.TextField(label="Fecha Emisión (YYYY-MM-DD)", width=190,
                                   value=date.today().isoformat(),
                                   border_color=COLOR_SECONDARY)
    nro_plan_tf    = ft.TextField(label="N° Planilla Inicial", width=140,
                                   value="1",
                                   keyboard_type=ft.KeyboardType.NUMBER,
                                   border_color=COLOR_SECONDARY)
    total_text     = ft.Text("TOTAL: S/. 0.00", size=16, weight=ft.FontWeight.BOLD,
                              color=COLOR_TEXT)
    trips_col      = ft.Column(spacing=6, scroll=ft.ScrollMode.AUTO, height=350)
    status_bar     = ft.Text("", color=COLOR_SUCCESS, size=13, italic=True)

    # ── RRHH ──
    queue_files_state: list = []  # lista de paths de archivos .xlsx
    queue_col      = ft.Column(spacing=6, scroll=ft.ScrollMode.AUTO, height=300)
    rrhh_status    = ft.Text("", color=COLOR_SUCCESS, size=13, italic=True)
    file_picker    = ft.FilePicker()

    page.overlay.append(file_picker)

    # ── NAVEGACIÓN ──
    state = {"current_view": "vendedor"}

    nav_vendedor_btn = ft.TextButton("🚗  Vendedor")
    nav_rrhh_btn     = ft.TextButton("📊  Panel RRHH")

    vendedor_view = ft.Column(visible=True, expand=True)
    rrhh_view     = ft.Column(visible=False, expand=True)

    def switch_view(view_name):
        state["current_view"] = view_name
        vendedor_view.visible = (view_name == "vendedor")
        rrhh_view.visible     = (view_name == "rrhh")
        nav_vendedor_btn.style = ft.ButtonStyle(color=COLOR_PRIMARY if view_name == "vendedor" else COLOR_SUBTEXT)
        nav_rrhh_btn.style     = ft.ButtonStyle(color=COLOR_PRIMARY if view_name == "rrhh" else COLOR_SUBTEXT)
        page.update()

    nav_vendedor_btn.on_click = lambda _: switch_view("vendedor")
    nav_rrhh_btn.on_click     = lambda _: switch_view("rrhh")

    # ════════════════════════════════════════════════════════
    # LÓGICA: PANTALLA VENDEDOR
    # ════════════════════════════════════════════════════════

    def recalcular_total():
        total = 0.0
        for row in trip_rows:
            try:
                total += float(row["monto_tf"].value or 0)
            except ValueError:
                pass
        total_text.value = f"TOTAL: S/. {total:,.2f}"
        page.update()

    def build_trip_row(dia=1, motivo="Visita a clientes", ruta="", lugar="SULLANA", monto=0.0):
        """Construye una fila de desplazamiento con los valores dados."""
        periodo = periodo_tf.value.strip()
        try:
            year, month = map(int, periodo.split("-"))
        except Exception:
            year, month = datetime.now().year, datetime.now().month

        days_in_month = (date(year, month % 12 + 1, 1) - date(year, month, 1)).days if month < 12 else 31

        dia_options = [ft.dropdown.Option(str(d)) for d in range(1, days_in_month + 1)]

        dia_dd    = ft.Dropdown(options=dia_options, value=str(dia), width=65, dense=True,
                                border_color=COLOR_SECONDARY)
        diaSem_lbl = ft.Text("---", size=11, color=COLOR_SUBTEXT, width=80)
        motivo_tf  = ft.TextField(value=motivo, width=140, dense=True,
                                   border_color=COLOR_SECONDARY)
        ruta_tf    = ft.TextField(value=ruta, hint_text="Ruta", width=160, dense=True,
                                   border_color=COLOR_SECONDARY)
        lugar_dd   = ft.Dropdown(options=[ft.dropdown.Option(l) for l in LUGARES],
                                  value=lugar if lugar in LUGARES else "OTRO",
                                  width=120, dense=True, border_color=COLOR_SECONDARY)
        monto_tf   = ft.TextField(value=str(monto), width=85, dense=True,
                                   keyboard_type=ft.KeyboardType.NUMBER,
                                   border_color=COLOR_SECONDARY)

        row_ref = {}

        def update_dia_label(e=None):
            try:
                d = int(dia_dd.value)
                d_obj = date(year, month, d)
                diaSem_lbl.value = DIAS_SEMANA[d_obj.weekday()]
            except Exception:
                diaSem_lbl.value = "---"
            page.update()

        dia_dd.on_change     = update_dia_label
        monto_tf.on_change   = lambda e: recalcular_total()
        update_dia_label()

        def delete_row(e):
            trip_rows.remove(row_ref)
            trips_col.controls.remove(row_container)
            recalcular_total()
            page.update()

        del_btn = ft.IconButton(icon=ft.icons.DELETE_OUTLINE, icon_color=COLOR_DANGER,
                                 tooltip="Eliminar fila", on_click=delete_row)

        row_container = ft.Container(
            content=ft.Row(
                controls=[dia_dd, diaSem_lbl, motivo_tf, ruta_tf, lugar_dd, monto_tf, del_btn],
                spacing=4, vertical_alignment=ft.CrossAxisAlignment.CENTER
            ),
            bgcolor=COLOR_CARD,
            border_radius=8,
            padding=ft.Padding.symmetric(horizontal=8, vertical=4),
            border=ft.border.all(1, "#E3EAF2"),
        )

        row_ref.update({
            "dia_dd": dia_dd,
            "diaSem_lbl": diaSem_lbl,
            "motivo_tf": motivo_tf,
            "ruta_tf": ruta_tf,
            "lugar_dd": lugar_dd,
            "monto_tf": monto_tf,
            "container": row_container,
            "year": year,
            "month": month,
        })
        return row_ref, row_container

    def cargar_rutas_predeterminadas(e=None):
        """Carga las rutas del vendedor seleccionado según el periodo."""
        vendor = vendedor_dd.value
        if not vendor or vendor == "NUEVO":
            return

        periodo = periodo_tf.value.strip()
        try:
            year, month = map(int, periodo.split("-"))
        except Exception:
            return

        trip_rows.clear()
        trips_col.controls.clear()

        defaults = vendors_data.get(vendor, {}).get("defaults", {})
        dnis_info = dnis_data.get(vendor, {})

        if isinstance(dnis_info, dict):
            dni_tf.value   = dnis_info.get("dni", "")
            banco_dd.value = dnis_info.get("banco", vendors_data.get(vendor, {}).get("banco", ""))
        else:
            dni_tf.value = str(dnis_info)

        # Generar todos los días hábiles del mes
        days_in_month = (date(year, month % 12 + 1, 1) - date(year, month, 1)).days if month < 12 else 31
        for day in range(1, days_in_month + 1):
            try:
                d_obj = date(year, month, day)
            except ValueError:
                continue
            weekday_name = DIAS_SEMANA[d_obj.weekday()]
            if weekday_name == "DOMINGO":
                continue
            default = defaults.get(weekday_name, {})
            row_ref, row_ctrl = build_trip_row(
                dia=day,
                motivo="Visita a clientes",
                ruta=default.get("ruta", ""),
                lugar=default.get("lugar", "SULLANA"),
                monto=default.get("monto", 0.0),
            )
            trip_rows.append(row_ref)
            trips_col.controls.append(row_ctrl)

        recalcular_total()
        page.update()

    def on_vendedor_change(e):
        v = vendedor_dd.value
        nuevo_vendor_tf.visible = (v == "NUEVO")
        if v != "NUEVO":
            cargar_rutas_predeterminadas()
        page.update()

    def agregar_fila(e):
        row_ref, row_ctrl = build_trip_row()
        trip_rows.append(row_ref)
        trips_col.controls.append(row_ctrl)
        recalcular_total()
        page.update()

    def generar_planilla(e):
        vendor = nuevo_vendor_tf.value.strip() if vendedor_dd.value == "NUEVO" else (vendedor_dd.value or "").strip()
        dni    = dni_tf.value.strip()
        banco  = banco_dd.value or ""
        periodo = periodo_tf.value.strip()
        fecha_em = fecha_em_tf.value.strip()
        nro_plan = nro_plan_tf.value.strip()

        if not vendor:
            status_bar.value = "⚠ Seleccione o escriba un vendedor."
            status_bar.color = COLOR_DANGER
            page.update()
            return
        if len(dni) != 8 or not dni.isdigit():
            status_bar.value = "⚠ El DNI debe tener exactamente 8 dígitos."
            status_bar.color = COLOR_DANGER
            page.update()
            return
        if not banco:
            status_bar.value = "⚠ Seleccione un banco."
            status_bar.color = COLOR_DANGER
            page.update()
            return
        if not trip_rows:
            status_bar.value = "⚠ No hay filas de desplazamiento."
            status_bar.color = COLOR_DANGER
            page.update()
            return

        rows = []
        for row in trip_rows:
            try:
                yr  = row["year"]
                mo  = row["month"]
                dia = int(row["dia_dd"].value)
                rows.append({
                    "dia":    dia,
                    "mes":    mo,
                    "ano":    yr,
                    "motivo": row["motivo_tf"].value or "Visita a clientes",
                    "ruta":   row["ruta_tf"].value or "",
                    "lugar":  row["lugar_dd"].value or "SULLANA",
                    "monto":  float(row["monto_tf"].value or 0),
                })
            except Exception:
                pass

        try:
            filepath = generate_planilla_excel(
                vendor, dni, banco, periodo, fecha_em, int(nro_plan), rows
            )
            status_bar.value = f"✔ Planilla guardada en Descargas:\n{os.path.basename(filepath)}"
            status_bar.color = COLOR_SUCCESS
        except Exception as ex:
            status_bar.value = f"✘ Error: {str(ex)}"
            status_bar.color = COLOR_DANGER

        page.update()

    # Poblar dropdown de vendedores
    vendedor_dd.options = [ft.dropdown.Option("NUEVO", "-- NUEVO VENDEDOR --")]
    for v in sorted(vendors_data.keys()):
        vendedor_dd.options.append(ft.dropdown.Option(v))
    vendedor_dd.on_change = on_vendedor_change
    periodo_tf.on_change  = cargar_rutas_predeterminadas

    # ════════════════════════════════════════════════════════
    # LÓGICA: PANTALLA RRHH
    # ════════════════════════════════════════════════════════

    def on_files_picked(e: ft.FilePickerResultEvent):
        if not e.files:
            return
        for f in e.files:
            if f.path and f.path not in queue_files_state:
                queue_files_state.append(f.path)
        refresh_queue_display()
        page.update()

    file_picker.on_result = on_files_picked

    def refresh_queue_display():
        queue_col.controls.clear()
        if not queue_files_state:
            queue_col.controls.append(
                ft.Text("No hay archivos en cola. Selecciona planillas para comenzar.",
                        color=COLOR_SUBTEXT, italic=True, size=13)
            )
        else:
            for path in queue_files_state:
                name = os.path.basename(path)
                def make_remove(p):
                    def remove(e):
                        queue_files_state.remove(p)
                        refresh_queue_display()
                        page.update()
                    return remove
                queue_col.controls.append(
                    ft.Container(
                        content=ft.Row([
                            ft.Icon(ft.icons.INSERT_DRIVE_FILE_OUTLINED, color=COLOR_SECONDARY, size=18),
                            ft.Text(name, expand=True, size=13),
                            ft.IconButton(icon=ft.icons.CLOSE, icon_color=COLOR_DANGER,
                                          tooltip="Quitar", on_click=make_remove(path))
                        ]),
                        bgcolor=COLOR_CARD,
                        border_radius=8,
                        padding=ft.Padding.symmetric(horizontal=8, vertical=4),
                        border=ft.border.all(1, "#E3EAF2"),
                    )
                )
        page.update()

    def consolidar(e):
        if not queue_files_state:
            rrhh_status.value = "⚠ No hay archivos en cola."
            rrhh_status.color = COLOR_DANGER
            page.update()
            return

        from logic import consolidate_planillas_excel
        try:
            filepath, logs = consolidate_planillas_excel(queue_files_state)
            rrhh_status.value = f"✔ Consolidado guardado en Descargas:\n{os.path.basename(filepath)}\n" + "\n".join(logs)
            rrhh_status.color = COLOR_SUCCESS
            queue_files_state.clear()
            refresh_queue_display()
        except Exception as ex:
            rrhh_status.value = f"✘ Error: {str(ex)}"
            rrhh_status.color = COLOR_DANGER
        page.update()

    refresh_queue_display()

    # ════════════════════════════════════════════════════════
    # CONSTRUCCIÓN DE VISTAS
    # ════════════════════════════════════════════════════════

    # ── Encabezado ──
    header = ft.Container(
        content=ft.Column([
            ft.Row([
                ft.Text("🚗  Planilla de Movilidad", size=22, weight=ft.FontWeight.BOLD, color="white"),
            ]),
            ft.Text("Corporación de Alimentos S.A.C.", size=12, color="#BBDEFB"),
            ft.Row([nav_vendedor_btn, nav_rrhh_btn]),
        ], spacing=4),
        bgcolor=COLOR_PRIMARY,
        padding=ft.Padding.symmetric(horizontal=20, vertical=14),
    )

    # ── VISTA VENDEDOR ──
    vendedor_view.controls = [
        ft.Container(
            content=ft.Column([
                ft.Text("Datos Generales", size=16, weight=ft.FontWeight.BOLD, color=COLOR_TEXT),
                ft.Divider(height=2, color="#E3EAF2"),
                ft.Row([vendedor_dd, nuevo_vendor_tf], wrap=True, spacing=8),
                ft.Row([dni_tf, banco_dd, nro_plan_tf], wrap=True, spacing=8),
                ft.Row([periodo_tf, fecha_em_tf], wrap=True, spacing=8),
            ], spacing=10),
            bgcolor=COLOR_CARD,
            border_radius=12,
            padding=16,
            margin=ft.Margin.symmetric(horizontal=12, vertical=8),
            shadow=ft.BoxShadow(blur_radius=8, color="#1A237E22"),
        ),
        ft.Container(
            content=ft.Column([
                ft.Row([
                    ft.Text("Desplazamientos y Pasajes", size=16,
                            weight=ft.FontWeight.BOLD, color=COLOR_TEXT),
                    ft.Button("+ Agregar Fila", on_click=agregar_fila,
                               bgcolor=COLOR_SECONDARY, color="white",
                               style=ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=8))),
                ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                ft.Row([
                    ft.Text("Día", size=11, color=COLOR_SUBTEXT, width=65),
                    ft.Text("Día Sem.", size=11, color=COLOR_SUBTEXT, width=80),
                    ft.Text("Motivo", size=11, color=COLOR_SUBTEXT, width=140),
                    ft.Text("Ruta", size=11, color=COLOR_SUBTEXT, width=160),
                    ft.Text("Lugar", size=11, color=COLOR_SUBTEXT, width=120),
                    ft.Text("Monto", size=11, color=COLOR_SUBTEXT, width=85),
                ], spacing=4),
                ft.Divider(height=1, color="#E3EAF2"),
                trips_col,
                ft.Divider(height=1, color="#E3EAF2"),
                total_text,
            ], spacing=8),
            bgcolor=COLOR_CARD,
            border_radius=12,
            padding=16,
            margin=ft.Margin.symmetric(horizontal=12, vertical=4),
            shadow=ft.BoxShadow(blur_radius=8, color="#1A237E22"),
        ),
        ft.Container(
            content=ft.Column([
                ft.Button(
                    "📥  Generar y Guardar Planilla",
                    on_click=generar_planilla,
                    bgcolor=COLOR_PRIMARY,
                    color="white",
                    width=320,
                    height=52,
                    style=ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=10)),
                ),
                status_bar,
            ], horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=10),
            margin=ft.Margin.symmetric(horizontal=12, vertical=8),
            alignment=ft.alignment.Alignment.CENTER,
        ),
    ]

    # ── VISTA RRHH ──
    rrhh_view.controls = [
        ft.Container(
            content=ft.Column([
                ft.Text("Subir Planillas de Vendedores", size=16,
                        weight=ft.FontWeight.BOLD, color=COLOR_TEXT),
                ft.Text("Selecciona los archivos Excel (.xlsx) generados por los vendedores.",
                        size=12, color=COLOR_SUBTEXT),
                ft.Button(
                    "📂  Seleccionar Archivos .xlsx",
                    on_click=lambda _: file_picker.pick_files(
                        allow_multiple=True,
                        allowed_extensions=["xlsx"],
                    ),
                    bgcolor=COLOR_SECONDARY,
                    color="white",
                    style=ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=8)),
                ),
            ], spacing=10),
            bgcolor=COLOR_CARD, border_radius=12, padding=16,
            margin=ft.Margin.symmetric(horizontal=12, vertical=8),
            shadow=ft.BoxShadow(blur_radius=8, color="#1A237E22"),
        ),
        ft.Container(
            content=ft.Column([
                ft.Text("Cola de Procesamiento", size=16,
                        weight=ft.FontWeight.BOLD, color=COLOR_TEXT),
                ft.Divider(height=1, color="#E3EAF2"),
                queue_col,
                ft.Divider(height=1, color="#E3EAF2"),
                ft.Row([
                    ft.Button(
                        "🗑  Vaciar Cola",
                        on_click=lambda _: [queue_files_state.clear(), refresh_queue_display()],
                        bgcolor="#EF9A9A", color="white",
                        style=ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=8)),
                    ),
                    ft.Button(
                        "📥  Generar y Descargar Consolidado",
                        on_click=consolidar,
                        bgcolor=COLOR_PRIMARY, color="white",
                        style=ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=8)),
                    ),
                ], spacing=10, wrap=True),
                rrhh_status,
            ], spacing=10),
            bgcolor=COLOR_CARD, border_radius=12, padding=16,
            margin=ft.Margin.symmetric(horizontal=12, vertical=4),
            shadow=ft.BoxShadow(blur_radius=8, color="#1A237E22"),
        ),
    ]

    # ── PÁGINA PRINCIPAL ──
    page.add(
        header,
        ft.Container(
            content=ft.Stack([vendedor_view, rrhh_view]),
            expand=True,
        )
    )

    switch_view("vendedor")


if __name__ == "__main__":
    ft.run(main)
