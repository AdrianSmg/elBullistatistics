# ============================================================
# Nombre del proyecto: elBullistatistics
# Archivo: app.py
# Descripción: Aplicación en Streamlit para generar informes estadísticos 
#              para elBulli1846 a partir de distintos ficheros de datos
#              generados por el sistema de reservas Clorian.
# Fecha de creación: Julio - Agosto 2025
# ============================================================

# --------------------------
# Librerías
# --------------------------

import streamlit as st
import pandas as pd
import altair as alt
import locale
import datetime
import tempfile
import time
from pdf import build_pdf
from babel.dates import format_date, format_datetime

# --------------------------
# Funciones
# --------------------------

def makePivot(df: pd.DataFrame, index_col: str, value_col: str, aggfunc: str = "sum", label_fmt: str | None = None) -> pd.DataFrame:
    
    pivot = (
        df.groupby(index_col, dropna=False, as_index=False)[value_col]
          .agg(aggfunc)
          .sort_values(index_col)
          .reset_index(drop=True)
    )

    def formatLabel(v):

        if label_fmt and pd.notnull(v):
            try:
                return v.strftime(label_fmt)
            except Exception:
                pass
        return "" if pd.isna(v) else str(v)

    pivot["label"] = pivot[index_col].apply(formatLabel)
    return pivot

def renderBlockWithTable(pivot_df: pd.DataFrame,
                            label_col: str,
                            value_col: str,
                            label_title: str,
                            secondlabel_title: str | None = "Nº PAX",
                            color: str = "#2db1fc",
                            height: int | None = None,
                            show_percent_labels: bool = True,
                            table_width_ratio=(2.5, 1.5),
                            y_order: list[str] | None = None,
                            _row_px: int = 26,
                            _min_h: int = 420,
                            _pad_px: int = 40,
                            value_format: str | None = None,
                            chart_type: str = "bar",
                            temporal_col: str | None = None
                            ) -> None:

    df = pivot_df.copy()
    total = float(df[value_col].sum())
    df["pct"] = (df[value_col] / total * 100).round(1) if total > 0 else 0.0
    df["Porcentaje"] = df["pct"].map(lambda x: f"{x:.1f}%")

    if y_order:
        df[label_col] = pd.Categorical(df[label_col], categories=y_order, ordered=True)
        df = df.sort_values(label_col).reset_index(drop=True)

    if height is None:
        nRows = len(df)
        height = max(_min_h, nRows * _row_px + _pad_px)

    if chart_type == "line":
        x_field = temporal_col if (temporal_col and temporal_col in df.columns and pd.api.types.is_datetime64_any_dtype(df[temporal_col])) else None
        if x_field is None and pd.api.types.is_datetime64_any_dtype(df.get(label_col, pd.Series([], dtype="datetime64[ns]"))):
            x_field = label_col
        if x_field is None:
            df["_parsed_time"] = pd.to_datetime(df[label_col], dayfirst=True, errors="coerce")
            x_field = "_parsed_time"

        chart = (
            alt.Chart(df)
            .mark_line(point=True, color=color)
            .encode(
                x=alt.X(f"{x_field}:T",
                        title=label_title,
                        axis=alt.Axis(format="%d/%m/%y")),
                y=alt.Y(f"{value_col}:Q", title=secondlabel_title),
                tooltip=[
                    alt.Tooltip(f"{x_field}:T", title=label_title, format="%d/%m/%Y"),
                    alt.Tooltip(f"{value_col}:Q", title=secondlabel_title),
                    alt.Tooltip("Porcentaje:N", title="Porcentaje"),
                ],
            )
            .properties(height=height)
        )
    else:
        if y_order:
            y_enc = alt.Y(
                f"{label_col}:N",
                title="",
                sort=y_order,
                scale=alt.Scale(domain=y_order, paddingInner=0.3, paddingOuter=0.15),
            )
        else:
            y_enc = alt.Y(
                f"{label_col}:N",
                title="",
                sort=None,
                scale=alt.Scale(paddingInner=0.3, paddingOuter=0.15),
            )

        bars = (
            alt.Chart(df)
            .mark_bar(size=18, color=color)
            .encode(
                x=alt.X(f"{value_col}:Q", title="Pax"),
                y=y_enc,
                tooltip=[
                    alt.Tooltip(f"{label_col}:N", title=label_title),
                    alt.Tooltip(f"{value_col}:Q", title=secondlabel_title),
                    alt.Tooltip("Porcentaje:N", title="Porcentaje"),
                ],
            )
            .properties(height=height)
        )

        chart = bars
        if show_percent_labels:
            labels = (
                alt.Chart(df)
                .mark_text(align="left", baseline="middle", dx=6)
                .encode(
                    x=f"{value_col}:Q",
                    y=f"{label_col}:N",
                    text="Porcentaje:N",
                )
            )
            chart = bars + labels

    table_df = df[[label_col, value_col]].copy()
    table_df.columns = [label_title, secondlabel_title]

    if value_format == "euro" and pd.api.types.is_numeric_dtype(table_df[secondlabel_title]):
        table_df[secondlabel_title] = table_df[secondlabel_title].apply(
            lambda x: f"{x:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")
        )

    with st.container(border=True):
        col1, col2 = st.columns(table_width_ratio)
        with col1:
            st.altair_chart(chart, use_container_width=True)
        with col2:
            st.dataframe(table_df, use_container_width=True, height=height, hide_index=True)

def infoBox(label: str, value, label_border_color="#2db1fc", value_bg_color="#cde8c1"):

    col1, col2 = st.columns([2, 1])
    with col1:
        st.markdown(
            f'<div style="border: 2px solid {label_border_color}; padding: 0.5rem; text-align: center;">{label}</div>',
            unsafe_allow_html=True
        )
    with col2:
        st.markdown(
            f"""
            <div style="
                background-color: {value_bg_color};
                padding: 0.5rem 1rem;
                border-radius: 4px;
                font-weight: bold;
                font-size: 1.1rem;
                text-align: center;
            ">
                {value}
            </div>
            """,
            unsafe_allow_html=True
        )

def fmt_euro(x):
            try:
                return f"{float(x):,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")
            except Exception:
                return x
            
def day_name_es(x):
    if pd.isnull(x):
        return None
    return format_datetime(pd.to_datetime(x), "EEEE", locale="es").capitalize()

# --------------------------
# Configuración de la página
# --------------------------

st.set_page_config(page_title="elBullistatistics", page_icon="resources/logo.png", layout="wide")

col1, col2, col3, col4 = st.columns([5, 1, 1, 2])
with col1:
    st.image("resources/image.png", width=400, use_container_width=False)
with col4:
    st.image("resources/water.png", width=250, use_container_width=False)

tab1, tab2, tab3, tab4 = st.tabs([
    "📁 Carga de archivos", 
    "📊 Informe generado",
    "⚙️ Exportar el informe",
    "📖 Manual del usuario"
])

# --------------------------
# Tab1: Carga de archivos
# --------------------------

with tab1:

    st.session_state.setdefault("reportReady", False)
    st.session_state.setdefault("fileName", "")
    st.session_state.setdefault("startDate", None)
    st.session_state.setdefault("endDate", None)
    st.session_state.setdefault("dfReservation", None)
    st.session_state.setdefault("dfOrigin", None)
    st.session_state.setdefault("dfClient", None)
    st.session_state.setdefault("dfStore", None)
    st.session_state.setdefault("dfVisit", None)
    st.session_state.setdefault("dfParking", None)

    # Formulario de creación del informe

    st.subheader("Generar informe estadístico")
    with st.form("report_generator"):

        fileName = st.text_input("Nombre del informe estadístico", help="Escribe el nombre deseado para el informe estadístico que se va a generar", key="fileName")
        colDay1, colDay2 = st.columns(2)
        with colDay1:
            startDate = st.date_input("Fecha de inicio", help="Selecciona una fecha inicial para el análisis", key="startDate")
        with colDay2:
            endDate = st.date_input("Fecha de fin", help="Selecciona una fecha final para el análisis", key="endDate")
        reportType = st.radio("¿Qué tipo de informe deseas generar?", ["📆 Informe mensual", "🗓️ Informe combinado de varios meses"], key="reportType", horizontal=True)
        devMode = st.checkbox("🔧 Modo desarrollador", value=False, help="Con el modo desarrollador pueden verse todas las tablas modificadas, con el objetivo de poder divisar posibles fallos.", key="devMode")
        reservationViewList = st.file_uploader("Subir archivo **reservationViewList**", type="xlsx")
        originSummary = st.file_uploader("Subir archivo **Resumen de Procedencias**", type="xlsx")
        clientList = st.file_uploader("Subir archivo **Listados de Clientes**", type="xlsx")
        storeRevenue = st.file_uploader("Subir archivo **Facturación de la Tienda**", type="xlsx")
        detailedVisitLog = st.file_uploader("Subir archivo **Diario de Visitas detallado**", type="xlsx")
        parkingSlotLog = st.file_uploader("Subir archivo **Diario de Visitas-Plazas Parking**", type="xlsx")
        st.markdown("<label style='font-size: 0.9rem;'><strong>(Opcional)</strong> Entradas no realmente gratuitas</label>", unsafe_allow_html=True)
        dfEmpty = pd.DataFrame({"Grupo": pd.Series(dtype="string"), "PAX": pd.Series(dtype="Int64")})
        dfDynamic = st.data_editor(dfEmpty, num_rows="dynamic", column_config={"Grupo": st.column_config.TextColumn("Grupo"), "PAX": st.column_config.NumberColumn("PAX", min_value=0, step=1),})
        realPaymentPax = dfDynamic["PAX"].sum()
        hasRows = not dfDynamic.dropna(how="all").empty

        st.divider()

        col1, col2, col3 = st.columns([3, 1, 3])
        with col2:
            generateButton = st.form_submit_button("✅ Generar informe")

        # Errores de campos obligatorios y éxito en la creación del informe

        if generateButton:
            if not fileName:
                st.warning("Debes asignar un nombre al informe.")
            elif startDate > endDate:
                st.error("La fecha de inicio no puede ser posterior a la de fin.")
            elif not reservationViewList or not originSummary or not clientList or not storeRevenue or not detailedVisitLog or not parkingSlotLog:
                st.warning("Debes subir todos los archivos requeridos.")
            else:
                steps = [
                    ("Cargando reservationListView...", lambda: pd.read_excel(reservationViewList), "dfReservation"),
                    ("Cargando procedencias...", lambda: pd.read_excel(originSummary, skiprows=5), "dfOrigin"),
                    ("Cargando perfiles...", lambda: pd.read_excel(clientList, sheet_name="PERFILES GENERAL", skiprows=3), "dfClient"),
                    ("Cargando grupos...", lambda: pd.read_excel(clientList, sheet_name="GRUPOS", skiprows=6), "dfGroup"),
                    ("Cargando ventas tienda...", lambda: pd.read_excel(storeRevenue), "dfStore"),
                    ("Cargando visitas detalladas...", lambda: pd.read_excel(detailedVisitLog, skiprows=5), "dfVisit"),
                    ("Cargando parking...", lambda: pd.read_excel(parkingSlotLog, skiprows=5), "dfParking"),
                ]

                totalSteps = len(steps) + 1
                progress = 0
                incr = int(100 / totalSteps)
                progress_text = st.empty()
                bar = st.progress(0, text="Iniciando...")

                try:
                    for msg, loader_fn, state_key in steps:
                        progress_text.markdown(f"**{msg}**")
                        with st.spinner(msg):
                            df = loader_fn()
                        st.session_state[state_key] = df

                        progress = min(progress + incr, 99)
                        bar.progress(progress, text=msg)

                    st.session_state["reportReady"] = True
                    progress_text.markdown("**Finalizando...**")
                    bar.progress(100, text="¡Completado!")
                    time.sleep(0.2)
                    bar.empty()
                    progress_text.empty()

                    st.success("Todos los campos están listos. **Informe generado** ✅")

                except Exception as e:
                    st.session_state["reportReady"] = False
                    bar.empty()
                    st.error(f"**Ha ocurrido un error al cargar los datos:** {e}")
        
# --------------------------
# Tab2: Informe generado
# --------------------------

with tab2:

    st.subheader("Informe estadístico generado")

    # Estado sin generar

    if not st.session_state.get("reportReady"):
        st.info("🔎 Sube los archivos y pulsa **«Generar informe»** en la pestaña de carga de archivos.")

    # Estado generado

    else:
        
        dfReservation = st.session_state["dfReservation"]
        dfOrigin = st.session_state["dfOrigin"]
        dfClient = st.session_state["dfClient"]
        dfGroup = st.session_state["dfGroup"]
        dfStore = st.session_state["dfStore"]
        dfVisit = st.session_state["dfVisit"]
        dfParking = st.session_state["dfParking"]
        
        startDate = pd.to_datetime(st.session_state["startDate"])
        endDate = pd.to_datetime(st.session_state["endDate"])

        # ----- Diario de visitas detallado -----
        
        dfVisit["Fecha visita"] = pd.to_datetime(dfVisit["Fecha visita"], format="%Y-%m-%d %H:%M:%S", errors="coerce")
        dfVisit["Fecha visita2"] = dfVisit["Fecha visita"].dt.normalize()
        dfVisit = dfVisit[(dfVisit["Fecha visita2"] >= startDate) & (dfVisit["Fecha visita2"] <= endDate)]
        valuesToDelete = ["Regala elBulli1846", "Regala Visita Guiada a elBulli1846", "Parking (3h)", "Regala Visita Guiada elBulli1846"]
        dfVisitCopy = dfVisit.loc[~dfVisit["Producto"].isin(valuesToDelete)].copy()
        dfVisitCopy["Hora"] = dfVisitCopy["Fecha visita"].dt.strftime("%H:%M")

        # Horas de acceso
        st.divider()

        st.markdown("<h3 style='color: #2db1fc; font-weight: bold;'>Horas de acceso</h3>", unsafe_allow_html=True)
        st.markdown("<h5 style='color: #292929; font-weight: bold;'>¿A qué hora se visita el museo?</h3>", unsafe_allow_html=True)
        pivotTime = makePivot(
            df=dfVisitCopy,
            index_col="Hora",
            value_col="Pax",
            aggfunc="sum",
            label_fmt="%H:%M"
        )
        renderBlockWithTable(
            pivot_df=pivotTime,
            label_col="label",
            value_col="Pax",
            label_title="Hora de acceso",
            color="#2db1fc",
            height=None,
            show_percent_labels=True
        )

        # Días de acceso
        st.divider()
        
        dfVisitCopy["Dia de la semana"] = dfVisitCopy["Fecha visita2"].apply(day_name_es)
        daysOrder = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
        
        st.markdown("<h3 style='color: #2db1fc; font-weight: bold;'>Días de acceso</h3>", unsafe_allow_html=True)
        st.markdown("<h5 style='color: #292929; font-weight: bold;'>¿Qué días de la semana se visita el museo?</h3>", unsafe_allow_html=True)
        pivotDay = makePivot(
            df=dfVisitCopy,
            index_col="Dia de la semana",
            value_col="Pax",
            aggfunc="sum",
            label_fmt=None
        )
        renderBlockWithTable(
            pivot_df=pivotDay,
            label_col="label",
            value_col="Pax",
            label_title="Día de la semana",
            color="#2db1fc",
            height=420,
            show_percent_labels=True,
            y_order=daysOrder
        )

        # Horas de acceso según el día de la semana
        st.divider()

        def highlightZeros(val):
            if val == 0:
                return "background-color: #ffcccc; color: red; font-weight: bold;"
            return ""
        def highlightHigherThan100(val):
            if val >= 100 and reportType == "🗓️ Informe combinado de varios meses":
                return "background-color: #ffe699; color: #b58900; font-weight: bold;"
            return ""
        def boldTotalRow(row):
            if row.iloc[0] == "TOTAL":
                return ["font-weight: bold;"] * len(row)
            return [""] * len(row)
        
        st.markdown("<h3 style='color: #2db1fc; font-weight: bold;'>Horas de acceso según día de la semana</h3>", unsafe_allow_html=True)
        st.markdown("<h5 style='color: #292929; font-weight: bold;'>¿Cúantos visitantes hay en cada franja de acceso?</h3>", unsafe_allow_html=True)
        pivotDayHour = pd.pivot_table(
            dfVisitCopy,
            index="Hora",
            columns="Dia de la semana",
            values="Pax",
            aggfunc="sum",
            fill_value=0
        )
        lowerOrder = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]
        cols_lower = [str(c).lower() for c in pivotDayHour.columns]
        cols_reorden = [pivotDayHour.columns[cols_lower.index(d)] for d in lowerOrder if d in cols_lower]
        pivotDayHour = pivotDayHour.reindex(columns=cols_reorden)
        pivotDayHour["TOTAL"] = pivotDayHour.sum(axis=1)
        fila_total = pivotDayHour.sum(axis=0)
        fila_total.name = "TOTAL"
        pivotDayHour = pivotDayHour.sort_index(
            key=lambda x: [
                (h.hour * 60 + h.minute) if isinstance(h, datetime.time) else 9999
                for h in x
            ]
        )
        pivotDayHour = pd.concat([pivotDayHour, fila_total.to_frame().T])
        pivotDayHour.index = [
            h.strftime("%H:%M") if hasattr(h, "strftime") else h
            for h in pivotDayHour.index
        ]
        tabla_mostrar = pivotDayHour.reset_index()
        tabla_mostrar.columns = [""] + list(tabla_mostrar.columns[1:])
        tabla_styled = (
            tabla_mostrar.style
            .applymap(highlightZeros, subset=tabla_mostrar.columns[1:])
            .applymap(highlightHigherThan100, subset=tabla_mostrar.columns[1:])
            .applymap(lambda x: "font-weight: bold;", subset=["TOTAL"])
        )
        tabla_styled = tabla_styled.apply(boldTotalRow, axis=1)
        rowHeight = 33
        nRows = tabla_mostrar.shape[0] + 2
        height = rowHeight * nRows
        with st.container(border=True):
            st.dataframe(
                tabla_styled,
                use_container_width=True,
                height=height,
                hide_index=True
            )
        
        # ----- ReservationListView -----

        dfReservation["Fecha reserva / compra"] = pd.to_datetime(dfReservation["Fecha reserva / compra"], dayfirst=True, errors="coerce")
        dfReservation["Fecha visita"] = pd.to_datetime(dfReservation["Fecha visita"], dayfirst=True, errors="coerce")
        dfReservation = dfReservation[(dfReservation["Fecha visita"] >= startDate) & (dfReservation["Fecha visita"] <= endDate)]
        dfReservation = dfReservation[~dfReservation["Producto"].isin(valuesToDelete)]
        reservationDay = dfReservation["Fecha reserva / compra"].dt.normalize()
        visitDay = dfReservation["Fecha visita"].dt.normalize()
        dfReservation["Fecha visita2"] = dfReservation["Fecha visita"].dt.normalize()
        dfReservation["Antelacion"] = (visitDay - reservationDay).dt.days
        
        # Antelación de compra
        st.divider()

        def classifyDays(days):
            if days == 0:
                return "0 días"
            elif days == 1:
                return "1 día"
            elif 2 <= days <= 5:
                return "2-5 días"
            elif 6 <= days <= 10:
                return "6-10 días"
            elif 11 <= days <= 20:
                return "11-20 días"
            elif 21 <= days <= 30:
                return "21-30 días"
            elif 31 <= days <= 60:
                return "31-60 días"
            elif 61 <= days <= 90:
                return "61-90 días"
            elif days > 90:
                return "+90 días"
            else:
                return "Error de antelación"

        st.markdown("<h3 style='color: #2db1fc; font-weight: bold;'>Antelación de compra</h3>", unsafe_allow_html=True)
        st.markdown("<h5 style='color: #292929; font-weight: bold;'>¿Con cuántos días de anticipación se adquieren las entradas al museo?</h3>", unsafe_allow_html=True)
        dfReservation["Antelacion clasificada"] = dfReservation["Antelacion"].apply(classifyDays)
        advanceOrder = ["0 días", "1 día", "2-5 días", "6-10 días", "11-20 días", "21-30 días", "31-60 días", "61-90 días", "+90 días"]
        pivotAdvance = makePivot(
            df=dfReservation,
            index_col="Antelacion clasificada",
            value_col="Tickets Válidos",
            aggfunc="sum",
            label_fmt=None
        )
        pivotAdvance["label"] = pd.Categorical(pivotAdvance["label"], categories=advanceOrder, ordered=True)
        pivotAdvance = pivotAdvance.sort_values("label")
        renderBlockWithTable(
            pivot_df=pivotAdvance,
            label_col="label",
            value_col="Tickets Válidos",
            label_title="Antelación",
            color="#2db1fc",
            height=420,
            show_percent_labels=True,
            y_order=advanceOrder
        )
        st.info("El dato de antelación se obtiene únicamente de la ReservationListView. Esta fuente puede ser inexacta, por lo que **el total de visitantes puede variar ligeramente respecto a otros cálculos**.")

        # Resumen de Procedencias
        st.divider()

        periods = pd.period_range(start=startDate, end=endDate, freq="M")
        dfOrigin["__periodo"] = pd.to_datetime(
            dfOrigin["Fecha visita"], format="%B %Y", errors="coerce"
        ).dt.to_period("M")
        dfOrigin = dfOrigin[dfOrigin["__periodo"].isin(periods)].copy()
        dfOrigin.drop(columns="__periodo", inplace=True)

        st.markdown("<h3 style='color: #2db1fc; font-weight: bold;'>Procedencia de los visitantes por países</h3>", unsafe_allow_html=True)
        st.markdown("<h5 style='color: #292929; font-weight: bold;'>¿De qué países provienen los visitantes del museo?</h3>", unsafe_allow_html=True)
        dfOrigin["Procedencia clasificada"] = dfOrigin["Procedencia"]
        dfOrigin["Cógido Postal Texto"] = dfOrigin["Código postal"].astype(str).str.strip().str[:5]
        dfOrigin.loc[(dfOrigin["Procedencia"] == "España") & (dfOrigin["Comunidad"] == "Catalunya") & (dfOrigin["Cógido Postal Texto"] == "17480"), "Procedencia clasificada"] = "Roses"
        dfOrigin.loc[(dfOrigin["Procedencia"] == "España") & (dfOrigin["Comunidad"] == "Catalunya") & (dfOrigin["Cógido Postal Texto"] != "17480"), "Procedencia clasificada"] = "Catalunya"
        pivotOrigin = makePivot(
            df=dfOrigin,
            index_col="Procedencia clasificada",
            value_col="Pax",
            aggfunc="sum",
            label_fmt=None
        )
        countriesOrder = ["Roses", "Catalunya", "España"]
        otherCountries = sorted([c for c in pivotOrigin["label"].unique() if c not in countriesOrder])
        finalCountriesOrder = countriesOrder + otherCountries
        pivotOrigin["label"] = pd.Categorical(pivotOrigin["label"], categories=finalCountriesOrder, ordered=True)
        pivotOrigin = pivotOrigin.sort_values("label")
        renderBlockWithTable(
            pivot_df=pivotOrigin,
            label_col="label",
            value_col="Pax",
            label_title="Procedencia",
            color="#2db1fc",
            height=None,
            show_percent_labels=True,
            y_order=finalCountriesOrder
        )
        st.info("El resumen de procedencias se basa en datos mensuales. Si el análisis abarca menos de un mes completo, la gráfica seguirá mostrando el total mensual asignado, por lo que los valores pueden exceder el periodo seleccionado.")

        # Perfil profesional de los visitantes
        st.divider()

        dfClient["FECHA"] = pd.to_datetime(dfClient["FECHA"], dayfirst=True, errors="coerce")
        dfClient = dfClient[(dfClient["FECHA"] >= startDate) & (dfClient["FECHA"] <= endDate)]

        st.markdown("<h3 style='color: #2db1fc; font-weight: bold;'>Perfil profesional de los visitantes</h3>", unsafe_allow_html=True)
        st.markdown("<h5 style='color: #292929; font-weight: bold;'>¿De qué sector profesional provienen los visitantes al museo?</h3>", unsafe_allow_html=True)
        pivotProfile = makePivot(
            df=dfClient,
            index_col="PERFIL PROFESIONAL",
            value_col="PAX",
            aggfunc="sum",
            label_fmt=None
        )
        renderBlockWithTable(
            pivot_df=pivotProfile,
            label_col="label",
            value_col="PAX",
            label_title="Perfil profesional",
            color="#2db1fc",
            height=None,
            show_percent_labels=True,
        )

        # Producto adquirido por el visitante
        st.divider()

        dfVisit.loc[dfVisit["Producto"] == "Visita exclusiva a elBulli1846", "Producto"] = "Visita guiada a elBulli1846"

        st.markdown("<h3 style='color: #2db1fc; font-weight: bold;'>Producto adquirido por el visitante</h3>", unsafe_allow_html=True)
        st.markdown("<h5 style='color: #292929; font-weight: bold;'>¿Cuántos visitantes adquieren cada tipo de producto?</h3>", unsafe_allow_html=True)
        pivotProduct = makePivot(
            df=dfVisit,
            index_col="Producto",
            value_col="Pax",
            aggfunc="sum",
            label_fmt=None
        )
        renderBlockWithTable(
            pivot_df=pivotProduct,
            label_col="label",
            value_col="Pax",
            label_title="Producto",
            color="#2db1fc",
            height=None,
            show_percent_labels=True,
        )

        # Colectivo del visitante
        st.divider()
        
        dfVisit = dfVisit[~dfVisit["Producto"].isin(valuesToDelete)]

        st.markdown("<h3 style='color: #2db1fc; font-weight: bold;'>Colectivo del visitante</h3>", unsafe_allow_html=True)
        st.markdown("<h5 style='color: #292929; font-weight: bold;'>¿A qué colectivo pertenecen los visitantes al museo?</h3>", unsafe_allow_html=True)
        pivotCollective = makePivot(
            df=dfVisit,
            index_col="Colectivo",
            value_col="Pax",
            aggfunc="sum",
            label_fmt=None
        )
        renderBlockWithTable(
            pivot_df=pivotCollective,
            label_col="label",
            value_col="Pax",
            label_title="Colectivo",
            color="#2db1fc",
            height=None,
            show_percent_labels=True,
        )
        
        # Información adicional del museo
        st.divider()

        st.markdown("<h3 style='color: #2db1fc; font-weight: bold;'>Información adicional del museo</h3>", unsafe_allow_html=True)
        
        dfVisit["Fecha visita"] = pd.to_datetime(dfVisit["Fecha visita"], errors="coerce").dt.normalize()
        pivot_dias = (
            dfVisit.groupby("Fecha visita", dropna=True, as_index=False)["Pax"]
            .sum(min_count=1)
            .sort_values("Fecha visita")
            .reset_index(drop=True)
        )
        openDays = int((pivot_dias["Pax"] > 0).sum())
        
        # --> Días abiertos
        st.markdown("<h5 style='color: #292929; font-weight: bold;'>¿Cuántos días se ha abierto el museo?</h3>", unsafe_allow_html=True)
        infoBox("Días abiertos", openDays)

        totalVisitors = int(dfVisit["Pax"].sum())
        visitAverage = int(totalVisitors / openDays)

        # --> Media de visitantes
        st.markdown("<h5 style='color: #292929; font-weight: bold; margin-top: 12px;'>¿Cuál ha sido la media de visitantes por día?</h3>", unsafe_allow_html=True)
        infoBox("Media de visitantes", visitAverage)

        # --> Número total de visitantes
        st.markdown("<h5 style='color: #292929; font-weight: bold; margin-top: 12px;'>¿Cuál ha sido el número total de visitantes?</h3>", unsafe_allow_html=True)
        infoBox("Número total de visitantes", totalVisitors)

        # ----- Diario detallado parking -----

        dfParking["Fecha visita"] = pd.to_datetime(dfParking["Fecha visita"], format="%Y-%m-%d %H:%M:%S", errors="coerce")
        dfParking["Fecha visita"] = dfParking["Fecha visita"].dt.normalize()
        dfParking = dfParking[(dfParking["Fecha visita"] >= startDate) & (dfParking["Fecha visita"] <= endDate)]
        
        totalParkings = int(dfParking["Plazas Parking"].sum())

        # --> Plazas de parking ocupadas
        st.markdown("<h5 style='color: #292929; font-weight: bold; margin-top: 12px;'>¿Cuántas plazas de párking se han ocupado?</h3>", unsafe_allow_html=True)
        infoBox("Número total de plazas de párking", totalParkings)

        st.markdown("<h5 style='color: #292929; font-weight: bold; margin-top: 12px;'>¿Cuál ha sido el promedio de invitaciones?</h3>", unsafe_allow_html=True)
        total_pax = int(dfVisit["Pax"].sum())
        freePax = int(dfVisit.loc[dfVisit["Importe (€)"] <= 0, "Pax"].sum())
        payPax  = int(dfVisit.loc[dfVisit["Importe (€)"] > 0, "Pax"].sum())
        freePct = (freePax / total_pax * 100) if total_pax > 0 else 0.0
        payPct = (payPax  / total_pax * 100) if total_pax > 0 else 0.0
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown('<div style="border: 2px solid #2db1fc; padding: 0.5rem; text-align: center;">Gratuita</div>', unsafe_allow_html=True)
        with col2:
            st.markdown(f'<div style="border: 2px solid #2db1fc; font-weight: bold; padding: 0.5rem; text-align: center;">{freePax}</div>', unsafe_allow_html=True)
        with col3:
            st.markdown(f'<div style="border: 2px solid #2db1fc; font-weight: bold; padding: 0.5rem; text-align: center;">{freePct:.1f}%</div>', unsafe_allow_html=True)
        col4, col5, col6 = st.columns(3)
        with col4:
            st.markdown('<div style="border: 2px solid #2db1fc; margin-top: 12px; padding: 0.5rem; text-align: center;">De pago</div>', unsafe_allow_html=True)
        with col5:
            st.markdown(f'<div style="border: 2px solid #2db1fc; font-weight: bold; margin-top: 12px; padding: 0.5rem; text-align: center;">{payPax}</div>', unsafe_allow_html=True)
        with col6:
            st.markdown(f'<div style="border: 2px solid #2db1fc; font-weight: bold; margin-top: 12px; padding: 0.5rem; text-align: center;">{payPct:.1f}%</div>', unsafe_allow_html=True)
        st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

        if hasRows:

            grupos = dfDynamic["Grupo"].dropna().tolist()
            grupos_str = ", ".join(grupos)

            message = (
                f"En el promedio de invitaciones, el número de gratuitas, **{freePax} PAX**, "
                f"incluye los **{realPaymentPax}** de {grupos_str}; que realmente no son gratuitas.  \n\n"
                f"El porcentaje real de invitaciones es:  \n"
                f"- **Gratuita** → {freePax - realPaymentPax} PAX ({freePct:.1f}%)  \n"
                f"- **De pago** → {payPax + realPaymentPax} PAX ({payPct:.1f}%)"
            )

            st.info(message)

        # Facturación de la tienda
        st.divider()

        st.markdown("<h3 style='color: #2db1fc; font-weight: bold;'>Facturación diaria de la tienda</h3>", unsafe_allow_html=True)
        st.markdown("<h5 style='color: #292929; font-weight: bold;'>¿Cuál ha sido la facturación diaria de la tienda?</h3>", unsafe_allow_html=True)
        
        dfStore["Fecha"] = pd.to_datetime(dfStore["Fecha"], dayfirst=True, errors="coerce")
        dfStore = dfStore[(dfStore["Fecha"] >= startDate) & (dfStore["Fecha"] <= endDate)]
        dfStore["TOTAL FACTURACIÓN TIENDA"] = pd.to_numeric(dfStore["TOTAL FACTURACIÓN TIENDA"], errors="coerce").fillna(0)
        dfStore = dfStore[dfStore["TOTAL FACTURACIÓN TIENDA"] > 0]

        pivotStore = makePivot(
            df=dfStore,
            index_col="Fecha",
            value_col="TOTAL FACTURACIÓN TIENDA",
            aggfunc="sum",
            label_fmt="%d/%m/%y"
        )
        renderBlockWithTable(
            pivot_df=pivotStore,
            label_col="label",
            secondlabel_title="Importe",
            value_col="TOTAL FACTURACIÓN TIENDA",
            label_title="Fecha",
            color="#2db1fc",
            height=720,
            show_percent_labels=False,
            value_format="euro",
            chart_type="line",
            temporal_col="Fecha", 
        )

        totalStore = int(dfStore["TOTAL FACTURACIÓN TIENDA"].sum())
        totalStoreFmt = f"{totalStore:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")

        if reportType == "📆 Informe mensual":
            infoBox("<b>Facturación total de la tienda</b>", totalStoreFmt)
        elif reportType == "🗓️ Informe combinado de varios meses":
            df_tmp = dfStore.copy()
            if not pd.api.types.is_datetime64_any_dtype(df_tmp["Fecha"]):
                df_tmp["Fecha"] = pd.to_datetime(df_tmp["Fecha"], errors="coerce")
            df_tmp = df_tmp[df_tmp["TOTAL FACTURACIÓN TIENDA"] > 0].copy()
            df_tmp["YM"] = df_tmp["Fecha"].dt.to_period("M")
            por_mes = df_tmp.groupby("YM")["TOTAL FACTURACIÓN TIENDA"].sum().sort_index()
            months = {1:"enero",2:"febrero",3:"marzo",4:"abril",5:"mayo",6:"junio",7:"julio",8:"agosto",9:"septiembre",10:"octubre",11:"noviembre",12:"diciembre"}

            def euro_fmt(x: float) -> str:
                return f"{x:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")

            for ym, importe in por_mes.items():
                yy, mm = ym.year, ym.month
                st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
                infoBox(f"Facturación de {months[mm]} de {yy}", euro_fmt(importe))

            st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
            infoBox("<b>Facturación total de la tienda</b>", totalStoreFmt)

        # Facturación en taquilla
        st.divider()

        st.markdown("<h3 style='color: #2db1fc; font-weight: bold;'>Facturación diaria de taquilla</h3>", unsafe_allow_html=True)
        st.markdown("<h5 style='color: #292929; font-weight: bold;'>¿Cuál ha sido la facturación diaria de taquilla?</h3>", unsafe_allow_html=True)
        
        dfVisit = dfVisit[dfVisit["Importe (€)"] > 0]

        pivotTickets = makePivot(
            df=dfVisit,
            index_col="Fecha visita",
            value_col="Importe (€)",
            aggfunc="sum",
            label_fmt="%d/%m/%y"
        )
        renderBlockWithTable(
            pivot_df=pivotTickets,
            label_col="label",
            secondlabel_title="Importe",
            value_col="Importe (€)",
            label_title="Fecha visita",
            color="#2db1fc",
            height=720,
            show_percent_labels=False,
            value_format="euro",
            chart_type="line",
            temporal_col="Fecha visita", 
        )

        totalTickets = int(dfVisit["Importe (€)"].sum())
        totalTicketsFmt = f"{totalTickets:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")

        if reportType == "📆 Informe mensual":
            infoBox("<b>Facturación total de la taquilla</b>", totalTicketsFmt)
        elif reportType == "🗓️ Informe combinado de varios meses":
            df_tmp = dfVisit.copy()
            if not pd.api.types.is_datetime64_any_dtype(df_tmp["Fecha visita"]):
                df_tmp["Fecha visita"] = pd.to_datetime(df_tmp["Fecha visita"], errors="coerce")
            df_tmp = df_tmp[df_tmp["Importe (€)"] > 0].copy()
            df_tmp["YM"] = df_tmp["Fecha visita"].dt.to_period("M")
            por_mes = df_tmp.groupby("YM")["Importe (€)"].sum().sort_index()
            months = {1:"enero",2:"febrero",3:"marzo",4:"abril",5:"mayo",6:"junio",7:"julio",8:"agosto",9:"septiembre",10:"octubre",11:"noviembre",12:"diciembre"}

            def euro_fmt(x: float) -> str:
                return f"{x:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")

            for ym, importe in por_mes.items():
                yy, mm = ym.year, ym.month
                st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
                infoBox(f"Facturación de {months[mm]} de {yy}", euro_fmt(importe))

            st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
            infoBox("<b>Facturación total de la taquilla</b>", totalTicketsFmt)

        #  Información adicional de ventas
        st.divider()
        
        st.markdown("<h3 style='color: #2db1fc; font-weight: bold;'>Información adicional de ventas</h3>", unsafe_allow_html=True)
        
        # --> Facturación según día de la semana
        st.markdown("<h5 style='color: #292929; font-weight: bold;'>¿Cuál ha sido la facturación en tienda según el día de la semana?</h3>", unsafe_allow_html=True)

        activeDays = pivotDay.loc[pivotDay["Pax"] > 0, "label"].tolist()
        daysOrder   = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
        orderActive = [d for d in daysOrder if d in activeDays]
        dfStore_week = dfStore.copy()
        dfStore_week["Dia de la semana"] = dfStore_week["Fecha"].apply(day_name_es)
        fact_store = (
            dfStore_week.groupby("Dia de la semana", as_index=False)["TOTAL FACTURACIÓN TIENDA"]
            .sum()
            .rename(columns={"TOTAL FACTURACIÓN TIENDA": "Fact. Total"})
        )
        pax_totales = pivotDay[["label", "Pax"]].rename(
            columns={"label": "Dia de la semana", "Pax": "Pax_total"}
        )
        fact_store["__dia_norm"] = fact_store["Dia de la semana"].astype(str).str.strip().str.lower()
        pax_totales["__dia_norm"] = pax_totales["Dia de la semana"].astype(str).str.strip().str.lower()

        tabla = fact_store.merge(
            pax_totales[["__dia_norm", "Pax_total"]],
            on="__dia_norm",
            how="left"
        ).drop(columns="__dia_norm")

        tabla = tabla[tabla["Dia de la semana"].isin(orderActive)].copy()
        tabla["Dia de la semana"] = pd.Categorical(tabla["Dia de la semana"], categories=orderActive, ordered=True)
        tabla = tabla.sort_values("Dia de la semana").reset_index(drop=True)

        tabla["Ticket medio"] = tabla.apply(
            lambda r: (r["Fact. Total"] / r["Pax_total"]) if pd.notnull(r["Pax_total"]) and r["Pax_total"] > 0 else 0.0,
            axis=1
        )

        tabla["Fact. Total"]  = tabla["Fact. Total"].apply(fmt_euro)
        tabla["Ticket medio"] = tabla["Ticket medio"].apply(fmt_euro)
        tableStorexDay = tabla.rename(columns={"Dia de la semana": "Día de la semana"})[["Día de la semana", "Fact. Total", "Ticket medio"]]

        with st.container(border=True):
            st.dataframe(tableStorexDay, use_container_width=True, hide_index=True)

        # --> Canal de venta más utilizado
        st.markdown("<h5 style='color: #292929; font-weight: bold;'>¿Cuál ha sido el canal de venta más utilizado para las entradas al museo?</h3>", unsafe_allow_html=True)

        dfVisit_canal = dfVisit[~dfVisit["Producto"].isin(valuesToDelete)].copy()
        tabla_canal = (
            dfVisit_canal.groupby("Canal de venta", as_index=False)
            .agg({"Pax": "sum", "Importe (€)": "sum"})
            .rename(columns={"Importe (€)": "Fact. Total"})
        )
        total_pax = float(tabla_canal["Pax"].sum())
        tabla_canal["%"] = (tabla_canal["Pax"] / total_pax * 100).round(1).astype(str) + "%"
        tabla_canal = tabla_canal.sort_values("Pax", ascending=False).reset_index(drop=True)
        tabla_canal["Fact. Total"] = tabla_canal["Fact. Total"].apply(fmt_euro)
        tabla_canal = tabla_canal[["Canal de venta", "Pax", "%", "Fact. Total"]]

        with st.container(border=True):
            st.dataframe(tabla_canal, use_container_width=True, hide_index=True)

        # --> Tickets medios
        st.markdown("<h5 style='color: #292929; font-weight: bold;'>¿Cuáles han sido los tickets medios?</h3>", unsafe_allow_html=True)
        
        ticketStore = totalStore / totalVisitors
        infoBox("Facturación en tienda", fmt_euro(ticketStore))
        st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
        ticketTickets = totalTickets / totalVisitors
        infoBox("Facturación en taquilla", fmt_euro(ticketTickets))
        st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
        averageGeneralTicket = ticketStore + ticketTickets
        infoBox("<b>Tienda + Taquilla</b>", fmt_euro(averageGeneralTicket))
        st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

        # Información adicional de grupos
        st.divider()
        
        dfGroup["FECHA"] = pd.to_datetime(dfGroup["FECHA"], dayfirst=True, errors="coerce")
        dfGroup = dfGroup[(dfGroup["FECHA"] >= startDate) & (dfGroup["FECHA"] <= endDate)]
        
        st.markdown("<h3 style='color: #2db1fc; font-weight: bold;'>Información adicional de grupos</h3>", unsafe_allow_html=True)

        groupSheet = dfGroup[["FECHA", "NOMBRE RESERVA", "PAX", "EMPRESA / OTRO TIPO GRUPO", "NOTAS"]]
        groupSheet["FECHA"] = groupSheet["FECHA"].dt.strftime("%d/%m/%Y")
        groupSheet = groupSheet.rename(columns={
            "FECHA": "Fecha",
            "NOMBRE RESERVA": "Nombre de la reserva",
            "PAX": "Nº PAX",
            "EMPRESA / OTRO TIPO GRUPO": "Empresa / Otros grupos",
            "NOTAS": "Observaciones"
        })
        groupSheet["Observaciones"] = groupSheet["Observaciones"].fillna("-")

        # --> Número total de grupos
        st.markdown("<h5 style='color: #292929; font-weight: bold;'>¿Cuántos grupos (+ 10 PAX) han visitado el museo?</h3>", unsafe_allow_html=True)
        groupNumber = len(dfGroup)
        infoBox("Número de grupos", groupNumber)
        st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

        # --> Total de visitantes en grupo
        st.markdown("<h5 style='color: #292929; font-weight: bold;'>¿Cuál es el total de visitantes en grupo?</h3>", unsafe_allow_html=True)
        totalGroups = dfGroup["PAX"].sum()
        infoBox("Total de visitantes en grupo", totalGroups)
        st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

        # --> Grupos que han visitado el museo
        st.markdown("<h5 style='color: #292929; font-weight: bold;'>¿Qué grupos han visitado el museo?</h3>", unsafe_allow_html=True)
        with st.container(border=True):
            st.dataframe(groupSheet, use_container_width=True, hide_index=True)

        # Modo desarrollador

        if st.session_state.get("devMode"):

            st.divider()
            st.subheader("Modo desarrollador")
            st.warning("⚠️ El modo desarrollador está activo")

            st.session_state["dfVisitCopy"] = dfVisitCopy
            st.session_state["dfReservation"] = dfReservation
            st.session_state["dfOrigin"] = dfOrigin
            st.session_state["dfClient"] = dfClient
            st.session_state["dfGroup"] = dfGroup
            st.session_state["dfVisit"] = dfVisit
            st.session_state["dfParking"] = dfParking

            st.markdown("<h5 style='color: #292929; font-weight: bold;'>dfVisitCopy</h3>", unsafe_allow_html=True)
            st.write(dfVisitCopy.head(len(dfVisitCopy)))
            st.markdown("<h5 style='color: #292929; font-weight: bold;'>dfReservation</h3>", unsafe_allow_html=True)
            st.write(dfReservation.head(len(dfReservation)))
            st.markdown("<h5 style='color: #292929; font-weight: bold;'>dfOrigin</h3>", unsafe_allow_html=True)
            st.write(dfOrigin.head(len(dfOrigin)))
            st.markdown("<h5 style='color: #292929; font-weight: bold;'>dfClient</h3>", unsafe_allow_html=True)
            st.write(dfClient.head(len(dfClient)))
            st.markdown("<h5 style='color: #292929; font-weight: bold;'>dfGroup</h3>", unsafe_allow_html=True)
            st.write(dfGroup.head(len(dfGroup)))
            st.markdown("<h5 style='color: #292929; font-weight: bold;'>dfVisit</h3>", unsafe_allow_html=True)
            st.write(dfVisit.head(len(dfVisit)))
            st.markdown("<h5 style='color: #292929; font-weight: bold;'>dfParking</h3>", unsafe_allow_html=True)
            st.write(dfParking.head(len(dfParking)))

# --------------------------
# Tab3: Exportar el informe
# --------------------------

with tab3:
    
    # Plantilla
    
    st.subheader("Plantilla del informe")
    with open("resources/Plantilla_Informe_Estadistico.docx", "rb") as file:
            st.download_button(
                label="📄 Descargar plantilla para Microsoft Word",
                data=file,
                file_name="Plantilla_Informe_Estadistico.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    
    # Generador de PDFs
    
    st.subheader("Generador de PDFs")
    if not st.session_state.get("reportReady"):
        st.info("🔎 Sube los archivos y pulsa **«Generar informe»** en la pestaña de carga de archivos para activar el **Generador de PDFs**.")
    else:

        with st.container(border=True):

            page1  = st.file_uploader("Subir captura de Horas de acceso", type=["png","jpg","jpeg"], key="p1")
            page2  = st.file_uploader("Subir captura de Días de acceso", type=["png","jpg","jpeg"], key="p2")
            page3  = st.file_uploader("Subir captura de Horas de acceso según día de la semana", type=["png","jpg","jpeg"], key="p3")
            page4  = st.file_uploader("Subir captura de Antelación de compra", type=["png","jpg","jpeg"], key="p4")
            page5  = st.file_uploader("Subir captura de Procedencia de los visitantes por países", type=["png","jpg","jpeg"], key="p5")
            page6  = st.file_uploader("Subir captura de Perfil profesional de los visitantes", type=["png","jpg","jpeg"], key="p6")
            page7  = st.file_uploader("Subir captura de Producto adquirido por el visitante", type=["png","jpg","jpeg"], key="p7")
            page8  = st.file_uploader("Subir captura de Colectivo del visitante", type=["png","jpg","jpeg"], key="p8")
            page9  = st.file_uploader("Subir captura de Información adicional del museo", type=["png","jpg","jpeg"], key="p9")
            page10  = st.file_uploader("Subir captura de Facturación diaria de la tienda", type=["png","jpg","jpeg"], key="p10")
            page11 = st.file_uploader("Subir captura de Facturación diaria de taquilla", type=["png","jpg","jpeg"], key="p11")
            page12 = st.file_uploader("Subir captura de Información adicional de ventas", type=["png","jpg","jpeg"], key="p12")
            page13 = st.file_uploader("Subir captura de Información adicional de grupos", type=["png","jpg","jpeg"], key="p13")
            uploads = [page1,page2,page3,page4,page5,page6,page7,page8,page9,page10,page11,page12,page13]

            st.divider()

            col1, col2, col3 = st.columns([3, 1, 3])
            with col2:
                if st.button("Generar y preparar descarga"):
                    import tempfile, os
                    imagePaths = []
                    for up in uploads:
                        if up is None:
                            continue
                        ext = os.path.splitext(up.name)[1].lower() or ".png"
                        with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as img_tmp:
                            img_tmp.write(up.read())
                            imagePaths.append(img_tmp.name)
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                        build_pdf(tmp.name, titulo="Informe Consolidado", file_name=fileName, total_visitors=totalVisitors, page_images=imagePaths)
                        temp_path = tmp.name
                    with open(temp_path, "rb") as f:
                        pdf_bytes = f.read()
                    st.download_button("⬇️ Descargar informe", data=pdf_bytes, file_name="informe.pdf", mime="application/pdf")
                    st.success("PDF generado. ¡Listo para descargar!")

# --------------------------
# Tab4: Manual del usuario
# --------------------------

with tab4:
    
    st.subheader("Manual en PDF")
    with open("resources/Plantilla_Informe_Estadistico.docx", "rb") as file:
            st.download_button(
                label="📄 Descargar manual en formato PDF",
                data=file,
                file_name="Plantilla_Informe_Estadistico.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

    st.subheader("Manual del usuario")