import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import datetime
import os # Aunque os no se usa directamente, lo mantenemos por si acaso

# --- Configuración de Página ---
st.set_page_config(page_title="Dashboard Comercial", layout="wide")

# --- Funciones de Ayuda ---
def format_currency(value):
    """Formatea un número como moneda COP."""
    try:
        return f"${float(value):,.0f}"
    except (ValueError, TypeError):
        return "$0" # O devuelve un valor por defecto apropiado

def safe_to_datetime(series):
    """Convierte una serie a datetime, manejando errores."""
    return pd.to_datetime(series, errors='coerce')

def safe_to_numeric(series):
    """Convierte una serie a numérico, manejando errores."""
    return pd.to_numeric(series, errors='coerce')

# --- Función para Pestañas de Resumen (Posibles, Aprovechables) ---
def render_summary_tab(df, date_col, value_col, tab_title, key_prefix):
    """
    Renderiza una pestaña con resumen por periodo, KPI y gráficos.

    Args:
        df (pd.DataFrame): DataFrame con los datos.
        date_col (str): Nombre de la columna de fecha.
        value_col (str): Nombre de la columna de valor numérico (e.g., Subtotal).
        tab_title (str): Título para la pestaña y gráficos.
        key_prefix (str): Prefijo único para las claves de los widgets de Streamlit.
    """
    st.subheader(f"📋 Datos y Resumen - {tab_title}")

    if df is None or df.empty:
        st.warning(f"No se encontraron datos para '{tab_title}'. Verifica la hoja en el Excel.")
        return

    # Mostrar datos crudos (primeras filas)
    st.dataframe(df.head())

    # --- Procesamiento ---
    df_processed = df.copy()
    df_processed[date_col] = safe_to_datetime(df_processed[date_col])
    df_processed[value_col] = safe_to_numeric(df_processed[value_col])
    df_processed = df_processed.dropna(subset=[date_col, value_col])

    if df_processed.empty:
        st.warning(f"No hay datos válidos (fecha y valor) para procesar en '{tab_title}'.")
        return

    # --- Selección de Agrupación ---
    opcion = st.selectbox(
        "Selecciona el período para agrupar:",
        ("Día", "Mes", "Año"),
        key=f'{key_prefix}_periodo_select'
    )

    # Crear columna 'Periodo'
    try:
        if opcion == "Día":
            df_processed['Periodo'] = df_processed[date_col].dt.date
        elif opcion == "Mes":
            df_processed['Periodo'] = df_processed[date_col].dt.to_period('M').astype(str)
        elif opcion == "Año":
            df_processed['Periodo'] = df_processed[date_col].dt.year
        df_processed['Periodo'] = df_processed['Periodo'].astype(str) # Asegurar que sea string para agrupar/filtrar
    except Exception as e:
        st.error(f"Error al crear columna 'Periodo': {e}")
        return

    # Agrupar por periodo
    resumen = df_processed.groupby('Periodo', as_index=False)[value_col].sum()
    resumen[f'{value_col} Pesos'] = resumen[value_col].apply(format_currency)
    resumen = resumen.sort_values('Periodo') # Ordenar para gráficos

    # --- Filtro Multiselect y KPI ---
    periodos_disponibles = resumen['Periodo'].unique().tolist()
    periodos_seleccionados = st.multiselect(
        f"Selecciona los {opcion.lower()}s que quieres analizar:",
        options=periodos_disponibles,
        default=periodos_disponibles,
        key=f'{key_prefix}_periodo_multi'
    )

    resumen_filtrado = resumen[resumen['Periodo'].isin(periodos_seleccionados)]
    total_seleccionado = resumen_filtrado[value_col].sum()

    st.metric(
        label=f"💰 Total {tab_title} ({opcion}s seleccionados)",
        value=format_currency(total_seleccionado)
    )

    # --- Visualizaciones (Primera Fila) ---
    col1, col2 = st.columns([1, 2])
    with col1:
        st.subheader(f"Total por {opcion}")
        st.dataframe(resumen_filtrado[['Periodo', f'{value_col} Pesos']])
    with col2:
        st.subheader(f"Evolución por {opcion}")
        if not resumen_filtrado.empty:
            fig_periodo = px.line(
                resumen_filtrado,
                x='Periodo',
                y=value_col,
                markers=True,
                title=f'Evolución {tab_title} por {opcion}'
            )
            fig_periodo.update_layout(
                xaxis_title=opcion,
                yaxis_title=f'Total {tab_title} ($)',
                yaxis_tickformat=',.0f'
            )
            st.plotly_chart(fig_periodo, use_container_width=True)
        else:
            st.info("No hay datos para mostrar en el gráfico con los filtros seleccionados.")

    # --- Visualizaciones (Segunda Fila) ---
    col3, col4 = st.columns(2) # Usamos col4 para mantener la estructura, aunque no se use explícitamente
    with col3:
        st.subheader(f"📈 Evolución Total {tab_title} en el Tiempo")

        # KPI General
        total_general = df_processed[value_col].sum()
        st.metric(label=f"💰 Total General {tab_title}", value=format_currency(total_general))

        # Gráfico de línea de evolución general (Corrección para fig2/equivalente)
        df_sorted = df_processed.sort_values(date_col)
        fig_evolucion = px.line(
            df_sorted,
            x=date_col,
            y=value_col,
            title=f'Evolución Total {tab_title} en el Tiempo'
        )
        fig_evolucion.update_layout(
            xaxis_title='Fecha',
            yaxis_title='Total ($)',
            yaxis_tickformat=',.0f'
        )
        st.plotly_chart(fig_evolucion, use_container_width=True)

# --- Función para Pestañas de Búsqueda y Exportación ---
def render_search_export_tab(df, date_col, value_col, cliente_col, comercial_col, cierre_col, req_col, obs_col, other_numeric_cols, pivot_agg_cols, tab_title, key_prefix):
    """
    Renderiza una pestaña con filtros, tabla de resultados, gráficos y opción de exportar.

    Args:
        df (pd.DataFrame): DataFrame con los datos.
        date_col (str): Nombre de la columna de fecha.
        value_col (str): Nombre de la columna principal de valor (e.g., Subtotal).
        cliente_col (str): Nombre de la columna de cliente.
        comercial_col (str): Nombre de la columna de comercial.
        cierre_col (str): Nombre de la columna de cierre facturación.
        req_col (str): Nombre de la columna de requerimiento especial.
        obs_col (str): Nombre de la columna de observaciones.
        other_numeric_cols (list): Lista de otras columnas numéricas para sumar en exportación.
        pivot_agg_cols (dict): Dict {col_name: agg_func} para la tabla dinámica (e.g., {'Peso CP': 'sum'}).
        tab_title (str): Título base para la pestaña.
        key_prefix (str): Prefijo único para las claves de los widgets de Streamlit.
    """
    st.subheader(f"🔍 {tab_title}")

    if df is None or df.empty:
        st.warning(f"No se encontraron datos para '{tab_title}'. Verifica la hoja en el Excel.")
        return

    df_original = df.copy() # Mantener original para filtros dinámicos
    df_original[date_col] = safe_to_datetime(df_original[date_col]) # Asegurar fecha para filtro
    df_original[value_col] = safe_to_numeric(df_original[value_col]) # Asegurar numérico para sumas

    # --- Filtros ---
    df_filtrado = df_original.copy()

    # Filtro Cliente (Texto)
    empresa_buscada = st.text_input("Ingrese el nombre del Cliente:", key=f'{key_prefix}_cliente_input')
    if empresa_buscada:
        try:
            df_filtrado = df_filtrado[df_filtrado[cliente_col].str.contains(empresa_buscada, case=False, na=False)]
        except KeyError:
            st.warning(f"Columna '{cliente_col}' no encontrada para filtrar por cliente.")
        except Exception as e:
            st.error(f"Error al filtrar por cliente: {e}")


    # Filtro Cierre Facturación (Selectbox) - Con estado de sesión
    if f"{key_prefix}_cierre_sel" not in st.session_state:
        st.session_state[f"{key_prefix}_cierre_sel"] = "Todos"
    try:
        cierres_disponibles = ["Todos"] + df_original[cierre_col].dropna().unique().tolist()
        # Asegurar que el valor guardado exista en las opciones actuales
        current_selection_cierre = st.session_state[f"{key_prefix}_cierre_sel"]
        if current_selection_cierre not in cierres_disponibles:
            current_selection_cierre = "Todos"

        cierre_seleccionado = st.selectbox(
            "Selecciona Cierre de Facturación:",
            cierres_disponibles,
            index=cierres_disponibles.index(current_selection_cierre),
            key=f'{key_prefix}_cierre_selectbox'
        )
        st.session_state[f"{key_prefix}_cierre_sel"] = cierre_seleccionado # Guardar selección
        if cierre_seleccionado != "Todos":
            df_filtrado = df_filtrado[df_filtrado[cierre_col] == cierre_seleccionado]
    except KeyError:
        st.warning(f"Columna '{cierre_col}' no encontrada para filtrar por cierre.")
    except Exception as e:
         st.error(f"Error al filtrar por cierre: {e}")

    # Filtro Comercial (Selectbox) - Con estado de sesión
    if f"{key_prefix}_comercial_sel" not in st.session_state:
        st.session_state[f"{key_prefix}_comercial_sel"] = "Todos"
    try:
        comerciales_disponibles = ["Todos"] + df_original[comercial_col].dropna().unique().tolist()
        current_selection_comercial = st.session_state[f"{key_prefix}_comercial_sel"]
        if current_selection_comercial not in comerciales_disponibles:
            current_selection_comercial = "Todos"

        comercial_seleccionado = st.selectbox(
            "Selecciona el Comercial:",
            comerciales_disponibles,
            index=comerciales_disponibles.index(current_selection_comercial),
            key=f'{key_prefix}_comercial_selectbox'
        )
        st.session_state[f"{key_prefix}_comercial_sel"] = comercial_seleccionado
        if comercial_seleccionado != "Todos":
            df_filtrado = df_filtrado[df_filtrado[comercial_col] == comercial_seleccionado]
    except KeyError:
         st.warning(f"Columna '{comercial_col}' no encontrada para filtrar por comercial.")
    except Exception as e:
         st.error(f"Error al filtrar por comercial: {e}")


    # Filtro Requerimiento Especial (Multiselect) - Con estado de sesión
    if f"{key_prefix}_req_sel" not in st.session_state:
        st.session_state[f"{key_prefix}_req_sel"] = []
    try:
        requerimientos_disponibles = df_original[req_col].dropna().unique().tolist()
        # Mantener solo selecciones previas que aún existen
        req_previos = [req for req in st.session_state[f"{key_prefix}_req_sel"] if req in requerimientos_disponibles]

        requerimiento_seleccionado = st.multiselect(
            "Selecciona Requerimiento(s) Especial(es):",
            requerimientos_disponibles,
            default=req_previos,
            key=f'{key_prefix}_req_multiselect'
        )
        st.session_state[f"{key_prefix}_req_sel"] = requerimiento_seleccionado
        if requerimiento_seleccionado:
            df_filtrado = df_filtrado[df_filtrado[req_col].isin(requerimiento_seleccionado)]
    except KeyError:
        st.warning(f"Columna '{req_col}' no encontrada para filtrar por requerimiento.")
    except Exception as e:
         st.error(f"Error al filtrar por requerimiento: {e}")

    # Filtro Observaciones (Multiselect) - Con estado de sesión
    if f"{key_prefix}_obs_sel" not in st.session_state:
        st.session_state[f"{key_prefix}_obs_sel"] = []
    try:
        observaciones_disponibles = df_original[obs_col].dropna().unique().tolist()
        obs_previas = [obs for obs in st.session_state[f"{key_prefix}_obs_sel"] if obs in observaciones_disponibles]

        observaciones_seleccionadas = st.multiselect(
            "Selecciona Observación(es):",
            observaciones_disponibles,
            default=obs_previas,
            key=f'{key_prefix}_obs_multiselect'
        )
        st.session_state[f"{key_prefix}_obs_sel"] = observaciones_seleccionadas
        if observaciones_seleccionadas:
            df_filtrado = df_filtrado[df_filtrado[obs_col].isin(observaciones_seleccionadas)]
    except KeyError:
        st.warning(f"Columna '{obs_col}' no encontrada para filtrar por observaciones.")
    except Exception as e:
         st.error(f"Error al filtrar por observaciones: {e}")

    # Filtro Fecha (Date Input)
    min_fecha_disp = df_original[date_col].min()
    max_fecha_disp = df_original[date_col].max()

    if pd.notna(min_fecha_disp) and pd.notna(max_fecha_disp):
      fecha_rango = st.date_input(
          "Selecciona un rango de fechas (opcional):",
          [], # Por defecto vacío
          min_value=min_fecha_disp.date(),
          max_value=max_fecha_disp.date(),
          key=f'{key_prefix}_date_input'
      )
      if fecha_rango and len(fecha_rango) == 2:
          fecha_inicio, fecha_fin = fecha_rango
          # Asegurar que las fechas de filtro sean datetime para comparar
          fecha_inicio_dt = pd.to_datetime(fecha_inicio)
          fecha_fin_dt = pd.to_datetime(fecha_fin) + pd.Timedelta(days=1) # Incluir el día final
          df_filtrado = df_filtrado[
              (df_filtrado[date_col] >= fecha_inicio_dt) &
              (df_filtrado[date_col] < fecha_fin_dt) # Usar < para incluir hasta fin del día
          ]
    else:
        st.info("No hay fechas válidas en los datos para habilitar el filtro de fecha.")


    # --- Resultados y Visualización ---
    if not df_filtrado.empty:
        st.write("### Resultados de la búsqueda")
        # Asegurar que value_col es numérico antes de formatear y mostrar
        df_display = df_filtrado.copy()
        if value_col in df_display.columns:
             df_display[value_col] = safe_to_numeric(df_display[value_col])
             df_display[f'{value_col} Formato'] = df_display[value_col].apply(format_currency)
        st.dataframe(df_display)

        # --- KPIs y Gráficos ---
        if value_col in df_filtrado.columns and pd.api.types.is_numeric_dtype(df_filtrado[value_col]):
            total_facturado_filtrado = df_filtrado[value_col].sum()
            st.metric(f"💰 Total {value_col} (Filtrado)", format_currency(total_facturado_filtrado))

            col_graf1, col_graf2 = st.columns(2)

            with col_graf1:
                # Gráfico Torta por Comercial
                if comercial_col in df_filtrado.columns:
                    df_comercial = df_filtrado.groupby(comercial_col, as_index=False)[value_col].sum()
                    if not df_comercial.empty:
                        fig_torta = px.pie(
                            df_comercial,
                            names=comercial_col,
                            values=value_col,
                            title=f'Distribución por {comercial_col}'
                        )
                        st.plotly_chart(fig_torta, use_container_width=True)
                else:
                    st.warning(f"Columna '{comercial_col}' no encontrada para gráfico de torta.")

            with col_graf2:
                 # Gráfico Barras por Comercial
                if comercial_col in df_filtrado.columns:
                    df_comercial = df_filtrado.groupby(comercial_col, as_index=False)[value_col].sum().sort_values(value_col, ascending=False)
                    if not df_comercial.empty:
                        fig_barras_com = px.bar(
                            df_comercial,
                            x=comercial_col,
                            y=value_col,
                            text_auto='.2s',
                            title=f'{value_col} por {comercial_col}'
                        )
                        fig_barras_com.update_traces(textposition='outside')
                        fig_barras_com.update_layout(yaxis_title=f'{value_col} ($)', uniformtext_minsize=8, uniformtext_mode='hide')
                        st.plotly_chart(fig_barras_com, use_container_width=True)
                else:
                    st.warning(f"Columna '{comercial_col}' no encontrada para gráfico de barras.")


            # Gráfico Pareto por Cliente
            if cliente_col in df_filtrado.columns:
                df_cliente = df_filtrado.groupby(cliente_col, as_index=False)[value_col].sum()
                df_cliente = df_cliente.sort_values(by=value_col, ascending=False)

                if not df_cliente.empty and df_cliente[value_col].sum() > 0: # Evitar división por cero
                    df_cliente['% Acumulado'] = df_cliente[value_col].cumsum() / df_cliente[value_col].sum() * 100

                    fig_pareto = px.bar(
                        df_cliente,
                        x=cliente_col,
                        y=value_col,
                        text_auto='.2s',
                        title=f'{value_col} por {cliente_col} (Pareto)'
                    )
                    fig_pareto.add_trace(go.Scatter(
                        x=df_cliente[cliente_col],
                        y=df_cliente['% Acumulado'],
                        mode='lines+markers',
                        name='% Acumulado',
                        yaxis='y2' # Usar eje secundario
                    ))
                    fig_pareto.update_layout(
                        yaxis_title=f'{value_col} ($)',
                        yaxis2=dict(
                            title='% Acumulado',
                            overlaying='y',
                            side='right',
                            range=[0, 105], # Rango 0-100%
                            tickformat=".0f", # Formato sin decimales para porcentaje
                            ticksuffix="%" # Añadir símbolo %
                        ),
                        xaxis={'categoryorder':'total descending'}, # Ordenar barras
                        legend=dict(yanchor="top", y=0.99, xanchor="left", x=0.01)
                    )
                    st.plotly_chart(fig_pareto, use_container_width=True)
                elif not df_cliente.empty:
                     st.info("No se puede generar Pareto (total es cero o negativo).")
            else:
                st.warning(f"Columna '{cliente_col}' no encontrada para gráfico Pareto.")

            # --- Exportación a Excel ---
            st.write("### Exportar Resultados Filtrados")
            columnas_disponibles_export = df_filtrado.columns.tolist()
            columnas_seleccionadas_export = st.multiselect(
                "Selecciona las columnas a exportar:",
                columnas_disponibles_export,
                default=columnas_disponibles_export,
                key=f'{key_prefix}_export_cols'
            )

            if columnas_seleccionadas_export:
                df_exportar = df_filtrado[columnas_seleccionadas_export]
                output = BytesIO()
                try:
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        # Hoja Consolidado
                        df_exportar.to_excel(writer, index=False, sheet_name='Consolidado')
                        workbook = writer.book
                        worksheet = writer.sheets['Consolidado']

                        # Formatos
                        header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'fg_color': '#D7E4BC', 'border': 1})
                        currency_format = workbook.add_format({'num_format': '$#,##0', 'border': 1})
                        default_format = workbook.add_format({'border': 1}) # Formato base con bordes

                        # Aplicar formato a encabezados y ajustar ancho
                        for col_num, value in enumerate(df_exportar.columns.values):
                            worksheet.write(0, col_num, value, header_format)
                            column_len = max(df_exportar[value].astype(str).map(len).max(), len(value))
                            worksheet.set_column(col_num, col_num, column_len + 2) # +2 para margen

                        # Aplicar formato a celdas (moneda o por defecto)
                        cols_currency = [value_col] + [col for col in other_numeric_cols if col in df_exportar.columns] # Columnas de moneda
                        for row_num in range(len(df_exportar)):
                             for col_num, col_name in enumerate(df_exportar.columns):
                                 cell_value = df_exportar.iloc[row_num, col_num]
                                 if pd.isna(cell_value):
                                     worksheet.write(row_num + 1, col_num, '', default_format) # Celda vacía con borde
                                 elif col_name in cols_currency and pd.api.types.is_number(cell_value):
                                     worksheet.write_number(row_num + 1, col_num, cell_value, currency_format)
                                 elif isinstance(cell_value, (datetime.date, datetime.datetime)):
                                      date_format = workbook.add_format({'num_format': 'yyyy-mm-dd', 'border': 1})
                                      worksheet.write_datetime(row_num + 1, col_num, cell_value, date_format)
                                 elif pd.api.types.is_number(cell_value):
                                      number_format = workbook.add_format({'num_format': '0', 'border': 1}) # Números sin formato moneda
                                      worksheet.write_number(row_num + 1, col_num, cell_value, number_format)
                                 else:
                                     worksheet.write_string(row_num + 1, col_num, str(cell_value), default_format) # Como texto

                        # Añadir fila de totales
                        row_total = len(df_exportar) + 1
                        worksheet.write(row_total, 0, 'TOTAL', header_format) # Etiqueta TOTAL
                        for col_num, col_name in enumerate(df_exportar.columns):
                            if col_name in cols_currency and pd.api.types.is_numeric_dtype(df_exportar[col_name].dropna()):
                                col_letter = chr(65 + col_num)
                                formula = f"=SUM({col_letter}2:{col_letter}{row_total})"
                                worksheet.write_formula(row_total, col_num, formula, currency_format)
                            # Opcional: Sumar otras columnas numéricas si se desea
                            # elif col_name in other_numeric_cols and pd.api.types.is_numeric_dtype(df_exportar[col_name].dropna()):
                            #    col_letter = chr(65 + col_num)
                            #    formula = f"=SUM({col_letter}2:{col_letter}{row_total})"
                            #    # Definir un formato numérico si es necesario
                            #    worksheet.write_formula(row_total, col_num, formula, default_format)

                        # Hoja Tabla Dinámica (simplificada)
                        pivot_sheet = workbook.add_worksheet('Tabla_Dinamica_Resumen')
                        pivot_group_col = 'Residuo' # Columna principal para agrupar (ajustar si es necesario)

                        if pivot_group_col in df_filtrado.columns:
                            # Definir las agregaciones deseadas
                            agg_dict = {value_col: 'sum'} # Sumar la columna de valor principal por defecto
                            valid_pivot_agg_cols = {}
                            for col, agg_func in pivot_agg_cols.items():
                                if col in df_filtrado.columns and pd.api.types.is_numeric_dtype(df_filtrado[col].dropna()):
                                    valid_pivot_agg_cols[col] = agg_func

                            agg_dict.update(valid_pivot_agg_cols) # Añadir otras agregaciones numéricas válidas

                            try:
                                pivot_data = df_filtrado.groupby(pivot_group_col, as_index=False).agg(agg_dict)
                                pivot_data = pivot_data.sort_values(value_col, ascending=False)

                                # Escribir tabla dinámica
                                pivot_sheet.write_row(0, 0, pivot_data.columns, header_format)
                                for idx, row in pivot_data.iterrows():
                                    for col_idx, col_name in enumerate(pivot_data.columns):
                                        cell_val = row[col_name]
                                        if col_name == value_col:
                                            pivot_sheet.write(idx + 1, col_idx, cell_val, currency_format)
                                        elif col_name in valid_pivot_agg_cols and pd.api.types.is_number(cell_val):
                                            # Podrías definir formatos específicos si es necesario
                                            pivot_sheet.write_number(idx + 1, col_idx, cell_val, default_format)
                                        else:
                                            pivot_sheet.write(idx + 1, col_idx, cell_val, default_format)
                                # Ajustar anchos en tabla dinámica
                                for col_num, value in enumerate(pivot_data.columns.values):
                                    column_len = max(pivot_data[value].astype(str).map(len).max(), len(value))
                                    pivot_sheet.set_column(col_num, col_num, column_len + 2)

                            except Exception as pivot_error:
                                pivot_sheet.write(0, 0, f"Error al generar tabla dinámica: {pivot_error}")

                        else:
                            pivot_sheet.write(0, 0, f"No se pudo generar la tabla dinámica.")
                            pivot_sheet.write(1, 0, f"Revisa que exista la columna '{pivot_group_col}'.")

                    # Botón de Descarga
                    st.download_button(
                        label=f"💾 Descargar {tab_title} Filtrado",
                        data=output.getvalue(),
                        file_name=f"{key_prefix}_filtrado.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as export_error:
                    st.error(f"Error al generar el archivo Excel: {export_error}")

            else:
                st.warning("Selecciona al menos una columna para exportar.")

        else:
            st.warning(f"La columna '{value_col}' necesaria para los cálculos no es numérica o no existe.")

    else:
        st.warning("No se encontraron resultados para los filtros seleccionados.")


# --- Estilos CSS ---
st.markdown("""
    <style>
        /* Estilo para Botones */
        .stButton>button {
            background-color: #1DB954; /* Verde neón tipo smartwatch */
            color: white;
            font-size: 16px;
            border-radius: 12px;
            padding: 10px 20px;
            border: none;
            transition: all 0.3s ease-in-out;
        }
        .stButton>button:hover {
            background-color: #17a2b8; /* Cambio de color */
            transform: scale(1.05);
        }
        /* Estilo para Métricas */
        div[data-testid="stMetricValue"] {
            font-size: 24px;
            font-weight: bold;
            color: #1DB954; /* Verde neón */
        }
        div[data-testid="stMetricLabel"] {
            font-size: 16px;
            font-weight: bold;
            color: white; /* Cambiado a blanco para mejor contraste en tema oscuro */
        }
        /* Estilo para Sidebar */
        section[data-testid="stSidebar"] {
            background-color: #181818 !important; /* Gris oscuro, !important puede ser necesario */
            border-radius: 10px;
            padding: 20px;
        }
         /* Estilo General del Dashboard (asegurar fondo oscuro) */
        body {
            color: white; /* Color de texto general */
        }
        /* Fondo oscuro para el área principal */
         .main .block-container {
            background-color: #121212;
             color: white;
         }
         /* Texto en widgets de sidebar blanco */
         section[data-testid="stSidebar"] .stTextInput label,
         section[data-testid="stSidebar"] .stFileUploader label,
         section[data-testid="stSidebar"] .stSelectbox label,
         section[data-testid="stSidebar"] .stDateInput label,
         section[data-testid="stSidebar"] .stMultiselect label,
         section[data-testid="stSidebar"] p,
         section[data-testid="stSidebar"] small {
              color: white !important;
          }
         /* Placeholder texto input sidebar gris claro */
          section[data-testid="stSidebar"] .stTextInput input {
              color: #ccc !important; /* O blanco si prefieres */
          }

    </style>
""", unsafe_allow_html=True)


# --- Carga de Archivo ---
archivo = st.sidebar.file_uploader("📂 Sube el archivo Excel (.xlsm)", type=["xlsm"])

# Diccionario para mapear nombres esperados a los nombres/índices reales de las hojas
# AJUSTA ESTOS NOMBRES SI LOS DE TU EXCEL SON DIFERENTES
sheet_names_map = {
    "posibles": "POSIBLES", # Nombre real de la hoja 0
    "enviados": "ENVIADOS", # Nombre real de la hoja 1
    "metas": "METAS",      # Nombre real de la hoja 2
    "asivamos": "ASI VAMOS", # Nombre real de la hoja 3
    "aprovechables": "APROVECHABLES" # Nombre real de la hoja 4
}
# También podrías usar índices si los nombres no son fiables:
# sheet_names_map = {
#     "posibles": 0,
#     "enviados": 1,
#     "metas": 2,
#     "asivamos": 3,
#     "aprovechables": 4
# }


# --- Procesamiento Principal ---
if archivo is not None:
    all_dfs = {}
    try:
        # Leer todas las hojas necesarias una sola vez
        excel_file = pd.ExcelFile(archivo)
        for key, sheet_identifier in sheet_names_map.items():
            try:
                 all_dfs[key] = excel_file.parse(sheet_identifier)
                 st.sidebar.success(f"✅ Hoja '{sheet_identifier}' cargada.")
            except Exception as e:
                 st.sidebar.warning(f"⚠️ No se pudo cargar la hoja '{sheet_identifier}': {e}. Se usará un DataFrame vacío.")
                 all_dfs[key] = pd.DataFrame() # Usar DF vacío si falla la carga

        # Asignar a variables específicas para mayor claridad (opcional)
        df_posibles = all_dfs.get("posibles")
        df_enviados = all_dfs.get("enviados")
        df_metas = all_dfs.get("metas")
        df_asivamos = all_dfs.get("asivamos")
        df_aprovechables = all_dfs.get("aprovechables")

        # --- Definición de Pestañas ---
        tab_keys = ["Posibles", "Enviados", "Metas", "Buscar RESPEL", "Así va facturación", "Aprovechables", "Buscar Aprovechables", "Resumen Comerciales"]
        tabs = st.tabs(tab_keys)

        # --- Contenido de Pestañas ---
        with tabs[0]: # Posibles
             # Ajusta los nombres de columna si son diferentes en tu hoja "Posibles"
             render_summary_tab(
                 df=df_posibles,
                 date_col='Fecha CC',
                 value_col='Subtotal',
                 tab_title='Posibles por Facturar',
                 key_prefix='posibles'
             )

        with tabs[1]: # Enviados
            st.subheader("📊 Datos y Resumen - Enviados a Facturar")
            if df_enviados is not None and not df_enviados.empty:
                st.dataframe(df_enviados.head())

                # Procesar datos específicos de Enviados
                df_enviados_proc = df_enviados.copy()
                 # AJUSTA los nombres de columna si son diferentes
                date_col_env = 'Dia'
                value_col_env = 'Subtotal'
                df_enviados_proc[date_col_env] = safe_to_datetime(df_enviados_proc[date_col_env])
                df_enviados_proc[value_col_env] = safe_to_numeric(df_enviados_proc[value_col_env])
                df_enviados_filtrado = df_enviados_proc.dropna(subset=[date_col_env, value_col_env])

                if not df_enviados_filtrado.empty:
                    # Selección del tipo de período
                    opcion_enviados = st.selectbox("Selecciona el período para agrupar:", ("Día", "Mes", "Año"), key='enviados_periodo_select')

                    # Generar columna de período
                    try:
                        if opcion_enviados == "Día":
                            df_enviados_filtrado['Periodo'] = df_enviados_filtrado[date_col_env].dt.date.astype(str)
                        elif opcion_enviados == "Mes":
                            df_enviados_filtrado['Periodo'] = df_enviados_filtrado[date_col_env].dt.to_period('M').astype(str)
                        elif opcion_enviados == "Año":
                            df_enviados_filtrado['Periodo'] = df_enviados_filtrado[date_col_env].dt.year.astype(str)
                    except Exception as e:
                        st.error(f"Error al crear columna 'Periodo' para Enviados: {e}")
                        st.stop() # Detener si hay error crítico aquí

                    # Agrupar datos
                    resumen_enviados = df_enviados_filtrado.groupby('Periodo', as_index=False)[value_col_env].sum().sort_values('Periodo')
                    resumen_enviados[f'{value_col_env} Pesos'] = resumen_enviados[value_col_env].apply(format_currency)

                    total_general_enviados = resumen_enviados[value_col_env].sum()

                    # Mostrar el resumen general y el total general
                    col5, col6 = st.columns([1, 2])
                    with col5:
                        st.subheader(f"Total enviados por {opcion_enviados}")
                        st.dataframe(resumen_enviados[['Periodo', f'{value_col_env} Pesos']])
                        st.success(f"✅ Total general enviados: **{format_currency(total_general_enviados)}**")

                        # Filtro adicional opcional
                        periodos_disponibles_env = resumen_enviados['Periodo'].tolist()
                        seleccion_periodos_env = st.multiselect(f"(Opcional) Selecciona {opcion_enviados.lower()}s para sumar:", periodos_disponibles_env, key='enviados_multi_select')
                        if seleccion_periodos_env:
                            resumen_filtrado_env = resumen_enviados[resumen_enviados['Periodo'].isin(seleccion_periodos_env)]
                            total_seleccionado_env = resumen_filtrado_env[value_col_env].sum()
                            st.info(f"🔎 Total en {opcion_enviados.lower()}s seleccionados: **{format_currency(total_seleccionado_env)}**")

                    with col6:
                        st.subheader(f"Evolución por {opcion_enviados}")
                        if not resumen_enviados.empty:
                            fig3 = px.line(resumen_enviados, x='Periodo', y=value_col_env, markers=True,
                                        title=f'Total enviados a facturar por {opcion_enviados}')
                            fig3.update_layout(xaxis_title=opcion_enviados, yaxis_title=f'Total {value_col_env} ($)', yaxis_tickformat=',.0f')
                            st.plotly_chart(fig3, use_container_width=True)

                    # Segunda fila: Evolución general (Corrección fig4)
                    col7, col8 = st.columns(2)
                    with col7:
                         st.subheader("📈 Evolución Total Enviados en el Tiempo")
                         total_gral_enviados_kpi = df_enviados_filtrado[value_col_env].sum()
                         st.metric(label=f"💰 Total General Enviados", value=format_currency(total_gral_enviados_kpi))

                         df_env_sorted = df_enviados_filtrado.sort_values(date_col_env)
                         fig4 = px.line( # Definición de fig4
                               df_env_sorted,
                               x=date_col_env,
                               y=value_col_env,
                               title='Evolución Total Enviados en el Tiempo'
                         )
                         fig4.update_layout(xaxis_title='Fecha', yaxis_title='Total ($)', yaxis_tickformat=',.0f')
                         st.plotly_chart(fig4, use_container_width=True) # Mostrar fig4

                else:
                    st.warning("No hay datos válidos (fecha y valor) para procesar en 'Enviados'.")
            else:
                st.warning("No se encontraron datos para 'Enviados'. Verifica la hoja en el Excel.")


        with tabs[2]: # Metas
            st.subheader("🎯 Comparativo Metas vs. Facturación Real")

            if df_metas is None or df_metas.empty or df_asivamos is None or df_asivamos.empty:
                 st.warning("Faltan datos de 'Metas' o 'AsiVamos' para el comparativo. Verifica las hojas en el Excel.")
            else:
                try:
                    # Asumiendo que df_metas tiene meses como columnas (Enero, Febrero...)
                    st.write("Metas Mensuales (del Excel):")
                    st.dataframe(df_metas)

                    # Transformar metas a formato largo
                    df_metas_long = df_metas.melt(var_name='Mes', value_name='Meta Facturación')
                    meses_ordenados = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
                                     'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
                    # Crear diccionario para mapeo seguro
                    mes_a_numero = {mes: i + 1 for i, mes in enumerate(meses_ordenados)}
                    df_metas_long['Mes_Numero'] = df_metas_long['Mes'].map(mes_a_numero)
                    # Filtrar filas donde el mes no fue reconocido (si melt() incluyó otras columnas)
                    df_metas_long = df_metas_long.dropna(subset=['Mes_Numero'])
                    df_metas_long['Mes_Numero'] = df_metas_long['Mes_Numero'].astype(int)

                    # Procesar df_asivamos (facturación real)
                    df_asivamos_proc = df_asivamos.copy()
                    # AJUSTA los nombres de columna si son diferentes
                    fecha_real_col = 'CreaFecha'
                    valor_real_col = 'Total'
                    df_asivamos_proc[fecha_real_col] = safe_to_datetime(df_asivamos_proc[fecha_real_col])
                    df_asivamos_proc[valor_real_col] = safe_to_numeric(df_asivamos_proc[valor_real_col])
                    df_asivamos_proc = df_asivamos_proc.dropna(subset=[fecha_real_col, valor_real_col])

                    # Agrupar facturación real por mes
                    df_asivamos_proc['Mes_Numero'] = df_asivamos_proc[fecha_real_col].dt.month
                    facturado_por_mes = df_asivamos_proc.groupby('Mes_Numero', as_index=False)[valor_real_col].sum()
                    facturado_por_mes.rename(columns={valor_real_col: 'Facturado'}, inplace=True)

                    # Unir metas y facturado
                    comparacion = pd.merge(df_metas_long, facturado_por_mes, on='Mes_Numero', how='left')
                    comparacion['Facturado'] = comparacion['Facturado'].fillna(0) # Meses sin facturación -> 0

                    # Calcular cumplimiento
                    # Evitar división por cero si la meta es 0
                    comparacion['% Cumplimiento'] = comparacion.apply(
                        lambda row: (row['Facturado'] / row['Meta Facturación'] * 100) if row['Meta Facturación'] else 0,
                        axis=1
                    )

                    # Formato para mostrar
                    comparacion['Meta Facturación $'] = comparacion['Meta Facturación'].apply(format_currency)
                    comparacion['Facturado $'] = comparacion['Facturado'].apply(format_currency)
                    comparacion['% Cumplimiento Format'] = comparacion['% Cumplimiento'].apply(lambda x: f"{x:.1f}%")

                    # Ordenar por mes y mostrar tabla
                    comparacion = comparacion.sort_values('Mes_Numero')
                    comparacion_final = comparacion[['Mes', 'Meta Facturación $', 'Facturado $', '% Cumplimiento Format']]
                    st.dataframe(comparacion_final)

                    # Gráficos
                    # Gráfico de barras comparativo
                    fig_comp = px.bar(
                        comparacion, x='Mes', y=['Meta Facturación', 'Facturado'],
                        barmode='group', title='Meta vs Facturación Real por Mes',
                        labels={'value': 'Monto ($)', 'variable': 'Tipo'}
                    )
                    fig_comp.update_layout(yaxis_tickformat=',.0f')
                    st.plotly_chart(fig_comp, use_container_width=True)

                    # Gráfico de línea de cumplimiento
                    fig_cump = px.line(
                        comparacion, x='Mes', y='% Cumplimiento', markers=True,
                        title='% de Cumplimiento de Meta por Mes'
                    )
                    fig_cump.update_layout(yaxis_title='Cumplimiento (%)', yaxis_ticksuffix='%', yaxis_range=[0, comparacion['% Cumplimiento'].max() * 1.1]) # Ajustar rango Y
                    st.plotly_chart(fig_cump, use_container_width=True)

                except KeyError as ke:
                    st.error(f"Error de columna procesando Metas/AsiVamos: {ke}. Verifica los nombres de columna en tu Excel.")
                except Exception as e:
                    st.error(f"Error procesando la pestaña Metas: {e}")


        with tabs[3]: # Buscar RESPEL (Posibles)
            render_search_export_tab(
                df=df_posibles,
                date_col='Fecha CC', # Ajusta si es necesario
                value_col='Subtotal', # Ajusta si es necesario
                cliente_col='Cliente', # Ajusta si es necesario
                comercial_col='Comercial', # Ajusta si es necesario
                cierre_col='CIERRE DE FACTURACIÓN', # Ajusta si es necesario
                req_col='REQUERIMIENTO ESPECIAL', # Ajusta si es necesario
                obs_col='OBSERVACIONES', # Ajusta si es necesario
                other_numeric_cols=['Vlr Unit', 'Total'], # Columnas adicionales para sumar en exportación
                pivot_agg_cols={}, # No se especificó 'Peso CP' aquí, así que solo agregamos Subtotal
                tab_title='Buscar en RESPEL (Posibles)',
                key_prefix='buscar_respel'
            )

        with tabs[4]: # Así va facturación
            st.subheader("📊 Resumen de Facturación por Comercial (Así Vamos)")
            if df_asivamos is not None and not df_asivamos.empty:
                 # AJUSTA los nombres de columna si son diferentes
                fecha_asivamos_col = 'CreaFecha'
                valor_asivamos_col = 'Total'
                comercial_asivamos_col = 'COMERCIAL' # Asegúrate que este es el nombre correcto

                df_asivamos_proc = df_asivamos.copy()
                df_asivamos_proc[fecha_asivamos_col] = safe_to_datetime(df_asivamos_proc[fecha_asivamos_col])
                df_asivamos_proc[valor_asivamos_col] = safe_to_numeric(df_asivamos_proc[valor_asivamos_col])

                # Filtro de fecha
                min_fecha_av = df_asivamos_proc[fecha_asivamos_col].min()
                max_fecha_av = df_asivamos_proc[fecha_asivamos_col].max()

                fecha_inicio_av, fecha_fin_av = (min_fecha_av, max_fecha_av) # Valores por defecto iniciales

                if pd.notna(min_fecha_av) and pd.notna(max_fecha_av):
                    fecha_inicio_av_sel, fecha_fin_av_sel = st.date_input(
                        label="Selecciona el rango de fechas:",
                        value=(min_fecha_av.date(), max_fecha_av.date()), # Valores por defecto
                        min_value=min_fecha_av.date(),
                        max_value=max_fecha_av.date(),
                        key='asivamos_date_range'
                    )
                    # Convertir selección a datetime para filtrar
                    fecha_inicio_dt_av = pd.to_datetime(fecha_inicio_av_sel)
                    fecha_fin_dt_av = pd.to_datetime(fecha_fin_av_sel) + pd.Timedelta(days=1) # Incluir día final

                    filtro_fecha_av = (df_asivamos_proc[fecha_asivamos_col] >= fecha_inicio_dt_av) & (df_asivamos_proc[fecha_asivamos_col] < fecha_fin_dt_av)
                    df_filtrado_av = df_asivamos_proc[filtro_fecha_av]
                else:
                    st.warning("No hay fechas válidas para filtrar en 'AsiVamos'. Mostrando todos los datos.")
                    df_filtrado_av = df_asivamos_proc # Mostrar todo si no hay fechas válidas

                if not df_filtrado_av.empty:
                    st.dataframe(df_filtrado_av.head())

                    # Agrupar por comercial
                    if comercial_asivamos_col in df_filtrado_av.columns and valor_asivamos_col in df_filtrado_av.columns:
                        facturacion_comercial_av = df_filtrado_av.groupby(comercial_asivamos_col)[valor_asivamos_col].sum().reset_index()
                        facturacion_comercial_av['Total Formateado'] = facturacion_comercial_av[valor_asivamos_col].apply(format_currency)
                        facturacion_comercial_av = facturacion_comercial_av.sort_values(valor_asivamos_col, ascending=False)

                        st.subheader("Facturación por Comercial (Filtrado)")
                        st.dataframe(facturacion_comercial_av[[comercial_asivamos_col, 'Total Formateado']])

                        # Gráfico de barras
                        fig_av = px.bar(
                            facturacion_comercial_av,
                            x=comercial_asivamos_col, y=valor_asivamos_col,
                            color=valor_asivamos_col, # Colorear por valor
                            title='🏆 Total Facturación por Comercial (Periodo Seleccionado)',
                            text=facturacion_comercial_av['Total Formateado'] # Mostrar valor formateado
                        )
                        fig_av.update_layout(
                            xaxis_title='Comercial', yaxis_title='Total Valor Facturado ($)',
                            yaxis_tickformat=',.0f', template='plotly_white',
                             uniformtext_minsize=8, uniformtext_mode='hide'
                        )
                        fig_av.update_traces(textposition='outside')
                        st.plotly_chart(fig_av, use_container_width=True)

                        # Total general filtrado
                        total_general_av = facturacion_comercial_av[valor_asivamos_col].sum()
                        st.metric(label="💰 Total Facturación (Periodo Seleccionado)", value=format_currency(total_general_av))
                    else:
                        st.warning(f"No se encontraron las columnas '{comercial_asivamos_col}' o '{valor_asivamos_col}' en 'AsiVamos' para agrupar.")

                else:
                    st.info("No se encontraron datos de 'AsiVamos' para el rango de fechas seleccionado.")

            else:
                st.warning("No se encontraron datos para 'AsiVamos'. Verifica la hoja en el Excel.")


        with tabs[5]: # Aprovechables (Resumen)
             # Ajusta los nombres de columna si son diferentes en tu hoja "Aprovechables"
             render_summary_tab(
                 df=df_aprovechables,
                 date_col='Fecha CC',
                 value_col='Subtotalmen', # Cambiado a Subtotalmen
                 tab_title='Aprovechables por Facturar',
                 key_prefix='aprovechables'
             )

        with tabs[6]: # Buscar Aprovechables
             # Ajusta los nombres de columna si son diferentes en tu hoja "Aprovechables"
            render_search_export_tab(
                df=df_aprovechables,
                date_col='Fecha CC',
                value_col='Subtotalmen', # Cambiado a Subtotalmen
                cliente_col='Cliente',
                comercial_col='Comercial',
                cierre_col='CIERRE DE FACTURACIÓN',
                req_col='REQUERIMIENTO ESPECIAL', # Asumiendo que existe, si no, pasa '' o None
                obs_col='OBSERVACIONES', # Asumiendo que existe, si no, pasa '' o None
                other_numeric_cols=['Vlr Unit', 'Total', 'Peso CP'], # Añadido Peso CP
                pivot_agg_cols={'Peso CP': 'sum'}, # Agregación para la tabla dinámica
                tab_title='Buscar en Aprovechables',
                key_prefix='buscar_aprov'
            )

        with tabs[7]: # Resumen Comerciales
            st.subheader("📄 Resumen General por Comerciales")
            st.info("Esta sección está pendiente de desarrollo. ¿Qué resumen te gustaría ver aquí?")
            # Aquí podrías, por ejemplo, combinar datos de 'Posibles', 'Enviados', 'Aprovechables'
            # y agruparlos por comercial para tener una visión completa.

    except FileNotFoundError:
         st.error("⛔ Error: El archivo Excel subido no se encontró o no se pudo abrir.")
    except ValueError as ve:
         st.error(f"⛔ Error al leer el archivo Excel: {ve}. Asegúrate de que el archivo no esté corrupto y las hojas existan.")
    except Exception as e:
        st.error(f"⛔ Ocurrió un error inesperado al procesar el archivo: {e}")
        st.exception(e) # Muestra el traceback completo en la app para depuración

else:
    st.info("🔄 Por favor, sube un archivo Excel (.xlsm) en la barra lateral para comenzar.")
    st.warning("Asegúrate de que el archivo contenga las hojas: " + ", ".join(sheet_names_map.values()))
