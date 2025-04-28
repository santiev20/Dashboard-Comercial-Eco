import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
import plotly.graph_objects as go
import datetime
import os



st.set_page_config(page_title="Dashboard Comercial", layout="wide")

# ⬆️ Subida de archivo (sin administrador)
archivo = st.sidebar.file_uploader("📂 Sube el archivo Excel (.xlsm)", type=["xlsm"])

if archivo is not None:
    try:
        # Cargar hojas desde el archivo subido
        df_posibles = pd.read_excel(archivo, sheet_name=0)
        df_enviados = pd.read_excel(archivo, sheet_name=1)
        df_metas = pd.read_excel(archivo, sheet_name=2)
        df_asivamos = pd.read_excel(archivo, sheet_name=3)
        df_aprovechables = pd.read_excel(archivo, sheet_name=4)

        # ... aquí sigue todo tu dashboard normal con las pestañas

    except Exception as e:
        st.error(f"⛔ Hubo un problema al leer el archivo: {e}")
else:
    st.warning("🔄 Por favor, sube un archivo Excel (.xlsm) en la barra lateral para comenzar.")

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
            color: white;
        }

        /* Estilo para Sidebar */
        section[data-testid="stSidebar"] {
            background-color: #181818; /* Gris oscuro */
            border-radius: 10px;
            padding: 20px;
        }

        /* Estilo General del Dashboard */
        body {
            background-color: #121212; /* Fondo oscuro */
            color: white;
            font-family: 'Arial', sans-serif;
        }
    </style>
""", unsafe_allow_html=True)

if archivo is not None:
    # Cargar archivo directamente desde el servidor
    df_posibles = pd.read_excel(archivo, sheet_name=0)
    df_enviados = pd.read_excel(archivo, sheet_name=1)

    # Crear pestañas para separar la visualización de cada hoja
    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
    "Posibles", 
    "Enviados", 
    "Metas", 
    "Buscar RESPEL", 
    "Así va facturación", 
    "Aprovechables y ECOFUEL",
    "Buscar Aprovechables",
    "Resumen Comerciales"
])

    # ----------- TAB 1: POSIBLES -----------
    with tab1:
        st.subheader("📋 Datos de la hoja 1 - POSIBLES")
        st.dataframe(df_posibles.head())

        # Procesar datos
        df_posibles['Fecha CC'] = pd.to_datetime(df_posibles['Fecha CC'], errors='coerce')
        df_posibles['Subtotal'] = pd.to_numeric(df_posibles['Subtotal'], errors='coerce')
        df_posibles_filtrado = df_posibles.dropna(subset=['Fecha CC', 'Subtotal'])

        # Selección de agrupación por período
        opcion = st.selectbox("Selecciona el período para agrupar:", ("Día", "Mes", "Año"), key='posibles')

        # Crear columna 'Periodo' según la opción elegida
        if opcion == "Día":
            df_posibles_filtrado['Periodo'] = df_posibles_filtrado['Fecha CC'].dt.date
        elif opcion == "Mes":
            df_posibles_filtrado['Periodo'] = df_posibles_filtrado['Fecha CC'].dt.to_period('M').astype(str)
        elif opcion == "Año":
            df_posibles_filtrado['Periodo'] = df_posibles_filtrado['Fecha CC'].dt.year

        # Agrupar por periodo y sumar los subtotales
        resumen_posibles = df_posibles_filtrado.groupby('Periodo', as_index=False)['Subtotal'].sum()
        resumen_posibles['Subtotal Pesos'] = resumen_posibles['Subtotal'].apply(lambda x: f"${x:,.0f}")

        # Crear una lista con los periodos disponibles para filtrar
        periodos_disponibles = resumen_posibles['Periodo'].unique().tolist()

        # Multiselect para que el usuario elija los periodos a analizar
        periodos_seleccionados = st.multiselect(f"Selecciona los {opcion.lower()}s que quieres analizar:",
                                                options=periodos_disponibles,
                                                default=periodos_disponibles)

        # Filtrar el dataframe según la selección del usuario
        resumen_filtrado = resumen_posibles[resumen_posibles['Periodo'].isin(periodos_seleccionados)]

        # Calcular el total facturado para los periodos seleccionados
        total_seleccionado = resumen_filtrado['Subtotal'].sum()

        # Mostrar el KPI con el total facturado
        st.metric(label=f"💰 Total Posibles por facturar en los {opcion.lower()}s seleccionados", value=f"${total_seleccionado:,.0f}")

        # Dividir en columnas para tabla y gráfico
        col1, col2 = st.columns([1, 2])

        with col1:
            st.subheader(f"Total Posibles por facturar por {opcion}")
            st.dataframe(resumen_filtrado[['Periodo', 'Subtotal Pesos']])

        with col2:
            fig = px.line(resumen_filtrado.sort_values('Periodo'),
                        x='Periodo',
                        y='Subtotal',
                        markers=True,
                        title=f'Evolución posibles por {opcion}')
            fig.update_layout(
                xaxis_title=opcion,
                yaxis_title='Total en posibles ($)',
                yaxis_tickformat=',.0f'
            )
            st.plotly_chart(fig)

        # Segunda fila de columnas
        col3, col4 = st.columns(2)

        with col3:
            st.subheader("📈 Evolución de posibles en el tiempo (POSIBLES)")
            
            # KPI - Total Facturado POSIBLES
            total_posibles = df_posibles_filtrado['Subtotal'].sum()
            st.metric(label="💰 Total por enviar a facturar (POSIBLES)", value=f"${total_posibles:,.0f}")
            
            # Gráfico de línea de evolución
            fig2.update_layout(
                xaxis_title='Fecha',
                yaxis_title='Total Facturado ($)',
                yaxis_tickformat=',.0f'
            )
            
            st.plotly_chart(fig2)
          
    # ----------- TAB 2: ENVIADOS -----------
    with tab2:
        st.subheader("📋 Datos de la hoja 2 - ENVIADOS")
        st.dataframe(df_enviados.head())

        # Procesar datos
        df_enviados['Dia'] = pd.to_datetime(df_enviados['Dia'], errors='coerce')
        df_enviados['Subtotal'] = pd.to_numeric(df_enviados['Subtotal'], errors='coerce')
        df_enviados_filtrado = df_enviados.dropna(subset=['Dia', 'Subtotal'])

        # Selección del tipo de período
        opcion_enviados = st.selectbox("Selecciona el período para agrupar:", ("Día", "Mes", "Año"), key='enviados')

        # Generar columna de período
        if opcion_enviados == "Día":
            df_enviados_filtrado['Periodo'] = df_enviados_filtrado['Dia'].dt.date.astype(str)
        elif opcion_enviados == "Mes":
            df_enviados_filtrado['Periodo'] = df_enviados_filtrado['Dia'].dt.to_period('M').astype(str)
        elif opcion_enviados == "Año":
            df_enviados_filtrado['Periodo'] = df_enviados_filtrado['Dia'].dt.year.astype(str)

        # Agrupar datos
        resumen_enviados = df_enviados_filtrado.groupby('Periodo', as_index=False)['Subtotal'].sum()
        resumen_enviados['Subtotal Pesos'] = resumen_enviados['Subtotal'].apply(lambda x: f"${x:,.0f}")

        # 👉 Total general de enviados (sin necesidad de seleccionar)
        total_general_enviados = resumen_enviados['Subtotal'].sum()

        # Mostrar el resumen general y el total general
        col5, col6 = st.columns([1, 2])

        with col5:
            st.subheader(f"Total enviados a facturar por {opcion_enviados}")
            st.dataframe(resumen_enviados[['Periodo', 'Subtotal Pesos']])

            # 🔹 Mostrar el total general automáticamente
            st.success(f"✅ Total general enviados a facturar por {opcion_enviados}: **${total_general_enviados:,.0f}**")

            # ➕ Filtro adicional si el usuario quiere ver solo algunos períodos
            periodos_disponibles = resumen_enviados['Periodo'].tolist()
            seleccion_periodos = st.multiselect(f"(Opcional) Selecciona los {opcion_enviados.lower()} que quieres sumar:", periodos_disponibles)

            if seleccion_periodos:
                resumen_filtrado = resumen_enviados[resumen_enviados['Periodo'].isin(seleccion_periodos)]
                total_seleccionado = resumen_filtrado['Subtotal'].sum()

                st.info(f"🔎 Total enviados a facturar en los {opcion_enviados.lower()} seleccionados: **${total_seleccionado:,.0f}**")

        with col6:
            fig3 = px.line(resumen_enviados, x='Periodo', y='Subtotal', markers=True,
                        title=f'Total enviados a facturar por {opcion_enviados}')
            fig3.update_layout(xaxis_title=opcion_enviados, yaxis_title='Total enviados a facturar ($)', yaxis_tickformat=',.0f')
            st.plotly_chart(fig3)

        col7, col8 = st.columns(2)

        with col7:
            st.subheader("📈 Evolución de ventas en el tiempo (ENVIADOS)")
            fig4.update_layout(xaxis_title='Fecha', yaxis_title='Total Facturado ($)', yaxis_tickformat=',.0f')
            st.plotly_chart(fig4)

    with tab3:    
        # Leemos la hoja 3 del archivo
        df_metas = pd.read_excel(archivo, sheet_name=2)
        df_asivamos = pd.read_excel(archivo, sheet_name=3)

        # Mostramos para verificar
        st.subheader("🎯 Metas de Facturación por Mes")
        st.dataframe(df_metas)

        # Convertimos las columnas de meses en filas para hacer merge después
        df_metas_long = df_metas.melt(var_name='Mes', value_name='Meta Facturación')

        # Convertimos los nombres de los meses a números (enero=1, etc.)
        meses_ordenados = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
                        'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
        df_metas_long['Mes_Numero'] = df_metas_long['Mes'].apply(lambda x: meses_ordenados.index(x) + 1)

        # Filtramos si es necesario (aquí solo se mantiene la variable para seguir el flujo)
        df_asivamos_filtrado = df_asivamos.copy()

        #Convertir a datetime
        df_asivamos_filtrado['CreaFecha'] = pd.to_datetime(df_asivamos_filtrado['CreaFecha'], errors='coerce')

        # Extraer el mes desde la columna 'CreaFecha'
        df_asivamos_filtrado['Mes_Numero'] = df_asivamos_filtrado['CreaFecha'].dt.month

        # Agrupar por Mes_Numero y sumar el Total
        enviados_por_mes = df_asivamos_filtrado.groupby('Mes_Numero', as_index=False)['Total'].sum()

        # Cambiar nombre a la columna para dejarlo más claro
        enviados_por_mes.rename(columns={'Total': 'Facturado'}, inplace=True)

        # Hacemos el merge de metas y enviados
        comparacion = pd.merge(df_metas_long, enviados_por_mes, on='Mes_Numero', how='left')

        # Rellenamos NaN por si hay meses sin facturación
        comparacion['Facturado'] = comparacion['Facturado'].fillna(0)

        # Calculamos el porcentaje de cumplimiento
        comparacion['% Cumplimiento'] = (comparacion['Facturado'] / comparacion['Meta Facturación']) * 100

        # Opcional: Formato en pesos para mostrar
        comparacion['Meta Facturación $'] = comparacion['Meta Facturación'].apply(lambda x: f"${x:,.0f}")
        comparacion['Facturado $'] = comparacion['Facturado'].apply(lambda x: f"${x:,.0f}")

        # Reordenamos columnas para visualización
        comparacion_final = comparacion[['Mes', 'Meta Facturación $', 'Facturado $', '% Cumplimiento']]

        st.subheader("📊 Comparativo de Facturación vs Meta")
        st.dataframe(comparacion_final)

        # Gráfico de barras de metas vs facturado
        fig_comp = px.bar(
            comparacion,
            x='Mes',
            y=['Meta Facturación', 'Facturado'],
            barmode='group',
            title='Meta vs Facturación Real por Mes'
        )

        fig_comp.update_layout(yaxis_title='Monto ($)', xaxis_title='Mes')
        st.plotly_chart(fig_comp)

        # Gráfico de línea o KPI de cumplimiento si quieres
        fig_cump = px.line(
            comparacion,
            x='Mes',
            y='% Cumplimiento',
            markers=True,
            title='% de Cumplimiento por Mes'
        )

        fig_cump.update_layout(yaxis_title='Porcentaje (%)', xaxis_title='Mes')
        st.plotly_chart(fig_cump)


                #================================= TAB 4: BUSCAR =================================

    # ===============================================
    # ✅ Preprocesamiento de Fecha antes del tab4
    # ===============================================
    df_posibles['Fecha CC'] = pd.to_datetime(df_posibles['Fecha CC'], errors='coerce')
    df_posibles['Año'] = df_posibles['Fecha CC'].dt.year
    df_posibles['Mes'] = df_posibles['Fecha CC'].dt.month
    df_posibles['Día'] = df_posibles['Fecha CC'].dt.day

    # ===============================================
# ✅ TAB 4 - Filtro de Cliente y Fechas
# ===============================================
    with tab4:
        st.subheader("🔍 Buscar Cliente y Filtrar por Cierre de Facturación y Comercial")

        df_filtrado = df_posibles.copy()  # Copia completa del DataFrame

        # CLIENTE
        empresa_buscada = st.text_input("Ingrese el nombre del Cliente:")
        if empresa_buscada:
            df_filtrado = df_filtrado[df_filtrado['Cliente'].str.contains(empresa_buscada, case=False, na=False)]
        else:
            df_filtrado = df_posibles.copy()  # Restaurar el DataFrame completo si no hay búsqueda

        # CIERRE DE FACTURACIÓN
        if "cierre_seleccionado" not in st.session_state:
            st.session_state["cierre_seleccionado"] = "Todos"

        cierres_disponibles = df_filtrado['CIERRE DE FACTURACIÓN'].dropna().unique().tolist()
        opciones_cierre = ["Todos"] + cierres_disponibles

        cierre_seleccionado = st.selectbox(
            "Selecciona el número de Cierre de Facturación:",
            opciones_cierre,
            index=opciones_cierre.index(st.session_state["cierre_seleccionado"])
            if st.session_state["cierre_seleccionado"] in opciones_cierre else 0
        )

        st.session_state["cierre_seleccionado"] = cierre_seleccionado
        if cierre_seleccionado != "Todos":
            df_filtrado = df_filtrado[df_filtrado['CIERRE DE FACTURACIÓN'] == cierre_seleccionado]

        # COMERCIAL
        if "comerciales_seleccionado" not in st.session_state:
            st.session_state["comerciales_seleccionado"] = "Todos"

        comerciales_disponibles = df_filtrado['Comercial'].dropna().unique().tolist()
        opciones_comercial = ["Todos"] + comerciales_disponibles

        comerciales_seleccionado = st.selectbox(
            "Selecciona el Comercial:",
            opciones_comercial,
            index=opciones_comercial.index(st.session_state["comerciales_seleccionado"])
            if st.session_state["comerciales_seleccionado"] in opciones_comercial else 0
        )

        st.session_state["comerciales_seleccionado"] = comerciales_seleccionado
        if comerciales_seleccionado != "Todos":
            df_filtrado = df_filtrado[df_filtrado['Comercial'] == comerciales_seleccionado]

        # REQUERIMIENTOS ESPECIALES
        if "requerimiento_seleccionado" not in st.session_state:
            st.session_state["requerimiento_seleccionado"] = []

        requerimientos_disponibles = df_filtrado['REQUERIMIENTO ESPECIAL'].dropna().unique().tolist()
        requerimientos_previos = [req for req in st.session_state["requerimiento_seleccionado"] if req in requerimientos_disponibles]
        
        requerimiento_seleccionado = st.multiselect(
            "Selecciona lo(s) requerimientos():", 
            requerimientos_disponibles, 
            default=requerimientos_previos
        )

        st.session_state["requerimiento_seleccionado"] = requerimiento_seleccionado
        if requerimiento_seleccionado:
            df_filtrado = df_filtrado[df_filtrado['REQUERIMIENTO ESPECIAL'].isin(requerimiento_seleccionado)]

        # OBSERVACIONES
        if "observaciones_seleccionadas" not in st.session_state:
            st.session_state["observaciones_seleccionadas"] = []

        observaciones_disponibles = df_filtrado['OBSERVACIONES'].dropna().unique().tolist()
        observaciones_previas = [obs for obs in st.session_state["observaciones_seleccionadas"] if obs in observaciones_disponibles]
        
        observaciones_seleccionadas = st.multiselect(
            "Selecciona la(s) Observación(es):", 
            observaciones_disponibles, 
            default=observaciones_previas
        )

        st.session_state["observaciones_seleccionadas"] = observaciones_seleccionadas
        if observaciones_seleccionadas:
            df_filtrado = df_filtrado[df_filtrado['OBSERVACIONES'].isin(observaciones_seleccionadas)]

        # FECHAS
        fecha_rango = st.date_input("Selecciona un rango de fechas (opcional):", [])



        if fecha_rango and len(fecha_rango) == 2:
            fecha_inicio, fecha_fin = fecha_rango
            df_filtrado = df_filtrado[
                (df_filtrado['Fecha CC'] >= pd.to_datetime(fecha_inicio)) &
                (df_filtrado['Fecha CC'] <= pd.to_datetime(fecha_fin))
            ]

        # ===============================================
        # ✅ Resultados y Exportación
        # ===============================================
        if not df_filtrado.empty:
            st.write("### Resultados de la búsqueda")
            st.dataframe(df_filtrado)

            # ================== KPI ==================
            if 'Subtotal' in df_filtrado.columns:
                df_comercial = df_filtrado.groupby('Comercial', as_index=False)['Subtotal'].sum()
                df_cliente = df_filtrado.groupby('Cliente', as_index=False)['Subtotal'].sum()
                total_facturado = df_comercial['Subtotal'].sum()
                total_factcli = df_cliente['Subtotal'].sum()

                st.metric("Total por facturar", f"${total_facturado:,.0f}")

                # ================== Gráfico de Torta ==================
                fig_torta = px.pie(
                    df_comercial,
                    names='Comercial',
                    values='Subtotal',
                    title='Distribución de Facturación por Comercial'
                )
                st.plotly_chart(fig_torta)

                # ================== Gráfico de Barras ==================
                fig_barras = px.bar(
                    df_comercial,
                    x='Comercial',
                    y='Subtotal',
                    text_auto='.2s',
                    title='Facturación por Comercial'
                )
                st.plotly_chart(fig_barras)

                # Ordenar los datos de mayor a menor por 'Subtotal'
                df_cliente = df_cliente.sort_values(by='Subtotal', ascending=False)

                # Calcular el porcentaje acumulado
                df_cliente['% Acumulado'] = df_cliente['Subtotal'].cumsum() / df_cliente['Subtotal'].sum() * 100


                # ================== Gráfico de Barras ==================
                fig_barras = px.bar(
                    df_cliente,
                    x='Cliente',
                    y='Subtotal',
                    text_auto='.2s',
                    title='Facturación por Empresa'
                )
                # Agregar línea de Pareto
                fig_barras.add_trace(go.Scatter(
                    x=df_cliente['Cliente'],
                    y=df_cliente['% Acumulado'],
                    mode='lines+markers',
                    name='% Acumulado',
                    yaxis='y2'
                ))

                # Configurar ejes secundarios
                fig_barras.update_layout(
                    yaxis2=dict(
                        title='% Acumulado',
                        overlaying='y',
                        side='right'
                    )
                )
                st.plotly_chart(fig_barras)

            # ================== Exportar a Excel ==================
            columnas_disponibles = df_filtrado.columns.tolist()
            columnas_seleccionadas = st.multiselect(
                "Selecciona las columnas a exportar:",
                columnas_disponibles,
                default=columnas_disponibles
            )

            if columnas_seleccionadas:
                df_exportar = df_filtrado[columnas_seleccionadas]

                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:

                    # ================== HOJA CONSOLIDADO ==================
                    df_exportar.to_excel(writer, index=False, sheet_name='Consolidado')

                    workbook = writer.book
                    worksheet = writer.sheets['Consolidado']

                    # FORMATO ENCABEZADOS
                    header_format = workbook.add_format({
                        'bold': True,
                        'text_wrap': True,
                        'valign': 'middle',
                        'fg_color': '#D7E4BC',
                        'border': 1
                    })

                    # FORMATO DE PESOS
                    formato_pesos = workbook.add_format({
                        'num_format': '"$"#,##0',
                        'border': 1
                    })

                    # Escribir encabezados y ajustar ancho
                    for col_num, value in enumerate(df_exportar.columns.values):
                        worksheet.write(0, col_num, value, header_format)
                        worksheet.set_column(col_num, col_num, 20)

                    columnas_pesos = ['Vlr Unit', 'Subtotal', 'Total']

                    for col in columnas_pesos:
                        if col in df_exportar.columns:
                            idx = df_exportar.columns.get_loc(col)
                            worksheet.set_column(idx, idx, 20, formato_pesos)

                    # ================== TOTALES ==================
                    row_total = len(df_exportar) + 1
                    for col_num, col_name in enumerate(df_exportar.columns):
                        if pd.api.types.is_numeric_dtype(df_exportar[col_name]):
                            col_letter = chr(65 + col_num)
                            formula = f"=SUM({col_letter}2:{col_letter}{len(df_exportar)+1})"

                            # Aplica formato de pesos si es Subtotal o Total
                            cell_format = formato_pesos if col_name in ['Subtotal', 'Total'] else None
                            worksheet.write_formula(row_total, col_num, formula, cell_format)

                    # ================== HOJA TABLA DINÁMICA ==================
                    pivot_sheet = workbook.add_worksheet('Tabla_Dinamica')

                    if {'Residuo', 'Subtotal'}.issubset(df_filtrado.columns):
                        pivot_data = df_filtrado.groupby('Residuo', as_index=False)['Subtotal'].sum()

                        # Escribir encabezados
                        pivot_sheet.write_row(0, 0, pivot_data.columns)

                        # Escribir datos
                        for idx, row in pivot_data.iterrows():
                            pivot_sheet.write(idx + 1, 0, row['Residuo'])
                            pivot_sheet.write(idx + 1, 1, row['Subtotal'], formato_pesos)
                    else:
                        pivot_sheet.write(0, 0, "No se pudo generar la tabla dinámica.")
                        pivot_sheet.write(1, 0, "Revisa que existan las columnas 'Residuo' y 'Subtotal'.")

                    writer.close()

                # ================== BOTÓN DE DESCARGA ==================
                st.download_button(
                    label="💾 Descargar Consolidado con Totales y Tabla Dinámica",
                    data=output.getvalue(),
                    file_name="consolidado_con_totales_y_tabla.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            else:
                st.warning("Selecciona al menos una columna para exportar.")

        else:
            st.warning("No se encontraron resultados para los filtros seleccionados.")

    with tab5:
        df_asivamos = pd.read_excel(archivo, sheet_name=3)
        
        # Asegurarse de que la columna Crear Fecha es tipo datetime
        df_asivamos['CreaFecha'] = pd.to_datetime(df_asivamos['CreaFecha'], errors='coerce')

        st.subheader("📅 Filtrar por Fecha de Creación")

        # Determinamos las fechas mínima y máxima del dataframe para sugerirlas por defecto
        min_fecha = df_asivamos['CreaFecha'].min().date()
        max_fecha = df_asivamos['CreaFecha'].max().date()

        # Date input para seleccionar el rango de fechas
        fecha_inicio, fecha_fin = st.date_input(
            label="Selecciona el rango de fechas",
            value=(min_fecha, max_fecha),
            min_value=min_fecha,
            max_value=max_fecha
        )

        # Filtramos el dataframe según el rango seleccionado
        filtro_fecha = (df_asivamos['CreaFecha'].dt.date >= fecha_inicio) & (df_asivamos['CreaFecha'].dt.date <= fecha_fin)
        df_filtrado = df_asivamos[filtro_fecha]

        st.subheader("📊 Facturación por Comercial (Filtrado)")
        st.dataframe(df_filtrado.head())

        # Agrupamos el total valor por comercial en el df filtrado
        facturacion_comercial = df_filtrado.groupby('COMERCIAL')['Total'].sum().reset_index()

        # Creamos una columna formateada para mostrar en la tabla
        facturacion_comercial['Total Formateado'] = facturacion_comercial['Total'].apply(lambda x: f"${x:,.0f}")

        # Mostramos la tabla con el total formateado
        st.dataframe(facturacion_comercial[['COMERCIAL', 'Total Formateado']])

        # Gráfico de barras con Plotly, usando la columna 'Total' y mostrando el texto formateado
        fig = px.bar(
            facturacion_comercial,
            x='COMERCIAL',
            y='Total',
            color='Total',
            title='🏆 Total Facturación por Comercial',
            text=facturacion_comercial['Total'].apply(lambda x: f"${x:,.0f}")
        )

        # Mejorar el diseño del gráfico
        fig.update_layout(
            xaxis_title='Comercial',
            yaxis_title='Total Valor Facturado ($)',
            template='plotly_white',
            uniformtext_minsize=8,
            uniformtext_mode='hide'
        )

        st.plotly_chart(fig, use_container_width=True)

        # Mostrar el total general en pesos
        total_general = facturacion_comercial['Total'].sum()
        st.metric(label="💰 Total Facturación General", value=f"${total_general:,.0f}")

    with tab6:
        df_aprovechables = pd.read_excel(archivo, sheet_name=4)
        st.subheader("📋 Aprovechables")
        st.dataframe(df_aprovechables.head())

        # Procesar datos
        df_aprovechables['Fecha CC'] = pd.to_datetime(df_aprovechables['Fecha CC'], errors='coerce')
        df_aprovechables['Subtotalmen'] = pd.to_numeric(df_aprovechables['Subtotalmen'], errors='coerce')
        df_aprovechables_filtrado = df_aprovechables.dropna(subset=['Fecha CC', 'Subtotalmen'])

        # Selección de agrupación por período
        opcion = st.selectbox("Selecciona el período para agrupar:", ("Día", "Mes", "Año"), key='Aprovechables')

        # Crear columna 'Periodo' según la opción elegida
        if opcion == "Día":
            df_aprovechables_filtrado['Periodo'] = df_aprovechables_filtrado['Fecha CC'].dt.date
        elif opcion == "Mes":
            df_aprovechables_filtrado['Periodo'] = df_aprovechables_filtrado['Fecha CC'].dt.to_period('M').astype(str)
        elif opcion == "Año":
            df_aprovechables_filtrado['Periodo'] = df_aprovechables_filtrado['Fecha CC'].dt.year

        # Agrupar por periodo y sumar los subtotales
        resumen_aprovechables = df_aprovechables_filtrado.groupby('Periodo', as_index=False)['Subtotalmen'].sum()
        resumen_aprovechables['Subtotalmen Pesos'] = resumen_aprovechables['Subtotalmen'].apply(lambda x: f"${x:,.0f}")

        # Crear una lista con los periodos disponibles para filtrar
        periodos_disponibles = resumen_aprovechables['Periodo'].unique().tolist()

        # Multiselect para que el usuario elija los periodos a analizar
        periodos_seleccionados = st.multiselect(
            f"Selecciona los {opcion.lower()}s que quieres analizar:",
            options=periodos_disponibles,
            default=periodos_disponibles
        )

        # Filtrar el dataframe según la selección del usuario
        resumen_filtrado = resumen_aprovechables[resumen_aprovechables['Periodo'].isin(periodos_seleccionados)]

        # Calcular el total facturado para los periodos seleccionados
        total_seleccionado = resumen_filtrado['Subtotalmen'].sum()

        # Mostrar el KPI con el total facturado
        st.metric(
            label=f"💰 Total Aprovechables por facturar en los {opcion.lower()}s seleccionados",
            value=f"${total_seleccionado:,.0f}"
        )

        # Dividir en columnas para tabla y gráfico
        col9, col10 = st.columns([1, 2])

        with col9:
            st.subheader(f"Total Aprovechables por facturar por {opcion}")
            st.dataframe(resumen_filtrado[['Periodo', 'Subtotalmen Pesos']])

        with col10:
            fig = px.line(
                resumen_filtrado.sort_values('Periodo'),
                x='Periodo',
                y='Subtotalmen',
                markers=True,
                title=f'Evolución aprovechables por {opcion}'
            )
            fig.update_layout(
                xaxis_title=opcion,
                yaxis_title='Total en aprovechables ($)',
                yaxis_tickformat=',.0f'
            )
            st.plotly_chart(fig)

        # Segunda fila de columnas
        col11, col12 = st.columns(2)

        with col11:
            st.subheader("📈 Evolución de aprovechables en el tiempo (APROVECHABLES)")

            # KPI - Total Facturado APROVECHABLES
            total_aprovechables = df_aprovechables_filtrado['Subtotalmen'].sum()
            st.metric(
                label="💰 Total por enviar a facturar (APROVECHABLES)",
                value=f"${total_aprovechables:,.0f}"
            )

            # Gráfico de línea de evolución
            fig2 = px.line(
                df_aprovechables_filtrado.sort_values('Fecha CC'),
                x='Fecha CC',
                y='Subtotalmen',
                title='Ventas en el Tiempo - APROVECHABLES'
            )
            fig2.update_layout(
                xaxis_title='Fecha',
                yaxis_title='Total Facturado ($)',
                yaxis_tickformat=',.0f'
            )

            st.plotly_chart(fig2)

    with tab7:
        st.subheader("🔍 Buscar Cliente y Filtrar por Cierre de Facturación y Comercial (Aprovechables)")

        # Buscar cliente en aprovechables
        empresa_buscada = st.text_input("Ingrese el nombre del Cliente:", key='busqueda_aprovechables')

        # Cierre de facturación en aprovechables
        cierres_disponibles = df_aprovechables['CIERRE DE FACTURACIÓN'].dropna().unique().tolist()
        cierre_seleccionado = st.selectbox(
            "Selecciona el número de Cierre de Facturación:",
            ["Todos"] + cierres_disponibles,
            key='cierre_aprovechables'
        )

        # Comercial en aprovechables
        comerciales_disponibles = df_aprovechables['Comercial'].dropna().unique().tolist()
        comercial_seleccionado = st.selectbox(
            "Selecciona el Comercial:",
            ["Todos"] + comerciales_disponibles,
            key='comercial_aprovechables'
        )

        # ✅ Opcional: Rango de fechas
        fecha_rango = st.date_input(
            "Selecciona un rango de fechas (opcional):",
            [],
            key='fecha_rango_aprovechables'
        )

        # 🎯 Filtrado según selección
        df_filtrado = df_aprovechables.copy()

        if empresa_buscada:
            df_filtrado = df_filtrado[df_filtrado['Cliente'].str.contains(empresa_buscada, case=False, na=False)]

        if cierre_seleccionado != "Todos":
            df_filtrado = df_filtrado[df_filtrado['CIERRE DE FACTURACIÓN'] == cierre_seleccionado]

        if comercial_seleccionado != "Todos":
            df_filtrado = df_filtrado[df_filtrado['Comercial'] == comercial_seleccionado]

        # ✅ Rango de fechas (si se selecciona)
        if fecha_rango and len(fecha_rango) == 2:
            fecha_inicio, fecha_fin = fecha_rango
            df_filtrado = df_filtrado[
                (df_filtrado['Fecha CC'] >= pd.to_datetime(fecha_inicio)) &
                (df_filtrado['Fecha CC'] <= pd.to_datetime(fecha_fin))
            ]

        # ===============================================
        # ✅ Resultados y Visualización de Datos
        # ===============================================
        if not df_filtrado.empty:
            st.write("### Resultados de la búsqueda - Aprovechables")
            st.dataframe(df_filtrado)

            # ========= KPI Total facturado ==========  
            df_comercial = df_filtrado.groupby('Comercial', as_index=False)['Subtotalmen'].sum()
            total_facturado = df_comercial['Subtotalmen'].sum()

            st.metric("💰 Total por facturar", f"${total_facturado:,.0f}")

            # ========= Gráfico de Torta ==========  
            fig_torta = px.pie(
                df_comercial,
                names='Comercial',
                values='Subtotalmen',
                title='Distribución de Facturación por Comercial'
            )
            st.plotly_chart(fig_torta, use_container_width=True, key="grafico_torta_aprovechables")

            # ========= Gráfico de Barras ==========  
            fig_barras = px.bar(
                df_comercial,
                x='Comercial',
                y='Subtotalmen',
                text_auto='.2s',
                title='Facturación por Comercial'
            )
            st.plotly_chart(fig_barras, use_container_width=True, key="grafico_barras_aprovechables")

            # ===============================================
            # ✅ Exportación a Excel con Totales y Tabla Dinámica
            # ===============================================
            columnas_disponibles = df_filtrado.columns.tolist()
            columnas_seleccionadas = st.multiselect(
                "Selecciona las columnas a exportar:",
                columnas_disponibles,
                default=columnas_disponibles,
                key='columnas_aprovechables'
            )

            if columnas_seleccionadas:
                df_exportar = df_filtrado[columnas_seleccionadas]

                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:

                    # ================== HOJA CONSOLIDADO ==================
                    df_exportar.to_excel(writer, index=False, sheet_name='Consolidado')

                    workbook = writer.book
                    worksheet = writer.sheets['Consolidado']

                    # FORMATO ENCABEZADOS
                    header_format = workbook.add_format({
                        'bold': True,
                        'text_wrap': True,
                        'valign': 'middle',
                        'fg_color': '#D7E4BC',
                        'border': 1
                    })

                    # FORMATO DE PESOS
                    formato_pesos = workbook.add_format({
                        'num_format': '"$"#,##0',
                        'border': 1
                    })

                    # Escribir encabezados y ajustar ancho
                    for col_num, value in enumerate(df_exportar.columns.values):
                        worksheet.write(0, col_num, value, header_format)
                        worksheet.set_column(col_num, col_num, 20)

                    columnas_pesos = ['Vlr Unit', 'Subtotalmen', 'Total']
                    for col in columnas_pesos:
                        if col in df_exportar.columns:
                            idx = df_exportar.columns.get_loc(col)
                            worksheet.set_column(idx, idx, 20, formato_pesos)

                    # ================== TOTALES ==================
                    row_total = len(df_exportar) + 1
                    for col_num, col_name in enumerate(df_exportar.columns):
                        if pd.api.types.is_numeric_dtype(df_exportar[col_name]):
                            col_letter = chr(65 + col_num)
                            formula = f"=SUM({col_letter}2:{col_letter}{len(df_exportar)+1})"

                            # Aplica formato de pesos si es Subtotalmen o Total
                            cell_format = formato_pesos if col_name in ['Subtotalmen', 'Total'] else None
                            worksheet.write_formula(row_total, col_num, formula, cell_format)

                    # ================== HOJA TABLA DINÁMICA ==================
                    pivot_sheet = workbook.add_worksheet('Tabla_Dinamica')

                    if {'Residuo', 'Subtotalmen', 'Peso CP'}.issubset(df_filtrado.columns):
                        pivot_data = df_filtrado.groupby('Residuo', as_index=False).agg({
                            'Subtotalmen': 'sum',
                            'Peso CP': 'sum'
                        })

                        # Escribir encabezados
                        pivot_sheet.write_row(0, 0, pivot_data.columns)

                        # Escribir datos
                        for idx, row in pivot_data.iterrows():
                            pivot_sheet.write(idx + 1, 0, row['Residuo'])
                            pivot_sheet.write(idx + 1, 1, row['Subtotalmen'], formato_pesos)
                            pivot_sheet.write_number(idx + 1, 2, row['Peso CP'])  # ❗ Sin formato de pesos
                    else:
                        pivot_sheet.write(0, 0, "No se pudo generar la tabla dinámica.")
                        pivot_sheet.write(1, 0, "Revisa que existan las columnas 'Residuo', 'Subtotalmen' y 'Peso CP'.")

                # ================== BOTÓN DE DESCARGA ==================
                st.download_button(
                    label="💾 Descargar Consolidado con Totales y Tabla Dinámica (Aprovechables)",
                    data=output.getvalue(),
                    file_name="aprovechables_consolidado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            else:
                st.warning("Selecciona al menos una columna para exportar.")
        else:
            st.warning("No se encontraron resultados para los filtros seleccionados en Aprovechables.")


else:
    st.info("Por favor, sube un archivo Excel para comenzar.")
