import streamlit as st
import pandas as pd
from datetime import datetime
import holidays

# ─────────────────────────────────────────────
# CONFIGURACIÓN DE PÁGINA
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Reporte Biométrico",
    page_icon="🕐",
    layout="wide"
)

st.title("🕐 Reporte Consolidado Biométrico")
st.markdown("Carga los archivos de **horario administrativo** y **biométrico** para generar el reporte.")

# ─────────────────────────────────────────────
# DICCIONARIO DE MESES EN ESPAÑOL (con .map)
# ─────────────────────────────────────────────
MESES_ES = {
    1:  "ENERO",
    2:  "FEBRERO",
    3:  "MARZO",
    4:  "ABRIL",
    5:  "MAYO",
    6:  "JUNIO",
    7:  "JULIO",
    8:  "AGOSTO",
    9:  "SEPTIEMBRE",
    10: "OCTUBRE",
    11: "NOVIEMBRE",
    12: "DICIEMBRE",
}

# ─────────────────────────────────────────────
# CARGA DE ARCHIVOS
# ─────────────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    st.subheader("📂 Archivo de Horario Administrativo")
    file_carga = st.file_uploader(
        "Sube CARGA_ADMINISTRATIVO.xlsx",
        type=["xlsx"],
        key="carga"
    )

with col2:
    st.subheader("📂 Archivo Biométrico")
    file_bio = st.file_uploader(
        "Sube BIOMETRICO.xlsx",
        type=["xlsx"],
        key="bio"
    )

# ─────────────────────────────────────────────
# PROCESAMIENTO (solo si ambos archivos están cargados)
# ─────────────────────────────────────────────
if file_carga and file_bio:
    with st.spinner("Procesando datos..."):

        # ── 1. LEER CARGA ADMINISTRATIVA ──────────────────────────────
        columnas_requeridas = [
            "DOCUMENTO", "NOMBRE", "SEDE", "FINI", "FFIN",
            "HORARIO_LUNES_1",    "HORARIO_LUNES_2",
            "HORARIO_MARTES_1",   "HORARIO_MARTES_2",
            "HORARIO_MIERCOLES_1","HORARIO_MIERCOLES_2",
            "HORARIO_JUEVES_1",   "HORARIO_JUEVES_2",
            "HORARIO_VIERNES_1",  "HORARIO_VIERNES_2",
            "HORARIO_SABADO_1"
        ]
        df = pd.read_excel(file_carga, usecols=columnas_requeridas)

        # ── 2. COMBINAR JORNADAS POR DÍA ─────────────────────────────
        dias = ['LUNES', 'MARTES', 'MIERCOLES', 'JUEVES', 'VIERNES']

        for dia in dias:
            col_1    = f'HORARIO_{dia}_1'
            col_2    = f'HORARIO_{dia}_2'
            nueva    = f'HORARIO_{dia}'

            def combinar_horas(row, c1=col_1, c2=col_2):
                h1 = str(row[c1]).strip()
                h2 = str(row[c2]).strip()
                if h2 == '-' or h2 == '' or 'nan' in h2.lower():
                    return h1
                try:
                    inicio = h1.split('-')[0].strip()
                    fin    = h2.split('-')[1].strip()
                    return f"{inicio} - {fin}"
                except:
                    return h1

            df[nueva] = df.apply(combinar_horas, axis=1)

        df['HORARIO_SABADO'] = df['HORARIO_SABADO_1']

        columnas_a_borrar = [f'HORARIO_{d}_{i}' for d in dias for i in [1, 2]] + ['HORARIO_SABADO_1']
        df_final_step = df.drop(columns=columnas_a_borrar)

        # ── 3. MELT (DESDINAMIZACIÓN) ─────────────────────────────────
        id_vars    = ['DOCUMENTO', 'NOMBRE', 'SEDE', 'FINI', 'FFIN']
        value_vars = ['HORARIO_LUNES','HORARIO_MARTES','HORARIO_MIERCOLES',
                      'HORARIO_JUEVES','HORARIO_VIERNES','HORARIO_SABADO']

        df_largo = df_final_step.melt(
            id_vars=id_vars,
            value_vars=value_vars,
            var_name='DIA',
            value_name='HORARIO'
        )
        df_largo['DIA'] = df_largo['DIA'].str.replace('HORARIO_', '')

        # ── 4. CÁLCULO HORAS BRUTAS Y DESCUENTO ALMUERZO ─────────────
        def calcular_horario_ajustado(rango_texto):
            try:
                partes = str(rango_texto).split('-')
                if len(partes) < 2:
                    return 0, 0
                inicio     = datetime.strptime(partes[0].strip(), '%H:%M')
                fin        = datetime.strptime(partes[1].strip(), '%H:%M')
                total_bruto = int((fin - inicio).total_seconds() / 3600)
                almuerzo   = max(0, total_bruto - 8)
                return total_bruto, almuerzo
            except:
                return 0, 0

        resultados = df_largo['HORARIO'].apply(
            lambda x: pd.Series(calcular_horario_ajustado(x))
        )
        df_largo['HORAS_TOTALES_BRUTAS'] = resultados[0]
        df_largo['DESCUENTO_ALMUERZO']   = resultados[1]
        df_largo['JORNADA_FINAL']        = df_largo['HORAS_TOTALES_BRUTAS'] - df_largo['DESCUENTO_ALMUERZO']

        # ── 5. CONSTRUCCIÓN CRONOGRAMA ────────────────────────────────
        df_largo['FINI'] = pd.to_datetime(df_largo['FINI'], errors='coerce', dayfirst=True)
        df_largo['FFIN'] = pd.to_datetime(df_largo['FFIN'], errors='coerce', dayfirst=True)

        fecha_limite_proyeccion = pd.Timestamp('2027-12-31')
        df_limpio = df_largo.dropna(subset=['FINI']).copy()
        df_limpio['FFIN'] = df_limpio['FFIN'].fillna(fecha_limite_proyeccion)
        df_limpio.loc[df_limpio['FFIN'] > fecha_limite_proyeccion, 'FFIN'] = fecha_limite_proyeccion

        mapa_dias = {0:'LUNES',1:'MARTES',2:'MIERCOLES',3:'JUEVES',4:'VIERNES',5:'SABADO'}
        cronograma_lista = []

        for _, row in df_limpio.iterrows():
            try:
                rango_fechas = pd.date_range(start=row['FINI'], end=row['FFIN'], freq='D')
                df_temp = pd.DataFrame({'FECHA': rango_fechas})
                df_temp['DIA_SEMANA'] = df_temp['FECHA'].dt.dayofweek.map(mapa_dias)
                df_temp = df_temp[df_temp['DIA_SEMANA'] == row['DIA']].copy()

                if not df_temp.empty:
                    df_temp['DOCUMENTO']          = row['DOCUMENTO']
                    df_temp['NOMBRE']             = row['NOMBRE']
                    df_temp['SEDE']               = row['SEDE']
                    df_temp['HORARIO_PROYECTADO'] = row['HORARIO']
                    df_temp['HORAS_TOTALES_BRUTAS']= row['HORAS_TOTALES_BRUTAS']
                    df_temp['DESCUENTO_ALMUERZO'] = row['DESCUENTO_ALMUERZO']
                    df_temp['HORAS_TURNO']        = row['JORNADA_FINAL']
                    cronograma_lista.append(df_temp)
            except Exception as e:
                st.warning(f"Error procesando fila de {row['NOMBRE']}: {e}")
                continue

        df_cronograma_final = pd.concat(cronograma_lista, ignore_index=True)

        # ── 6. RECARGOS NOCTURNOS PROYECTADOS ────────────────────────
        def calcular_recargo_nocturno(horario_texto, inicio_nocturno=19):
            try:
                if not isinstance(horario_texto, str) or '-' not in horario_texto:
                    return 0.0
                partes  = horario_texto.split('-')
                fmt     = '%H:%M'
                inicio  = datetime.strptime(partes[0].strip(), fmt)
                fin     = datetime.strptime(partes[1].strip(), fmt)
                umbral  = inicio.replace(hour=inicio_nocturno, minute=0, second=0)
                if fin <= umbral:
                    return 0.0
                if inicio < umbral and fin > umbral:
                    return round((fin - umbral).total_seconds() / 3600, 2)
                if inicio >= umbral:
                    return round((fin - inicio).total_seconds() / 3600, 2)
                return 0.0
            except:
                return 0.0

        df_cronograma_final['RECARGOS_PROYECTADOS'] = df_cronograma_final['HORARIO_PROYECTADO'].apply(
            lambda x: calcular_recargo_nocturno(x, inicio_nocturno=19)
        )

        # ── 7. SEPARAR HORA INICIO / SALIDA + LLAVE ──────────────────
        split_col = df_cronograma_final['HORARIO_PROYECTADO'].str.split('-', expand=True)
        df_cronograma_final['HORA_INICIO'] = split_col[0].str.strip()
        df_cronograma_final['HORA_SALIDA'] = split_col[1].str.strip()

        df_cronograma_final['llave'] = (
            df_cronograma_final['FECHA'].dt.strftime('%d/%m/%Y') + "-" +
            df_cronograma_final['DOCUMENTO'].astype(str)
        )

        columnas_ordenadas = ['llave','FECHA','DIA_SEMANA','DOCUMENTO','NOMBRE','SEDE','HORA_INICIO','HORA_SALIDA']
        columnas_calculos  = ['HORAS_TOTALES_BRUTAS','DESCUENTO_ALMUERZO','HORAS_TURNO','RECARGOS_PROYECTADOS']
        df_final = df_cronograma_final[columnas_ordenadas + columnas_calculos].copy()

        # ── 8. LEER BIOMÉTRICO ────────────────────────────────────────
        biometrico = pd.read_excel(
            file_bio,
            usecols=["fecha","Documento","hora_entrada","hora_salida"],
            skiprows=1
        )
        biometrico['fecha'] = pd.to_datetime(biometrico['fecha'], format='%d/%m/%Y')
        biometrico.insert(
            0, 'llave',
            biometrico['fecha'].dt.strftime('%d/%m/%Y') + '-' + biometrico['Documento'].astype(str)
        )

        # ── 9. CRUCE ──────────────────────────────────────────────────
        df_cruce = df_final[["llave","FECHA","DOCUMENTO","NOMBRE","SEDE",
                             "HORA_INICIO","HORA_SALIDA","DESCUENTO_ALMUERZO"]].copy()

        # DICCIONARIO MESES → reemplaza dt.month_name(locale="es_ES")
        df_cruce.insert(1, "MES", df_cruce["FECHA"].dt.month.map(MESES_ES))

        df_cruce = pd.merge(
            df_cruce,
            biometrico[['llave','hora_entrada','hora_salida']],
            on='llave',
            how='left'
        )

        # ── 10. HORAS LABORADAS ───────────────────────────────────────
        festivos_co = holidays.Colombia(years=range(2025, 2028))

        def calcular_riguroso_con_festivos(row):
            es_semana_santa = (row['FECHA'] >= pd.Timestamp("2026-03-29")) and \
                              (row['FECHA'] <= pd.Timestamp("2026-04-05"))
            es_festivo = row['FECHA'] in festivos_co

            if es_festivo or es_semana_santa:
                try:
                    base  = "2026-01-01 "
                    p_in  = pd.to_datetime(base + str(row['HORA_INICIO']))
                    p_out = pd.to_datetime(base + str(row['HORA_SALIDA']))
                    if p_out > p_in:
                        bruto = (p_out - p_in).total_seconds() / 3600
                        desc  = float(row['DESCUENTO_ALMUERZO'])
                        turno = bruto - desc
                        return round(max(0, turno), 2)
                    return 0.0
                except:
                    return 0.0

            if row['hora_entrada'] == 'SIN MARCA' or row['hora_salida'] == 'SIN MARCA' \
               or pd.isna(row['hora_entrada']):
                return 0.0
            try:
                base      = "2026-01-01 "
                proj_in   = pd.to_datetime(base + str(row['HORA_INICIO']))
                proj_out  = pd.to_datetime(base + str(row['HORA_SALIDA']))
                real_in   = pd.to_datetime(base + str(row['hora_entrada']))
                real_out  = pd.to_datetime(base + str(row['hora_salida']))
                entrada_v = max(real_in,  proj_in)
                salida_v  = min(real_out, proj_out)
                if salida_v > entrada_v:
                    horas_b  = (salida_v - entrada_v).total_seconds() / 3600
                    desc     = float(row['DESCUENTO_ALMUERZO'])
                    resultado = horas_b - desc if horas_b >= 6 and not (desc > 3) else horas_b
                    return round(max(0, resultado), 2)
                return 0.0
            except:
                return 0.0

        df_cruce['HORAS_LABORADAS'] = df_cruce.apply(calcular_riguroso_con_festivos, axis=1)

        # ── 10.5 AJUSTAR A 8 JORNADAS DE MAS DE 8 HORAS ───────────────────────────────────────
        df_cruce['HORAS_LABORADAS'] = df_cruce['HORAS_LABORADAS'].clip(upper=8.0)


        # ── 11. RECARGOS REALES ───────────────────────────────────────
        def calcular_recargos_reales(row):
            if row['hora_entrada'] == 'SIN MARCA' or pd.isna(row['hora_entrada']):
                return 0.0
            try:
                h_salida_prog = pd.to_datetime(row['HORA_SALIDA'], format='%H:%M').time()
                if h_salida_prog <= datetime.strptime("19:00", "%H:%M").time():
                    return 0.0
                base           = "2026-01-01 "
                inicio_recargo = pd.to_datetime(base + "19:00")
                fin_recargo    = pd.to_datetime(base + "22:00")
                real_in        = pd.to_datetime(base + str(row['hora_entrada']))
                real_out       = pd.to_datetime(base + str(row['hora_salida']))
                proj_out       = pd.to_datetime(base + str(row['HORA_SALIDA']))
                inicio_v       = max(real_in,  inicio_recargo)
                fin_v          = min(real_out, proj_out, fin_recargo)
                if fin_v > inicio_v:
                    return round((fin_v - inicio_v).total_seconds() / 3600, 2)
                return 0.0
            except:
                return 0.0

        df_cruce['TOTAL_HORAS_RECARGO'] = df_cruce.apply(calcular_recargos_reales, axis=1)

    # ─────────────────────────────────────────────
    # DESCARGA DEL EXCEL FINAL
    # ─────────────────────────────────────────────
    import io
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_final.to_excel(writer,  sheet_name='Planilla_horario',  index=False)
        df_cruce.to_excel(writer,  sheet_name='Cruce_biometrico',  index=False)
    buffer.seek(0)

    st.download_button(
        label="⬇️ Descargar Reporte Excel",
        data=buffer,
        file_name="Reporte_Consolidado_Biometrico.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("⬆️ Sube ambos archivos para comenzar el procesamiento.")
