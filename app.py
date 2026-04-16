import streamlit as st
import pandas as pd
import re
from datetime import datetime
import io

st.set_page_config(page_title="Validador de Empleados", page_icon="📋", layout="centered")

st.title("📋 Validador de Empleados")
st.markdown("Sube el archivo `.xlsm` para validar y corregir los datos.")

campos_obligatorios = [
    "Id empleado", "Situación", "Nombres", "Apellido paterno", "Apellido materno",
    "Sexo", "Fecha de nacimiento", "Estado civil", "Numero de teléfono 1",
    "Numero de teléfono 2", "Comuna", "Ciudad", "Region", "Nombre Calle",
    "Numero Calle", "Departamento", "Id nación", "Email institucional",
    "Email personal", "Nivel de estudio", "Profesión", "Licencia de conducir",
    "Id banco", "Cuenta del banco", "Id forma de pago", "Id AFP",
    "Estado de jubilación", "¿Es expatriado?", "Sistema de pensiones",
    "ID INSTITUCION DE SALUD", "Monto cotizado en la Isapre en UF",
    "Moneda de la cotización", "Tramo de asignación familiar",
    "¿Supervisa otros empleados?", "¿Es un perfil solo aprobador?",
    "Número del contrato", "Tipo del contrato", "Fecha de inicio del contrato",
    "Fecha de término del contrato", "Sueldo base", "Cargo", "Id centro de costo",
    "Id sede donde se desempeña", "¿Realiza trabajo pesado?",
    "Porcentaje de cotización por trabajo pesado", "Id sindicato",
    "¿Jornada parcial?", "Permite ausencias en días inhábiles",
    "Horas de trabajo semanales", "Distribución de jornada",
    "¿Cotiza seguro de cesantía?", "Fecha de incorporación al seguro de cesantía",
    "Id empresa", "Id plantilla grupal", "Causal de término del contrato",
    "Fecha de reconocimiento de vacaciones",
    "Número de meses reconocidos con otro empleador", "Nivel SENCE", "Factor SENCE",
    "Pauta contable", "Agrupación de seguridad", "Área", "¿Descansa domingos?",
    "¿Cotiza previsión y salud?", "Empleado con perfil privado", "Código interno",
    "Talla de ropa", "Talla de zapatos", "Detalle contrato", "Supervisor",
    "Modalidad del contrato", "Turno", "Zona extrema", "Permisos administrativos",
    "Unidad de permisos administrativos", "Categoría INE", "Notas",
    "Centro de distribucion", "Fecha primera renovación", "Fecha segunda renovación",
    "Fecha de inicio de vacaciones"
]

campos_fecha = [
    "Fecha de nacimiento", "Fecha de inicio del contrato",
    "Fecha de término del contrato", "Fecha de incorporación al seguro de cesantía",
    "Fecha de reconocimiento de vacaciones", "Fecha adicional 1", "Fecha adicional 2",
    "Fecha de afiliación a AFP", "Fecha primera renovación",
    "Fecha segunda renovación", "Fecha de inicio de vacaciones"
]

estados_civiles_validos = ["S", "C", "V", "D", "U"]

formatos_posibles = [
    "%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d",
    "%Y/%m/%d", "%d-%m-%y", "%d/%m/%y", "%m/%d/%Y",
]

def convertir_fecha(valor):
    if pd.isna(valor) or str(valor).strip() == "":
        return valor
    valor_str = str(valor).strip()
    try:
        datetime.strptime(valor_str, "%d-%m-%Y")
        return valor_str
    except ValueError:
        pass
    for fmt in formatos_posibles:
        try:
            return datetime.strptime(valor_str, fmt).strftime("%d-%m-%Y")
        except ValueError:
            continue
    return valor

def validar_email(valor):
    if pd.isna(valor) or str(valor).strip() == "":
        return True
    v = str(valor).strip()
    if "@" not in v:
        return False
    partes = v.split("@")
    if len(partes) != 2 or not partes[0] or not partes[1]:
        return False
    return "." in partes[1]

def validar_corregir_id(valor):
    if pd.isna(valor) or str(valor).strip() == "":
        return valor
    id_str = str(valor).strip()
    if len(id_str) == 9:
        return id_str
    elif len(id_str) == 10 and id_str[0] == "0":
        return id_str[1:]
    elif len(id_str) == 10 and id_str[0] != "0":
        return id_str
    return valor

def limpiar_direccion(valor):
    if pd.isna(valor) or str(valor).strip() == "":
        return valor
    return re.sub(r'[^a-zA-ZáéíóúÁÉÍÓÚñÑ0-9 ]', '', str(valor).strip())

def convertir_email_minuscula(valor):
    if pd.isna(valor) or str(valor).strip() == "":
        return valor
    return str(valor).strip().lower()

def procesar_archivo(uploaded_file):
    df = pd.read_excel(uploaded_file, sheet_name="Empleados", dtype=str)
    total_original = len(df)

    if "Situación" in df.columns:
        df = df[df["Situación"].str.strip().str.upper() == "A"]
    filas_eliminadas = total_original - len(df)

    # Corregir Id empleado
    if "Id empleado" in df.columns:
        df["Id empleado"] = df["Id empleado"].apply(validar_corregir_id)

    # Reemplazar centro de costo y sede
    if "Id centro de costo" in df.columns:
        df["Id centro de costo"] = "sinDefinir"
    if "Id sede donde se desempeña" in df.columns:
        df["Id sede donde se desempeña"] = "sinDefinir"

    # Convertir fechas
    for campo in campos_fecha:
        if campo in df.columns:
            df[campo] = df[campo].apply(convertir_fecha)

    # Limpiar caracteres inválidos en dirección
    campos_direccion = ["Nombre Calle", "Numero Calle", "Departamento"]
    for campo in campos_direccion:
        if campo in df.columns:
            df[campo] = df[campo].apply(limpiar_direccion)

    # Rellenar comuna con ceros a la izquierda hasta 5 caracteres
    if "Comuna" in df.columns:
        df["Comuna"] = df["Comuna"].apply(lambda x: str(x).strip().zfill(5) if pd.notna(x) and str(x).strip() != "" else x)

    # Convertir emails a minúscula
    campos_email = ["Email institucional", "Email personal"]
    for campo in campos_email:
        if campo in df.columns:
            df[campo] = df[campo].apply(convertir_email_minuscula)

    # Validaciones
    errores = []
    for idx, fila in df.iterrows():
        num_fila = idx + 2
        campos_vacios = []
        errores_fila = []
        for campo in campos_obligatorios:
            if campo not in df.columns:
                continue
            valor = fila[campo]
            if pd.isna(valor) or str(valor).strip() == "":
                campos_vacios.append(campo)
                continue
            if campo == "Sexo" and str(valor).strip().upper() not in ["M", "F"]:
                errores_fila.append(f"Sexo (valor: '{valor}' debe ser M o F)")
            if campo == "Estado civil" and str(valor).strip().upper() not in estados_civiles_validos:
                errores_fila.append(f"Estado civil (valor: '{valor}' debe ser S, C, V, D o U)")
            if campo in ["Email institucional", "Email personal"] and not validar_email(valor):
                errores_fila.append(f"{campo} (valor: '{valor}' no tiene formato valido)")
            if campo == "Id empleado":
                id_str = str(valor).strip()
                if len(id_str) not in [9, 10]:
                    errores_fila.append(f"Id empleado (valor: '{valor}' debe tener 9 o 10 caracteres)")
        if campos_vacios or errores_fila:
            errores.append((num_fila, campos_vacios, errores_fila))

    return df, errores, total_original, filas_eliminadas

archivo = st.file_uploader("Selecciona el archivo Excel", type=["xlsm", "xlsx"])

if archivo:
    with st.spinner("Procesando archivo..."):
        try:
            df, errores, total_original, filas_eliminadas = procesar_archivo(archivo)

            col1, col2, col3 = st.columns(3)
            col1.metric("Filas originales", total_original)
            col2.metric("Filas eliminadas (no A)", filas_eliminadas)
            col3.metric("Filas con errores", len(errores))

            if errores:
                st.warning(f"Se encontraron errores en {len(errores)} fila(s)")
                with st.expander("Ver detalle de errores"):
                    reporte = ""
                    for fila, vacios, errs in errores:
                        reporte += f"**Fila {fila}:**\n"
                        for campo in vacios:
                            reporte += f"- '{campo}' está vacío\n"
                        for error in errs:
                            reporte += f"- {error}\n"
                        reporte += "\n"
                    st.markdown(reporte)
            else:
                st.success("Todo correcto! No hay errores en el archivo.")

            reporte_txt = "REPORTE DE VALIDACION\n" + "=" * 60 + "\n\n"
            if errores:
                reporte_txt += f"ERRORES ENCONTRADOS - {len(errores)} fila(s) con problemas\n\n"
                for fila, vacios, errs in errores:
                    reporte_txt += f"Fila {fila}:\n"
                    for campo in vacios:
                        reporte_txt += f"  - '{campo}' esta vacio\n"
                    for error in errs:
                        reporte_txt += f"  - {error}\n"
                    reporte_txt += "\n"
            else:
                reporte_txt += "Todo correcto! No hay errores.\n"

            buffer = io.BytesIO()
            df.to_excel(buffer, index=False)
            buffer.seek(0)

            st.markdown("---")
            st.markdown("### Descargar resultados")
            col_a, col_b = st.columns(2)
            with col_a:
                st.download_button(
                    label="📥 Descargar Excel corregido",
                    data=buffer,
                    file_name="importacion_corregido.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with col_b:
                st.download_button(
                    label="📄 Descargar reporte de errores",
                    data=reporte_txt.encode("utf-8"),
                    file_name="reporte_errores.txt",
                    mime="text/plain"
                )

        except Exception as e:
            st.error(f"Error al procesar el archivo: {e}")
