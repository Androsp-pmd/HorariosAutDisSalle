import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

def parsear_horario_visual(file):
    # Cargar Excel leyendo solo los valores
    wb = openpyxl.load_workbook(file, data_only=True)
    
    nombre_hoja = "HorarioSem"
    if nombre_hoja not in wb.sheetnames:
        st.error(f"No se encontró la hoja '{nombre_hoja}'.")
        return None
    
    ws = wb[nombre_hoja]
    datos_lista = []
    # Columnas: B=2, C=3, D=4, E=5, F=6
    dias = {2: "Lunes", 3: "Martes", 4: "Miércoles", 5: "Jueves", 6: "Viernes"}
    
    bloques = []
    for row in range(1, ws.max_row + 1):
        cell_value = str(ws.cell(row=row, column=1).value).upper()
        if "SEMESTRE" in cell_value and "HORARIOS 202602" not in cell_value:
            bloques.append({"nombre": cell_value, "inicio": row + 2, "fin": row + 14})

    for b in bloques:
        for fila in range(b["inicio"], b["fin"] + 1):
            hora = ws.cell(row=fila, column=1).value
            if not hora: continue
            
            for col_idx, dia_nombre in dias.items():
                contenido = ws.cell(row=fila, column=col_idx).value
                if contenido:
                    lineas = [l.strip() for l in str(contenido).split('\n')]
                    while len(lineas) < 4: lineas.append("PENDIENTE")
                    
                    datos_lista.append({
                        "Semestre": b["nombre"], "Dia": dia_nombre, "Hora": hora,
                        "Codigo": lineas[0], "Asignatura": lineas[1],
                        "Profesor": lineas[2], "Salon": lineas[3]
                    })
                    
    return pd.DataFrame(datos_lista)

# --- INTERFAZ STREAMLIT ---
st.set_page_config(page_title="Gestor de Horarios Universitarios", layout="wide")
st.title("📅 Generador de Horarios por Profesor")

archivo_subido = st.file_uploader("Elige tu archivo Excel (.xlsx)", type=["xlsx"])

if archivo_subido:
    df = parsear_horario_visual(archivo_subido)
    
    if df is not None and not df.empty:
        etiquetas_ignorar = ["PENDIENTE", "POR ASIGNAR", "-", "TBD", ""]
        df_reales = df[~df['Profesor'].isin(etiquetas_ignorar)]
        
        # Validación de Cruces: Solo si el código de asignatura es diferente
        cruces = df_reales[df_reales.groupby(['Dia', 'Hora', 'Profesor'])['Codigo'].transform('nunique') > 1]

        if not cruces.empty:
            st.error("⚠️ Se detectaron conflictos de asignación (Diferentes materias a la misma hora).")
            reporte_df = cruces[['Profesor', 'Dia', 'Hora', 'Asignatura', 'Semestre']].sort_values(by=['Profesor', 'Dia', 'Hora'])
            st.dataframe(reporte_df)
            
            st.download_button(
                label="📥 Descargar Reporte de Errores", 
                data="CONFLICTOS:\n" + reporte_df.to_string(index=False), 
                file_name="Reporte_Errores.txt"
            )
        
        elif df_reales.empty:
            st.warning("No se encontraron profesores reales para procesar (solo hay PENDIENTES o TBD).")
            
        else:
            st.success("✅ ¡Base de datos validada! Generando archivos...")
            
            output = BytesIO()
            horas_orden = df['Hora'].unique()
            
            # Intentar crear el Excel
            try:
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    for profe in df_reales['Profesor'].unique():
                        df_profe = df_reales[df_reales['Profesor'] == profe]
                        
                        # Usamos pivot_table con 'first' por si hay materias espejo (mismo código)
                        cal = df_profe.pivot_table(
                            index='Hora', 
                            columns='Dia', 
                            values='Asignatura', 
                            aggfunc=lambda x: " / ".join(set(x))
                        )
                        
                        # Reindexar para asegurar cuadrícula completa
                        cal = cal.reindex(index=horas_orden, columns=["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]).fillna("-")
                        
                        # Limpiar nombre de pestaña (máx 31 caracteres, sin caracteres prohibidos)
                        nombre_pestana = "".join(c for c in str(profe)[:30] if c not in r'[]:*?/\\')
                        cal.to_excel(writer, sheet_name=nombre_pestana)

                st.download_button(
                    label="📥 Descargar Horarios Individuales (Excel)",
                    data=output.getvalue(),
                    file_name="Horarios_Profesores_Generados.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Error al generar el Excel: {e}")
    else:
        st.info("El archivo parece estar vacío o no contiene la palabra 'SEMESTRE' en la columna A.")
