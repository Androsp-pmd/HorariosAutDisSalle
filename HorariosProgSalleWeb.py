import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

# --- 1. MOTOR DE PROCESAMIENTO ---

def parsear_horario_visual(file):
    # Cargar Excel leyendo solo los valores, no las fórmulas
    wb = openpyxl.load_workbook(file, data_only=True)
    
    # Validar que exista la pestaña correcta
    nombre_hoja = "HorarioSem"
    if nombre_hoja not in wb.sheetnames:
        st.error(f"No se encontró la hoja '{nombre_hoja}'. Revisa el nombre de la pestaña en tu Excel.")
        return None
    
    ws = wb[nombre_hoja]
    datos_lista = []
    dias = {2: "Lunes", 3: "Martes", 4: "Miércoles", 5: "Jueves", 6: "Viernes"}
    
    # Localizar automáticamente los bloques de cada semestre
    bloques = []
    for row in range(1, ws.max_row + 1):
        cell_value = str(ws.cell(row=row, column=1).value).upper()
        # Evita leer el título principal duplicado
        if "SEMESTRE" in cell_value and "HORARIOS 202602" not in cell_value:
            bloques.append({"nombre": cell_value, "inicio": row + 2, "fin": row + 14})

    # Escanear celdas
    for b in bloques:
        for fila in range(b["inicio"], b["fin"] + 1):
            hora = ws.cell(row=fila, column=1).value
            if not hora: continue
            
            for col_idx, dia_nombre in dias.items():
                contenido = ws.cell(row=fila, column=col_idx).value
                if contenido:
                    # Dividir la celda por saltos de línea (Alt+Enter)
                    lineas = [l.strip() for l in str(contenido).split('\n')]
                    
                    # Rellenar datos faltantes para mantener la estructura
                    while len(lineas) < 4: lineas.append("PENDIENTE")
                    
                    datos_lista.append({
                        "Semestre": b["nombre"], "Dia": dia_nombre, "Hora": hora,
                        "Codigo": lineas[0], "Asignatura": lineas[1],
                        "Profesor": lineas[2], "Salon": lineas[3]
                    })
                    
    return pd.DataFrame(datos_lista)

# --- 2. INTERFAZ WEB (STREAMLIT) ---

# Configuración de la página
st.set_page_config(page_title="Gestor de Horarios Universitarios", layout="wide")
st.title("📅 Generador de Horarios por Profesor")
st.markdown("Sube tu plantilla general de Excel para validar cruces de asignación y generar los calendarios individuales.")

# Zona de carga
archivo_subido = st.file_uploader("Elige tu archivo Excel (.xlsx)", type=["xlsx"])

if archivo_subido:
    # Procesar el archivo
    df = parsear_horario_visual(archivo_subido)
    
    if df is not None:
        # Definir qué textos NO son profesores reales
        etiquetas_ignorar = ["PENDIENTE", "POR ASIGNAR", "-", "TBD", ""]
        df_reales = df[~df['Profesor'].isin(etiquetas_ignorar)]
        
        # Buscar cruces (mismo día, hora y profesor)
        cruces = df_reales[df_reales.groupby(['Dia', 'Hora', 'Profesor'])['Codigo'].transform('nunique') > 1]

        # --- LÓGICA DE DECISIÓN ---
        if not cruces.empty:
            st.error("⚠️ Se detectaron conflictos de asignación. Corrige el archivo original y vuelve a subirlo.")
            
            # Formatear el reporte de errores
            reporte_df = cruces[['Profesor', 'Dia', 'Hora', 'Asignatura', 'Semestre']].sort_values(by=['Profesor', 'Dia', 'Hora'])
            reporte_txt = "ALERTA DE CONFLICTOS DE HORARIO\n" + "="*35 + "\n\n" + reporte_df.to_string(index=False)
            
            st.download_button(
                label="📥 Descargar Reporte de Errores (TXT)", 
                data=reporte_txt, 
                file_name="Reporte_Errores_Cruces.txt"
            )
            
            # Mostrar la tabla en pantalla para revisión rápida
            st.dataframe(reporte_df)
            
        else:
            st.success("✅ ¡Base de datos validada sin cruces! Ya puedes descargar los horarios.")
            
            # Generar Excel de salida en memoria
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Capturar el orden exacto de las horas para mantener la estética uniforme
                horas_orden = df['Hora'].unique()
                
                for profe in df_reales['Profesor'].unique():
                    df_profe = df[df['Profesor'] == profe]
                    
                    # Pivotar datos al formato calendario
                    cal = df_profe.pivot(index='Hora', columns='Dia', values='Asignatura')
                    
                    # Aplicar el reindexado para asegurar que todos tengan la misma grilla (L-V, todas las horas)
                    cal = cal.reindex(index=horas_orden, columns=["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]).fillna("-")
                    
                    # Guardar en pestaña limpiando caracteres no permitidos por Excel
                    nombre_pestana = str(profe)[:30].replace('/', '-')
                    cal.to_excel(writer, sheet_name=nombre_pestana)
            
            st.download_button(
                label="📥 Descargar Horarios Individuales (Excel)",
                data=output.getvalue(),
                file_name="Horarios_Profesores_Generados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
