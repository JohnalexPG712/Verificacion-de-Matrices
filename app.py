import streamlit as st
import pandas as pd
import os
from datetime import datetime
import io

def validar_matrices_streamlit(rpt_matrices_file, justificacion_file):
    """
    Función de validación adaptada para Streamlit
    """
    try:
        # Cargar archivos
        rpt_matrices = pd.read_excel(rpt_matrices_file, skiprows=4, header=0)
        justificacion = pd.read_excel(justificacion_file)
        
        # Limpiar espacios en blanco
        rpt_matrices.iloc[:, 1] = rpt_matrices.iloc[:, 1].astype(str).str.strip()
        rpt_matrices.iloc[:, 4] = rpt_matrices.iloc[:, 4].astype(str).str.strip()
        justificacion['CODIGO MATRIZ'] = justificacion['CODIGO MATRIZ'].astype(str).str.strip()
        justificacion['CODIGO COMPONENTE'] = justificacion['CODIGO COMPONENTE'].astype(str).str.strip()
        
        # Validación
        resultados = []
        errores_detalle = []
        
        for idx, fila_just in justificacion.iterrows():
            codigo_matriz = fila_just['CODIGO MATRIZ']
            codigo_componente = fila_just['CODIGO COMPONENTE']
            
            coincidencias = rpt_matrices[
                (rpt_matrices.iloc[:, 1] == codigo_matriz) & 
                (rpt_matrices.iloc[:, 4] == codigo_componente)
            ]
            
            if len(coincidencias) > 0:
                fila_rpt = coincidencias.iloc[0]
                consumo_just = fila_just['CONSUMO DEL COMPONENTE']
                consumo_rpt = fila_rpt.iloc[6]
                desperdicio_just = fila_just['CONSUMO DE DESPERDICIO']
                desperdicio_rpt = fila_rpt.iloc[7]
                
                consumo_coincide = consumo_just == consumo_rpt
                desperdicio_coincide = desperdicio_just == desperdicio_rpt
                todo_valido = consumo_coincide and desperdicio_coincide
                
                resultado = '✅ CORRECTO' if todo_valido else '❌ INCONSISTENTE'
                
                if not todo_valido:
                    errores_detalle.append({
                        'CODIGO_MATRIZ': codigo_matriz,
                        'CODIGO_COMPONENTE': codigo_componente,
                        'PROBLEMA_CONSUMO': not consumo_coincide,
                        'CONSUMO_JUST': consumo_just,
                        'CONSUMO_RPT': consumo_rpt,
                        'PROBLEMA_DESPERDICIO': not desperdicio_coincide,
                        'DESPERDICIO_JUST': desperdicio_just,
                        'DESPERDICIO_RPT': desperdicio_rpt
                    })
                
                resultados.append({
                    'CODIGO_MATRIZ': codigo_matriz,
                    'CODIGO_COMPONENTE': codigo_componente,
                    'CONSUMO_JUST': consumo_just,
                    'CONSUMO_RPT': consumo_rpt,
                    'DESPERDICIO_JUST': desperdicio_just,
                    'DESPERDICIO_RPT': desperdicio_rpt,
                    'CONSUMO_COINCIDE': '✅' if consumo_coincide else '❌',
                    'DESPERDICIO_COINCIDE': '✅' if desperdicio_coincide else '❌',
                    'RESULTADO': resultado
                })
            else:
                resultados.append({
                    'CODIGO_MATRIZ': codigo_matriz,
                    'CODIGO_COMPONENTE': codigo_componente,
                    'CONSUMO_JUST': fila_just['CONSUMO DEL COMPONENTE'],
                    'CONSUMO_RPT': 'N/A',
                    'DESPERDICIO_JUST': fila_just['CONSUMO DE DESPERDICIO'],
                    'DESPERDICIO_RPT': 'N/A',
                    'CONSUMO_COINCIDE': 'N/A',
                    'DESPERDICIO_COINCIDE': 'N/A',
                    'RESULTADO': '🔍 NO ENCONTRADO'
                })
        
        return pd.DataFrame(resultados), errores_detalle
        
    except Exception as e:
        st.error(f"❌ Error en la validación: {e}")
        return None, None

def main():
    st.set_page_config(
        page_title="Validador de Matrices",
        page_icon="🚀",
        layout="wide"
    )
    
    st.title("🚀 Verificador de Matrices")
    st.markdown("Valida Archivos **Informe Jasper** vs **Justificación Matrices** Usuario")
    
    # Sidebar para subir archivos
    st.sidebar.header("📁 Cargar Archivos")
    
    rpt_file = st.sidebar.file_uploader(
        "Subir Rpt_Matrices.xlsx" Jasper, 
        type=['xlsx'],
        help="Archivo con estructura: EMPRESA, CODÍGO PT, DESCRIPCIÓN PT., etc."
    )
    
    justificacion_file = st.sidebar.file_uploader(
        "Subir Justificación Matrices Nuevas.xlsx", 
        type=['xlsx'],
        help="Archivo con estructura: CODIGO MATRIZ, CODIGO COMPONENTE, CONSUMO, etc."
    )
    
    # Botón de validación
    if st.sidebar.button("🚀 Ejecutar Validación", type="primary"):
        if rpt_file is not None and justificacion_file is not None:
            with st.spinner("Validando archivos..."):
                df_resultados, errores = validar_matrices_streamlit(rpt_file, justificacion_file)
                
                if df_resultados is not None:
                    # Mostrar estadísticas
                    st.success("✅ Validación completada")
                    
                    col1, col2, col3 = st.columns(3)
                    
                    total_validos = len(df_resultados[df_resultados['RESULTADO'] == '✅ CORRECTO'])
                    total_inconsistentes = len(df_resultados[df_resultados['RESULTADO'] == '❌ INCONSISTENTE'])
                    total_no_encontrados = len(df_resultados[df_resultados['RESULTADO'] == '🔍 NO ENCONTRADO'])
                    
                    with col1:
                        st.metric("✅ CORRECTO", total_validos)
                    with col2:
                        st.metric("❌ Inconsistentes", total_inconsistentes)
                    with col3:
                        st.metric("🔍 No Encontrados", total_no_encontrados)
                    
                    # Mostrar detalles de errores
                    if errores:
                        st.error("### 🔴 Inconsistencias Encontradas")
                        for i, error in enumerate(errores, 1):
                            with st.expander(f"{i}. {error['CODIGO_MATRIZ']} | {error['CODIGO_COMPONENTE']}"):
                                if error['PROBLEMA_CONSUMO']:
                                    st.write(f"**Consumo:** {error['CONSUMO_JUST']} vs {error['CONSUMO_RPT']} "
                                           f"(diferencia: {abs(error['CONSUMO_JUST'] - error['CONSUMO_RPT'])})")
                                if error['PROBLEMA_DESPERDICIO']:
                                    st.write(f"**Desperdicio:** {error['DESPERDICIO_JUST']} vs {error['DESPERDICIO_RPT']} "
                                           f"(diferencia: {abs(error['DESPERDICIO_JUST'] - error['DESPERDICIO_RPT'])})")
                    
                    # Mostrar tabla de resultados
                    st.subheader("📋 Resultados Detallados")
                    st.dataframe(df_resultados, use_container_width=True)
                    
                    # Descargar resultados
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    output_filename = f'Resultados_Validacion_{timestamp}.xlsx'
                    
                    # Convertir a Excel para descarga
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_resultados.to_excel(writer, index=False, sheet_name='Resultados')
                    
                    st.download_button(
                        label="📥 Descargar Resultados en Excel",
                        data=output.getvalue(),
                        file_name=output_filename,
                        mime="application/vnd.ms-excel"
                    )
                    
        else:
            st.warning("⚠️ Por favor sube ambos archivos para continuar")
    
    # Instrucciones en el sidebar
    st.sidebar.markdown("---")
    st.sidebar.subheader("📝 Instrucciones")
    st.sidebar.markdown("""
    1. Sube **Rpt_Matrices.xlsx** Jasper
    2. Sube **Justificaciones Matrices Nuevas.xlsx**  
    3. Haz clic en **Ejecutar Validación**
    4. Revisa resultados y descarga el reporte
    """)

if __name__ == "__main__":
    main()
