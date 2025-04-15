import streamlit as st
import pandas as pd
from pandas.api.types import is_numeric_dtype
from datetime import datetime
import os

# Configuración inicial
st.set_page_config(
    page_title="Sistema Médico",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded" 
)

# Menú lateral
with st.sidebar:
    st.image("ENCABEZADO.jpeg", width=150)
    st.markdown("## Menú Principal")
    
    # Opciones del menú con íconos
    menu_option = st.radio(
        "  ",
        options=["Inicio", "valoracion", "densitometria", "audiometria", "certificado"],
        format_func=lambda x: {
            "Inicio": "INICIO",
            "valoracion": "VALORACIÓN",
            "densitometria": "DENCITOMETRÍA",
            "audiometria": "AUDIOMETRÍA",
            "certificado": "CERTIFICADO"
        }[x]
    )
    
    # Actualizar el estado según la selección del menú
    if menu_option != "Inicio":
        st.session_state.current_form = menu_option.lower().replace("/", "").replace(" ", "_")
        st.session_state.datos_guardados = None
    else:
        st.session_state.current_form = None



# Inicialización de estados
if "current_form" not in st.session_state:
    st.session_state.current_form = None
    st.session_state.datos_guardados = None
# Función para guardar datos ---
def guardar_en_excel(datos, nombre_archivo):
    """Guarda los datos en un archivo Excel, agregándolos al final si el archivo ya existe"""
    if not os.path.exists('registros'):
        os.makedirs('registros')
    
    filename = f"registros/{nombre_archivo}.xlsx"
    
    try:
        # Si el archivo existe, cargamos los datos existentes
        if os.path.exists(filename):
            df_existente = pd.read_excel(filename, engine='openpyxl')
            df_nuevo = pd.DataFrame([datos])
            df_final = pd.concat([df_existente, df_nuevo], ignore_index=True)
        else:
            # Si no existe, creamos uno nuevo
            df_final = pd.DataFrame([datos])
        
        # Guardamos el DataFrame
        df_final.to_excel(filename, index=False, engine='openpyxl')
        return filename
    
    except Exception as e:
        st.error(f"Error al guardar los datos: {e}")
        return None


# --- Formulario de VALORACIÓN ---
if st.session_state.current_form == "valoracion":
    st.image("ESCUDO.png")
    with st.form("form_valoracion"):
        
        st.header("HOJA DE VALORACIÓN-APOYO A LA VIGILANCIA DE LA SALUD")
        # Sección A
        st.header("A. DATOS DEL ESTABLECIMIENTO - EMPRESA Y USUARIO", divider="blue")
        col1, col2,col3 = st.columns(3)
        with col1:
            institucion = st.text_input("INSTITUCIÓN DEL SISTEMA O NOMBRE DE LA EMPRESA")
        with col2:
            ruc = st.text_input("RUC*",help ="Número de RUC obligatorio")
        with col3:
            ciu = st.text_input("CIU") 
        col1, col2,col3,col4 = st.columns(4)
        with col1:
            establecimiento = st.text_input("ESTABLECIMIENTO DE SALUD")
        with col2:
            historia = st.text_input("NÚMERO DE HISTORIA CLÍNICA")
        with col3:
            archivo = st.text_input("NÚMERO DE ARCHIVO")
        with col4:
            numero_movil = st.text_input("NÚMERO MOVIL")
        col1, col2,col3,col4,col5,col6,col7 = st.columns(7)
        with col1:
            primer_apellido = st.text_input("PRIMER APELLIDO")
        with col2:
            segundo_apellido = st.text_input("SEGUNDO APELLIDO")
        with col3: 
            primer_nombre = st.text_input("PRIMER NOMBRE")
        with col4:
            segundo_nombre = st.text_input("SEGUNDO NOMBRE")
        with col5:
            sexo = st.selectbox("SEXO", ["Masculino", "Femenino"])
        with col6:
            edad = st.number_input("EDAD", min_value=0,value=0,key="edad")
        with col7:
            puesto_trabajo = st.text_input("PUESTO DE TRABAJO")
        # SECCION B
        st.header("B. MOTIVO DE CONSULTA", divider="blue")
        descripcion_consulta = st.text_area("DESCRIBA EL MOTIVO DE CONSULTA")
        # Sección C:
        st.header("C. ANTECEDENTES PERSONALES", divider="blue")
        antecedentes_clinicos_quirurgicos = st.text_area("ANTECEDENTES CLÍNICOS Y QUIRÚRGICOS")
        st.subheader("HÁBITOS TÓXICOS-CONSUMOS NOCIVOS")
        # Consumo de Tabaco
        col1, col2, col3, col4,col5 = st.columns(5)
        with col1:
            consume_tabaco = st.radio("CONSUMO DE TABACO", ["No", "Sí"], key="tabaco_radio", horizontal=True)
        with col2:
            tiempo_consumo_tabaco = st.number_input("TIEMPO DE CONSUMO TABACO", min_value=0,value=0,key="tiempo_tabaco")
        with col3:
            cantidad_tabaco = st.number_input("CANTIDAD DE CIGARILLOS",min_value=0, value=0, key="cantidad_tabaco")
        with col4:
            ex_consumidor_tabaco = st.radio("EX CONSUMIDOR TABACO",["No", "Sí"],key="ex_tabaco",horizontal=True)
        with col5:
            tiempo_abstinencia_tabaco = st.number_input("TIEMPO ABSTINENCIA TABACO",min_value=0, value=0,key="abstinencia_tabaco")
        # Consumo de Alcohol
        col1, col2, col3, col4,col5 = st.columns(5)
        with col1:
            consume_alcohol = st.radio("CONSUMO DE ALCOHOL", ["No", "Sí"], key="alcohol_radio", horizontal=True)
        with col2:
            tiempo_consumo_alcohol = st.number_input("TIEMPO CONSUMO ALCOHOL",min_value=0,value=0,key="tiempo_alcohol")
        with col3:
            cantidad_alcohol = st.number_input("CANTIDAD ALCOHOL", min_value=0,value=0,key="cantidad_alcohol")
        with col4:
            ex_consumidor_alcohol = st.radio("EX CONSUMIDOR ALCOHOL",["No", "Sí"], key="ex_alcohol",horizontal=True)
        with col5:
            tiempo_abstinencia_alcohol = st.number_input("TIEMPO ABSTINENCIA ALCOHOL",min_value=0, value=0,key="abstinencia_alcohol")     
        # Consumo de otras drogas
        col1, col2, col3, col4,col5 ,col6= st.columns(6)
        with col1:
            consume_otras_drogas = st.radio("OTRAS DROGAS", ["No", "Sí"], key="otras_drogas_radio", horizontal=True)
        with col2:
            otras_drogas = st.text_input("¿CUÁL?",key="cual_drogas")
        with col3:
            tiempo_consumo_otras_drogas = st.number_input("TIEMPO CONSUMO",min_value=0,value=0,key="tiempo_otras_drogas")
        with col4:    
            cantidad_otras_drogas = st.number_input("CANTIDAD", min_value=0,value=0,key="cantidad_otras_drogas")
        with col5:
            ex_consumidor_otras_drogas = st.radio("EX CONSUMIDOR",["No", "Sí"], key="ex_otras_drogas",horizontal=True)
        with col6:
            tiempo_abstinencia_otras_drogas = st.number_input("TIEMPO ABSTINENCIA",min_value=0, value=0,key="abstinencia_otras_drogas")
        st.subheader("ESTILO DE VIDA")
        # Actividad física
        col1, col2, col3,col4 = st.columns(4)
        with col1:
            actividad_física = st.radio("ACTIVIDAD FÍSICA", ["No", "Sí"], key="actividad_radio", horizontal=True)
        with col2:
            cual_actividad_fisica=st.text_input("¿CUÁL?", key="cual_actividad")
        with col3:
            tiempo_actividad_fisica = st.number_input("Tiempo (día)",min_value=0, value=0,key="tiempo_actividad_fisica")
        with col4:
            cantidad_actividad_fisica = st.number_input("Cantidad (unidad)",min_value=0, value=0,key="cantidad_actividad_fisica")
        # MEDICACIÓN HABITUAL
        col1, col2, col3 , col4 = st.columns(4)
        with col1:
            medicacion_habitual = st.radio("MEDICACIÓN HABITUAL", ["No", "Sí"], key="medicacion_habitual_radio", horizontal=True)
        with col2:
            cual_medicacion_habitual = st.text_input("¿CUÁL?",key="cual_medicamento")
        with col3:
            tiempo_medicacion_habitual = st.number_input("Tiempo (día)",min_value=0, value=0,key="tiempo_medicacion_habitual")
        with col4:
            cantidad_medicacion_habitual = st.number_input("Cantidad (unidad)",min_value=0, value=0,key="cantidad_medicacion_habitual")
        # INCIDENTES
        st.subheader("INCIDENTES")
        principales_incidentes= st.text_area("Describir los principales incidentes suscitados",key="principales_incidentes")
        # ACCIDENTES
        st.subheader("ACCIDENTES DE TRABAJO")
        col1, col2, col3 = st.columns(3)
        with col1:
            calif_iess_accidente = st.radio("FUE CALIFICADO POR EL INSTITUTO DE SEGURIDAD SOCIAL CORRESPONDIENTE: ", ["No", "Sí"], key="calif_accidente_radio", horizontal=True)
        with col2:
            especificar_accidentes=st.text_input("ESPECIFICAR", key="especificar_accidentes")
        with col3:
            fecha_accidente = st.date_input("Fecha de valoración*", value=datetime.now(),key="fecha_accidente")
        observaciones_accidentes= st.text_area("Observaciones",key="observaciones_accidentes")
        st.subheader("ENFERMEDADES PROFESIONALES ")
        col1, col2, col3 = st.columns(3)
        with col1:
            calif_iess_enfermedades = st.radio("FUE CALIFICADO POR EL INSTITUTO DE SEGURIDAD SOCIAL CORRESPONDIENTE: ", ["No", "Sí"], key="calif_enfermedades_radio", horizontal=True)
        with col2:
            especificar_enfermedades=st.text_input("ESPECIFICAR",key="especificar_enfermedades")
        with col3:
            fecha_enfermedades = st.date_input("Fecha de valoración*", value=datetime.now(),key="fecha_enfermedades")
        observaciones_enfermedades= st.text_area("Observaciones",key="observaciones_enfermedades")
        #Seccion D:
        st.subheader("D. ANTECEDENTES FAMILIARES (DETALLAR EL PARENTESCO)", divider="blue")
        opciones_antecedentes = [
            "1. ENFERMEDAD CARDIO-VASCULAR",
            "2. ENFERMEDAD METABÓLICA",
            "3. ENFERMEDAD NEUROLÓGICA",
            "4. ENFERMEDAD ONCOLÓGICA",
            "5. ENFERMEDAD INFECCIOSA",
            "6.  ENFERMEDAD HEREDITARIA / CONGÉNITA",
            "7. DISCAPACIDADES",
            "8. OTROS"
        ]
        antecedentes_familiares = st.multiselect(
            "DESCRIBIR ABAJO ANOTANDO EL NÚMERO",
            opciones_antecedentes
        )
        descripcion_antecedentes_familiares = st.text_area("Descripción")
        #SECCION E
        st.subheader("E.  FACTORES DE RIESGOS DEL PUESTO DE TRABAJO", divider="blue")
        puesto_trabajo_area=st.text_input("PUESTO DE TRABAJO / ÁREA")
        col1,col2= st.columns(2)
        with col1:
            # Riesgo fisico
            opciones_riesgo_fisico = [
                "Temperaturas altas",
                "Temperaturas bajas",
                "Radiación Ionizante",
                "Radiación No Ionizante",
                "Ruido",
                "Vibración",
                "Iluminación",
                "Ventilación",
                "Fluido eléctrico",
                "otros"
            ]
            riesgo_fisico = st.multiselect(
                "Seleccione los factores de riesgo físico:",
                opciones_riesgo_fisico
            )
            # Riesgo Mecanico
            opciones_riesgo_mecanico = [
                "Atrapamiento entre máquinas",
                "Atrapamiento entre superficies",
                "Atrapamiento entre objetos",
                "Caída de objetos",
                "Caídas al mismo nivel",
                "Caídas a diferente nivel",
                "Contacto eléctrico",
                "Contacto con superficies de trabajos",
                "Proyección de partículas - fragmentos",
                "Proyección de fluidos",
                "Pinchazos",
                "Cortes",
                "Atropellamientos por vehículos",
                "Choques /colisión vehicular",
                "Otros"
            ]
            riesgo_mecanico = st.multiselect(
                "Seleccione los factores de riesgo mecánico:",
                opciones_riesgo_mecanico
            )     
            # Riesgo quimico
            opciones_riesgo_quimico = [
                "Sólidos",
                "Polvos ",
                "Humos",
                "líquidos ",
                "vapores",
                "Aerosoles",
                "Neblinas ",
                "Gaseosos",
                "Otros"
            ]
            riesgo_quimico = st.multiselect(
                "Seleccione los factores de riesgo químico:",
                opciones_riesgo_quimico
            )
        with col2:
            # Riesgo biologico
            opciones_riesgo_biologico = [
                "Virus ",
                "Hongos",
                "Bacterias ",
                "Parásitos ",
                "Exposición a vectores",
                "Exposición a animales selváticos ",
                "Otros"
            ]
            riesgo_biologico = st.multiselect(
                "Seleccione los factores de riesgo biológico:",
                opciones_riesgo_biologico
            )
            # Riesgo ergonomico
            opciones_riesgo_ergonomico = [
                "Manejo manual de cargas",
                "Movimiento repetitivos",
                "Posturas forzadas",
                "Trabajos con PVD",
                "Otros"
            ]
            riesgo_ergonomico = st.multiselect(
                "Seleccione los factores de riesgo ergonómico:",
                opciones_riesgo_ergonomico
            )
            # Riesgo psicosocial
            opciones_riesgo_psicosocial = [
                "Monotonía del trabajo ",
                "Sobrecarga laboral",
                "Minuciosidad de la tarea ",
                "Alta responsabilidad",
                "Autonomía  en la toma de decisiones",
                "Supervisión y estilos de dirección deficiente",
                "Conflicto de rol",
                "Falta de Claridad en las funciones",
                "Incorrecta distribución del trabajo ",
                "Turnos rotativos",
                "Relaciones interpersonales ",
                "inestabilidad laboral",
                "Otros"
            ]
            riesgo_psicosocial = st.multiselect(
                "Seleccione los factores de riesgo psicosocial:",
                opciones_riesgo_psicosocial
            )
        medidas_preventivas=st.text_area("Medidas preventivas")
        #SECCION F:
        st.subheader("F. ENFERMEDAD ACTUAL", divider="blue")
        enfrmedad_actual = st.text_area("Descripción", key="enfrmedad_actual")
        #SECCION G
        st. subheader("G. REVISIÓN DE ÓRGANOS Y SISTEMAS", divider="blue")
        opciones_revision = [
            "1. PIEL - ANEXOS",
            "2. ÓRGANOS DE LOS SENTIDOS",
            "3. RESPIRATORIO",
            "4. CARDIO-VASCULAR",
            "5. DIGESTIVO",
            "6. GENITO - URINARIO",
            "7. MÚSCULO ESQUELÉTICO",
            "8. ENDOCRINO",
            "9. HEMO LINFÁTICO",
            "10. NERVIOSO"
        ]
        revision_organos_sistemas = st.multiselect(
            "EN CASO DE EXISTIR PATOLOGÍA, SELECCIONE Y DESCRIBRA ABAJO ANOTANDO EL NUMERAL",
            opciones_revision
        )
        patologia=st.text_area("Descripción",key="patologia_descripcion")
        #SECCION H
        st.subheader("H. CONSTANTES VITALES Y ANTROPOMETRÍA", divider="blue")
        col1,col2,col3 =st.columns(3)
        with col1:
            presion_arterial=st.text_input("PRESIÓN ARTERIAL (mmHg)")
        with col2:
            temperatura=st.text_input("TEMPERATURA (°C)")
        with col3:
            frecuncia_cardiaca=st.text_input("FRECUENCIA CARDIACA (Lat/min)")
        col1,col2,col3 =st.columns(3)
        with col1:    
            saturacion_oxigeno=st.text_input("SATURACIÓN DE OXÍGENO (O2%)")
        with col2: 
            frecuencia_respiratoria=st.text_input("FRECUENCIA RESPIRATORIA (fr/min)")
        with col3:
            perimetro_abdominal=st.number_input("PERÍMETRO ABDOMINAL (cm)")
        col1,col2,col3,col4 =st.columns([3,3,3,1])
        with col1:
            talla=st.number_input("TALLA (cm)")
        with col2:
            peso=st.number_input("PESO (Kg)")
        with col4:
            if st.form_submit_button("IMC", use_container_width=True):
                if peso > 0 and talla > 0:
                    st.session_state.imc = peso / ((talla/100)**2)
                else:
                    st.session_state.imc = 0
        with col3:   
            # Mostrar resultado
            imc = st.session_state.get('imc', 0)
            st.metric(
                label="ÍNDICE DE MASA CORPORAL (Kg/m²)",
                value=f"{imc:.1f}" if imc > 0 else "_"
            )   
        
        #SECCION I
        st.subheader("I. EXAMEN FÍSICO REGIONAL", divider="blue")
        col1,col2,col3,col4,col5,col6= st.columns(6)
        with col1:
                # PIEL
            opciones_piel = [
                "a. Cicatrices",
                "b. Tatuajes",
                "c. Piel  y faneras"]
            examen_fisico_piel = st.multiselect(
                "1 - PIEL",
                opciones_piel)
        with col2:
                # OJOS
            opciones_ojos = [
                "a. Párpados",
                "b. Conjuntivas",
                "c. Pupilas",
                "d. Córnea",
                "e. Motilidad"]
            examen_fisico_ojos = st.multiselect(
                "2 - OJOS",
                opciones_ojos)
        with col3:
                # OIDO
            opciones_oido = [
                "a. C. auditivo externo",
                "b. Pabellón",
                "c. Tímpanos"]
            examen_fisico_oido = st.multiselect(
                "3 - OÍDO",
                opciones_oido)
        with col4:
                # ORO FARINGE
            opciones_oro_faringe = [
                "a. Labios",
                "b. Lengua",
                "c. Faringec",
                "d. Amígdalas",
                "e. Dentadura"]
            examen_fisico_oro_faringe = st.multiselect(
                "4 - ORO FARINGE",
                opciones_oro_faringe)
        with col5:
                # NARIZ
            opciones_nariz = [
                "a. Tabique",
                "b. Cornetes",
                "c. Mucosas",
                "d. Senos paranasales"]
            examen_fisico_nariz = st.multiselect(
                "5 - NARÍZ",
                opciones_nariz)
        with col6:
                # CUELLO
            opciones_cuello = [
                "a. Tiroides / masas",
                "b. Movilidad"]
            examen_fisico_cuello = st.multiselect(
                "6 - CUELLO",
                opciones_cuello)
            
        col1,col2,col3,col4,col5,col6= st.columns(6)
        with col1:
                # TORAX
            opciones_torax = [
                "a. Mamas",
                "b. Corazón",
                "c. Pulmones",
                "d. Parrilla costal"]
            examen_fisico_torax = st.multiselect(
                "7 - TÓRAX",
                opciones_torax)
        with col2:
                # ABDOMEN
            opciones_abdomen = [
                "a. Vísceras",
                "b. Pared abdominal"]
            examen_fisico_abdomen = st.multiselect(
                "8 - ABDOMEN",
                opciones_abdomen)
        with col3:
                # COLUMNA
            opciones_columna = [
                "a. Flexibilidad",
                "b. Desviación",
                "c. Dolor"]
            examen_fisico_columna = st.multiselect(
                "9 - COLUMNA",
                opciones_columna)
        with col4:
                # PELVIS
            opciones_pelvis = [
                "a. Pelvis",
                "b. Genitales"]
            examen_fisico_pelvis = st.multiselect(
                "10 - PELVIS",
                opciones_pelvis)
        with col5:
                # EXTREMIDADES
            opciones_extremidades = [
                "a. Vascular",
                "b. Miembros superiores",
                "c. Miembros inferiores"]
            examen_fisico_extremidades = st.multiselect(
                "11 - EXTREMIDADES",
                opciones_extremidades)
        with col6:
                # NEUROLOGICO
            opciones_neurologico = [
                "a. Fuerza ",
                "b. Sensibilidad",
                "c. Marcha",
                "d. Reflejos"]
            examen_fisico_neurologico = st.multiselect(
                "12 - NEUROLÓGICO",
                opciones_neurologico)
        descripcion_examen = st.text_area("SI EXISTE EVIDENCIA DE PATOLOGÍA SELECIONE Y DESCRIBIR EN LA SIGUIENTE SECCIÓN COLOCANDO EL NUMERAL", key="des_examen")
        #SECCION J
        st.subheader("J. RESULTADOS DE EXÁMENES GENERALES Y ESPECÍFICOS DE ACUERDO AL RIESGO Y PUESTO DE TRABAJO (IMAGEN, LABORATORIO Y OTROS)", divider="blue")
        col1,col2,col3=st.columns([1,2,2])
        with col1:
            st.metric(label="EXAMEN",value="EKG")
        with col2:
            fecha_examen_ekg=st.date_input("Fecha", value=datetime.now(),key="fecha_examen_ekg")
        with col3:
            resultado_examen_ekg=st.text_input("Resultado", key="resiltado_ekg")
        col1,col2,col3=st.columns([1,2,2])
        with col1:
            st.metric(label="EXAMEN",value="DCO")
        with col2:
            fecha_examen_dco=st.date_input("Fecha", value=datetime.now(),key="fecha_examen_dco")
        with col3:
            resultado_examen_dco=st.text_input("Resultado",key="resultado_dco")
        #SECCION K
        st.subheader("K. DIAGNÓSTICO", divider="blue")
        st.subheader("EVALUACIÓN MÉDICA OCUPACIONAL")
        col1,col2,col3,col4=st.columns(4)
        with col1:
            descripcion_diagnostico=st.text_input("Descripción", key="des_diagnostico")
        with col2:
            cie_diagnostico=st.text_input("CIE",key="cie_diagnostico")
        with col3:
            pre_diagnostico=st.text_input("PRE",key="pre_diagnostico")
        with col4:
            def_diagnostico=st.text_input("DEF",key="def_diagnostico")
        #SECCION L
        st.subheader("L. APTITUD MÉDICA PARA EL TRABAJO", divider="blue")
        
        opciones_aptitud = [
                "APTO",
                "APTO EN OBSERVACIÓN",
                "APTO CON LIMITACIONES ",
                "NO APTO "]
        aptitud_medica = st.multiselect(
                "ESTADO",
            opciones_aptitud)
        observaciones_aptitud=st.text_input("Observación",key="observaciones_aptitud")
        limitaciones_aptitud=st.text_input("Limitación",key="limitaciones_aptitud")
        #SECCION M
        st.subheader("M. RECOMENDACIONES Y/O TRATAMIENTO", divider="blue")
        recomendaciones_tratamiento=st.text_area("Descripción",key="recomendaciones_tratamiento")
        #SECCION N
        
        st.markdown("---")
        st.info(""" 
        CERTIFICO QUE LO ANTERIORMENTE EXPRESADO EN RELACIÓN A MI ESTADO DE SALUD ES VERDAD. 
        SE ME HA INFORMADO LAS MEDIDAS PREVENTIVAS A TOMAR PARA DISMINUIR O MITIGAR LOS 
        RIESGOS RELACIONADOS CON MI ACTIVIDAD LABORAL.""")
        
        st.subheader("N. DATOS DEL PROFESIONAL", divider="blue")
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1:
            fecha_dato_profecional=st.date_input("Fecha", value=datetime.now(),key="fecha_dato_profecional")
        with col2:
            hora_dato_profecional=st.text_input("Hora",key="hora_dato")
        with col3:
            codigo_dato_profecional=st.text_input("Código",key="codigo_dato")
        with col4:
            firma_dato_profecional=st.text_input("Firma y sello",key="firma_dato")
        # Botón para guardar
        guardar = st.form_submit_button("💾 Guardar Datos")
        if guardar:
            if ruc:  # Validar solo el RUC como obligatorio
                datos = {
                    # Sección A
                    "institucion": institucion,
                    "ciu": ciu,
                    "ruc": ruc,
                    "establecimiento": establecimiento,
                    "historia": historia,
                    "archivo": archivo,
                    "primer_apellido": primer_apellido,
                    "segundo_apellido": segundo_apellido,
                    "primer_nombre" : primer_nombre,
                    "segundo_nombre": segundo_nombre,
                    "sexo":sexo,
                    "puesto_trabajo":puesto_trabajo,
                    "numero_movil":numero_movil,
                    "edad":edad,
                    # Sección B
                    "descripcion_consulta":descripcion_consulta,
                    # Sección C
                    "antecedentes_clinicos_quirurgicos": antecedentes_clinicos_quirurgicos,
                    "consume_tabaco": consume_tabaco,
                    "tiempo_consumo_tabaco": tiempo_consumo_tabaco,
                    "cantidad_tabaco": cantidad_tabaco,
                    "ex_consumidor_tabaco": ex_consumidor_tabaco,
                    "tiempo_abstinencia_tabaco": tiempo_abstinencia_tabaco,
                    "consume_alcohol":consume_alcohol,
                    "tiempo_consumo_alcohol": tiempo_consumo_alcohol,
                    "cantidad_alcohol": cantidad_alcohol,
                    "ex_consumidor_alcohol": ex_consumidor_alcohol,
                    "tiempo_abstinencia_alcohol": tiempo_abstinencia_alcohol,
                    "consume_otras_drogas": consume_otras_drogas,
                    "otras_drogas": otras_drogas,
                    "tiempo_consumo_otras_drogas": tiempo_consumo_otras_drogas,
                    "cantidad_otras_drogas": cantidad_otras_drogas,
                    "ex_consumidor_otras_drogas":ex_consumidor_otras_drogas,
                    "tiempo_abstinencia_otras_drogas": tiempo_abstinencia_otras_drogas,
                    "actividad_física":actividad_física,
                    "cual_actividad_fisica":cual_actividad_fisica,
                    "tiempo_actividad_fisica":tiempo_actividad_fisica,
                    "cantidad_actividad_fisica":cantidad_actividad_fisica,
                    "medicacion_habitual":medicacion_habitual,
                    "cual_medicacion_habitual":cual_medicacion_habitual,
                    "tiempo_medicacion_habitual": tiempo_medicacion_habitual,
                    "cantidad_medicacion_habitual": cantidad_medicacion_habitual,
                    "principales_incidentes": principales_incidentes,
                    "calif_iess_accidente":calif_iess_accidente,
                    "especificar_accidentes": especificar_accidentes,
                    "fecha_accidente":fecha_accidente,
                    "observaciones_accidentes": observaciones_accidentes,
                    "calif_iess_enfermedades":calif_iess_enfermedades,
                    "especificar_enfermedades":especificar_enfermedades,
                    "observaciones_enfermedades":observaciones_enfermedades,
                    #SECCION D
                    "antecedentes_familiares:": ", ".join(antecedentes_familiares),
                    "descripcion_antecedentes_familiares": descripcion_antecedentes_familiares,
                    #SECCION E
                    "factores_riesgo_fisico": ", ".join(riesgo_fisico),
                    "factores_riesgo_mecanico": ", ".join(riesgo_mecanico),
                    "factores_riesgo_quimico": ", ".join(riesgo_quimico),
                    "factores_riesgo_biologico": ", ".join(riesgo_biologico),
                    "factores_riesgo_ergonomico": ", ".join(riesgo_ergonomico),
                    "factores_riesgo_psicosocial": ", ".join(riesgo_psicosocial),
                    "puesto_trabajo_area": puesto_trabajo_area,
                    "medidas_preventivas": medidas_preventivas,
                    #SECCION F
                    "enfrmedad_actual": enfrmedad_actual,
                    #SECCION G
                    "revision_organos_sistemas": ", ".join(revision_organos_sistemas),
                    "patologia":patologia,
                    #SECCION H
                    "presion_arterial":presion_arterial,
                    "temperatura":temperatura,
                    "frecuncia_cardiaca":frecuncia_cardiaca,
                    "saturacion_oxigeno":saturacion_oxigeno,
                    "frecuencia_respiratoria":frecuencia_respiratoria,
                    "peso":peso,
                    "talla":talla,
                    "imc":imc,
                    "perimetro_abdominal":perimetro_abdominal,
                    #SECCION I
                    "examen_fisico_piel":", ".join(examen_fisico_piel),
                    "examen_fisico_ojos":", ".join(examen_fisico_ojos),
                    "examen_fisico_oido":", ".join(examen_fisico_oido),
                    "examen_fisico_oro_faringe":", ".join(examen_fisico_oro_faringe),
                    "examen_fisico_nariz":", ".join(examen_fisico_nariz),
                    "examen_fisico_cuello":", ".join(examen_fisico_cuello),
                    "examen_fisico_torax":", ".join(examen_fisico_torax),
                    "examen_fisico_abdomen":", ".join(examen_fisico_abdomen),
                    "examen_fisico_columna":", ".join(examen_fisico_columna),
                    "examen_fisico_pelvis":", ".join(examen_fisico_pelvis),
                    "examen_fisico_extremidades":", ".join(examen_fisico_extremidades),
                    "examen_fisico_neurologico":", ".join(examen_fisico_neurologico),
                    "descripcion_examen":descripcion_examen,
                    #SECCION J
                    "fecha_examen_ekg": fecha_examen_ekg,
                    "resultado_examen_ekg": resultado_examen_ekg,
                    "fecha_examen_dco": fecha_examen_dco,
                    "resultado_examen_dco": resultado_examen_dco,
                    #SECCION K
                    "descripcion_diagnostico": descripcion_diagnostico,
                    "cie_diagnostico": cie_diagnostico,
                    "pre_diagnostico":pre_diagnostico,
                    "def_diagnostico":def_diagnostico,
                    #SECCION L
                    "aptitud_medica":", ".join(aptitud_medica),
                    "observaciones_aptitud":observaciones_aptitud,
                    "limitaciones_aptitud":limitaciones_aptitud,
                    #SECCION M
                    "recomendaciones_tratamiento": recomendaciones_tratamiento,
                    #SECCION N
                    "fecha_dato_profecional":fecha_dato_profecional,
                    "hora_dato_profecional": hora_dato_profecional,
                    "codigo_dato_profecional":codigo_dato_profecional,
                    "firma_dato_profecional": firma_dato_profecional,
                    # Metadatos
                    "Fecha_registro": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }
                
                # Guardar en Excel
                archivo_guardado = guardar_en_excel(datos, "valoracion_medica")
                if archivo_guardado:
                    st.success("¡Datos guardados exitosamente!")
            else:
                st.error("Ingrese el RUC por favor")

# --- Formulario de DENSITOMETRÍA ---
elif st.session_state.current_form == "densitometria": 
    st.image("ESCUDO.png")
    with st.form("form_densitometria"):
        
        st.header("DENSITOMETRIA OSEA PORTATIL",divider="blue")
        ruc_den=st.text_input("RUC",key="ruc_den")
        #DENSITOMETRIA OSEA PORTATIL
        col1,col2=st.columns([3,1])
        with col1:
            st.image("DENSITOMETRIA.png",width=650)
        with col2:
            st.image("ESQUELETO.png",width=215)  
            opciones_dolor = [
                    "CADERA",
                    "CLAVICULA",
                    "CODO",
                    "COLUMNA CERVICAL",
                    "CRÁNEO",
                    "CUBITO",
                    "CULUMNA LUMBAR",
                    "ESTERNÓN",
                    "FÉMUR",
                    "HOMBRO",
                    "HÚMERO",
                    "MANO",
                    "MUÑECA",
                    "PARRILLA COSTAL",
                    "PELVIS",
                    "PERONÉ",
                    "PIE",
                    "RADIO",
                    "RODILLA",
                    "RÓTULA",
                    "TIBIA",
                    "TOBILLO"]
            dolor = st.multiselect(
             "SELECIONE DONDE LE DUELE",
            opciones_dolor)
        #PARAMTEROS DE VALORACIÓN DE SENSITOMETRÍA OSEA        
        st.header("PARAMTEROS DE VALORACIÓN DE SENSITOMETRÍA OSEA", divider="blue")
        col1,col2,col3=st.columns([1,2,3])
        with col1:
            densitometria_osea=st.text_area("RESULTADO",key="densitometria_osea")
        with col2:
            st.info("""
                    - **NORMAL** > -1 a 3  
                    - **OSTEOPENIA** -1 a -2.5  
                    - **OSTEOPOROSIS** -2.51 a -5
                    """)
        with col3:
            st.image("NORMAL.png")
        #RASTREO ECOGRAFICO
        st.header("RASTREO ECOGRAFICO",divider="blue")
        st.info(""" INDICACIONES: Tomar agua 5 litros aproximadamente, para ecografía de próstata y útero """)
        col1, col2, col3, col4, col5, = st.columns(5)
        with col1:
            higado=st.text_input("HIGADO",key="higado")
        with col2:
            riñon=st.text_input("RIÑON",key="riñon")
        with col3:
            vesicula=st.text_input("VESICULA",key="vesicula")
        with col4:
            prostata=st.text_input("PROSTATA",key="prostata")
        with col5:
            utero=st.text_input("UTERO",key="utero")
        
        #TRATAMIENO
        st.header("TRATAMIENTO",divider="blue")
        # 1. Inicializar tabla si no existe
        if 'productos_df' not in st.session_state:
            st.session_state.productos_df = pd.DataFrame(columns=[
                'CANTIDAD', 'PRODUCTO', 'PVP', 'TOTAL'
            ])
        total_general = 0.0

        # 2. Controles para agregar productos
        col1, col2, col3 = st.columns([1, 3, 1])
        with col1:
            cantidad = st.number_input("Cantidad", min_value=0, value=0, key='prod_cantidad')
        with col2:
            producto = st.text_input("Producto", key='prod_descripcion')
        with col3:
            pvp = st.number_input("PVP ($)", min_value=0.0, value=0.0, step=0.01, key='prod_pvp')
            
        if st.form_submit_button("Agregar Producto", use_container_width=True):
            if producto.strip() != "":
                total = cantidad * pvp if cantidad > 0 else pvp
                nuevo_registro = pd.DataFrame([[cantidad if cantidad > 0 else 1, producto, pvp, total]], 
                                            columns=['CANTIDAD', 'PRODUCTO', 'PVP', 'TOTAL'])
                st.session_state.productos_df = pd.concat([st.session_state.productos_df, nuevo_registro], ignore_index=True)

        # 3. Mostrar tabla (solo eliminación de filas)
        if not st.session_state.productos_df.empty:
            edited_df = st.data_editor(
                st.session_state.productos_df,
                column_config={
                    "CANTIDAD": st.column_config.NumberColumn(format="%d", disabled=True),
                    "PRODUCTO": st.column_config.TextColumn(disabled=True),
                    "PVP": st.column_config.NumberColumn(format="$%.2f", disabled=True),
                    "TOTAL": st.column_config.NumberColumn(format="$%.2f", disabled=True)
                },
                num_rows="dynamic",
                hide_index=True,
                use_container_width=True,
                key="productos_editor",
                disabled=["CANTIDAD", "PRODUCTO", "PVP", "TOTAL"]
            )
            
            # Actualizar si se eliminan filas
            if len(edited_df) < len(st.session_state.productos_df):
                st.session_state.productos_df = edited_df.copy()
            
            # Calcular total general
            total_general = st.session_state.productos_df['TOTAL'].sum()
            col1, col2 = st.columns([4, 1])
            with col2:
                    # Botón para actualizar (usando form_submit_button)
                actualizar = st.form_submit_button("Actualizar", 
                                                    help="Actualizar el total general",
                                                    use_container_width=True)
                if actualizar:
                    pass
            with col1:
                st.metric("TOTAL GENERAL", f"${total_general:,.2f}")             
        else:
            st.info("No se han agregado productos")
            

        submitted = st.form_submit_button("💾 Guardar Datos")
        if submitted:
            if ruc_den:
                datos = {
                    "ruc": ruc_den,
                    "dolor":", ".join(dolor),
                    "densitometria_osea":densitometria_osea,
                    "higado":higado,
                    "riñon":riñon,
                    "vesicula":vesicula,
                    "prostata":prostata,
                    "utero":utero,
                    "productos": st.session_state.productos_df.to_dict('records'),
                    "total_venta": total_general  
                }
                archivo_guardado = guardar_en_excel(datos, "densitometria")
                if archivo_guardado:
                    st.session_state.datos_guardados = archivo_guardado
                    st.success("¡Datos agregados exitosamente!")
            else:
                st.warning("Ingrese el RUC por favor")


# --- Formulario de AUDIOMETRÍA ---
elif st.session_state.current_form == "audiometria":
    st.image("ESCUDO.png")
    
    with st.form("form_audiometria"):
        #AUDIOMETRÍA OCUPACIONAL
        st.header("AUDIOMETRÍA OCUPACIONAL", divider="blue")
        #SECCION 1
        col1, col2 = st.columns([1,4])
        with col1:
            ruc_aud=st.text_input("RUC",key="ruc_aud")
        with col2:    
            consulta_audiometria=st.text_input("Consulta",key="consulta_audiometria")
        col1, col2, col3 = st. columns(3)
        with col1:
            opciones_audiometria = [
                    "Otalgia",
                    "Otorrea",
                    "Sensación de oido tapado",
                    "Vértigo",
                    "Acúfnos",
                    "Prurito"]
            audiometria_sintomas = st.multiselect("A presentado",opciones_audiometria)
        with col2:
            oido_mejor = st.selectbox("¿Por cuál oído oye mejor?", ["AMBOS","OI", "OD"])
        with col3:
            familiar_hipoacusia=st.selectbox("¿Algún familiar con hipoacusia?", ["NO", "SI"])
        col1, col2, col3 = st.columns(3)
        with col1:
            expuesto_ruidos=st.selectbox("¿Está expuesto a ruidos fuertes?", ["NO", "SI"])
        with col2:
            cirugias_oido=st.selectbox("¿Cirugías en el oido?", ["NO", "SI"])
        with col3:
            usa_audifonos=st.selectbox("¿Usa o ha usado audifonos?", ["NO", "SI"])
        col1, col2, col3 = st.columns(3)
        with col1:
            opcines_audifonos=[
                "Intrauriculares",
                "Retroauriculares",
                "Marca"]
            tipo_audifono=st.multiselect("Tipo de audifono",opcines_audifonos)
        with col2:
            tiempo_uso=st.number_input("Tiempo de uso", min_value=0,value=0)
        with col3:
            otros_audiometria=st.text_input("Otros")
        #SECCION 2
        st.header("DATOS CLINICOS", divider="blue")
        col1, col2 = st. columns(2)
        with col1:
            ostocopia_oido_derecho=st.text_input("Otoscopia Oido Derecho", key="ostocopia_oido_derecho")
            st.image("AO.png")
        with col2:
            ostocopia_oido_izquierdo=st.text_input("Otoscopia Oido Derecho", key="ostocopia_oido_izquierdo")
            st.image("AO.png")
        col1, col2, col3 = st.columns([1,0.5,1])
        with col1:
           pta_oido_dercho=st.text_input("PTA Oido Derecho",key="pta_oido_dercho")
        with col2:
            opciones_weber=[
                "500",
                "1000",
                "2000",
                "4000"] 
            weber=st.multiselect("WEBER",opciones_weber)
        with col3:
            pta_oido_izquierdo=st.text_input("PTA Oido Izquierdo",key="pta_oido_izquierdo")
        
        st.header("ESPIROMETRÍA OCUPACIONAL")
        col1, col2, col3 = st.columns([1,1,2])
        with col1:
            fuma=st.selectbox("Fuma", ["NO","SI"])
        with col2:
            expuesto_humo=st.selectbox("Expuesto a humo del tabaco", ["NO","SI"])
        with col3:
            expuesto_contaminacion=st.selectbox("Expuesto a contaminación ambiental en forma de gases o polvo",["NO","SI"])
        st.subheader("Enfermdades diagnosticadas")
        col1, col2, col3, col4, col5 =st.columns([1,1,1,1,2])
        with col1:
            sinusitis=st.selectbox("Sinusitis", ["NO","SI"])
        with col2:
            asma=st.selectbox("Asma", ["NO","SI"])
        with col3:
            efisema=st.selectbox("Efisema", ["NO","SI"])
        with col4:
            epoc=st.selectbox("EPOC", ["NO","SI"])
        with col5:
            equipo_auditivo=st.selectbox("Utiliza equipo de proteción auditiva", ["NO","SI"])
    
        submitted = st.form_submit_button("💾 Guardar Datos")
        if submitted:
            if ruc_aud:
                datos = {
                    #SECCION 1
                    "ruc": ruc_aud,
                    "consulta_audiometria":consulta_audiometria,
                    "audiometria_sintomas":", ".join(opciones_audiometria),
                    "oido_mejor":oido_mejor,
                    "familiar_hipoacusia":familiar_hipoacusia,
                    "expuesto_ruidos":expuesto_ruidos,
                    "cirugias_oido":cirugias_oido,
                    "tipo_audifono":", ".join(tipo_audifono),
                    "tiempo_uso":tiempo_uso,
                    "otros_audiometria":otros_audiometria,
                    #SECCION 2
                    "ostocopia_oido_derecho":ostocopia_oido_derecho,
                    "ostocopia_oido_izquierdo":ostocopia_oido_izquierdo,
                    "pta_oido_dercho":pta_oido_dercho,
                    "pta_oido_izquierdo":pta_oido_izquierdo,
                    "weber":", ".join(weber),
                    #SECCION 3
                    "fuma":fuma,
                    "expuesto_humo": expuesto_humo,
                    "expuesto_contaminacion": expuesto_contaminacion,
                    "sinusitis":sinusitis,
                    "asma":asma,
                    "efisema":efisema,
                    "epoc":epoc,
                    "equipo_auditivo":equipo_auditivo    
                }
                archivo_guardado = guardar_en_excel(datos, "audiometria")
                if archivo_guardado:
                    st.session_state.datos_guardados = archivo_guardado
                    st.success("¡Datos agregados exitosamente!")
            else:
                st.warning("Agrege el RUC por favor")


# --- Formulario de CERTIFICADO ---
elif st.session_state.current_form == "certificado":
    st.image("ESCUDO.png")
    st.header("CERTIFICADO DE APTITUD LABORAL")
    
    with st.form("form_certificado"):
        
        #SECCIÓN A
        fecha_emision = st.date_input("Fecha de emisión", value=datetime.now())
        st.subheader("A. DATOS DEL ESTABLECIMIENTO- EMPRESA Y USUARIO", divider="blue")
        col1, col2,col3, col4=st.columns([2,1,1,1])
        with col1:
            institucion_certificado=st.text_input("INSTITUCIÓN DEL SISTEMA O NOMBRE DE LA EMPRESA")
        with col2:
            ruc_certificado=st.text_input("RUC")
        with col3:
            ciu_certificado=st.text_input("CIU")
        with col4:
            establecimiento_certificado=st.text_input("ESTABLECIMIENTTO DE SALUD")   
        col1, col2,col3=st.columns([1,1,1])
        with col1:
            historia_certificado=st.text_input("NÚMERO DE HISTORIA CLINICA")
        with col2:
            archivo_certificado=st.text_input("NÚMERO DE ARCHIVO")
        with col3:
            cargo_certificado=st.text_input("CARGO/OCUPACIÓN")
        col1, col2,col3, col4, col5 = st.columns([1,1,1,1,1])
        with col1:
            primer_apellido_certificado=st.text_input("PRIMER APELLIDO")
        with col2:
            segundo_apellido_certificado=st.text_input("SEGUNDO APELLIDO")
        with col3:
            primer_nombre_certificado=st.text_input("PRIMER NOMBRE")
        with col4:
            segundo_nombre_certificado=st.text_input("SEGUNDO NOMBRE")
        with col5:
           sexo_certificado=st.selectbox("SEXO",["MASCULINO","FEMENINO"] )
        #SECCION B
        st.header("B. DATOS GENERALES", divider="blue")
        col1, col2 = st.columns([1,4])
        with col1:
            fecha_certificado=st.date_input("Fecha de emisión", value=datetime.now(),key="fecha_certificado")
        with col2:
            evaluacion= st.radio("EVALUACIÓN", ["INGRESO", "PERIODICO","REINGRESO","SALIDA"], key="evaluacion", horizontal=True)
        #SECCION C
        st.subheader("C. CONCEPTO PARA APTITUD LABORAL",divider="blue")
        st.info("Después de la valoración médica ocupacional se certifica que la persona en meción, es calificada como")
        estado= st.radio("ESTADO", ["APTO", "APTO OBSERVACION","APTO CON LIMITACIONES","NO APTO"], key="estado", horizontal=True)
        #SECCION D
        st.subheader("D. CONDICIONES DE SALUD AL MOMENTO",divider="blue")
        st.info("Después de la valoración médica ocupacional se certifica las condicions de salud al momento")
        condiciones= st.radio("ESTADO", ["SATISFACTORIO", "NO SATISFACTORIO"], key="condiciones", horizontal=True)
        observaciones_certificado=st.text_area("OBSERVACIONES RELACIONADAS CON LAS CONDICIONES DE SALUD AL MOMENTO DEL RETIRO:",key="observaciones_certificado")
        #SECCION E
        st.subheader("E. RECOMNDACIONES",divider="blue")
        recomendaciones_certificado=st.text_area(" ")
        
        st.info("Con este documento certifico que el trabajador se ha sometido a la evaluación médica requerida para (el ingreso/la ejecución/ el reintegro y retiro) al puesto laboral y se ha informado sobre los riesgos relacionados con el trabajo emitiendo recomendacions relacionadas con su estado de salud")
        st.info("La presente certificación se expide con base en la historia del usuario(a), la cuál tiene carácter de confidencial")
        
        col1, col2, col3, col4=st.columns([1,1,1,3])
        with col1:
            st.text_area("NOMBRE Y APELLIDO")
        with col2:
            st.text_area("CÓDIGO")
        with col3:
            st.text_area("FIRMA Y SELLO")
        with col4:
            st.text_area("FIRMA DEL USUARIO")
            
        submitted = st.form_submit_button("💾 Guardar Datos")
        if submitted:
            if ruc_certificado:
                datos = {
                    #SECCION 1
                    "ruc": ruc_certificado,
                      
                }
                archivo_guardado = guardar_en_excel(datos, "certificado")
                if archivo_guardado:
                    st.session_state.datos_guardados = archivo_guardado
                    st.success("¡Datos agregados exitosamente!")
            else:
                st.warning("Agrege el RUC por favor")
                
                    
# Mensaje inicial
if st.session_state.current_form is None:
    st.info("Buen día ¿Cómo podemos ayudarte el día de hoy?")
    st.image("INICIO.jpeg") 
  
# Nota al pie
st.markdown("---")
st.caption("SISTEMA DE REGISTRO| FUNADACIÓN HOSPITALARIA SEMINARIO MENOR SAN LUÍS v2.0 | © 2025 |")