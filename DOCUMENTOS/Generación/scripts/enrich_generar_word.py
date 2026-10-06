import os
import re

word_script_path = r"c:\GRESLY\DOCUMENTOS\Generación\scripts\generar_word.py"
with open(word_script_path, "r", encoding="utf-8") as f:
    code = f.read()

# 1. New TOC Items with exact subheadings
new_toc = """    toc_items = [
        ("INTRODUCCIÓN", "1", True, 0, 4),
        ("CAPÍTULO I: EL PROBLEMA DE INVESTIGACIÓN", "3", True, 0, 4),
        ("1.1. Descripción de la realidad problemática", "3", False, 1, 3),
        ("1.2. Formulación del problema (Preguntas norteadoras)", "4", False, 1, 3),
        ("1.3. Objetivos de la investigación", "5", False, 1, 3),
        ("1.4. Supuestos fenomenológicos de la investigación", "5", False, 1, 3),
        ("1.5. Justificación e importancia de la investigación", "6", False, 1, 3),
        ("1.6. Delimitaciones y limitaciones del estudio", "7", False, 1, 3),
        ("1.7. Viabilidad y consideraciones bioéticas", "8", False, 1, 4),
        ("CAPÍTULO II: MARCO CONTEXTUAL", "9", True, 0, 4),
        ("2.1. Descripción geográfica, territorial y ambiental del área de estudio", "9", False, 1, 3),
        ("2.2. Características demográficas y estructura poblacional", "10", False, 1, 3),
        ("2.3. Características socioculturales, lingüísticas y cosmovisión andina", "11", False, 1, 3),
        ("2.4. Dinámica socioeconómica, vulnerabilidad y nivel de aseguramiento público", "12", False, 1, 3),
        ("2.5. Características sanitarias, cartera de servicios y capacidad resolutiva institucional", "13", False, 1, 3),
        ("2.6. Articulación dialéctica del escenario físico-social con las dimensiones existenciales del usuario (Lebenswelt)", "14", False, 1, 4),
        ("REFERENCIAS BIBLIOGRÁFICAS", "16", True, 0, 4),
        ("ANEXOS", "18", True, 0, 4),
        ("Anexo 1: Guía de entrevista a profundidad fenomenológica", "18", False, 1, 3),
        ("Anexo 2: Matriz de consistencia cualitativa fenomenológica", "19", False, 1, 3),
        ("Anexo 3: Formato de juicio de expertos", "20", False, 1, 3),
        ("Anexo 4: Consentimiento informado", "21", False, 1, 3),
        ("Anexo 5: Carta de asesoría formal", "22", False, 1, 3),
    ]"""

pattern_toc = r"    toc_items = \[\n.*?    \]\n"
code = re.sub(pattern_toc, new_toc + "\n", code, flags=re.DOTALL)

# 2. Comprehensive Capítulo I text replacing lines from sec_cap1 to sec_cap2
new_cap1_block = '''    # ==========================================================================
    # SECCIÓN 4: CAPÍTULO I - EL PROBLEMA DE INVESTIGACIÓN
    # ==========================================================================
    sec_cap1 = doc.add_section(WD_SECTION.NEW_PAGE)
    sec_cap1.page_width = Cm(21.0)
    sec_cap1.page_height = Cm(29.7)
    sec_cap1.top_margin = Cm(3.0)
    sec_cap1.left_margin = Cm(3.5)
    sec_cap1.right_margin = Cm(2.5)
    sec_cap1.bottom_margin = Cm(2.5)
    sec_cap1.different_first_page_header_footer = False
    sec_cap1.header.is_linked_to_previous = False
    sec_cap1.footer.is_linked_to_previous = False
    sec_cap1.footer.paragraphs[0].text = ""
    set_section_pgnum(sec_cap1, start=None, fmt='decimal')
    setup_header(sec_cap1, "CAPÍTULO I: EL PROBLEMA DE INVESTIGACIÓN")
    
    add_heading_1(doc, "CAPÍTULO I: EL PROBLEMA DE INVESTIGACIÓN")
    add_heading_2(doc, "1.1. Descripción de la realidad problemática")
    add_body_p(doc, "En el marco epistemológico de la investigación cualitativa, tal como sostienen Hernández-Sampieri y Mendoza [5], el planteamiento del problema adopta una naturaleza puramente inductiva y fenomenológica: no parte del aislamiento artificial de variables cuantitativas ni del contraste de hipótesis numéricas preconcebidas, sino de la inmersión vivencial directa en el ambiente natural donde acontecen los fenómenos humanos. Plantear y delimitar un problema cualitativo fenomenológico constituye un proceso reflexivo, dinámico y abierto orientado a capturar la esencia, las emociones, los sentires corpóreos y los significados intersubjetivos que las personas otorgan a sus vivencias en el momento del acto asistencial, siguiendo los fundamentos de Husserl [7] y van Manen [8].")
    add_body_p(doc, "Para tal fin, la metodología demanda un contacto directo y vivencial con la realidad estudiada en su escenario cotidiano. En el presente estudio, el origen de la indagación se sustenta en la observación reflexiva y la experiencia acumulada por la investigadora durante sus prácticas preprofesionales de la Escuela Profesional de Enfermería de la Universidad Nacional de San Cristóbal de Huamanga (UNSCH) en los consultorios del Centro de Salud Belén.")
    add_body_p(doc, "En estricta observancia del canon metodológico de la cátedra de EN 486 (Dr. Manglio Aguirre Andrade), la caracterización de la realidad problemática se articula mediante la técnica del embudo en tres niveles de profundidad:")
    
    add_heading_3(doc, "Contexto Internacional (Macro)")
    add_body_p(doc, "A escala mundial, los sistemas de salud enfrentan una crisis estructural caracterizada por la fragmentación asistencial y la despersonalización del acto de cuidado, según lo advertido por la Organización Mundial de la Salud (OMS), la OCDE y el Banco Mundial [1]. El predominio de un paradigma biomédico tecnicista y la sobrecarga burocrática en el primer nivel de atención han reducido la evaluación de los servicios a meros indicadores de rendimiento cuantitativo, invisibilizando la vivencia afectiva del paciente, el sufrimiento somático de la espera y los determinantes socioculturales de la salud. De acuerdo con Donabedian [9, 10] y Watson [12], la calidad del cuidado no reside únicamente en la corrección instrumental del procedimiento terapéutico, sino medularmente en el proceso interpersonal: la calidez de la mirada, la capacidad de escucha, el respeto a la dignidad y la reciprocidad humanizada entre el profesional y el ser cuidado.")
    
    add_heading_3(doc, "Contexto Nacional (Meso)")
    add_body_p(doc, "En el Perú, la Política Nacional de Calidad en Salud del Ministerio de Salud [2] reconoce que la persistencia de barreras de acceso geográficas, organizacionales, económicas y culturales debilita gravemente la capacidad resolutiva del primer nivel asistencial (categorías I-1 a I-4), donde recae más del 78% de las atenciones asistenciales del país. A pesar de los esfuerzos normativos, persisten cuellos de botella crónicos: colas en la madrugada para conseguir un cupo, desabastecimiento de medicamentos esenciales del petitorio oficial y una marcada frialdad comunicativa en el trato al usuario. Esta brecha motivó al Estado peruano a priorizar la investigación en salud a través de la Línea 10 de Investigación en Salud al 2030: “Sistemas y servicios de salud, acceso y cobertura universal” (Resolución Ministerial N° 424-2025/MINSA) [3], orientada a generar evidencia empírica directa para dignificar la atención y garantizar la cobertura sanitaria universal.")
    
    add_heading_3(doc, "Contexto Regional y Local (Micro: Centro de Salud Belén)")
    add_body_p(doc, "En el ámbito regional, el estudio se articula de forma vinculante con la Prioridad 10 de Investigación en Salud de Ayacucho 2025–2030, aprobada por la Dirección Regional de Salud de Ayacucho [4]: “Organización, gestión, calidad y accesibilidad de los servicios de salud”. En este marco, el Centro de Salud Belén (Categoría I-3, Microred Huamanga) representa un punto neurálgico asistencial que atiende a una población adscrita de 18,450 habitantes en el distrito de Ayacucho, a 2,760 metros sobre el nivel del mar.")
    add_body_p(doc, "En este establecimiento aflora una dramática distancia entre la lógica asistencial programática y la vivencia sentida del usuario andino:")
    add_bullet_item(doc, "Vulnerabilidad socioeconómica: ", "El 68.2% de la población usuaria se encuentra en situación de pobreza o extrema pobreza según el SISFOH, y más del 94.5% depende del Seguro Integral de Salud (SIS) en su modalidad subsidiada.")
    add_bullet_item(doc, "Identidad lingüística y cultural: ", "El 72.0% de los usuarios presenta bilingüismo activo quechua chanka-castellano. Para esta población, la salud se vivencia a través de códigos ancestrales de afecto y reciprocidad comunitaria: la acogida hospitalaria (allin chaskiy), el saludo cálido y respetuoso (napaykuy) y el respeto mutuo (respetanakuy), conforme a los hallazgos de Aguirre-Andrade y colaboradores [18].")
    add_bullet_item(doc, "La experiencia de la desatención: ", "Cuando los profesionales —condicionados por la prisa protocolar, la sobrecarga laboral o el monolingüismo hispanohablante— omiten el saludo, evitan el contacto visual o se muestran indiferentes ante el dolor (nanay), el usuario experimenta sensaciones de desamparo, intimidación y violencia simbólica, replegándose en el silencio defensivo (upallay). Asimismo, las madrugadas gélidas a la intemperie (5 °C a 8 °C) desde las 4:00 AM para obtener un cupo generan agotamiento somático y menoscabo de la dignidad de las madres gestantes y adultos mayores.")
    add_body_p(doc, "El problema medular de investigación radica, por consiguiente, en develar y comprender cómo los propios usuarios perciben, vivencian y significan la calidad de atención y el acceso asistencial en el Centro de Salud Belén desde la profundidad de su mundo de la vida.")
    
    add_heading_2(doc, "1.2. Formulación del problema (Preguntas norteadoras)")
    add_body_p(doc, "En coherencia con los preceptos del diseño fenomenológico articulados por Hernández-Sampieri y Mendoza [5] y el principio de inmutabilidad del entorno sellado, se formulan las preguntas norteadoras de la investigación:")
    
    add_heading_3(doc, "Problema General (Pregunta Principal)")
    add_body_p(doc, "¿Cuál es el significado, estructura y esencia de la experiencia vivida por los usuarios respecto al fenómeno del acceso y la calidad de atención en los servicios de salud del Centro de Salud Belén, Ayacucho 2026?")
    
    add_heading_3(doc, "Problemas Específicos (Preguntas Específicas)")
    add_num_item(doc, 1, "¿Cómo vivencian los usuarios la dimensión temporal (tiempo vivido / Lebenszeit) en las filas de madrugada (4:00 AM) y las horas transcurridas en sala de espera antes de ser atendidos?")
    add_num_item(doc, 2, "¿De qué manera experimentan los usuarios la dimensión espacial (espacio vivido / Lebensraum) respecto al hacinamiento, comodidad física y privacidad dentro de los consultorios?")
    add_num_item(doc, 3, "¿Cuáles son las vivencias somáticas y emocionales vinculadas a la dimensión corporal (cuerpo vivido / Leib), tales como frío matutino, fatiga física, hambre y dolor (nanay) frente a la espera asistencial?")
    add_num_item(doc, 4, "¿Cómo perciben los usuarios la dimensión relacional (relación humana vivida / Mitwelt) respecto a la empatía, calidez, escucha activa y comunicación intercultural en quechua chanka brindada por el equipo de salud?")
    add_num_item(doc, 5, "¿Cuál es la esencia compartida y las divergencias vivenciales que configuran el significado integral atribuido por los usuarios a una atención de calidad y digna en el Centro de Salud Belén?")
    
    add_heading_2(doc, "1.3. Objetivos de la investigación")
    add_body_p(doc, "Guardando estricta simetría lógica biunívoca (1:1) con las preguntas norteadoras y empleando verbos en infinitivo propios de la indagación cualitativa comprensiva, se establecen los siguientes objetivos:")
    
    add_heading_3(doc, "Objetivo General")
    add_body_p(doc, "Explorar, comprender y describir el significado, estructura y esencia de la experiencia vivida por los usuarios respecto al acceso y calidad de atención en el Centro de Salud Belén, Ayacucho 2026.")
    
    add_heading_3(doc, "Objetivos Específicos")
    add_num_item(doc, 1, "Explorar y comprender las vivencias de los usuarios respecto a la dimensión temporal (tiempo vivido) en las filas matutinas de madrugada y el tiempo de espera en el establecimiento.")
    add_num_item(doc, 2, "Describir y comprender la experiencia de los usuarios respecto a la dimensión espacial (espacio vivido) en cuanto al confort, hacinamiento y privacidad de los ambientes asistenciales.")
    add_num_item(doc, 3, "Identificar y comprender las vivencias somáticas de la dimensión corporal (cuerpo vivido), analizando la resistencia al frío matutino, fatiga, hambre y dolor del paciente.")
    add_num_item(doc, 4, "Comprender y analizar la experiencia de la dimensión relacional (relación humana vivida) en el encuentro asistencial, focalizando la empatía, trato digno y diálogo en lengua quechua.")
    add_num_item(doc, 5, "Sintetizar la estructura esencial y las divergencias fenomenológicas que configuran una atención digna y de calidad en el Centro de Salud Belén.")
    
    add_heading_2(doc, "1.4. Supuestos fenomenológicos de la investigación")
    add_body_p(doc, "En el paradigma cualitativo fenomenológico, en concordancia con Hernández-Sampieri y Mendoza [5] y van Manen [8], la investigación no formula hipótesis cuantitativas para contrastación estadística paramétrica. En su lugar, se establecen supuestos fenomenológicos inductivos, los cuales constituyen proposiciones teóricas preliminares que orientan la mirada de la investigadora sin prejuzgar ni constreñir las vivencias que emergerán de los informantes durante la reducción fenomenológica (epojé):")
    
    add_heading_3(doc, "Supuesto General")
    add_body_p(doc, "La calidad de atención de salud constituye para el usuario del Centro de Salud Belén una vivencia intersubjetiva profunda donde la calidez humana, el reconocimiento de su dignidad andina, la comunicación empática en lengua originaria y la oportunidad resolutiva del padecimiento prevalecen sobre la mera dimensión instrumental o protocolar del acto asistencial.")
    
    add_heading_3(doc, "Supuestos Específicos del Mundo de la Vida (Lebenswelt)")
    add_num_item(doc, 1, "La prolongada espera a la intemperie y la incertidumbre en la obtención de cupos genera en el usuario una vivencia de desvalorización de su tiempo vital y angustia existencial, transformando el acceso en una experiencia de sacrificio físico.", bold_prefix="Supuesto sobre la Temporalidad: ")
    add_num_item(doc, 2, "El espacio físico precario, la falta de asientos adecuados y la ausencia de privacidad acústica y visual en el consultorio restringen la intimidad del paciente, inhibiendo la revelación libre de sus síntomas y temores íntimos.", bold_prefix="Supuesto sobre la Espacialidad: ")
    add_num_item(doc, 3, "El cuerpo del usuario somatiza la desatención institucional a través del impacto del frío matutino andino (5 °C a 8 °C), el agotamiento muscular y la agudización del dolor físico (nanay), transformando la espera en una vivencia corporalmente aflictiva.", bold_prefix="Supuesto sobre la Corporalidad: ")
    add_num_item(doc, 4, "La interacción dialógica intercultural mediada por la calidez del saludo (napaykuy) y la acogida en quechua chanka (allin chaskiy) mitiga el temor institucional, restablece la confianza terapéutica y resignifica positivamente la calidad del servicio.", bold_prefix="Supuesto sobre la Relacionalidad: ")
    add_num_item(doc, 5, "La esencia del cuidado digno radica en la experiencia compartida de sentirse respetado como persona humana integral, emergiendo divergencias vivenciales moduladas por la edad y el grado de vulnerabilidad socioeconómica del usuario.", bold_prefix="Supuesto sobre la Esencia y Divergencias: ")
    
    add_heading_2(doc, "1.5. Justificación e importancia de la investigación")
    add_body_p(doc, "Conforme a las directrices epistemológicas de la cátedra de EN 486 (UNSCH), fundamentadas en la concepción científica y rigurosa de Mario Bunge [15], y los criterios de relevancia de Hernández-Sampieri y Mendoza [5], el proyecto se sustenta en cuatro dimensiones justificatorias:")
    
    add_heading_3(doc, "Justificación Teórica")
    add_body_p(doc, "El estudio llena un vacío epistémico crítico en la literatura regional sobre calidad de atención en salud pública. Tradicionalmente, la evaluación de servicios se ha sustentado en modelos estandarizados cuantitativos (como SERVQUAL) [19, 20] que reducen la satisfacción a medias aritméticas, omitiendo las complejidades subjetivas y la cosmovisión del habitante andino. Esta investigación articula dialógicamente el Enfoque de Calidad en Salud de Donabedian [9, 10], la Teoría del Acceso de Penchansky [11] y la Teoría del Cuidado Humano de Watson [12] con las cuatro dimensiones existenciales de van Manen [8], construyendo un marco comprensivo pionero con pertinencia cultural andina en la UNSCH.")
    
    add_heading_3(doc, "Justificación Práctica")
    add_body_p(doc, "La investigación responde a una necesidad institucional urgente. Al desentrañar los nudos críticos del acceso y el trato humano desde el relato directo de los usuarios, los hallazgos aportarán insumos concretos para que la Jefatura del Centro de Salud Belén y la Dirección de la Red de Salud Huamanga rediseñen el flujo de admisión matutina, erradiquen las filas a la intemperie en la madrugada y elaboren protocolos de triaje y acogida intercultural con pertinencia lingüística en quechua chanka.")
    
    add_heading_3(doc, "Justificación Metodológica")
    add_body_p(doc, "Aporta al campo de la investigación en enfermería un protocolo fenomenológico riguroso, validado por juicio de expertos y estructurado bajo los cuatro existenciales del Lebenswelt. Diseña y valida una guía de entrevista semiestructurada bilingüe (Quechua Chanka-Español) y una matriz de categorización temática cualitativa auditadas bajo los criterios de rigor de Guba y Lincoln [13]: credibilidad, transferibilidad, consistencia y confirmabilidad, estableciendo un precedente metodológico replicable para futuras tesis de la Facultad de Ciencias de la Salud de la UNSCH.")
    
    add_heading_3(doc, "Justificación Social y Sanitaria")
    add_body_p(doc, "El proyecto encarna un imperativo ético de equidad en salud. Visibiliza el sufrimiento silenciado de las madres de familia en CRED, gestantes y adultos mayores en situación de pobreza (68.2% SISFOH) y alta dependencia del SIS (94.5%). Responde de manera vinculante a la Línea 10 Nacional de Investigación del MINSA al 2030 [3] y a la Prioridad 10 Regional de Investigación de la DIRESA Ayacucho 2025–2030 [4], contribuyendo a transformar la atención de salud en un espacio de dignidad humana y justicia social.")
    
    add_heading_2(doc, "1.6. Delimitaciones y limitaciones del estudio")
    add_body_p(doc, "Para preservar las fronteras del entorno de investigación y garantizar el aislamiento metodológico, se declaran formalmente las delimitaciones del alcance y las limitaciones del estudio:")
    
    add_heading_3(doc, "Delimitación Espacial y Territorial")
    add_body_p(doc, "La investigación se circunscribe de forma estricta a los ambientes físicos del Centro de Salud Belén (Categoría I-3), abarcando la vereda exterior de espera matutina, las salas interiores de espera y los consultorios asistenciales de Medicina General, Enfermería (CRED, Inmunizaciones), Obstetricia y Odontología, ubicado en el distrito de Ayacucho, provincia de Huamanga, departamento de Ayacucho, a 2,760 msnm.")
    
    add_heading_3(doc, "Delimitación Temporal")
    add_body_p(doc, "El horizonte temporal de recolección y análisis vivencial corresponde al año fiscal y académico 2026.")
    
    add_heading_3(doc, "Delimitación Teórica y Epistemológica")
    add_body_p(doc, "El marco teórico se limita de manera exclusiva a los modelos de Calidad Asistencial de Donabedian [9], la Teoría de Acceso de Penchansky [11], la Teoría del Cuidado Humano de Watson [12] y la fenomenología hermenéutica de Husserl [7] y van Manen [8]. Todo constructo positivista ajeno a estas fronteras queda excluido.")
    
    add_heading_3(doc, "Delimitación Conceptual y Temática")
    add_body_p(doc, "El fenómeno se investiga a través del prisma exclusivo de las cuatro dimensiones existenciales del mundo de la vida (Lebenswelt): Temporalidad (Lebenszeit), Espacialidad (Lebensraum), Corporalidad (Leib) y Relacionalidad (Mitwelt), más la síntesis estructural esencial.")
    
    add_heading_3(doc, "Delimitación Social y Poblacional")
    add_body_p(doc, "La unidad de estudio está constituida por personas usuarias mayores de 18 años, hispanohablantes o quechuahablantes, que acuden como usuarios externos a los servicios asistenciales del establecimiento. Se excluye al personal administrativo y asistencial, cuya vivencia laboral constituye un objeto de estudio independiente.")
    
    add_heading_3(doc, "Limitaciones Metodológicas del Estudio")
    add_bullet_item(doc, "Naturaleza no probabilística del diseño: ", "Los hallazgos fenomenológicos no persiguen la generalización estadística numérica de frecuencias poblacionales, sino la transferibilidad conceptual contextualizada según los criterios de rigor de Guba y Lincoln [13].")
    add_bullet_item(doc, "Efecto de reactividad y pudor: ", "La presencia de la investigadora y el uso de grabadoras de audio podrían generar reticencia inicial en ciertos informantes quechuahablantes. Dicha limitación se mitiga mediante una inmersión prolongada en el campo, el establecimiento de rapport empático y el diálogo en su lengua materna.")
    add_bullet_item(doc, "Condicionamiento biometeorológico: ", "Las variaciones estacionales de frío y lluvia pueden modular la vivencia somática, lo que demanda registrar notas de campo reflexivas precisas en cada sesión de entrevista.")
    
    add_heading_2(doc, "1.7. Viabilidad y consideraciones bioéticas")
    add_body_p(doc, "La presente investigación es plenamente viable puesto que cuenta con acceso institucional favorable a las instalaciones del Centro de Salud Belén, disponibilidad de recursos humanos calificados, respaldo académico tutorial de la cátedra de EN 486 (Dr. Manglio Aguirre Andrade) y financiamiento asegurado en su totalidad con recursos propios de la investigadora principal.")
    add_body_p(doc, "En estricta observancia del rigor bioético, el protocolo será sometido al Comité Institucional de Ética en Investigación (CIEI) de la Universidad Nacional de San Cristóbal de Huamanga (UNSCH) y coordinado formalmente con la Dirección Regional de Salud (DIRESA) de Ayacucho y la Red de Salud Huamanga antes de cualquier abordaje de campo.")
    add_body_p(doc, "El estudio se rige bajo los principios de la Declaración de Helsinki [16], las pautas internacionales del CIOMS [17] y el Informe Belmont:")
    add_bullet_item(doc, "Principio de Autonomía: ", "Se garantizará la voluntariedad absoluta a través de la suscripción formal del Consentimiento Informado (Anexo 4), explicándose de forma detallada y comprensible la naturaleza de la investigación en quechua chanka o castellano.")
    add_bullet_item(doc, "Principio de Beneficencia y No Maleficencia: ", "Se asegurará un ambiente de escucha digno y contenedor, evitando cualquier revictimización o angustia emocional durante el relato vivencial. Las entrevistas podrán suspenderse si el informante lo solicita.")
    add_bullet_item(doc, "Principio de Justicia y Confidencialidad: ", "Se mantendrá el anonimato absoluto mediante el uso de códigos alfanuméricos despersonalizados (ejemplo: Informante 01 -- [INF-01-MED]) en todas las etapas de transcripción, categorización en ATLAS.ti v9 y redacción del informe final.")
'''

# Replace Section 4 in code
pattern_sec4 = r"    # ==========================================================================\n    # SECCIÓN 4: CAPÍTULO I.*?(?=    # ==========================================================================\n    # SECCIÓN 5: CAPÍTULO II)"
code = re.sub(pattern_sec4, new_cap1_block, code, flags=re.DOTALL)

with open(word_script_path, "w", encoding="utf-8") as f:
    f.write(code)

print("generar_word.py actualizado con Capítulo I completo y detallado.")
