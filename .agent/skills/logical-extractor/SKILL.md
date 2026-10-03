---
name: logical-extractor
description: Extrae y blinda las relaciones causa-efecto, datos empíricos duros, citas y tesis central de cada párrafo antes de cualquier edición o parafraseo. Basado en el patrón analyze_prose de Fabric.
---

# Logical Extractor: Extractor y Blindador Lógico Forense

Inspirado en el patrón open-source `analyze_prose` del proyecto Fabric (Daniel Miessler), esta habilidad actúa como un **auditor lógico-semántico preliminar**. Su objetivo fundamental es deconstruir cualquier texto académico o analítico, aislar sus componentes esenciales de verdad y crear una **Matriz de Invarianza Semántica** que blinde los hechos antes de que cualquier proceso de reescritura o parafraseo tenga lugar.

---

## 🎯 Propósito y Filosofía
En la edición y parafraseo asistido por IA, el riesgo más grave no es el estilo, sino la **alucinación por dilución causal**: cuando la IA reescribe una frase y altera sutilmente el sujeto, invierte la causa y el efecto, o suprime un dato numérico o territorial indispensable.

`logical-extractor` previene este error ejecutando una disección analítica previa, garantizando que el núcleo lógico quede completamente blindado.

---

## 🛠️ Procedimiento de Extracción Paso a Paso

Cuando se invoque este skill sobre un texto o fragmento, ejecuta en orden estricto los siguientes pasos:

### 1. Deconstrucción Párrafo a Párrafo
Divide el texto de entrada en unidades de pensamiento discretas (párrafos o proposiciones principales) y evalúa su propósito comunicativo.

### 2. Identificación de la Tesis Central y Proposiciones Nucleares
* **Tesis / Idea Fuerza:** ¿Cuál es la aserción principal indiscutible que sostiene el fragmento?
* **Sub-argumentos de apoyo:** ¿Cuáles son las premisas que fundamentan dicha aserción?

### 3. Blindaje de Relaciones Causa-Efecto (Nexo Causal)
Identifica explícitamente cada cadena de causalidad presente en el texto:
* **Condición / Causa ($A$):** Qué fenómeno o evento desencadena el efecto.
* **Mecanismo / Nexo Causal ($\to$):** A través de qué vía o proceso se produce la consecuencia.
* **Consecuencia / Efecto ($B$):** Cuál es el resultado empírico o conceptual observado.
* *Regla de blindaje:* La posterior reescritura jamás debe convertir una causalidad en correlación débil, ni alterar el sentido de la implicación ($A \implies B$).

### 4. Aislamiento de Datos Duros y Evidencia Inmutable
Extrae y cataloga todos los elementos fácticos que no admiten sinónimos ni aproximaciones:
* **Cifras y Datos Numéricos:** Porcentajes (ej. 68.2%, 94.5%, 72%), mediciones (2,760 msnm, 5 °C a 8 °C), tiempos (4:00 a.m., 15 minutos), tamaños de muestra (12 a 18 informantes).
* **Normativa y Documentos Oficiales:** Resoluciones (RM N° 424-2025/MINSA), líneas prioritarias (Línea 10, Prioridad 10), nombres de políticas o instituciones (MINSA, DIRESA Ayacucho, UNSCH).
* **Entidades y Topónimos Geográficos:** Centro de Salud Belén, Microred Huamanga, Barrio de Belén, río Alameda, cerro Acuchimay, distritos y provincias.
* **Términos Conceptuales Específicos o Étnicos:** Vocablos en Quechua Chanka (*nanay*, *onqoy*, *llakikuy*, *napaykuy*, *allin chaskiy*, *respetanakuy*, *upallay*) y constructos teóricos (*Lebenswelt*, *epojé*, *cuidado transpersonal*).
* **Citas y Atribuciones:** Apellidos de autores y años (Hernández-Sampieri, Donabedian, Watson, van Manen, Guba y Lincoln, Bunge) y referencias Vancouver.

### 5. Generación de la Matriz de Invarianza Semántica
El resultado del skill debe tabularse para ser consumido inmediatamente por el reescritor:

```markdown
### MATRIZ DE INVARIANZA SEMÁNTICA (Logical Extractor)

| Componente | Elemento Extraído del Texto Original | Regla de Invarianza para la Reescritura |
|---|---|---|
| **Tesis Central** | [Idea directriz fundamental] | Prohibido diluir o contradecir la conclusión |
| **Relaciones Causa-Efecto** | [Causa A] ➔ [Efecto B] | El vínculo de causalidad debe mantenerse inequívoco |
| **Datos Duros Inmutables** | [Cifras, %, fechas, unidades] | Prohibido redondear, cambiar o suprimir números |
| **Citas y Nombres Propios** | [Autores, leyes, entidades, lugares] | Mantener ortografía y citas Vancouver intactas |
| **Léxico Especializado / Quechua** | [Términos técnicos, vocabulario andino] | Prohibido sustituir por palabras genéricas en español |
```

---

## 📋 Reglas de Salida del Skill
1. **Precisión Quirúrgica:** No inventar premisas que no figuren en el texto de entrada.
2. **Neutralidad Analítica:** No calificar de bueno o malo el estilo; el rol del extractor es únicamente blindar la lógica y la evidencia.
3. **Paso Previo Obligatorio:** Ninguna labor de parafraseo debe iniciarse sin haber generado o verificado primero esta matriz de invariabilidad.
