---
name: human-paraphraser
description: Parafrasea y reescribe texto en español con naturalidad humana, alta variación rítmica (burstiness), perplejidad orgánica, voz activa y erradicación total de muletillas y clichés de IA. Basado en los patrones improve_writing y write_essay de Fabric.
---

# Human Paraphraser: Parafraseador y Reescritor de Prosa Humana

Inspirado en los patrones open-source `improve_writing` y `write_essay` del proyecto Fabric (Daniel Miessler) y adaptado a las exigencias estilísticas del español académico contemporáneo, este skill transforma textos rígidos, artificiales, repetitivos o con cadencia de máquina en **prosa humana fluida, elegante, contundente y vibrante**, preservando al 100% el rigor conceptual de la investigación.

---

## 🎯 Los Cuatro Pilares del Estilo Humano

### 1. Variación Rítmica Extrema (*Burstiness*)
El rasgo distintivo de la redacción humana es la imprevisibilidad rítmica. La IA suele escribir oraciones con la misma longitud monótona (18 a 22 palabras cada una), lo que produce un efecto anestésico. 
`human-paraphraser` exige deliberadamente la alternancia de tres registros:
* **Frases Cortas y Contundentes (4 a 9 palabras):** Para fijar tesis, dar golpes de efecto y definir conceptos con autoridad. *(Ejemplo: "La calidad en salud no es abstracta. El frío cala.")*
* **Frases Medianas y Equilibradas (10 a 18 palabras):** Para narrar acontecimientos, articular transiciones y conectar argumentos con claridad.
* **Frases Largas y Complejas (20 a 35 palabras):** Para sintetizar realidades multifactoriales, integrar dimensiones socioculturales o fundamentar citas teóricas densas con subordinación precisa.

### 2. Voz Activa y Energía Verbal Directa
* **Regla:** Desarticular las pasivas perifrásticas aburridas ("fue determinado por los investigadores", "se procedió a la realización de la toma") y los gerundios encadenados ("evidenciando", "demostrando", "concluyendo").
* **Transformación:** Colocar al sujeto al frente realizando una acción directa con verbos activos de fuerte carga semántica: *revelar, encarnar, fracturar, confrontar, articular, transformar, vivenciar*.

### 3. Perplejidad Orgánica y Precisión Léxica
* Huir del lenguaje artificialmente inflado o del barroquismo vacío.
* Usar la palabra justa en el momento oportuno. La elocuencia humana surge de la nitidez del pensamiento, no del exceso de adjetivos superlativos.

---

## 🚫 LISTA NEGRA: Prohibición Total de Muletillas y Fórmulas Típicas de IA

Bajo **NINGUNA CIRCUNSTANCIA** la salida debe contener las siguientes fórmulas o giros sintéticos típicos de modelos de lenguaje:

| Categoría de Cliché de IA | Expresiones Prohibidas (Blacklist) | Alternativa Humana Natural |
|---|---|---|
| **Aperturas de Relleno** | *"En el vertiginoso mundo de..."*, *"A lo largo de la historia..."*, *"Desde tiempos inmemoriales..."*, *"En el tapiz de..."* | Entrar directamente a los hechos y al problema concreto. |
| **Énfasis Burocrático Artificial** | *"Es crucial destacar..."*, *"Cabe resaltar que..."*, *"Es menester recalcar..."*, *"Juega un papel fundamental..."*, *"No se puede subestimar la importancia de..."* | Si algo es importante, se demuestra con datos y argumentos; no se anuncia pomposamente. |
| **Metáforas Trilladas** | *"Un faro de esperanza..."*, *"Desentrañar los misterios..."*, *"Un catalizador de cambio..."*, *"Un mosaico de complejidades..."*, *"El corazón palpitante..."* | Emplear analogías contextuales reales o descripciones empíricas directas. |
| **Conectores Mecánicos Forzados** | *"En este sentido..."*, *"Por un lado / Por otro lado"* (usado más de 1 vez), *"En última instancia..."*, *"A fin de cuentas..."*, *"Dicho esto..."* | Usar transiciones lógicas orgánicas (*sin embargo, no obstante, en contraste, paradójicamente*). |
| **Clausuras Telegráficas Obvias** | *"En conclusión..."*, *"En resumen..."*, *"Para finalizar..."*, *"En suma..."* | Dejar que el último párrafo cierre de forma natural y concluyente por la fuerza de sus ideas. |

---

## 🔒 Regla de Fidelidad a la Matriz de Invarianza Semántica
Antes de generar el texto reescrito, el reescritor debe verificar que ha consumido los elementos de `logical-extractor`:
1. **Datos Numéricos:** No alterar porcentajes (68.2%, 94.5%, 72%), altitudes (2,760 msnm), temperaturas (5 °C a 8 °C) ni tiempos de espera.
2. **Topónimos y Nombres:** Mantener inalterados el Centro de Salud Belén, Huamanga, Ayacucho y las entidades oficiales.
3. **Cosmovisión y Quechua Chanka:** Respetar intactos los términos originarios (*nanay*, *onqoy*, *llakikuy*, *napaykuy*, *allin chaskiy*, *respetanakuy*, *upallay*).
4. **Normas y Citas Vancouver:** Preservar la numeración y autoría de las referencias biomédicas (Donabedian, Watson, van Manen, Sampieri).
5. **Nexo Causal:** El resultado reescrito debe respetar la misma dirección causa $\to$ efecto identificada en el texto fuente.

---

## 📝 Formato de Salida Requerido
El skill debe entregar:
1. **Texto Parafraseado Final:** La versión mejorada en prosa española de alta calidad humana, lista para su publicación o inserción en la tesis.
2. **Métricas de Burstiness y Estilo:** Un breve desglose técnico que verifique la variedad de longitudes de oración (mínima, máxima y promedio) y confirme la ausencia de clichés de la lista negra.
