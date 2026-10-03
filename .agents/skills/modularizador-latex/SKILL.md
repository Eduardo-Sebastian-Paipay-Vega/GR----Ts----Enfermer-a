---
name: modularizador-latex
description: Lee capítulos extensos en LaTeX, los divide en subcapítulos modulares (.tex) y expande su contenido manteniendo la integridad de las referencias y la compilación.
---

# Instrucciones del Agente Especialista en LaTeX

Cuando el usuario solicite dividir o detallar un capítulo en LaTeX, ejecuta este flujo de trabajo en orden estricto:

## Paso 1: Lectura y Análisis
1. Lee el archivo `.tex` principal indicado por el usuario.
2. Identifica la estructura jerárquica actual (`\chapter`, `\section`, `\subsection`).
3. Extrae y registra en memoria todas las etiquetas (`\label{}`), referencias (`\ref{}`, `\autoref{}`) y citas bibliográficas (`\cite{}`) para asegurar que ninguna se pierda durante la división.

## Paso 2: Creación de la Estructura Modular
1. Crea una subcarpeta para alojar los subcapítulos (por ejemplo, `subcapitulos/`).
2. Divide el contenido del capítulo original en archivos independientes por cada `\section` (ejemplo: `sec_01_introduccion.tex`, `sec_02_metodologia.tex`).
3. Reemplaza el contenido en el archivo original del capítulo por comandos `\input{subcapitulos/nombre_archivo.tex}` (evita usar `\include{}` para no forzar saltos de página innecesarios).

## Paso 3: Expansión y Detalle Controlado
Al agregar información o desarrollar un subcapítulo específico:
1. **Regla de fragmento puro:** Los archivos de subcapítulos NUNCA deben contener `\documentclass`, `\usepackage`, `\begin{document}` ni `\end{document}`. Deben iniciar directamente con `\section{...}`.
2. **Coherencia de etiquetas:** Conserva las etiquetas originales. Si agregas nuevas ecuaciones, tablas o figuras, utiliza un prefijo claro basado en la sección (por ejemplo, `\label{eq:sec1_balance}`, `\label{tab:sec2_datos}`).
3. **Límites de alcance:** Desarrolla únicamente el tema asignado a ese archivo específico sin repetir definiciones que pertenezcan a secciones anteriores.

## Paso 4: Verificación
1. Confirma que todas las llaves `{}` y entornos (`\begin{...}` / `\end{...}`) estén correctamente cerrados en cada archivo creado.
2. Entrega un resumen de los archivos generados y las etiquetas actualizadas.
