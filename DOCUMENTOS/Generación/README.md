# Suite de Generación de Documentos LaTeX - UNSCH / GRESLY

Esta suite proporciona el entorno completo y estandarizado para la redacción, compilación y exportación a PDF de proyectos de investigación, tesis, artículos y documentación académica de la **Facultad de Ciencias de la Salud - Escuela Profesional de Enfermería (UNSCH)**.

---

## 📁 Estructura del Módulo

```text
DOCUMENTOS\Generación\
│
├── compilar.bat                 # Acceso directo para compilar con doble clic desde Windows
├── README.md                    # Esta guía de usuario
│
├── assets/
│   └── escudo_unsch.jpeg        # Escudo oficial de la UNSCH para carátulas
│
├── build/                       # Carpeta de salida para los PDFs y archivos auxiliares
│   ├── protocolo_cualitativo.pdf
│   └── articulo_cientifico.pdf
│
├── config/
│   └── preambulo_unsch.tex      # Configuración central (fuente Times 12pt, márgenes 3.5/3/2.5/2.5, interlineado 1.5)
│
├── templates/
│   ├── protocolo_cualitativo.tex # Plantilla oficial según la "Guía de Proyecto Cualitativo" (con carátula y anexos)
│   └── articulo_cientifico.tex   # Plantilla para artículos científicos resumidos
│
└── scripts/
    ├── compilar.ps1             # Compilador avanzado multi-pasada en PowerShell
    ├── compilar.bat             # Compilador por lotes
    ├── limpiar.ps1              # Elimina archivos auxiliares (.aux, .log, .toc, etc.)
    └── md_a_latex.ps1           # Convierte notas Markdown (.md) a PDF usando Pandoc + LaTeX
```

---

## 🚀 Guía Rápida de Uso

### 1. Compilación con Doble Clic (Forma más simple)
Haz doble clic sobre el archivo `compilar.bat` ubicado en esta carpeta. Compilará automáticamente la plantilla oficial `protocolo_cualitativo.tex` en dos pasadas y dejará el PDF listo en la carpeta `build\`.

---

### 2. Compilación desde PowerShell

Abre PowerShell en la carpeta `DOCUMENTOS\Generación\scripts`:

```powershell
# Compilación estándar (2 pasadas para resolver carátula, índices y tablas):
.\compilar.ps1

# Compilar y abrir el PDF inmediatamente en tu visor predeterminado:
.\compilar.ps1 -Open

# Compilar otro archivo específico:
.\compilar.ps1 -File "..\templates\articulo_cientifico.tex" -Open

# Compilar y limpiar los archivos temporales (.aux, .log, .toc):
.\compilar.ps1 -Clean
```

---

### 3. Convertir documentos Markdown de `DOCUMENTOS\MDs` a PDF

Para convertir cualquiera de los resúmenes o sílabos en Markdown a PDF con estilo LaTeX:

```powershell
# Modo interactivo (muestra la lista numerada de archivos disponibles):
.\md_a_latex.ps1

# Modo directo con un archivo específico:
.\md_a_latex.ps1 -InputFile "..\..\MDs\2 INV. CIENTIFICA.md" -Open
```

---

### 4. Limpieza de Archivos Temporales

Si deseas dejar limpia la carpeta `build\` eliminando los archivos intermedios generados por LaTeX:

```powershell
.\limpiar.ps1
```

---

## ⚙️ Normativa Académica Implementada en `config\preambulo_unsch.tex`

1. **Tipografía:** Times New Roman 12 pt (`mathptmx`).
2. **Márgenes oficiales:**
   - Superior: 3.0 cm
   - Izquierdo: 3.5 cm
   - Derecho: 2.5 cm
   - Inferior: 2.5 cm
3. **Interlineado:** 1.5 líneas (`\onehalfspacing`).
4. **Sangría:** 1.27 cm (estilo APA).
5. **Idioma:** Español con reglas tipográficas (`babel-spanish`).
6. **Encabezados:** Numeración superior derecha, estilo formal `fancyhdr`.
7. **Estructura:** Índice de contenidos, lista de tablas, lista de figuras, capítulos numerados, referencias y anexos formateados.
