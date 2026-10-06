# 🛠️ GUÍA DE INSTALACIÓN, CONFIGURACIÓN Y EJECUCIÓN DEL REPOSITORIO
**Proyecto:** *Calidad de Atención en Usuarios del Centro de Salud Belén, Ayacucho 2026*  
**Cátedra:** Proyecto de Investigación en Salud (EN 486) – Escuela Profesional de Enfermería (UNSCH)  
**Autora:** Bach. Gresly Lucero Pariona Palomino  
**Asesor:** Dr. Manglio Aguirre Andrade  

---

## 📋 Requisitos Previos del Sistema

Para clonar, editar, compilar el PDF en LaTeX y generar los documentos Word (.docx) automatizados, se requieren las siguientes herramientas:

| Herramienta | Versión Recomendada | Finalidad |
|---|---|---|
| **Git** | 2.40 o superior | Control de versiones y clonación del repositorio |
| **Python** | 3.10 a 3.13 (64-bit) | Ejecución del generador `generar_word.py` y verificación de paginación |
| **LaTeX (MiKTeX o TeX Live)** | 2023 o superior | Motor `pdflatex` para compilar el protocolo oficial en PDF |
| **PowerShell** | 5.1 o PowerShell 7 | Ejecución de scripts automatizados de compilación (`.ps1`) |
| **Microsoft Word** | Office 2016, 2019, 2021 o 365 | Edición de `EDITA.docx` y cálculo exacto de páginas mediante COM |
| **Editor de Código** | VS Code / Antigravity / Cursor | Edición de Markdown, LaTeX y reglas de inteligencia artificial |

---

## 🚀 Paso a Paso: Instalación y Puesta en Marcha

### Paso 1: Clonar el Repositorio de GitHub
Abre tu terminal (PowerShell o Git Bash) y clona el proyecto en tu máquina local:

```bash
git clone https://github.com/Eduardo-Sebastian-Paipay-Vega/GR----Ts----Enfermer-a.git
cd GR----Ts----Enfermer-a
```

*(Si ya te encuentras en la carpeta local `c:\GRESLY`, no necesitas volver a clonarlo).*

---

### Paso 2: Configurar el Entorno de Python e Instalar Dependencias
Se recomienda utilizar un entorno virtual para aislar las librerías necesarias:

#### En Windows (PowerShell / CMD):
```powershell
# 1. Crear el entorno virtual (opcional pero recomendado)
python -m venv .venv

# 2. Activar el entorno virtual
.venv\Scripts\Activate.ps1
# (Si usas CMD clásico: .venv\Scripts\activate.bat)

# 3. Instalar las dependencias oficiales
pip install -r requirements.txt
```

Las dependencias clave instaladas son:
* **`python-docx`**: Motor de ensamblado y formato para los documentos Word (.docx).
* **`PyMuPDF` (`fitz`)**: Extracción y análisis de texto de PDFs para auditoría de paginación.
* **`pywin32`**: Integración con Microsoft Word vía COM para renderizado y calibración física.

---

### Paso 3: Verificar el Motor LaTeX (`pdflatex`)
Asegúrate de que `pdflatex` esté disponible en las variables de entorno de tu sistema:

```powershell
pdflatex --version
```

> [!TIP]
> Si utilizas **MiKTeX**, asegúrate de habilitar en sus opciones: *“Always install missing packages on-the-fly”* (Instalar paquetes faltantes automáticamente) para que descargue sin pausas paquetes como `tabularx`, `booktabs` o `geometry`.

---

## ⚡ Guía de Ejecución y Compilación

### 1. Compilación del Documento PDF Oficial (LaTeX)
El protocolo maestro de tesis se encuentra modularizado en LaTeX dentro de `DOCUMENTOS/Generación/templates/subcapitulos/`. Para compilarlo a PDF con dos pasadas automáticas (para resolver índices y referencias cruzadas):

#### Desde PowerShell:
```powershell
powershell -ExecutionPolicy Bypass -File DOCUMENTOS\Generación\scripts\compilar.ps1
```

#### Desde la Línea de Comandos Clásica (CMD):
```cmd
DOCUMENTOS\Generación\scripts\compilar.bat
```

* **Salida generada:**  
  [`DOCUMENTOS/Generación/build/protocolo_cualitativo.pdf`](file:///c:/GRESLY/DOCUMENTOS/Generación/build/protocolo_cualitativo.pdf)  
  *(Documento PDF oficial de 32 páginas con carátula UNSCH, márgenes reglamentarios y referencias Vancouver).*

---

### 2. Generación Automatizada del Documento Word (.docx)
El repositorio cuenta con un generador Python (`generar_word.py`) que reconstruye el documento Word aplicando los estilos oficiales de la UNSCH (Times New Roman 12 pt, interlineado 1.5, sangría 1.0 cm, tablas con bordes gris tenue y encabezados institucionales):

```powershell
python DOCUMENTOS\Generación\scripts\generar_word.py
```

* **Salidas generadas:**  
  1. [`plantilla/AQUI/EDITA.docx`](file:///c:/GRESLY/plantilla/AQUI/EDITA.docx) *(Copia de trabajo editable para el usuario).*  
  2. [`DOCUMENTOS/Generación/protocolo_cualitativo.docx`](file:///c:/GRESLY/DOCUMENTOS/Generación/protocolo_cualitativo.docx) *(Copia maestra institucional).*

---

### 3. Verificación de Paginación y Encabezados
Para validar que los números de página del Índice General coincidan exactamente con la página física impresa de cada título en Word:

```powershell
python DOCUMENTOS\Generación\scripts\check_headers.py
```

---

### 4. Actualización del Grafo de Conocimiento (`graphify`)
Si dispones de la herramienta de grafos de conocimiento `graphify`, puedes sincronizar los nodos y relaciones tras modificar cualquier archivo:

```powershell
graphify update .
```

---

## 📂 Mapa de Carpetas y Rutas de Entrega

| Directorio / Archivo | Función y Contenido |
|---|---|
| [`README.md`](file:///c:/GRESLY/README.md) | Portada principal y resumen ejecutivo del repositorio |
| [`AGENTS.md`](file:///c:/GRESLY/AGENTS.md) | Reglamento y directrices obligatorias para la cátedra EN 486 |
| [`RUTA_COGNITIVA_IA.md`](file:///c:/GRESLY/RUTA_COGNITIVA_IA.md) | Guía de razonamiento metodológico para agentes y LLMs |
| [`DOCUMENTOS/MDs/01_EPISTEMOLOGIA_Y_NORMATIVA/`](file:///c:/GRESLY/DOCUMENTOS/MDs/01_EPISTEMOLOGIA_Y_NORMATIVA/) | Epistemología (Bunge), líneas MINSA 2030, DIRESA y sílabo |
| [`DOCUMENTOS/MDs/02_CANON_APOYO_DOCENTE_DR_MANGLIO/`](file:///c:/GRESLY/DOCUMENTOS/MDs/02_CANON_APOYO_DOCENTE_DR_MANGLIO/) | **Modelos de Oro:** Tesis aprobada (Ruiz) y artículos del Dr. Manglio |
| [`DOCUMENTOS/MDs/03_PROYECTO_ACTIVO_CS_BELEN/`](file:///c:/GRESLY/DOCUMENTOS/MDs/03_PROYECTO_ACTIVO_CS_BELEN/) | Protocolo Maestro Integral de Tesis de Gresly Pariona |
| [`DOCUMENTOS/MDs/04_GUIAS_Y_METRICAS/`](file:///c:/GRESLY/DOCUMENTOS/MDs/04_GUIAS_Y_METRICAS/) | Rúbricas Vancouver (0–20 pts) y purga de sesgos positivistas |
| [`DOCUMENTOS/Generación/templates/subcapitulos/`](file:///c:/GRESLY/DOCUMENTOS/Generación/templates/subcapitulos/) | Código fuente LaTeX modular (`sec_00_*.tex` a `sec_07_*.tex`) |
| [`plantilla/AQUI/EDITA.docx`](file:///c:/GRESLY/plantilla/AQUI/EDITA.docx) | **Archivo Word editable principal listo para imprimir o enviar** |
| [`DOCUMENTOS/Generación/build/protocolo_cualitativo.pdf`](file:///c:/GRESLY/DOCUMENTOS/Generación/build/protocolo_cualitativo.pdf) | **Archivo PDF compilado oficial listo para sustentar** |

---

## ❓ Preguntas Frecuentes y Solución de Problemas (Troubleshooting)

### Error 1: `Execution of scripts is disabled on this system` en PowerShell
**Causa:** Política restrictiva de ejecución de scripts de Windows.  
**Solución:** Ejecuta el script añadiendo el parámetro `-ExecutionPolicy Bypass`:
```powershell
powershell -ExecutionPolicy Bypass -File .\DOCUMENTOS\Generación\scripts\compilar.ps1
```

### Error 2: `Permission denied: 'plantilla/AQUI/EDITA.docx'` al ejecutar `generar_word.py`
**Causa:** El archivo Word está abierto en Microsoft Word u otro programa editor.  
**Solución:** Cierra Microsoft Word y vuelve a ejecutar el comando en la terminal.

### Error 3: Paquetes faltantes en LaTeX (`! LaTeX Error: File 'tabularx.sty' not found`)
**Causa:** Instalación básica de LaTeX sin paquetes adicionales.  
**Solución:** 
* En **MiKTeX**: Abre la consola de MiKTeX y ve a *Settings ➔ You can choose whether missing packages are to be installed automatically ➔ Yes*.
* En **TeX Live / Linux**: Ejecuta `sudo tlmgr install tabularx booktabs geometry titlesec caption fancyhdr microtype`.

---

## 🤝 Soporte y Asesoría
* **Investigadora Principal:** Bach. Gresly Lucero Pariona Palomino *(gresly.pariona@unsch.edu.pe)*
* **Asesor Metodológico:** Dr. Manglio Aguirre Andrade *(Facultad de Ciencias de la Salud – UNSCH)*
