import urllib.request
import urllib.error
import ssl
import hashlib
import os
import sys
import json
import time

if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml,application/pdf,*/*;q=0.9",
    "Accept-Language": "es-ES,es;q=0.9,en;q=0.8"
}

ssl_ctx = ssl.create_default_context()
ssl_ctx.check_hostname = False
ssl_ctx.verify_mode = ssl.CERT_NONE

downloads = [
    {
        "id": "ANT-01",
        "category": "Antecedente de Investigación (Tesis Pregrado)",
        "author": "Velarde Quispedina, A.",
        "year": "2023",
        "title": "Asociación entre calidad de atención y satisfacción de los pacientes que acuden al servicio de salud",
        "institution": "Universidad María Auxiliadora (UMA)",
        "url": "https://repositorio.uma.edu.pe/bitstream/handle/20.500.12970/3207/TESIS-VELARDE%20QUISPEDINA.pdf",
        "dest_dir": "C:/GRESLY/DOCUMENTOS/ANTECEDENTES",
        "filename": "Velarde_Quispedina_2023_Calidad_Satisfaccion.pdf"
    },
    {
        "id": "ANT-02",
        "category": "Antecedente de Investigación (Trabajo Académico)",
        "author": "Ascama Calderon, M.",
        "year": "2022",
        "title": "Calidad de atención de enfermería en pacientes que acuden al servicio de emergencia",
        "institution": "Universidad María Auxiliadora (UMA)",
        "url": "https://repositorio.uma.edu.pe/bitstream/handle/20.500.12970/1828/TRABAJO%20ACADEMICO%20-ASCAMA%20CALDERON.pdf",
        "dest_dir": "C:/GRESLY/DOCUMENTOS/ANTECEDENTES",
        "filename": "Ascama_Calderon_2022_Enfermeria_Emergencia.pdf"
    },
    {
        "id": "ANT-03",
        "category": "Antecedente de Investigación (Tesis Maestría)",
        "author": "Toledo, C.",
        "year": "2018",
        "title": "Percepción de la calidad en la atención de los usuarios externos del centro de salud",
        "institution": "Universidad Nacional Autónoma de Nicaragua (UNAN)",
        "url": "https://repositorio.unan.edu.ni/id/eprint/7642/1/t753.pdf",
        "dest_dir": "C:/GRESLY/DOCUMENTOS/ANTECEDENTES",
        "filename": "Toledo_2018_Calidad_Usuarios_Externos.pdf"
    },
    {
        "id": "ANT-04",
        "category": "Antecedente de Investigación (Tesis Pregrado)",
        "author": "González, E.",
        "year": "2021",
        "title": "Tiempo de espera y percepción de la calidad de atención del usuario en el servicio de emergencia",
        "institution": "Universidad Nacional del Callao (UNAC)",
        "url": "https://repositorio.unac.edu.pe/backend/api/core/bitstreams/9b501d21-eee5-42b1-9707-d7a42dc4291c/content",
        "dest_dir": "C:/GRESLY/DOCUMENTOS/ANTECEDENTES",
        "filename": "Gonzalez_2021_Tiempo_Espera_Calidad.pdf"
    },
    {
        "id": "NORM-01",
        "category": "Normativa Técnica Sanitaria",
        "author": "Ministerio de Salud del Perú [MINSA]",
        "year": "2011 / 2019",
        "title": "Guía Técnica para la Evaluación de la Satisfacción del Usuario Externo en Establecimientos de Salud y Servicios Médicos de Apoyo (SERVQUAL Modificado)",
        "institution": "Ministerio de Salud del Perú",
        "url": "https://web.insnsb.gob.pe/wp-content/uploads/2019/08/RM-527-2011-MINSA-GUIA-TECNICA-SATISFACCION-USUARIO-EXTERNO.pdf",
        "dest_dir": "C:/GRESLY/DOCUMENTOS/REFERENCIAS_TEORICAS",
        "filename": "MINSA_RM_527_2011_Guia_Tecnica_Satisfaccion_Servqual.pdf"
    },
    {
        "id": "TEOR-04",
        "category": "Marco Internacional de Calidad",
        "author": "Organización Mundial de la Salud [OMS], OCDE, Banco Mundial",
        "year": "2018",
        "title": "Delivering quality health services: a global imperative for universal health coverage",
        "institution": "Organización Mundial de la Salud (OMS / WHO IRIS)",
        "url": "https://iris.who.int/bitstream/handle/10665/272465/9789241513906-eng.pdf",
        "dest_dir": "C:/GRESLY/DOCUMENTOS/REFERENCIAS_TEORICAS",
        "filename": "OMS_2018_Delivering_Quality_Health_Services.pdf"
    }
]

print(f"Descargando {len(downloads)} fuentes primarias y antecedentes en PDF...")

results = []

for item in downloads:
    os.makedirs(item["dest_dir"], exist_ok=True)
    target_path = os.path.join(item["dest_dir"], item["filename"])
    print(f"\nProcesando {item['id']} - {item['author']} ({item['year']}):")
    print(f"  URL: {item['url']}")
    
    t0 = time.time()
    req = urllib.request.Request(item["url"], headers=headers)
    success = False
    size = 0
    sha256 = ""
    error_msg = ""
    
    try:
        with urllib.request.urlopen(req, timeout=35, context=ssl_ctx) as resp:
            data = resp.read()
            size = len(data)
            with open(target_path, "wb") as f:
                f.write(data)
            sha256 = hashlib.sha256(data).hexdigest()
            elapsed = time.time() - t0
            print(f"  [OK] Descargado: {size/1024:.1f} KB en {elapsed:.2f}s -> {target_path}")
            success = True
    except Exception as e:
        error_msg = str(e)
        print(f"  [FAIL] Error en descarga: {e}")
        
    results.append({
        "id": item["id"],
        "author": item["author"],
        "year": item["year"],
        "title": item["title"],
        "institution": item["institution"],
        "url": item["url"],
        "target_path": target_path,
        "success": success,
        "size_bytes": size,
        "sha256": sha256,
        "error": error_msg
    })

meta_path = "C:/GRESLY/DOCUMENTOS/descargas_referencias_metadatos.json"
with open(meta_path, "w", encoding="utf-8") as f:
    json.dump(results, f, indent=2, ensure_ascii=False)

print(f"\nFinalizado. Resumen guardado en {meta_path}")
