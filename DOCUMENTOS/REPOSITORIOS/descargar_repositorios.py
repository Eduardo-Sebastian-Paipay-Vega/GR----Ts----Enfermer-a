import urllib.request
import urllib.error
import ssl
import hashlib
import os
import re
import json
import time
import sys

# Ensure UTF-8 output for Windows console
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass


repos = [
    # Repositorios Nacionales (Perú)
    {
        "id": "REP-PE-01",
        "slug": "01_alicia_concytec",
        "name": "ALICIA (CONCYTEC)",
        "category": "Repositorios Nacionales (Perú)",
        "url": "https://alicia.concytec.gob.pe",
        "description": "El agregador nacional que unifica todas las tesis y artículos de universidades e institutos de investigación del país.",
        "enfoque_salud": "Búsqueda integrada de producción científica peruana, tesis de pregrado y posgrado en salud, financiamiento e investigadores RENACYT."
    },
    {
        "id": "REP-PE-02",
        "slug": "02_renati_sunedu",
        "name": "RENATI (SUNEDU)",
        "category": "Repositorios Nacionales (Perú)",
        "url": "https://renati.sunedu.gob.pe",
        "description": "Registro Nacional de Trabajos de Investigación, indispensable para buscar tesis de pregrado y posgrado aprobadas.",
        "enfoque_salud": "Verificación oficial de grados, títulos y antecedentes de investigación aprobados a nivel universitario nacional."
    },
    {
        "id": "REP-PE-03",
        "slug": "03_repositorio_minsa",
        "name": "Repositorio Institucional MINSA",
        "category": "Repositorios Nacionales (Perú)",
        "url": "https://repositorio.minsa.gob.pe",
        "description": "Base de datos oficial del Ministerio de Salud con documentos técnicos, investigaciones y reportes sobre políticas de acceso en Perú.",
        "enfoque_salud": "Normas técnicas de salud (NTS), guías de práctica clínica (GPC), planes de contingencia, directivas sanitarias y estadísticas nacionales."
    },
    {
        "id": "REP-PE-04",
        "slug": "04_repositorio_ins",
        "name": "Repositorio INS (Instituto Nacional de Salud)",
        "category": "Repositorios Nacionales (Perú)",
        "url": "https://repositorio.ins.gob.pe",
        "description": "Centrado en investigaciones de salud pública y epidemiología a nivel nacional.",
        "enfoque_salud": "Vigilancia epidemiológica, salud intercultural, nutrición (CENAN), salud ocupacional (CENSOPAS), biomedicina y control de brotes."
    },
    {
        "id": "REP-PE-05",
        "slug": "05_cybertesis_unmsm",
        "name": "Cybertesis UNMSM",
        "category": "Repositorios Nacionales (Perú)",
        "url": "https://cybertesis.unmsm.edu.pe",
        "description": "Repositorio de la Universidad Nacional Mayor de San Marcos, con un alto volumen histórico en su Facultad de Medicina.",
        "enfoque_salud": "Tesis históricas y de vanguardia de Medicina San Fernando, Enfermería, Farmacia y Bioquímica, y Obstetricia."
    },
    {
        "id": "REP-PE-06",
        "slug": "06_repositorio_upch",
        "name": "Repositorio Institucional UPCH",
        "category": "Repositorios Nacionales (Perú)",
        "url": "https://repositorio.upch.edu.pe",
        "description": "Universidad Peruana Cayetano Heredia, principal referente académico privado en ciencias de la salud.",
        "enfoque_salud": "Ensayos clínicos, salud pública, epidemiología clínica, enfermería y medicina tropical de excelencia internacional."
    },
    {
        "id": "REP-PE-07",
        "slug": "07_repositorio_unsch",
        "name": "Repositorio Institucional UNSCH",
        "category": "Repositorios Nacionales (Perú)",
        "url": "https://repositorio.unsch.edu.pe",
        "description": "Universidad Nacional de San Cristóbal de Huamanga, clave para estudios de acceso a la salud en Ayacucho y zonas andinas.",
        "enfoque_salud": "Tesis locales de Enfermería, Obstetricia, Medicina y Biología aplicadas a la realidad de las comunidades de Ayacucho y Huamanga."
    },
    {
        "id": "REP-PE-08",
        "slug": "08_repositorio_ucv",
        "name": "Repositorio UCV",
        "category": "Repositorios Nacionales (Perú)",
        "url": "https://repositorio.ucv.edu.pe",
        "description": "La Universidad César Vallejo concentra la mayor cantidad de tesis sobre gestión pública y medición de calidad percibida en hospitales.",
        "enfoque_salud": "Modelos Servqual, gestión de la calidad hospitalaria, clima laboral, satisfacción del usuario externo y gerencia en salud."
    },
    {
        "id": "REP-PE-09",
        "slug": "09_repositorio_usmp",
        "name": "Repositorio USMP",
        "category": "Repositorios Nacionales (Perú)",
        "url": "https://repositorio.usmp.edu.pe",
        "description": "Destacado por su alta producción de tesis de maestría en gerencia y administración de servicios de salud.",
        "enfoque_salud": "Administración hospitalaria, auditoría médica, calidad del cuidado, gestión directiva e innovación en centros de salud."
    },
    {
        "id": "REP-PE-10",
        "slug": "10_tesis_pucp",
        "name": "Repositorio de Tesis PUCP",
        "category": "Repositorios Nacionales (Perú)",
        "url": "https://tesis.pucp.edu.pe",
        "description": "Excelente fuente para estudios con enfoques cualitativos, de políticas públicas y derechos ciudadanos en salud.",
        "enfoque_salud": "Investigación cualitativa (fenomenología, hermenéutica, etnografía), bioética, determinantes sociales, derechos humanos y políticas sanitarias."
    },

    # Repositorios Internacionales
    {
        "id": "REP-INT-01",
        "slug": "11_scielo",
        "name": "SciELO (Scientific Electronic Library Online)",
        "category": "Repositorios Internacionales",
        "url": "https://scielo.org",
        "description": "La hemeroteca científica de acceso abierto más grande de Iberoamérica.",
        "enfoque_salud": "Revistas indexadas de salud pública, enfermería, epidemiología y ciencias médicas en español y portugués a texto completo."
    },
    {
        "id": "REP-INT-02",
        "slug": "12_bvs_salud",
        "name": "BVS (Biblioteca Virtual en Salud)",
        "category": "Repositorios Internacionales",
        "url": "https://bvsalud.org",
        "description": "Red coordinada por BIREME/OPS, especializada exclusivamente en literatura de ciencias de la salud de América Latina y el Caribe.",
        "enfoque_salud": "LILACS, DeCS (Descriptores en Ciencias de la Salud), bases de datos OPS/OMS, guías clínicas e intervenciones de enfermería basadas en evidencia."
    },
    {
        "id": "REP-INT-03",
        "slug": "13_lareferencia",
        "name": "LA Referencia",
        "category": "Repositorios Internacionales",
        "url": "https://www.lareferencia.info",
        "description": "Red que agrupa repositorios nacionales de acceso abierto de toda América Latina, ideal para comparar tesis entre diferentes países.",
        "enfoque_salud": "Estudios comparativos regionales, metadatos abiertos de producción académica y científica latinoamericana."
    },
    {
        "id": "REP-INT-04",
        "slug": "14_iris_paho",
        "name": "IRIS (OPS)",
        "category": "Repositorios Internacionales",
        "url": "https://iris.paho.org",
        "description": "Repositorio Institucional para Compartir Información de la Organización Panamericana de la Salud. Contiene directrices y evaluaciones regionales de servicios médicos.",
        "enfoque_salud": "Informes epidemiológicos de las Américas, directrices de salud comunitaria, guías de atención primaria y metas sanitarias regionales."
    },
    {
        "id": "REP-INT-05",
        "slug": "15_pubmed_central",
        "name": "PubMed Central (PMC)",
        "category": "Repositorios Internacionales",
        "url": "https://www.ncbi.nlm.nih.gov/pmc",
        "description": "Archivo gratuito de la Biblioteca Nacional de Medicina de EE. UU. (Recomendado buscar como \"health care quality access\" o \"patient satisfaction\").",
        "enfoque_salud": "Estándar de oro mundial biomédico: artículos revisados por pares a texto completo, metanálisis, ensayos controlados y estudios cualitativos de rigor."
    },
    {
        "id": "REP-INT-06",
        "slug": "16_redalyc",
        "name": "Redalyc",
        "category": "Repositorios Internacionales",
        "url": "https://www.redalyc.org",
        "description": "Sistema de información que alberga miles de revistas médicas y de ciencias sociales de América Latina, España y Portugal.",
        "enfoque_salud": "Acceso abierto diamante, investigación en salud holística, ciencias de la conducta, sociología médica y enfermería comunitaria."
    },
    {
        "id": "REP-INT-07",
        "slug": "17_dialnet",
        "name": "Dialnet",
        "category": "Repositorios Internacionales",
        "url": "https://dialnet.unirioja.es",
        "description": "El principal portal bibliográfico de literatura científica en español, imprescindible para tesis doctorales y artículos académicos de España y LatAm.",
        "enfoque_salud": "Monografías, actas de congresos, tesis doctorales y revistas de humanidades médicas, enfermería y bioética en español."
    },
    {
        "id": "REP-INT-08",
        "slug": "18_teseo",
        "name": "TESEO",
        "category": "Repositorios Internacionales",
        "url": "https://www.educacion.gob.es/teseo",
        "description": "Base de datos centralizada de las tesis doctorales sustentadas en las universidades de España.",
        "enfoque_salud": "Referencia metodológica de alto nivel en doctorados de enfermería, salud pública, epidemiología y gestión asistencial española."
    },
    {
        "id": "REP-INT-09",
        "slug": "19_who_iris",
        "name": "WHO IRIS",
        "category": "Repositorios Internacionales",
        "url": "https://iris.who.int",
        "description": "Repositorio global de la Organización Mundial de la Salud, útil para estudios sobre determinantes macroeconómicos y sociales del acceso.",
        "enfoque_salud": "Directrices globales de la OMS, CIE-10/CIE-11, reportes mundiales de salud, estadísticas sanitarias globales y recomendaciones de salud pública."
    },
    {
        "id": "REP-INT-10",
        "slug": "20_openaire",
        "name": "OpenAIRE",
        "category": "Repositorios Internacionales",
        "url": "https://explore.openaire.eu",
        "description": "Portal europeo que federa miles de repositorios y revistas de acceso abierto a nivel mundial.",
        "enfoque_salud": "Ciencia abierta a nivel global, datasets de investigación médica vinculados a proyectos europeos e internacionales."
    }
]

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "es-ES,es;q=0.9,en;q=0.8"
}

output_dir = os.path.join(os.path.dirname(__file__), "descargas")
os.makedirs(output_dir, exist_ok=True)

# Context with lenient SSL verification for public government portals if chain errors occur
ssl_ctx = ssl.create_default_context()
ssl_ctx.check_hostname = False
ssl_ctx.verify_mode = ssl.CERT_NONE

results = []

print(f"Iniciando descarga y verificación de los {len(repos)} repositorios de información...")

for idx, r in enumerate(repos, 1):
    print(f"[{idx}/{len(repos)}] Procesando {r['name']} ({r['url']})...")
    req = urllib.request.Request(r["url"], headers=headers)
    t0 = time.time()
    status_code = None
    title = ""
    file_size = 0
    sha256 = ""
    filename = f"{r['slug']}.html"
    filepath = os.path.join(output_dir, filename)
    error_msg = ""
    
    try:
        with urllib.request.urlopen(req, timeout=20, context=ssl_ctx) as response:
            status_code = response.getcode()
            content = response.read()
            elapsed = time.time() - t0
            file_size = len(content)
            
            # Save downloaded HTML
            with open(filepath, "wb") as f:
                f.write(content)
            
            # Compute SHA256
            sha256 = hashlib.sha256(content).hexdigest()
            
            # Extract title if possible
            try:
                decoded = content.decode('utf-8', errors='ignore')
                match = re.search(r'<title>(.*?)</title>', decoded, re.IGNORECASE | re.DOTALL)
                if match:
                    title = " ".join(match.group(1).strip().split())
            except Exception as e:
                title = "(Error extrayendo título)"
                
            print(f"   [OK] Descargado ({status_code}) - {file_size/1024:.1f} KB en {elapsed:.2f}s | Titulo: {title[:50]}...")
    except urllib.error.HTTPError as e:
        status_code = e.code
        error_msg = f"HTTP Error {e.code}: {e.reason}"
        print(f"   [WARN] HTTP Error: {error_msg}")
        # Try to read body even on error
        try:
            content = e.read()
            file_size = len(content)
            with open(filepath, "wb") as f:
                f.write(content)
            sha256 = hashlib.sha256(content).hexdigest()
        except:
            pass
    except urllib.error.URLError as e:
        error_msg = f"URL Error: {e.reason}"
        print(f"   [FAIL] URL Error: {error_msg}")
    except Exception as e:
        error_msg = f"Error: {str(e)}"
        print(f"   [FAIL] General Error: {error_msg}")
        
    result_data = {
        "id": r["id"],
        "name": r["name"],
        "category": r["category"],
        "url": r["url"],
        "description": r["description"],
        "enfoque_salud": r["enfoque_salud"],
        "filename": filename,
        "local_path": filepath,
        "status_code": status_code,
        "title": title,
        "file_size_bytes": file_size,
        "sha256": sha256,
        "error": error_msg
    }
    results.append(result_data)

# Save JSON metadata
json_path = os.path.join(output_dir, "repositorios_metadatos.json")
with open(json_path, "w", encoding="utf-8") as jf:
    json.dump(results, jf, indent=2, ensure_ascii=False)

print(f"\nProceso finalizado. Metadatos guardados en {json_path}")
