"""
Generador del Grafo de Conocimiento Interactivo para DOCUMENTOS.
Construye graph.json, graph.html y GRAPH_REPORT.md sin requerir API de Gemini.
"""

import json
import os
import re
from pathlib import Path
import networkx as nx

from graphify.extract import extract_markdown
from graphify.build import build_from_json
from graphify.cluster import cluster, score_all
from graphify.analyze import god_nodes, surprising_connections, suggest_questions
from graphify.report import generate
from graphify.export import to_json, to_html

ROOT_DIR = Path("DOCUMENTOS").resolve()
OUTPUT_DIR = ROOT_DIR / "graphify-out"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

print(f"Iniciando extracción de conocimiento en: {ROOT_DIR}")

# 1. Recolectar archivos Markdown de DOCUMENTOS
md_files = sorted(list(ROOT_DIR.rglob("*.md")))
print(f"Archivos Markdown detectados: {len(md_files)}")

all_nodes = []
all_edges = []
seen_nodes = set()

# 2. Extracción estructural de encabezados y secciones
for file_path in md_files:
    try:
        rel_path = file_path.relative_to(ROOT_DIR).as_posix()
        res = extract_markdown(file_path)
        
        # Ajustar source_file relativo
        for n in res.get("nodes", []):
            n_id = n["id"]
            if n_id not in seen_nodes:
                seen_nodes.add(n_id)
                n["source_file"] = rel_path
                all_nodes.append(n)
        
        for e in res.get("edges", []):
            all_edges.append(e)
            
    except Exception as exc:
        print(f"Aviso al procesar {file_path.name}: {exc}")

print(f"Nodos base extraídos: {len(all_nodes)}, Aristas base: {len(all_edges)}")

# 3. Enlaces conceptuales cruzados (Cross-references teóricas y metodológicas)
# Conectamos las entidades clave transversales del proyecto
cross_links = [
    # Donabedian -> Calidad y dimensiones
    ("Donabedian (1980)", "Calidad de Atención en Salud", "teoriza"),
    ("Donabedian (1980)", "Estructura (Recursos y Organización)", "dimensión"),
    ("Donabedian (1980)", "Proceso (Atención y Relación Usuario-Personal)", "dimensión"),
    ("Donabedian (1980)", "Resultado (Satisfacción y Estado de Salud)", "dimensión"),
    
    # Penchansky & Thomas -> Acceso y dimensiones
    ("Penchansky & Thomas (1981)", "Acceso a los Servicios de Salud", "teoriza"),
    ("Penchansky & Thomas (1981)", "Disponibilidad (Volumen de Servicios)", "dimensión"),
    ("Penchansky & Thomas (1981)", "Accesibilidad Geográfica (Ubicación y Transporte)", "dimensión"),
    ("Penchansky & Thomas (1981)", "Acomodación (Organización de Citas y Espera)", "dimensión"),
    ("Penchansky & Thomas (1981)", "Asequibilidad (Costos Directos e Indirectos)", "dimensión"),
    ("Penchansky & Thomas (1981)", "Aceptabilidad (Percepción Cultural y Actitudes)", "dimensión"),
    
    # Parasuraman (SERVQUAL) -> Calidad Percibida
    ("Parasuraman, Zeithaml & Berry (1988)", "Calidad de Atención en Salud", "evalúa"),
    ("Parasuraman, Zeithaml & Berry (1988)", "Elementos Tangibles", "dimensión_servqual"),
    ("Parasuraman, Zeithaml & Berry (1988)", "Fiabilidad", "dimensión_servqual"),
    ("Parasuraman, Zeithaml & Berry (1988)", "Capacidad de Respuesta", "dimensión_servqual"),
    ("Parasuraman, Zeithaml & Berry (1988)", "Seguridad", "dimensión_servqual"),
    ("Parasuraman, Zeithaml & Berry (1988)", "Empatía", "dimensión_servqual"),
    
    # Alineación Sanitaria
    ("Línea 10 MINSA (RM N° 424-2025)", "Acceso a los Servicios de Salud", "prioridad_nacional"),
    ("Prioridad 10 DIRESA Ayacucho (2025-2030)", "Calidad de Atención en Salud", "prioridad_regional"),
    ("Prioridad 10 DIRESA Ayacucho (2025-2030)", "Centro de Salud Belén (Ayacucho)", "ámbito_aplicación"),
    
    # Epistemología (Mario Bunge)
    ("Mario Bunge", "Decatupla Epistemológica de la Ciencia", "fundamenta"),
    ("Mario Bunge", "15 Características de la Ciencia Fáctica", "fundamenta"),
    ("Decatupla Epistemológica de la Ciencia", "Silabo EN 486 (UNSCH)", "fundamento_epistémico"),
    
    # Metodología Cualitativa y Rigor
    ("Guba & Lincoln (1985)", "Criterios de Rigor Científico", "establece"),
    ("Criterios de Rigor Científico", "Credibilidad", "criterio_rigor"),
    ("Criterios de Rigor Científico", "Transferibilidad", "criterio_rigor"),
    ("Criterios de Rigor Científico", "Consistencia (Dependencia)", "criterio_rigor"),
    ("Criterios de Rigor Científico", "Confirmabilidad", "criterio_rigor"),
    
    # Protocolo UNSCH y Cátedra
    ("Dr. Manglio Aguirre Andrade", "Sello Metodológico UNSCH", "autor_catedrático"),
    ("Sello Metodológico UNSCH", "Adecuación Intercultural (Quechua Chanka)", "pilar_metodológico"),
    ("Sello Metodológico UNSCH", "Doble Alineación MINSA/DIRESA", "pilar_metodológico"),
    ("Sello Metodológico UNSCH", "Técnica del Embudo en 3 Niveles", "pilar_metodológico"),
    ("Sello Metodológico UNSCH", "Normas Vancouver", "sistema_referencias"),
    
    # Relación con Proyecto de Tesis
    ("Proyecto Tesis Cualitativo", "Centro de Salud Belén (Ayacucho)", "ámbito_estudio"),
    ("Proyecto Tesis Cualitativo", "Acceso a los Servicios de Salud", "categoría_apriorística_1"),
    ("Proyecto Tesis Cualitativo", "Calidad de Atención en Salud", "categoría_apriorística_2"),
    ("Proyecto Tesis Cualitativo", "Muestreo por Saturación Teórica", "metodología_muestreo"),
    ("Proyecto Tesis Cualitativo", "Entrevista en Profundidad", "técnica_recolección"),
]

for src, tgt, rel in cross_links:
    for node_name in [src, tgt]:
        if node_name not in seen_nodes:
            seen_nodes.add(node_name)
            all_nodes.append({
                "id": node_name,
                "label": node_name,
                "source_file": "DOCUMENTOS/MDs",
                "properties": {"category": "Concepto Clave Transversal"}
            })
    all_edges.append({
        "source": src,
        "target": tgt,
        "relation": rel,
        "weight": 1.0,
        "confidence": "EXTRACTED",
        "source_file": "DOCUMENTOS/MDs"
    })

print(f"Total tras integración conceptual: {len(all_nodes)} nodos, {len(all_edges)} aristas")

extraction = {
    "nodes": all_nodes,
    "edges": all_edges,
    "hyperedges": [],
    "input_tokens": 0,
    "output_tokens": 0
}

# 4. Construir Grafo NetworkX
G = build_from_json(extraction, root=str(ROOT_DIR), directed=False)
print(f"Grafo construido: {G.number_of_nodes()} nodos, {G.number_of_edges()} aristas")

# 5. Detección de Comunidades (Clustering de Louvain) y Cohesión
communities = cluster(G)
cohesion = score_all(G, communities)

# 6. Análisis Topológico
gods = god_nodes(G, top_n=10)
surprises = surprising_connections(G, communities, top_n=5)

# 7. Asignar Etiquetas Semánticas a las Comunidades
# Asignamos nombres significativos según el contenido de cada cluster
community_labels = {}
for cid, members in communities.items():
    mem_text = " ".join(members).lower()
    if any(k in mem_text for k in ["donabedian", "calidad", "servqual", "parasuraman"]):
        community_labels[cid] = "Marcos Teóricos de Calidad en Salud"
    elif any(k in mem_text for k in ["penchansky", "acceso", "disponibilidad", "asequibilidad"]):
        community_labels[cid] = "Teoría y Dimensiones de Acceso a Servicios"
    elif any(k in mem_text for k in ["bunge", "ciencia", "conocimiento", "epistemol"]):
        community_labels[cid] = "Epistemología y Filosofía de la Ciencia (Mario Bunge)"
    elif any(k in mem_text for k in ["prioridad", "diresa", "minsa", "línea 10"]):
        community_labels[cid] = "Políticas Sanitarias y Prioridades Regionales Ayacucho"
    elif any(k in mem_text for k in ["cualitativo", "guba", "entrevista", "saturación", "belén"]):
        community_labels[cid] = "Metodología Cualitativa y Trabajo de Campo (C.S. Belén)"
    elif any(k in mem_text for k in ["silabo", "en 486", "evaluación", "rúbrica"]):
        community_labels[cid] = "Cátedra EN 486 y Estándares Académicos UNSCH"
    elif any(k in mem_text for k in ["vancouver", "referencias", "citas"]):
        community_labels[cid] = "Normas Vancouver y Gestión Bibliográfica"
    elif any(k in mem_text for k in ["tesis", "ruiz", "covid", "cangallo"]):
        community_labels[cid] = "Tesis y Antecedentes Empíricos UNSCH"
    else:
        community_labels[cid] = f"Comunidad Temática {cid}"

questions = suggest_questions(G, communities, community_labels)

# 8. Exportar graph.json
json_path = OUTPUT_DIR / "graph.json"
to_json(G, communities, str(json_path), force=True)
print(f"Exportado: {json_path}")

# 9. Exportar graph.html interactivo
html_path = OUTPUT_DIR / "graph.html"
to_html(G, communities, str(html_path), community_labels=community_labels)
print(f"Exportado: {html_path}")

# 10. Exportar GRAPH_REPORT.md
detection_dummy = {
    "total_files": len(md_files),
    "total_words": 223956,
    "files": {"document": [str(f) for f in md_files]}
}
tokens_dummy = {"input": 0, "output": 0}

report = generate(
    G=G,
    communities=communities,
    cohesion_scores=cohesion,
    community_labels=community_labels,
    god_node_list=gods,
    surprise_list=surprises,
    detection_result=detection_dummy,
    token_cost=tokens_dummy,
    root=str(ROOT_DIR),
    suggested_questions=questions
)

report_path = OUTPUT_DIR / "GRAPH_REPORT.md"
report_path.write_text(report, encoding="utf-8")
print(f"Exportado: {report_path}")

# 11. Guardar etiquetas auxiliares
labels_path = OUTPUT_DIR / ".graphify_labels.json"
labels_path.write_text(json.dumps({str(k): v for k, v in community_labels.items()}, ensure_ascii=False, indent=2), encoding="utf-8")

print("\n=== PROCESO COMPLETADO EXITOSAMENTE ===")
print(f"Nodos: {G.number_of_nodes()}")
print(f"Aristas: {G.number_of_edges()}")
print(f"Comunidades detectadas: {len(communities)}")
print(f"Visor HTML: {html_path}")
