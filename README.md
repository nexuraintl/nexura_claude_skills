# nexura_claude_skills

Skills de Claude para el proyecto de **Gobernanza GCP** de NEXURA: verificación de cumplimiento, corrección de repositorios y generación de documentación para microservicios FastAPI sobre Cloud Run.

## Skills incluidas

| Skill | Dónde se usa | Qué hace | Se activa con |
|---|---|---|---|
| [`gobernanza-gcp`](./gobernanza-gcp/SKILL.md) | Claude.ai (Proyectos) | Verifica el cumplimiento de un microservicio contra los estándares (GOB-GCP-STD-01, GOB-GCP-GOB-01, GOB-GCP-GOB-04) y, si aprueba, genera el MANUAL del servicio, la ficha de incorporación al gateway y/o el RFC de PROD | "Documenta el microservicio [nombre]", "verifica si cumple con gobernanza", "genera el manual del servicio" |
| [`gobernanza-gcp-code`](./gobernanza-gcp-code/SKILL.md) | Claude Code | Analiza un repositorio de microservicio (nuevo o existente) contra el estándar GOB-GCP-STD-01 y aplica correcciones en el código, con confirmación del desarrollador antes de tocar nada | "Revisa este repo", "verifica cumplimiento", "aplica el estándar", "migra este servicio" |
| [`docx-nexura`](./docx-nexura/SKILL.md) | Claude.ai (Proyectos) | Genera los `.docx` con el formato corporativo oficial de NEXURA (header, footer, TOC, estilos). Es la skill que usa `gobernanza-gcp` para producir los documentos finales — no se invoca sola normalmente | "Genera el documento", "crea el .docx", "escribe el estándar" |

### Cómo se relacionan

```
gobernanza-gcp (Claude.ai)          gobernanza-gcp-code (Claude Code)
  Fase 1: verifica cumplimiento       Fase 1: analiza el repo
  Fase 2: documenta ──uses──▶ docx-nexura   Fase 2: corrige el código
```

- **`gobernanza-gcp`** y **`docx-nexura`** trabajan juntas en Claude.ai: la primera verifica y decide qué documentos generar, la segunda los redacta con el formato NEXURA.
- **`gobernanza-gcp-code`** es independiente y corre en Claude Code, directamente sobre el repositorio del microservicio (no genera `.docx`, corrige código).

## Cómo usarlas

### En Claude.ai (`gobernanza-gcp`, `docx-nexura`)

1. Crear (o abrir) un Proyecto de Claude.ai para Gobernanza GCP.
2. En la configuración del proyecto, agregar ambas skills (`gobernanza-gcp` y `docx-nexura`) desde este repositorio.
3. Subir como archivos del proyecto los archivos clave del microservicio a documentar (ver la sección "Qué leer del microservicio" en [`gobernanza-gcp/SKILL.md`](./gobernanza-gcp/SKILL.md)): `api/main.py`, `api/core/config.py`, `openapi/openapi.yaml`, `docs/MANUAL.md`, `.env.example`, etc.
4. Pedir: *"Documenta el microservicio [nombre]"* o *"Verifica y documenta [nombre]"*.
5. Claude ejecuta primero la verificación de cumplimiento y presenta un reporte. Si el servicio está apto, con tu confirmación genera los documentos usando `docx-nexura`.

Si no tenés los archivos a mano, Claude te va a pedir los datos mínimos (nombre del servicio, proyecto GCP, SAs, variables de entorno, endpoints, base de datos).

### En Claude Code (`gobernanza-gcp-code`)

1. Copiar la carpeta `gobernanza-gcp-code/` a `.claude/skills/` del repositorio del microservicio (o al directorio de skills que uses en Claude Code).
2. Pararse en la **raíz del repositorio** del microservicio (la skill asume que corre desde ahí).
3. Pedir: *"Revisa este repo"*, *"verifica cumplimiento"*, *"aplica el estándar"* o *"migra este servicio"*.
4. Claude ejecuta la Fase 1 (análisis estructural y de contenido) y presenta el reporte de hallazgos.
5. Solo después de tu confirmación aplica las correcciones (Fase 2). Nunca corrige sin mostrar el reporte primero.

Aplica tanto a repos nuevos (creados desde [`template_standard_CloudRun_Repo`](https://github.com/nexuraintl/template_standard_CloudRun_Repo)) como a repos existentes que haya que migrar al estándar.

## Estándares de referencia

| Código | Área |
|---|---|
| GOB-GCP-STD-01 | Estructura de repositorio, logging, Correlation-ID, versionado de endpoints, MANUAL.md |
| GOB-GCP-GOB-01 | Configuración del API Gateway (ESPv2, `CONSTANT_ADDRESS`, audiences, `operationId`) |
| GOB-GCP-GOB-04 | Proceso de aprobación para PROD (RFC GS-F-007_V4.0, ticket mesa de servicios) |

## Mantenimiento

Cada skill vive en su propia carpeta con un único `SKILL.md` que define metadata (`name`, `description`) y las instrucciones completas. Para actualizar una skill, editar su `SKILL.md` y volver a sincronizarla en Claude.ai / Claude Code.
