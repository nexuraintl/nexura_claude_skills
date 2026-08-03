---
name: docx-nexura
description: >
  Genera documentos Word (.docx) con el formato corporativo oficial de NEXURA para el proyecto
  Gobernanza GCP. SIEMPRE usar este skill cuando se solicite crear cualquier documento de gobernanza,
  estándar técnico, manual, guía o entregable formal para NEXURA — incluso si el usuario solo dice
  "genera el documento", "crea el .docx" o "escribe el estándar". Este skill reemplaza al skill
  genérico `docx` para todo trabajo dentro del proyecto Gobernanza GCP. Incluye estructura base
  completa (header, footer, TOC, estilos, helpers) lista para ejecutar con Node.js.
---

# docx-nexura — Formato corporativo NEXURA

Genera documentos Word listos para entregar con el formato oficial de NEXURA.
**No requiere instrucciones de estilo adicionales** — todo está codificado en el script base.

---

## Flujo obligatorio

1. Leer este SKILL.md completo antes de generar cualquier script.
2. Copiar el script base desde `scripts/base.js` (ver abajo).
3. Rellenar el contenido del documento (secciones, tablas, código, listas).
4. Ejecutar con Node.js.
5. Validar con el script de validación.
6. Convertir a PDF para preview si se solicita.
7. Entregar con `present_files`.

```bash
node /home/claude/scripts/<nombre_documento>.js
python /mnt/skills/public/docx/scripts/office/validate.py /mnt/user-data/outputs/<archivo>.docx
# Preview opcional:
python /mnt/skills/public/docx/scripts/office/soffice.py --headless --convert-to pdf /mnt/user-data/outputs/<archivo>.docx
```

---

## Convenciones de nombre de archivo

Formato: `AAAAMMDD_nombre-del-documento_v1.docx`

Ejemplos:
- `20260619_estandar-repositorio_v1.docx`
- `20260619_proceso-aprobacion-microservicios_v1.docx`

---

## Paleta de colores NEXURA

| Token | Hex | Uso |
|-------|-----|-----|
| `RED` | `C00000` | H1, H2, headers de tabla, línea del encabezado |
| `RED_LIGHT` | `FBE5E5` | Fondo de headers de tabla |
| `GRAY_FOOTER` | `595959` | Texto de footer, placeholder de logo |
| `CODE_BG` | `F2F2F2` | Fondo de bloques de código |
| `WHITE` | `FFFFFF` | Fondo de celdas de datos en tablas |

---

## Estructura fija de todos los documentos NEXURA

1. **Encabezado de página** — `[ESPACIO PARA LOGO]` a la izquierda (itálica, gris), línea roja inferior, nombre del documento a la derecha con tab stop MAX.
2. **Footer** — nombre del documento a la izquierda, número de página a la derecha.
3. **Portada** — bloque de metadata (tabla 2 columnas: campo/valor) antes del TOC.
4. **Tabla de Contenido** — `headingStyleRange: "1-3"`, `hyperlink: true`, `updateFields: true`. Presionar Ctrl+A → F9 en Word para actualizar.
5. **Cuerpo** — H1 y H2 en rojo `C00000`, fuente Arial.
6. **Bloques de código** — tabla single-column, `Courier New` 18pt, fondo `F2F2F2`, sin bordes interiores.

---

## Script base completo

Guardar como `/home/claude/scripts/<nombre>.js` y modificar solo las secciones marcadas con `// <<< EDITAR`.

```javascript
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, TabStopType, TabStopPosition,
  TableOfContents, LevelFormat
} = require('docx');
const fs = require('fs');

// ─── CONSTANTES DE MARCA ───────────────────────────────────────────────────
const RED        = 'C00000';
const RED_LIGHT  = 'FBE5E5';
const GRAY       = '595959';
const CODE_BG    = 'F2F2F2';
const WHITE      = 'FFFFFF';

// ─── METADATOS DEL DOCUMENTO ──────────────────────────────────────────────
// <<< EDITAR
const DOC_TITLE    = 'Título del Documento';   // Aparece en header, footer y portada
const DOC_CODE     = 'GOB-GCP-XXX-00';         // Código del documento
const DOC_VERSION  = '1.0';
const DOC_DATE     = '2026-MM-DD';
const DOC_AUTHOR   = 'Santiago Valenzuela';
const DOC_PROJECT  = 'Gobernanza GCP — Microservicios sobre Cloud Run';
const DOC_TASK     = '#XX Nombre de la tarea del checklist';
const DOC_SCOPE    = 'Descripción del alcance del documento';
const OUTPUT_PATH  = '/mnt/user-data/outputs/AAAAMMDD_nombre_v1.docx';
// <<< FIN EDITAR

// ─── HELPERS ──────────────────────────────────────────────────────────────

/** Borde de tabla estándar */
const stdBorder = (color = 'CCCCCC') => ({
  style: BorderStyle.SINGLE, size: 1, color
});

/** Bordes completos para celdas de datos */
const dataBorders = {
  top:    stdBorder(), bottom: stdBorder(),
  left:   stdBorder(), right:  stdBorder()
};

/** Celda de header de tabla (fondo rojo claro, texto rojo oscuro negrita) */
function headerCell(text, widthDxa) {
  return new TableCell({
    width: { size: widthDxa, type: WidthType.DXA },
    shading: { fill: RED_LIGHT, type: ShadingType.CLEAR },
    borders: {
      top:    stdBorder(RED), bottom: stdBorder(RED),
      left:   stdBorder(RED), right:  stdBorder(RED)
    },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: AlignmentType.LEFT,
      children: [new TextRun({ text, bold: true, color: RED, font: 'Arial', size: 20 })]
    })]
  });
}

/** Celda de dato estándar */
function dataCell(text, widthDxa, bold = false) {
  return new TableCell({
    width: { size: widthDxa, type: WidthType.DXA },
    shading: { fill: WHITE, type: ShadingType.CLEAR },
    borders: dataBorders,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      children: [new TextRun({ text: String(text), bold, font: 'Arial', size: 20 })]
    })]
  });
}

/** Tabla de 2 columnas: campo (negrita) / valor */
function twoColTable(rows, col1Width = 3000, col2Width = 6360) {
  return new Table({
    width: { size: col1Width + col2Width, type: WidthType.DXA },
    columnWidths: [col1Width, col2Width],
    rows: rows.map(([campo, valor]) => new TableRow({
      children: [dataCell(campo, col1Width, true), dataCell(valor, col2Width)]
    }))
  });
}

/** Tabla con headers (primera fila = headers, resto = datos) */
function dataTable(headers, rows, colWidths) {
  const totalWidth = colWidths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: totalWidth, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map((h, i) => headerCell(h, colWidths[i]))
      }),
      ...rows.map(row => new TableRow({
        children: row.map((cell, i) => dataCell(cell, colWidths[i]))
      }))
    ]
  });
}

/** Bloque de código (Courier New, fondo gris, sin bordes interiores) */
function codeBlock(lines) {
  const noBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    borders: {
      top:    stdBorder('AAAAAA'), bottom: stdBorder('AAAAAA'),
      left:   stdBorder('AAAAAA'), right:  stdBorder('AAAAAA'),
      insideH: noBorder, insideV: noBorder
    },
    rows: lines.map(line => new TableRow({
      children: [new TableCell({
        width: { size: 9360, type: WidthType.DXA },
        shading: { fill: CODE_BG, type: ShadingType.CLEAR },
        borders: noBorders,
        margins: { top: 60, bottom: 60, left: 160, right: 160 },
        children: [new Paragraph({
          children: [new TextRun({ text: line, font: 'Courier New', size: 18 })]
        })]
      })]
    }))
  });
}

/** Párrafo vacío (separador) */
const emptyPara = () => new Paragraph({ children: [new TextRun('')] });

/** H1 rojo */
const h1 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  children: [new TextRun({ text, color: RED, bold: true, font: 'Arial', size: 32 })]
});

/** H2 rojo */
const h2 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  children: [new TextRun({ text, color: RED, bold: true, font: 'Arial', size: 28 })]
});

/** Párrafo body normal */
const p = (text) => new Paragraph({
  children: [new TextRun({ text, font: 'Arial', size: 22 })]
});

/** Párrafo de advertencia (⚠) */
const warn = (text) => new Paragraph({
  children: [new TextRun({ text: `⚠  ${text}`, font: 'Arial', size: 22, italics: true })]
});

// ─── PORTADA ──────────────────────────────────────────────────────────────

const portada = [
  twoColTable([
    ['Código del documento', DOC_CODE],
    ['Versión',              DOC_VERSION],
    ['Fecha de emisión',     DOC_DATE],
    ['Clasificación',        'Uso interno'],
    ['Elaborado por',        DOC_AUTHOR],
    ['Proyecto',             DOC_PROJECT],
    ['Tareas del checklist que cubre', DOC_TASK],
    ['Aplica a',             DOC_SCOPE],
  ]),
  emptyPara(),
  h1('Control de versiones'),
  dataTable(
    ['Versión', 'Fecha', 'Descripción', 'Autor'],
    [[DOC_VERSION, DOC_DATE, 'Versión inicial.', DOC_AUTHOR]],
    [1200, 1800, 4560, 1800]
  ),
  emptyPara(),
];

// ─── TABLA DE CONTENIDO ───────────────────────────────────────────────────

const toc = [
  new TableOfContents('Tabla de Contenido', {
    hyperlink: true,
    headingStyleRange: '1-3',
    // Actualizar en Word con Ctrl+A → F9
  }),
  emptyPara(),
];

// ─── ENCABEZADO DE PÁGINA ─────────────────────────────────────────────────

const pageHeader = new Header({
  children: [
    new Paragraph({
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: RED, space: 1 } },
      tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
      children: [
        new TextRun({ text: '[ESPACIO PARA LOGO]', italics: true, color: GRAY, font: 'Arial', size: 18 }),
        new TextRun({ text: `\t${DOC_TITLE}`, color: GRAY, font: 'Arial', size: 18 }),
      ]
    })
  ]
});

// ─── FOOTER ───────────────────────────────────────────────────────────────

const pageFooter = new Footer({
  children: [
    new Paragraph({
      tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
      children: [
        new TextRun({ text: DOC_TITLE, color: GRAY, font: 'Arial', size: 18 }),
        new TextRun({ text: '\tPágina ', color: GRAY, font: 'Arial', size: 18 }),
        new TextRun({ children: [PageNumber.CURRENT], color: GRAY, font: 'Arial', size: 18 }),
        new TextRun({ text: ' de ', color: GRAY, font: 'Arial', size: 18 }),
        new TextRun({ children: [PageNumber.TOTAL_PAGES], color: GRAY, font: 'Arial', size: 18 }),
      ]
    })
  ]
});

// ─── ESTILOS ──────────────────────────────────────────────────────────────

const styles = {
  default: {
    document: { run: { font: 'Arial', size: 22 } }
  },
  paragraphStyles: [
    {
      id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
      run: { size: 32, bold: true, color: RED, font: 'Arial' },
      paragraph: { spacing: { before: 300, after: 200 }, outlineLevel: 0 }
    },
    {
      id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
      run: { size: 28, bold: true, color: RED, font: 'Arial' },
      paragraph: { spacing: { before: 240, after: 160 }, outlineLevel: 1 }
    },
    {
      id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
      run: { size: 24, bold: true, font: 'Arial' },
      paragraph: { spacing: { before: 180, after: 120 }, outlineLevel: 2 }
    },
  ]
};

// ─── NUMERACIÓN (bullets / listas) ────────────────────────────────────────

const numbering = {
  config: [
    {
      reference: 'bullets',
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } }
      }]
    },
    {
      reference: 'numbers',
      levels: [{
        level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } }
      }]
    }
  ]
};

// ═══════════════════════════════════════════════════════════════════════════
// CONTENIDO DEL DOCUMENTO — EDITAR AQUÍ
// ═══════════════════════════════════════════════════════════════════════════

const content = [
  // <<< EDITAR: reemplazar con el contenido real del documento
  h1('1. Objetivo'),
  p('Texto del objetivo.'),
  emptyPara(),

  h1('2. Alcance'),
  p('Texto del alcance.'),
  emptyPara(),

  h1('3. Sección de ejemplo con tabla'),
  dataTable(
    ['Columna A', 'Columna B', 'Columna C'],
    [
      ['Valor 1', 'Valor 2', 'Valor 3'],
      ['Valor 4', 'Valor 5', 'Valor 6'],
    ],
    [3120, 3120, 3120]
  ),
  emptyPara(),

  h1('4. Ejemplo de bloque de código'),
  codeBlock([
    'gcloud run services update-traffic SERVICE_NAME \\',
    '  --to-revisions=REVISION=100 \\',
    '  --region=us-central1 \\',
    '  --project=PROJECT_ID',
  ]),
  emptyPara(),
  // <<< FIN EDITAR
];

// ═══════════════════════════════════════════════════════════════════════════
// ENSAMBLADO Y EXPORTACIÓN
// ═══════════════════════════════════════════════════════════════════════════

const doc = new Document({
  styles,
  numbering,
  sections: [{
    properties: {
      page: {
        size:   { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    headers:  { default: pageHeader },
    footers:  { default: pageFooter },
    children: [
      ...portada,
      ...toc,
      ...content,
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(OUTPUT_PATH, buffer);
  console.log(`✅ Documento generado: ${OUTPUT_PATH}`);
}).catch(err => {
  console.error('❌ Error:', err);
  process.exit(1);
});
```

---

## Uso de los helpers

### Párrafos y títulos
```javascript
h1('1. Título principal')        // H1 rojo
h2('1.1 Subtítulo')              // H2 rojo
p('Texto normal')                // body Arial 11pt
warn('Advertencia importante')   // ⚠ itálica
emptyPara()                      // separador vertical
```

### Listas (bullets / numeradas)
```javascript
// Bullet
new Paragraph({
  numbering: { reference: 'bullets', level: 0 },
  children: [new TextRun({ text: 'Item', font: 'Arial', size: 22 })]
})
// Numerada
new Paragraph({
  numbering: { reference: 'numbers', level: 0 },
  children: [new TextRun({ text: 'Paso 1', font: 'Arial', size: 22 })]
})
```

### Tablas
```javascript
// 2 columnas campo/valor (metadatos, configuraciones)
twoColTable([['Campo', 'Valor'], ['Otro campo', 'Otro valor']])

// Tabla con headers
dataTable(
  ['Col 1', 'Col 2', 'Col 3'],       // headers
  [['a', 'b', 'c'], ['d', 'e', 'f']], // filas
  [3120, 3120, 3120]                  // anchos en DXA — deben sumar 9360
)

// Bloque de código
codeBlock(['línea 1', 'línea 2', 'línea 3'])
```

---

## Reglas críticas

- **Anchos de columna**: siempre deben sumar `9360` (content width de Letter con 1" margins).
- **Listas**: nunca usar caracteres unicode directamente (`•`). Usar `numbering.config`.
- **TOC**: se ve vacío en PDF/preview — funciona correctamente en Word (Ctrl+A → F9).
- **Saltos de línea**: nunca usar `\n` dentro de `TextRun`. Usar `emptyPara()` o `codeBlock()` con array de líneas.
- **El script base ya incluye**: portada, control de versiones, header, footer, TOC, estilos, helpers. No reescribir estos componentes — solo editar el bloque `content`.

---

## Comandos de referencia rápida

```bash
# Ejecutar
node /home/claude/scripts/<nombre>.js

# Validar
python /mnt/skills/public/docx/scripts/office/validate.py /mnt/user-data/outputs/<archivo>.docx

# Convertir a PDF para preview
python /mnt/skills/public/docx/scripts/office/soffice.py --headless --convert-to pdf /mnt/user-data/outputs/<archivo>.docx

# Preview de páginas (requiere PDF previo)
pdftoppm -jpeg -r 150 /mnt/user-data/outputs/<archivo>.pdf /tmp/page
ls /tmp/page-*.jpg
```
