const fs = require('fs');
const path = require('path');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle, AlignmentType, PageBreak, ExternalHyperlink } = require('docx');
const ExcelJS = require('exceljs');
const pptxgen = require('pptxgenjs');

const out = __dirname;
const trip = {
  title: 'Azores: São Miguel + Terceira',
  dates: '12–19 de octubre de 2026',
  nights: 7,
  travelers: 2,
  budget: 5000,
  route: 'Madrid → Lisboa → Ponta Delgada → Terceira → Lisboa → Madrid',
  days: [
    ['12 oct., lunes', 'Madrid → Ponta Delgada', 'Vuelo con escala probable en Lisboa; coche y alojamiento en Ponta Delgada.'],
    ['13 oct., martes', 'São Miguel: Sete Cidades', 'Vista do Rei, Lagoa do Canário y tarde libre en Ponta Delgada.'],
    ['14 oct., miércoles', 'São Miguel: Lagoa do Fogo', 'Lagoa do Fogo, Caldeira Velha y Ribeira Grande.'],
    ['15 oct., jueves', 'São Miguel: Furnas', 'Lagoa das Furnas, Terra Nostra y aguas termales.'],
    ['16 oct., viernes', 'São Miguel → Terceira', 'Vuelo regional, coche y paseo por Angra do Heroísmo y Monte Brasil.'],
    ['17 oct., sábado', 'Terceira interior', 'Algar do Carvão, Furnas do Enxofre y Serra do Cume.'],
    ['18 oct., domingo', 'Terceira costa', 'Biscoitos, piscinas naturales o cetáceos si la meteorología permite.'],
    ['19 oct., lunes', 'Terceira → Madrid', 'Regreso con escala probable; devolver el coche con margen.']
  ],
  budgetRows: [
    ['Vuelos Madrid–Azores–Madrid', 1500, 'Estimado'],
    ['Vuelo regional entre islas', 260, 'Estimado'],
    ['Alojamiento, 7 noches', 1450, 'Estimado'],
    ['Coches de alquiler y combustible', 750, 'Estimado'],
    ['Comidas', 780, 'Estimado'],
    ['Termas, excursiones y actividades', 350, 'Estimado'],
    ['Seguro de viaje', 110, 'Estimado'],
    ['Contingencia', 300, 'Estimado']
  ],
  sources: [
    ['Turismo oficial de Azores', 'https://www.visitazores.com/es/', 'Destino, islas y experiencias'],
    ['Planificar viaje en Azores', 'https://www.visitazores.com/es/planificar-viaje-azores', 'Cómo llegar, transporte y alojamiento'],
    ['Senderos oficiales', 'https://trails.visitazores.com/es', 'Rutas de senderismo'],
    ['Observación de cetáceos', 'https://www.visitazores.com/es/experiencias-azores/observacion-de-ballenas', 'Actividad y operadores'],
    ['Azores Airlines', 'https://www.azoresairlines.pt/en/', 'Vuelos y equipaje'],
    ['Atlânticoline', 'https://www.atlanticoline.pt/', 'Ferries y horarios'],
    ['Equipaje TAP', 'https://www.flytap.com/en-es/baggage/hold-baggage', 'Maleta facturada'],
    ['Documentación UE', 'https://europa.eu/youreurope/citizens/travel/entry-exit/eu-citizen/index_en.htm', 'DNI o pasaporte válido'],
    ['Roaming UE', 'https://europa.eu/youreurope/citizens/travel/connected/electronic-communications/index_en.htm', 'Comunicaciones']
  ]
};

const total = trip.budgetRows.reduce((sum, row) => sum + row[1], 0);
const colors = { navy: '153243', teal: '007C83', coral: 'E86A5B', cream: 'F7F2E8', ink: '24323D', gray: '66737D', white: 'FFFFFF', green: '2E8B72' };
const dxa = 9360;
const borders = { top: { style: BorderStyle.SINGLE, size: 1, color: 'D6DDE1' }, bottom: { style: BorderStyle.SINGLE, size: 1, color: 'D6DDE1' }, left: { style: BorderStyle.SINGLE, size: 1, color: 'D6DDE1' }, right: { style: BorderStyle.SINGLE, size: 1, color: 'D6DDE1' } };

function cell(text, width, header = false) {
  return new TableCell({ width: { size: width, type: WidthType.DXA }, borders, shading: header ? { fill: colors.navy, type: ShadingType.CLEAR } : undefined, margins: { top: 90, bottom: 90, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: String(text), bold: header, color: header ? colors.white : colors.ink, font: 'Aptos', size: 20 })] })] });
}
function table(headers, rows, widths) {
  return new Table({ width: { size: dxa, type: WidthType.DXA }, columnWidths: widths, rows: [new TableRow({ children: headers.map((h, i) => cell(h, widths[i], true)) }), ...rows.map(row => new TableRow({ children: row.map((v, i) => cell(v, widths[i])) }))] });
}
function p(text, opts = {}) { return new Paragraph({ spacing: { after: 120 }, ...opts, children: [new TextRun({ text, font: 'Aptos', size: 22, color: colors.ink })] }); }
function link(label, url) { return new Paragraph({ spacing: { after: 80 }, children: [new ExternalHyperlink({ link: url, children: [new TextRun({ text: label, style: 'Hyperlink', font: 'Aptos', size: 20 })] })] }); }

async function createDocx() {
  const doc = new Document({
    styles: { default: { document: { run: { font: 'Aptos', size: 22 } } }, paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', run: { font: 'Aptos Display', size: 34, bold: true, color: colors.navy }, paragraph: { spacing: { before: 300, after: 180 }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', run: { font: 'Aptos Display', size: 27, bold: true, color: colors.teal }, paragraph: { spacing: { before: 240, after: 140 }, outlineLevel: 1 } }
    ] },
    sections: [{ properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1100, right: 1440, bottom: 1100, left: 1440 } } }, children: [
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 900, after: 200 }, children: [new TextRun({ text: 'AZORES', font: 'Aptos Display', size: 58, bold: true, color: colors.teal })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 250 }, children: [new TextRun({ text: 'São Miguel + Terceira', font: 'Aptos Display', size: 34, bold: true, color: colors.navy })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 600 }, children: [new TextRun({ text: trip.dates + '  |  2 viajeros  |  Madrid', font: 'Aptos', size: 24, color: colors.gray })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: colors.coral, space: 1 } }, spacing: { after: 400 }, children: [new TextRun({ text: 'Naturaleza volcánica, mar abierto y dos islas con ritmos distintos.', font: 'Aptos', size: 25, italic: true, color: colors.ink })] }),
      p('Presupuesto objetivo: 5.000 € para dos personas. La propuesta combina una isla principal de naturaleza y termas con una segunda isla de patrimonio, costa y gastronomía.'),
      new Paragraph({ children: [new PageBreak()] }),
      new Paragraph({ text: 'Resumen del viaje', heading: HeadingLevel.HEADING_1 }),
      table(['Dato', 'Propuesta'], [['Fechas', trip.dates], ['Duración', '7 noches'], ['Ruta', trip.route], ['Distribución', 'São Miguel: 4 noches · Terceira: 3 noches'], ['Transporte', 'Vuelos + coche de alquiler en cada isla'], ['Estado general', 'Estimaciones pendientes de cotización y disponibilidad']], [2500, 6860]),
      new Paragraph({ text: 'Itinerario diario', heading: HeadingLevel.HEADING_1 }),
      table(['Día', 'Zona', 'Plan'], trip.days, [1500, 2700, 5160]),
      new Paragraph({ text: 'Rutas y logística', heading: HeadingLevel.HEADING_1 }),
      p('Ruta aérea principal: Madrid–Ponta Delgada y Terceira–Madrid, previsiblemente con escala en Lisboa. Ruta interinsular: vuelo regional Ponta Delgada–Terceira. Los horarios, escalas y tarifas deben comprobarse en la reserva.'),
      p('Se recomienda recoger coche en ambos aeropuertos. En carretera, reservar margen para miradores, niebla y lluvia. La alternativa marítima es Atlânticoline, pero la operativa concreta de octubre debe verificarse antes de sustituir el vuelo.'),
      new Paragraph({ text: 'Presupuesto', heading: HeadingLevel.HEADING_1 }),
      table(['Partida', 'Importe estimado (€)', 'Estado'], trip.budgetRows.map(r => [r[0], r[1].toLocaleString('es-ES'), r[2]]), [5300, 2000, 2060]),
      p('Total estimado: ' + total.toLocaleString('es-ES') + ' €. Diferencia frente al objetivo: ' + (trip.budget - total).toLocaleString('es-ES') + ' €. Los importes son orientativos y no representan disponibilidad ni precio garantizado.'),
      new Paragraph({ text: 'Equipaje y requisitos', heading: HeadingLevel.HEADING_1 }),
      p('Equipaje previsto: una maleta facturada convencional por persona y un artículo personal. TAP publica como referencia, según tarifa y ruta, hasta 23 kg y 158 cm lineales para la maleta facturada; las tarifas económicas pueden no incluirla. La franquicia de Azores Airlines debe revisarse en cada tarifa.'),
      p('Para viajeros españoles: DNI o pasaporte válido el día del viaje. Portugal pertenece a la UE y al espacio Schengen. Se recomienda llevar Tarjeta Sanitaria Europea y seguro con asistencia, equipaje y cancelación. El roaming europeo se aplica sujeto a la política de uso razonable.'),
      p('Octubre: clima históricamente templado y cambiante, con lluvia, viento y niebla posibles. Llevar calzado con agarre, capa impermeable y ropa adaptable. La previsión meteorológica de los días del viaje debe consultarse cerca de la salida.'),
      new Paragraph({ text: 'Fuentes y verificación', heading: HeadingLevel.HEADING_1 }),
      p('Fuentes consultadas el 17/09/2026:'),
      ...trip.sources.map(s => link(s[0] + ' — ' + s[2], s[1])),
      p('Pendiente antes de reservar: cotización real de vuelos y coches, maleta incluida, horarios del vuelo interinsular, alojamiento, termas, excursiones de cetáceos, política de cancelación y previsión meteorológica.')
    ] }]
  });
  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(path.join(out, 'Azores_Madrid_12-19_octubre_2026.docx'), buffer);
}

async function createXlsx() {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Planificador de viajes'; wb.created = new Date('2026-09-17');
  const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.navy } };
  const titleFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.teal } };
  function setup(ws, widths) { ws.properties.defaultRowHeight = 20; ws.views = [{ state: 'frozen', ySplit: 3 }]; widths.forEach((w, i) => ws.getColumn(i + 1).width = w); }
  function title(ws, text, end = 4) { ws.mergeCells(1, 1, 1, end); const c = ws.getCell(1, 1); c.value = text; c.font = { name: 'Aptos Display', size: 18, bold: true, color: 'FFFFFF' }; c.fill = titleFill; c.alignment = { vertical: 'middle' }; ws.getRow(1).height = 30; }
  function headers(ws, values) { const row = ws.getRow(3); values.forEach((v, i) => { const c = row.getCell(i + 1); c.value = v; c.font = { name: 'Aptos', bold: true, color: 'FFFFFF' }; c.fill = headerFill; c.alignment = { wrapText: true }; }); }
  const summary = wb.addWorksheet('Resumen'); setup(summary, [24, 75]); title(summary, 'AZORES | Resumen', 2); headers(summary, ['Dato', 'Valor']); [['Fechas', trip.dates], ['Viajeros', '2 personas españolas'], ['Salida', 'Madrid'], ['Duración', '7 noches'], ['Ruta', trip.route], ['Presupuesto objetivo', trip.budget], ['Total estimado', { formula: "='Presupuesto'!B12" }], ['Estado', 'Estimado; verificar antes de reservar']].forEach((r, i) => { summary.getCell(i + 4, 1).value = r[0]; summary.getCell(i + 4, 2).value = r[1]; }); summary.getCell(9, 2).numFmt = '#,##0 €'; summary.getCell(10, 2).numFmt = '#,##0 €';
  const budget = wb.addWorksheet('Presupuesto'); setup(budget, [42, 20, 18, 35]); title(budget, 'AZORES | Presupuesto para 2 viajeros', 4); headers(budget, ['Partida', 'Importe (€)', 'Estado', 'Notas']); trip.budgetRows.forEach((r, i) => { const row = i + 4; budget.getCell(row, 1).value = r[0]; budget.getCell(row, 2).value = r[1]; budget.getCell(row, 3).value = r[2]; budget.getCell(row, 4).value = 'Estimación orientativa; cotizar antes de reservar'; }); budget.getCell(12, 1).value = 'TOTAL ESTIMADO'; budget.getCell(12, 2).value = { formula: '=SUM(B4:B11)' }; budget.getCell(12, 1).font = { bold: true }; budget.getCell(12, 2).font = { bold: true }; budget.getCell(4, 2).numFmt = '#,##0 €'; budget.getColumn(2).numFmt = '#,##0 €';
  const itinerary = wb.addWorksheet('Itinerario'); setup(itinerary, [18, 28, 75]); title(itinerary, 'AZORES | Itinerario diario', 3); headers(itinerary, ['Día', 'Zona / etapa', 'Plan']); trip.days.forEach(r => itinerary.addRow(r)); itinerary.eachRow((row, index) => { if (index > 3) row.alignment = { wrapText: true, vertical: 'top' }; });
  const routes = wb.addWorksheet('Rutas'); setup(routes, [28, 28, 22, 35]); title(routes, 'AZORES | Rutas y alternativas', 4); headers(routes, ['Trayecto', 'Medio', 'Duración', 'Notas']); [['Madrid–Ponta Delgada', 'Avión, escala probable en Lisboa', 'Pendiente', 'Comprobar tarifa, horarios y maleta'], ['Ponta Delgada–Terceira', 'Vuelo regional', 'Pendiente', 'Preferible al ferry para asegurar octubre'], ['Terceira–Madrid', 'Avión, escala probable', 'Pendiente', 'Devolver coche con margen']].forEach(r => routes.addRow(r));
  const luggage = wb.addWorksheet('Equipaje'); setup(luggage, [34, 30, 36]); title(luggage, 'AZORES | Equipaje y documentación', 3); headers(luggage, ['Elemento', 'Dato', 'Estado / fuente']); [['Maleta facturada', '1 por persona; peso y precio dependen de tarifa', 'Pendiente; revisar compañía'], ['Referencia TAP', 'Hasta 23 kg y 158 cm lineales según tarifa/ruta', 'Confirmar en reserva'], ['DNI o pasaporte', 'Válido el día del viaje', 'Confirmado por Your Europe'], ['Tarjeta Sanitaria Europea', 'Recomendable', 'Recomendación'], ['Seguro', 'Asistencia médica, equipaje y cancelación', 'Recomendado']].forEach(r => luggage.addRow(r));
  const links = wb.addWorksheet('Enlaces'); setup(links, [30, 58, 42]); title(links, 'AZORES | Enlaces útiles', 3); headers(links, ['Recurso', 'URL', 'Uso']); trip.sources.forEach(r => { const row = links.addRow(r); row.getCell(2).value = { text: r[1], hyperlink: r[1] }; row.getCell(2).font = { color: '0563C1', underline: true }; });
  const sources = wb.addWorksheet('Fuentes'); setup(sources, [30, 22, 80]); title(sources, 'AZORES | Fuentes y estado', 3); headers(sources, ['Fuente', 'Fecha consulta', 'Observación']); trip.sources.forEach(r => sources.addRow([r[0], '17/09/2026', r[2] + ' | ' + r[1]]));
  for (const ws of wb.worksheets) { ws.eachRow(row => row.eachCell(c => { c.font = { name: 'Aptos', size: 10, ...(c.font || {}) }; c.alignment = { vertical: 'top', wrapText: true, ...(c.alignment || {}) }; })); }
  await wb.xlsx.writeFile(path.join(out, 'Azores_Madrid_12-19_octubre_2026.xlsx'));
}

async function createPptx() {
  const ppt = new pptxgen(); ppt.layout = 'LAYOUT_WIDE'; ppt.author = 'Planificador de viajes'; ppt.subject = trip.title; ppt.title = trip.title; ppt.company = 'Plan de viaje';
  const W = 13.333, H = 7.5;
  function base(slide, titleText, subtitle = '') { slide.background = { color: colors.cream }; slide.addShape(ppt.ShapeType.rect, { x: 0, y: 0, w: W, h: 0.18, fill: { color: colors.coral }, line: { color: colors.coral } }); slide.addText(titleText, { x: 0.65, y: 0.55, w: 8.8, h: 0.55, fontFace: 'Aptos Display', fontSize: 28, bold: true, color: colors.navy, margin: 0 }); if (subtitle) slide.addText(subtitle, { x: 0.68, y: 1.15, w: 11.5, h: 0.35, fontFace: 'Aptos', fontSize: 12, color: colors.gray, margin: 0 }); slide.addText('AZORES  /  12–19 OCT 2026', { x: 10.1, y: 0.62, w: 2.55, h: 0.25, fontFace: 'Aptos', fontSize: 9, bold: true, color: colors.teal, align: 'right', margin: 0 }); }
  function pill(slide, x, y, w, text, fill = colors.teal) { slide.addShape(ppt.ShapeType.roundRect, { x, y, w, h: 0.42, rectRadius: 0.06, fill: { color: fill }, line: { color: fill } }); slide.addText(text, { x, y: y + 0.08, w, h: 0.2, fontFace: 'Aptos', fontSize: 11, bold: true, color: colors.white, align: 'center', margin: 0 }); }
  let s = ppt.addSlide(); s.background = { color: colors.navy }; s.addShape(ppt.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: colors.navy }, line: { color: colors.navy } }); s.addShape(ppt.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 0.2, fill: { color: colors.coral }, line: { color: colors.coral } }); s.addText('AZORES', { x: 0.75, y: 1.1, w: 7, h: 0.8, fontFace: 'Aptos Display', fontSize: 54, bold: true, color: '8FE0D2', margin: 0 }); s.addText('São Miguel + Terceira', { x: 0.78, y: 2.05, w: 8.4, h: 0.65, fontFace: 'Aptos Display', fontSize: 31, bold: true, color: colors.white, margin: 0 }); s.addText('7 noches para alternar volcanes, termas, costa y patrimonio atlántico.', { x: 0.8, y: 3.05, w: 6.8, h: 0.65, fontFace: 'Aptos', fontSize: 19, color: 'D8E7E4', margin: 0 }); pill(s, 0.8, 4.35, 2.15, '12–19 OCT 2026', colors.coral); pill(s, 3.15, 4.35, 1.45, '2 viajeros', colors.teal); s.addShape(ppt.ShapeType.arc, { x: 8.4, y: 1.05, w: 3.2, h: 3.2, line: { color: colors.coral, width: 3, transparency: 15 }, adjustPoint: 0.25 }); s.addShape(ppt.ShapeType.arc, { x: 9.2, y: 2.0, w: 2.4, h: 2.4, line: { color: '8FE0D2', width: 2, transparency: 10 }, adjustPoint: 0.25 }); s.addText('MADRID  →  AZORES', { x: 8.3, y: 5.75, w: 3.6, h: 0.3, fontFace: 'Aptos', fontSize: 12, color: '8FE0D2', align: 'center', margin: 0 });
  s = ppt.addSlide(); base(s, 'La idea del viaje', 'Una combinación compacta y variada, pensada para moverse con coche y dejar margen al tiempo atlántico.'); [['4 noches', 'São Miguel', 'Lagunas, Furnas, termas y miradores'], ['3 noches', 'Terceira', 'Angra, volcanes, costa y gastronomía'], ['5.000 €', 'Objetivo total', 'Presupuesto orientativo para dos']].forEach((r, i) => { const x = 0.75 + i * 4.15; s.addShape(ppt.ShapeType.rect, { x, y: 2.0, w: 3.45, h: 2.3, fill: { color: i === 1 ? 'D7EEE9' : colors.white }, line: { color: 'D6DDE1', width: 1 } }); s.addText(r[0], { x: x + 0.25, y: 2.35, w: 2.9, h: 0.55, fontFace: 'Aptos Display', fontSize: 28, bold: true, color: colors.coral, margin: 0 }); s.addText(r[1], { x: x + 0.25, y: 3.05, w: 2.9, h: 0.35, fontFace: 'Aptos Display', fontSize: 18, bold: true, color: colors.navy, margin: 0 }); s.addText(r[2], { x: x + 0.25, y: 3.55, w: 2.9, h: 0.45, fontFace: 'Aptos', fontSize: 12, color: colors.gray, margin: 0 }); });
  s = ppt.addSlide(); base(s, 'La ruta', 'El vuelo interinsular reduce el riesgo logístico frente a depender de una conexión marítima no confirmada.'); const route = [['MAD', 'Madrid'], ['LIS', 'Lisboa'], ['PDL', 'Ponta Delgada'], ['TER', 'Terceira'], ['LIS', 'Lisboa'], ['MAD', 'Madrid']]; route.forEach((r, i) => { const x = 0.85 + i * 2.35; if (i < route.length - 1) s.addShape(ppt.ShapeType.line, { x: x + 0.52, y: 3.1, w: 1.85, h: 0, line: { color: i === 2 ? colors.coral : colors.teal, width: 2, dash: i === 2 ? 'dash' : 'solid', beginArrowType: 'none', endArrowType: 'triangle' } }); s.addShape(ppt.ShapeType.ellipse, { x, y: 2.68, w: 1.05, h: 0.85, fill: { color: i === 2 || i === 3 ? colors.coral : colors.teal }, line: { color: colors.white, width: 2 } }); s.addText(r[0], { x, y: 2.93, w: 1.05, h: 0.2, fontFace: 'Aptos', fontSize: 15, bold: true, color: colors.white, align: 'center', margin: 0 }); s.addText(r[1], { x: x - 0.25, y: 3.8, w: 1.55, h: 0.3, fontFace: 'Aptos', fontSize: 11, color: colors.ink, align: 'center', margin: 0 }); }); s.addText('Tramo interinsular: vuelo regional Ponta Delgada → Terceira', { x: 2.2, y: 5.15, w: 8.8, h: 0.35, fontFace: 'Aptos', fontSize: 15, bold: true, color: colors.navy, align: 'center', margin: 0 });
  s = ppt.addSlide(); base(s, 'Ocho días, dos ritmos', 'El plan deja las actividades de mar y montaña intercambiables según nubosidad, viento y lluvia.'); trip.days.forEach((r, i) => { const col = i % 2, row = Math.floor(i / 2), x = 0.8 + col * 6.1, y = 1.75 + row * 1.25; s.addShape(ppt.ShapeType.rect, { x, y, w: 5.55, h: 0.92, fill: { color: i === 4 ? 'D7EEE9' : colors.white }, line: { color: 'D6DDE1', width: 0.8 } }); s.addText(r[0], { x: x + 0.18, y: y + 0.17, w: 1.2, h: 0.2, fontFace: 'Aptos', fontSize: 10, bold: true, color: colors.coral, margin: 0 }); s.addText(r[1], { x: x + 1.45, y: y + 0.14, w: 3.8, h: 0.23, fontFace: 'Aptos Display', fontSize: 14, bold: true, color: colors.navy, margin: 0 }); s.addText(r[2], { x: x + 1.45, y: y + 0.47, w: 3.8, h: 0.22, fontFace: 'Aptos', fontSize: 10, color: colors.gray, margin: 0 }); });
  s = ppt.addSlide(); base(s, 'Presupuesto bajo control', 'Estimación de trabajo; la disponibilidad real debe cotizarse antes de reservar.'); s.addChart(ppt.ChartType.doughnut, [{ name: 'Coste', labels: trip.budgetRows.map(r => r[0]), values: trip.budgetRows.map(r => r[1]) }], { x: 0.7, y: 1.7, w: 5.6, h: 4.8, holeSize: 58, chartColors: [colors.teal, colors.coral, 'F3B562', '6C8E8E', 'A9C5BA', '9A7B67', '496A81', 'CBD5D1'], showLegend: false, showTitle: false, showValue: false, chartArea: { fill: { color: colors.cream }, line: { color: colors.cream } } }); s.addText(total.toLocaleString('es-ES') + ' €', { x: 2.05, y: 3.42, w: 2.9, h: 0.5, fontFace: 'Aptos Display', fontSize: 27, bold: true, color: colors.navy, align: 'center', margin: 0 }); s.addText('estimación total', { x: 2.05, y: 3.98, w: 2.9, h: 0.25, fontFace: 'Aptos', fontSize: 11, color: colors.gray, align: 'center', margin: 0 }); trip.budgetRows.slice(0, 6).forEach((r, i) => { const y = 1.85 + i * 0.63; s.addShape(ppt.ShapeType.rect, { x: 7.0, y: y + 0.03, w: 0.15, h: 0.15, fill: { color: ['007C83','E86A5B','F3B562','6C8E8E','A9C5BA','9A7B67'][i] }, line: { color: 'FFFFFF', transparency: 100 } }); s.addText(r[0], { x: 7.3, y, w: 3.4, h: 0.22, fontFace: 'Aptos', fontSize: 12, color: colors.ink, margin: 0 }); s.addText(r[1].toLocaleString('es-ES') + ' €', { x: 10.8, y, w: 1.15, h: 0.22, fontFace: 'Aptos', fontSize: 12, bold: true, color: colors.navy, align: 'right', margin: 0 }); });
  s = ppt.addSlide(); base(s, 'Antes de reservar', 'Cinco comprobaciones que convierten la idea en un viaje ejecutable.'); [['01', 'Vuelos', 'Horarios, escalas y precio final desde Madrid'], ['02', 'Maletas', 'Franquicia incluida en cada tarifa'], ['03', 'Coches', 'Seguro, depósito y política de combustible'], ['04', 'Tiempo', 'Previsión de montaña y salidas al mar'], ['05', 'Flexibilidad', 'Cancelación y cambios en alojamiento/actividades']].forEach((r, i) => { const x = 0.8 + (i % 3) * 4.05, y = 1.85 + Math.floor(i / 3) * 2.25; s.addText(r[0], { x, y, w: 0.65, h: 0.5, fontFace: 'Aptos Display', fontSize: 26, bold: true, color: colors.coral, margin: 0 }); s.addText(r[1], { x: x + 0.85, y: y + 0.03, w: 2.8, h: 0.3, fontFace: 'Aptos Display', fontSize: 18, bold: true, color: colors.navy, margin: 0 }); s.addText(r[2], { x: x + 0.85, y: y + 0.52, w: 3.0, h: 0.45, fontFace: 'Aptos', fontSize: 12, color: colors.gray, margin: 0 }); }); s.addText('DNI o pasaporte válido · euro · roaming UE · seguro recomendado', { x: 0.85, y: 6.35, w: 11.6, h: 0.3, fontFace: 'Aptos', fontSize: 13, bold: true, color: colors.teal, align: 'center', margin: 0 });
  await ppt.writeFile({ fileName: path.join(out, 'Azores_Madrid_12-19_octubre_2026.pptx') });
}

(async () => { await createDocx(); await createXlsx(); await createPptx(); console.log('Entregables generados:', fs.readdirSync(out).filter(f => /Azores_Madrid/.test(f)).join(', ')); })().catch(error => { console.error(error); process.exitCode = 1; });
