const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType
} = require('docx');

const NAVY = '081D4D';
const ORANGE = 'FF6A42';
const FONT = 'Gill Sans MT';

function emptyPara(pt) {
  return new Paragraph({ spacing: { before: 0, after: (pt || 6) * 20 } });
}

function orangeLine() {
  return new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: ORANGE, space: 1 } },
    spacing: { before: 0, after: 80 }
  });
}

function sectionHeader(text) {
  return new Paragraph({
    spacing: { before: 240, after: 80 },
    children: [new TextRun({ text: text.toUpperCase(), bold: true, size: 22, color: NAVY, font: FONT })]
  });
}

function buildPersonalTable(lines) {
  const border = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
  const borders = { top: border, bottom: border, left: border, right: border };
  const rows = lines.filter(l => l.includes(':')).map(l => {
    const idx = l.indexOf(':');
    const label = l.slice(0, idx).trim();
    const value = l.slice(idx + 1).trim();
    return new TableRow({
      children: [
        new TableCell({
          borders, width: { size: 2800, type: WidthType.DXA },
          shading: { fill: 'EEF2F8', type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: label, bold: true, size: 19, font: FONT, color: NAVY })] })]
        }),
        new TableCell({
          borders, width: { size: 6200, type: WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: value, size: 19, font: FONT })] })]
        })
      ]
    });
  });
  if (!rows.length) return null;
  return new Table({ width: { size: 9000, type: WidthType.DXA }, columnWidths: [2800, 6200], rows });
}

function parseReport(raw) {
  const lines = raw.split('\n');
  const sections = {};
  let current = 'HEADER';
  sections[current] = [];
  const KEYS = [
    'PERSÖNLICHE ANGABEN','PERSOENLICHE ANGABEN','PERSONAL DETAILS',
    'AUSBILDUNG','EDUCATION',
    'VERGÜTUNG','VERGUETUNG','COMPENSATION',
    'KARRIERE ZUSAMMENFASSUNG','CAREER SUMMARY',
    'KANDIDATENBEWERTUNG','FACHLICHES RESÜMEE','FACHLICHES RESUME','BEWERTUNG','BEWERBERMOTIVATION','MOTIVATION',
    'BERUFSERFAHRUNG','BERUFLICHER WERDEGANG','WORK EXPERIENCE','PROFESSIONAL EXPERIENCE',
    'ANMERKUNGEN ZUM WERDEGANG'
  ];
  for (const line of lines) {
    const upper = line.trim().toUpperCase();
    const matched = KEYS.find(k => upper === k || upper.startsWith(k));
    if (matched && line.trim().length > 0) {
      current = line.trim();
      sections[current] = [];
    } else {
      sections[current].push(line);
    }
  }
  return sections;
}

function buildDoc(reportText, candidateName, position, client, datum) {
  const sections = parseReport(reportText);
  const children = [];

  // Cover
  children.push(emptyPara(20));
  children.push(new Paragraph({
    spacing: { before: 0, after: 60 },
    children: [new TextRun({ text: (candidateName || 'KANDIDAT').toUpperCase(), bold: true, size: 40, color: NAVY, font: FONT })]
  }));
  children.push(new Paragraph({
    spacing: { before: 0, after: 80 },
    children: [new TextRun({ text: 'VERTRAULICHER KANDIDATENBERICHT', bold: true, size: 26, color: ORANGE, font: FONT })]
  }));
  if (position) {
    children.push(new Paragraph({
      spacing: { before: 0, after: 40 },
      children: [new TextRun({ text: position.toUpperCase(), bold: true, size: 22, color: NAVY, font: FONT })]
    }));
  }
  if (client) {
    children.push(new Paragraph({
      spacing: { before: 0, after: 40 },
      children: [new TextRun({ text: client, size: 20, color: '444444', font: FONT })]
    }));
  }
  children.push(new Paragraph({
    spacing: { before: 0, after: 200 },
    children: [new TextRun({ text: datum || '', size: 18, color: '888888', font: FONT })]
  }));
  children.push(orangeLine());
  children.push(emptyPara(8));
  children.push(new Paragraph({
    spacing: { before: 0, after: 240 },
    children: [new TextRun({
      text: 'Dieser Vertrauliche Bericht enthält zum Teil Informationen, die uns unter Zusicherung strengster Vertraulichkeit mitgeteilt wurden. Entsprechend unseren berufsethischen Prinzipien müssen wir Sie dazu verpflichten, nur einer begrenzten Auswahl von Personen, die sich direkt mit der Auswertung befassen, Einsicht in diese Berichte zu gewähren. Der Inhalt muss auch jeglichen Drittpersonen gegenüber geheim gehalten werden.',
      size: 17, italics: true, color: '555555', font: FONT
    })]
  }));

  // Sections
  for (const [key, lines] of Object.entries(sections)) {
    if (key === 'HEADER') continue;
    const content = lines.map(l => l.trim()).filter(Boolean);
    if (!content.length) continue;

    children.push(sectionHeader(key));
    children.push(orangeLine());

    const keyUpper = key.toUpperCase();
    const isPersonal = keyUpper.includes('PERSÖN') || keyUpper.includes('PERSOEN') || keyUpper.includes('PERSONAL');

    if (isPersonal) {
      const tbl = buildPersonalTable(content);
      if (tbl) { children.push(tbl); children.push(emptyPara(10)); continue; }
    }

    for (const line of content) {
      const isBullet = /^[-•–]/.test(line);
      const isCareerLine = /\d{4}/.test(line) && line.includes('|');
      const isSubheading = line === line.toUpperCase() && line.length > 4 && !/\d/.test(line);
      const isBold = /^(Hauptverantwortlichkeiten|Key Achievements|Verantwortlichkeiten):/.test(line);

      if (isCareerLine) {
        children.push(emptyPara(6));
        children.push(new Paragraph({
          spacing: { before: 60, after: 40 },
          children: [new TextRun({ text: line, bold: true, size: 20, color: NAVY, font: FONT })]
        }));
      } else if (isBullet) {
        children.push(new Paragraph({
          spacing: { before: 40, after: 40 },
          indent: { left: 360 },
          children: [new TextRun({ text: '• ' + line.replace(/^[-•–]\s*/, ''), size: 19, font: FONT })]
        }));
      } else if (isBold) {
        children.push(new Paragraph({
          spacing: { before: 100, after: 40 },
          children: [new TextRun({ text: line, bold: true, size: 19, color: NAVY, font: FONT })]
        }));
      } else if (isSubheading) {
        children.push(new Paragraph({
          spacing: { before: 160, after: 60 },
          children: [new TextRun({ text: line, bold: true, size: 20, color: NAVY, font: FONT })]
        }));
      } else {
        children.push(new Paragraph({
          spacing: { before: 0, after: 100 },
          children: [new TextRun({ text: line, size: 19, font: FONT })]
        }));
      }
    }
  }

  // Footer
  children.push(emptyPara(30));
  children.push(orangeLine());
  children.push(new Paragraph({
    spacing: { before: 120, after: 0 },
    children: [new TextRun({ text: 'Vorbereitet von: Dr. Sami Hamid | Managing Partner | Signium Austria', size: 18, color: NAVY, font: FONT })]
  }));
  children.push(new Paragraph({
    spacing: { before: 40, after: 0 },
    children: [new TextRun({ text: 't +43 664 4568862 | sami.hamid@signium.com', size: 17, color: '555555', font: FONT })]
  }));

  return new Document({
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 }
        }
      },
      children
    }]
  });
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method Not Allowed' });

  try {
    const { text, candidateName, position, client, datum } = req.body;
    const doc = buildDoc(text, candidateName, position, client, datum);
    const buffer = await Packer.toBuffer(doc);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename="Signium_Kandidatenbericht.docx"');
    res.send(buffer);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}
