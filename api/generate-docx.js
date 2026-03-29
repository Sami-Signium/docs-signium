const { Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle, HeadingLevel, LevelFormat } = require('docx');

const NAVY = '081D4D';
const ORANGE = 'FF6A42';

function parseReport(text) {
  const sections = [];
  const lines = text.split('\n');
  let current = { title: null, lines: [] };
  
  const sectionTitles = [
    'PERSÖNLICHE ANGABEN', 'PERSOENLICHE ANGABEN',
    'AUSBILDUNG UND QUALIFIKATIONEN', 'AUSBILDUNG',
    'VERGÜTUNG UND VERFÜGBARKEIT', 'VERGUETUNG',
    'KARRIERE ZUSAMMENFASSUNG', 'KARRIERE',
    'KANDIDATENBEWERTUNG',
    'FACHLICHES RESÜMEE', 'FACHLICHES RESUME',
    'BEWERTUNG', 'BEWERBERMOTIVATION', 'MOTIVATION',
    'BERUFSERFAHRUNG', 'BERUFLICHER WERDEGANG',
    'ANMERKUNGEN ZUM WERDEGANG'
  ];

  for (const line of lines) {
    const trimmed = line.trim();
    const isSection = sectionTitles.some(t => trimmed.toUpperCase() === t.toUpperCase() || trimmed.toUpperCase().startsWith(t.toUpperCase()));
    if (isSection && trimmed.length > 0) {
      if (current.lines.length > 0 || current.title) sections.push(current);
      current = { title: trimmed, lines: [] };
    } else {
      current.lines.push(line);
    }
  }
  if (current.lines.length > 0 || current.title) sections.push(current);
  return sections;
}

function makeParagraph(text, opts = {}) {
  return new Paragraph({
    alignment: opts.center ? AlignmentType.CENTER : AlignmentType.LEFT,
    spacing: { before: opts.before || 0, after: opts.after || 120 },
    border: opts.bottomBorder ? {
      bottom: { style: BorderStyle.SINGLE, size: 6, color: NAVY, space: 1 }
    } : undefined,
    children: [new TextRun({
      text: text || '',
      bold: opts.bold || false,
      size: opts.size || 22,
      color: opts.color || '000000',
      font: 'Gill Sans MT',
    })]
  });
}

export default async function handler(req, res) {
  if (req.method === 'OPTIONS') {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    return res.status(200).end();
  }

  if (req.method !== 'POST') return res.status(405).end();

  try {
    const { text } = req.body;
    const lines = text.split('\n');
    const children = [];

    // Cover area — first few lines
    let bodyStart = 0;
    const coverLines = [];
    for (let i = 0; i < Math.min(20, lines.length); i++) {
      const l = lines[i].trim();
      if (l === '---') { bodyStart = i + 1; break; }
      coverLines.push(l);
      bodyStart = i + 1;
    }

    // Cover page
    children.push(new Paragraph({ spacing: { before: 2000 } }));
    for (const cl of coverLines) {
      if (!cl) continue;
      const isName = coverLines.indexOf(cl) === 0;
      const isTitle = cl === 'VERTRAULICHER KANDIDATENBERICHT' || cl === 'VERTRAULICHER KANDIDATENBERICHT';
      children.push(makeParagraph(cl, {
        center: true,
        bold: isName || isTitle,
        size: isName ? 36 : isTitle ? 28 : 22,
        color: isName || isTitle ? NAVY : '444444',
        before: isName ? 400 : 120,
        after: 120
      }));
    }

    // Divider
    children.push(new Paragraph({
      spacing: { before: 400, after: 400 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: ORANGE, space: 1 } }
    }));

    // Body
    const bodyLines = lines.slice(bodyStart);
    const sectionKeywords = ['PERSÖNLICHE', 'AUSBILDUNG', 'VERGÜTUNG', 'KARRIERE', 'KANDIDATEN', 'FACHLICHES', 'BEWERTUNG', 'BEWERBERMOTIVATION', 'BERUFSERFAHRUNG', 'BERUFLICHER', 'ANMERKUNGEN', 'MOTIVATION'];

    for (const line of bodyLines) {
      const trimmed = line.trim();
      if (!trimmed || trimmed === '---') {
        children.push(new Paragraph({ spacing: { before: 0, after: 80 } }));
        continue;
      }

      const isMainSection = sectionKeywords.some(k => trimmed.toUpperCase().startsWith(k));
      const isBullet = trimmed.startsWith('-') || trimmed.startsWith('•');
      const isSubHeader = trimmed.endsWith(':') && trimmed.length < 60 && !isBullet;

      if (isMainSection) {
        children.push(new Paragraph({
          spacing: { before: 400, after: 160 },
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: NAVY, space: 1 } },
          children: [new TextRun({ text: trimmed, bold: true, size: 26, color: NAVY, font: 'Gill Sans MT' })]
        }));
      } else if (isSubHeader) {
        children.push(new Paragraph({
          spacing: { before: 200, after: 80 },
          children: [new TextRun({ text: trimmed, bold: true, size: 22, color: '333333', font: 'Gill Sans MT' })]
        }));
      } else if (isBullet) {
        children.push(new Paragraph({
          spacing: { before: 40, after: 40 },
          indent: { left: 360 },
          children: [new TextRun({ text: trimmed.replace(/^[-•]\s*/, '• '), size: 20, font: 'Gill Sans MT' })]
        }));
      } else {
        children.push(new Paragraph({
          spacing: { before: 60, after: 60 },
          children: [new TextRun({ text: trimmed, size: 20, font: 'Gill Sans MT' })]
        }));
      }
    }

    const doc = new Document({
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

    const buffer = await Packer.toBuffer(doc);

    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename="Signium_Kandidatenbericht.docx"');
    res.send(buffer);

  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}
