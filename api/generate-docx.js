const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType
} = require('docx');

// Exact values from real Signium Quintin Stephen report
const NAVY   = '102E66';  // berschrift2 color
const DARK   = '414042';  // body dark
const BODY   = '262626';  // body text
const LIGHT  = '595959';  // lighter body
const FONT_H = 'Calibri'; // headings
const FONT_B = 'Calibri Light'; // body

function sp(before, after) {
  return { spacing: { before: before || 0, after: after !== undefined ? after : 0 } };
}

function sectionHeading(text) {
  // Matches berschrift2: Calibri, sz=36 (18pt), color=102E66, bold
  return new Paragraph({
    ...sp(400, 120),
    children: [new TextRun({
      text: text.toUpperCase(),
      bold: true,
      size: 36,
      color: NAVY,
      font: FONT_H
    })]
  });
}

function subHeading(text) {
  return new Paragraph({
    ...sp(240, 80),
    children: [new TextRun({
      text: text.toUpperCase(),
      bold: true,
      size: 24,
      color: NAVY,
      font: FONT_H
    })]
  });
}

function bodyPara(text, opts) {
  opts = opts || {};
  return new Paragraph({
    ...sp(opts.before || 0, opts.after !== undefined ? opts.after : 120),
    alignment: opts.justify ? AlignmentType.BOTH : AlignmentType.LEFT,
    children: [new TextRun({
      text: text,
      size: opts.size || 22,
      bold: opts.bold || false,
      italics: opts.italic || false,
      color: opts.color || BODY,
      font: opts.font || FONT_B
    })]
  });
}

function bulletPara(text) {
  const clean = text.replace(/^[-\u2022\u2013]\s*/, '');
  return new Paragraph({
    ...sp(40, 40),
    indent: { left: 360, hanging: 220 },
    children: [
      new TextRun({ text: '\u2022  ', size: 22, color: DARK, font: FONT_B }),
      new TextRun({ text: clean, size: 22, color: BODY, font: FONT_B })
    ]
  });
}

function thinRule() {
  // Simple thin gray line - like in the real report between sections
  return new Paragraph({
    ...sp(0, 160),
    border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: 'CCCCCC', space: 1 } },
    children: [new TextRun({ text: '', size: 2 })]
  });
}

function buildPersonalTable(lines) {
  const noBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  const bottomBorder = { style: BorderStyle.SINGLE, size: 1, color: 'E0E0E0' };
  const cellBorders = {
    top: noBorder, bottom: bottomBorder,
    left: noBorder, right: noBorder,
    insideH: noBorder, insideV: noBorder
  };

  const rows = lines.filter(l => l.includes(':')).map(l => {
    const idx = l.indexOf(':');
    const label = l.slice(0, idx).trim();
    const value = l.slice(idx + 1).trim();
    return new TableRow({
      children: [
        new TableCell({
          borders: cellBorders,
          width: { size: 2800, type: WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 0, right: 120 },
          children: [new Paragraph({
            children: [new TextRun({ text: label, bold: true, size: 20, color: DARK, font: FONT_H })]
          })]
        }),
        new TableCell({
          borders: cellBorders,
          width: { size: 6200, type: WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 0, right: 0 },
          children: [new Paragraph({
            children: [new TextRun({ text: value, size: 20, color: BODY, font: FONT_B })]
          })]
        })
      ]
    });
  });

  if (!rows.length) return null;
  const noBorderT = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  return new Table({
    width: { size: 9000, type: WidthType.DXA },
    columnWidths: [2800, 6200],
    borders: { top: noBorderT, bottom: noBorderT, left: noBorderT, right: noBorderT, insideH: noBorderT, insideV: noBorderT },
    rows
  });
}

const SECTION_KEYS = [
  'PERSÖNLICHE ANGABEN','PERSOENLICHE ANGABEN','PERSONAL DETAILS','PERSONAL DATA',
  'AUSBILDUNG UND QUALIFIKATIONEN','AUSBILDUNG','EDUCATION & QUALIFICATIONS','EDUCATION',
  'VERGÜTUNG UND VERFÜGBARKEIT','VERGUETUNG','COMPENSATION & AVAILABILITY','COMPENSATION',
  'KARRIERE ZUSAMMENFASSUNG','CAREER SUMMARY',
  'KANDIDATENBEWERTUNG','CANDIDATE ASSESSMENT',
  'FACHLICHES RESÜMEE','FACHLICHES RESUME','PROFESSIONAL SUMMARY',
  'BEWERTUNG','PERSONALITY','PERSÖNLICHKEIT',
  'BEWERBERMOTIVATION','MOTIVATION',
  'BERUFSERFAHRUNG','BERUFLICHER WERDEGANG','WORK EXPERIENCE','PROFESSIONAL EXPERIENCE',
  'ANMERKUNGEN ZUM WERDEGANG'
];

const SUB_KEYS = [
  'FACHLICHES RESÜMEE','FACHLICHES RESUME','PROFESSIONAL SUMMARY',
  'BEWERTUNG','PERSONALITY','PERSÖNLICHKEIT',
  'BEWERBERMOTIVATION','MOTIVATION'
];

function parseReport(raw) {
  const lines = raw.split('\n');
  const result = [];
  let current = { key: 'HEADER', lines: [] };
  for (const line of lines) {
    const t = line.trim();
    const u = t.toUpperCase();
    const matched = SECTION_KEYS.find(k => u === k || u.startsWith(k + ' ') || u.startsWith(k + ':'));
    if (matched && t.length > 0) {
      result.push(current);
      current = { key: t, lines: [] };
    } else {
      current.lines.push(line);
    }
  }
  result.push(current);
  return result;
}

function isSubSection(line) {
  const u = line.trim().toUpperCase();
  return SUB_KEYS.some(s => u === s || u.startsWith(s + ' ') || u.startsWith(s + ':'));
}

function buildDoc(reportText, candidateName, position, client, datum) {
  const sections = parseReport(reportText);
  const children = [];

  // ── COVER ──
  // Candidate name - large, matches Titleheader style
  children.push(new Paragraph({ ...sp(120, 80) }));
  children.push(new Paragraph({
    ...sp(0, 40),
    children: [new TextRun({
      text: (candidateName || 'KANDIDAT').toUpperCase(),
      bold: true,
      size: 56,
      color: DARK,
      font: FONT_H
    })]
  }));

  // Report type - matches Coverdoctitle: navy, 15pt
  children.push(new Paragraph({
    ...sp(0, 200),
    children: [new TextRun({
      text: 'VERTRAULICHER KANDIDATENBERICHT',
      size: 30,
      color: NAVY,
      font: FONT_H
    })]
  }));

  // Position - bold dark
  if (position) {
    children.push(new Paragraph({
      ...sp(0, 60),
      children: [new TextRun({ text: position.toUpperCase(), bold: true, size: 24, color: DARK, font: FONT_H })]
    }));
  }

  // Client
  if (client && client !== 'Vertraulich') {
    children.push(new Paragraph({
      ...sp(0, 40),
      children: [new TextRun({ text: client, size: 22, color: DARK, font: FONT_B })]
    }));
  }

  // Date
  children.push(new Paragraph({
    ...sp(0, 400),
    children: [new TextRun({ text: datum || '', size: 22, color: DARK, font: FONT_B })]
  }));

  // Thin rule
  children.push(thinRule());

  // Confidentiality text - italic, small, justified
  children.push(new Paragraph({
    ...sp(0, 300),
    alignment: AlignmentType.BOTH,
    children: [new TextRun({
      text: 'Dieser Vertrauliche Bericht enthält zum Teil Informationen, die uns unter Zusicherung strengster Vertraulichkeit mitgeteilt wurden. Entsprechend unseren berufsethischen Prinzipien müssen wir Sie dazu verpflichten, nur einer begrenzten Auswahl von Personen, die sich direkt mit der Auswertung befassen, Einsicht in diese Berichte zu gewähren. Der Inhalt muss auch jeglichen Drittpersonen gegenüber geheim gehalten werden. Es dürfen keinerlei Referenzen ohne Zustimmung des Kandidaten oder unsererseits eingeholt werden.',
      size: 17, italics: true, color: '666666', font: FONT_B
    })]
  }));

  // ── SECTIONS ──
  for (const section of sections) {
    if (section.key === 'HEADER') continue;
    const content = section.lines.map(l => l.trim()).filter(Boolean);
    if (!content.length) continue;

    const ku = section.key.toUpperCase();
    const isPersonal = ku.includes('PERSÖN') || ku.includes('PERSOEN') || ku.includes('PERSONAL');
    const isKandidaten = ku.includes('KANDIDATEN') || ku.includes('CANDIDATE ASSESSMENT');

    children.push(sectionHeading(section.key));
    children.push(thinRule());

    if (isKandidaten) {
      let curSub = null;
      let subLines = [];
      const flush = () => {
        if (!curSub || !subLines.length) return;
        children.push(subHeading(curSub));
        for (const l of subLines) {
          if (!l.trim()) continue;
          if (/^[-\u2022]/.test(l)) children.push(bulletPara(l));
          else children.push(bodyPara(l, { justify: true, after: 160 }));
        }
        subLines = [];
      };
      for (const line of content) {
        if (isSubSection(line)) { flush(); curSub = line.trim(); }
        else subLines.push(line);
      }
      flush();
      continue;
    }

    if (isPersonal) {
      const tbl = buildPersonalTable(content);
      if (tbl) { children.push(tbl); children.push(new Paragraph({ ...sp(0, 80) })); continue; }
    }

    for (const line of content) {
      const isBullet = /^[-\u2022\u2013]/.test(line);
      const isCareer = /\d{4}/.test(line) && line.includes('|');
      const isBoldLabel = /^(Hauptverantwortlichkeiten|Key Achievements|Verantwortlichkeiten|Responsibilities):?$/.test(line.trim());
      const isItalic = line.startsWith('*') && line.endsWith('*');

      if (isCareer) {
        children.push(new Paragraph({
          ...sp(200, 40),
          children: [new TextRun({ text: line, bold: true, size: 22, color: NAVY, font: FONT_H })]
        }));
      } else if (isBoldLabel) {
        children.push(new Paragraph({
          ...sp(120, 40),
          children: [new TextRun({ text: line, bold: true, size: 20, color: DARK, font: FONT_H })]
        }));
      } else if (isBullet) {
        children.push(bulletPara(line));
      } else if (isItalic) {
        children.push(bodyPara(line.replace(/^\*|\*$/g, ''), { italic: true, color: LIGHT, after: 80 }));
      } else {
        children.push(bodyPara(line, { justify: true }));
      }
    }
  }

  // ── FOOTER ──
  children.push(new Paragraph({ ...sp(400, 0) }));
  children.push(thinRule());
  children.push(new Paragraph({
    ...sp(120, 0),
    children: [new TextRun({ text: 'Vorbereitet von: Dr. Sami Hamid  |  Managing Partner  |  Signium Austria', size: 18, bold: true, color: NAVY, font: FONT_H })]
  }));
  children.push(new Paragraph({
    ...sp(40, 0),
    children: [new TextRun({ text: 't +43 664 4568862  |  sami.hamid@signium.com', size: 17, color: LIGHT, font: FONT_B })]
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
    if (!text) return res.status(400).json({ error: 'No text provided' });
    const doc = buildDoc(text, candidateName, position, client, datum);
    const buffer = await Packer.toBuffer(doc);
    const safeName = (candidateName || 'Kandidat').replace(/\s+/g, '_');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${safeName}_Signium_Bericht.docx"`);
    res.send(buffer);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}
