const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

export const config = { api: { bodyParser: { sizeLimit: '10mb' } } };

// ── XML helpers ───────────────────────────────────────────────────────────────

function xmlEscape(str) {
  if (!str) return '';
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

// Build a paragraph using an existing style from the template
function makePara(styleId, text, extraRpr) {
  const rpr = extraRpr || '';
  return `<w:p>
    <w:pPr><w:pStyle w:val="${styleId}"/></w:pPr>
    <w:r>${rpr}<w:t xml:space="preserve">${xmlEscape(text)}</w:t></w:r>
  </w:p>`;
}

function makeNormalPara(text, opts) {
  opts = opts || {};
  let rpr = '<w:rPr>';
  if (opts.bold) rpr += '<w:b/>';
  if (opts.italic) rpr += '<w:i/>';
  if (opts.color) rpr += `<w:color w:val="${opts.color}"/>`;
  if (opts.size) rpr += `<w:sz w:val="${opts.size}"/><w:szCs w:val="${opts.size}"/>`;
  rpr += '</w:rPr>';
  const jc = opts.justify ? '<w:jc w:val="both"/>' : '';
  return `<w:p>
    <w:pPr>${jc}</w:pPr>
    <w:r>${rpr}<w:t xml:space="preserve">${xmlEscape(text)}</w:t></w:r>
  </w:p>`;
}

function emptyPara() {
  return '<w:p><w:pPr></w:pPr></w:p>';
}

// ── Report parser ─────────────────────────────────────────────────────────────

const SECTION_KEYS = [
  'PERSÖNLICHE ANGABEN','PERSOENLICHE ANGABEN','PERSONAL DETAILS',
  'AUSBILDUNG UND QUALIFIKATIONEN','AUSBILDUNG','EDUCATION & QUALIFICATIONS','EDUCATION',
  'VERGÜTUNG UND VERFÜGBARKEIT','VERGUETUNG','COMPENSATION & AVAILABILITY','COMPENSATION',
  'KARRIERE ZUSAMMENFASSUNG','CAREER SUMMARY',
  'KANDIDATENBEWERTUNG','CANDIDATE ASSESSMENT','CANDIDATE EVALUATION',
  'FACHLICHES RESÜMEE','FACHLICHES RESUME','PROFESSIONAL SUMMARY',
  'BEWERTUNG','PERSONALITY','PERSÖNLICHKEIT',
  'BEWERBERMOTIVATION','MOTIVATION',
  'BERUFSERFAHRUNG','BERUFLICHER WERDEGANG','WORK EXPERIENCE','PROFESSIONAL EXPERIENCE',
  'ANMERKUNGEN ZUM WERDEGANG'
];

const SUB_KEYS = [
  'FACHLICHES RESÜMEE','FACHLICHES RESUME','PROFESSIONAL SUMMARY',
  'BEWERTUNG','PERSONALITY','PERSÖNLICHKEIT',
  'BEWERBERMOTIVATION','MOTIVATION','KANDIDATENMOTIVATION'
];

function parseReport(raw) {
  const lines = raw.split('\n');
  const result = [];
  let current = { key: 'HEADER', lines: [] };
  for (const line of lines) {
    const t = line.trim();
    const u = t.toUpperCase();
    const matched = SECTION_KEYS.find(k => u === k || u.startsWith(k + ':'));
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
  return SUB_KEYS.some(s => u === s || u.startsWith(s + ':'));
}

// ── Build document XML body ───────────────────────────────────────────────────

function buildBodyXml(reportText, candidateName, position, client, datum) {
  const sections = parseReport(reportText);
  const parts = [];

  // ── COVER ──
  parts.push(makePara('Titleheader', (candidateName || 'KANDIDAT').toUpperCase()));
  parts.push(makePara('Coverdoctitle', 'VERTRAULICHER KANDIDATENBERICHT'));
  if (position) parts.push(makePara('Coverdate', position.toUpperCase()));
  if (client && client !== 'Vertraulich') parts.push(makePara('Coverdate', client));
  parts.push(makePara('Coverdate', datum || ''));
  parts.push(emptyPara());

  // Confidentiality text
  parts.push(makeNormalPara(
    'Dieser Vertrauliche Bericht enthält zum Teil Informationen, die uns unter Zusicherung strengster Vertraulichkeit mitgeteilt wurden. Entsprechend unseren berufsethischen Prinzipien müssen wir Sie dazu verpflichten, nur einer begrenzten Auswahl von Personen, die sich direkt mit der Auswertung befassen, Einsicht in diese Berichte zu gewähren. Der Inhalt muss auch jeglichen Drittpersonen gegenüber geheim gehalten werden. Es dürfen keinerlei Referenzen ohne Zustimmung des Kandidaten oder unsererseits eingeholt werden.',
    { italic: true, justify: true, color: '595959', size: 18 }
  ));
  parts.push(emptyPara());

  // ── SECTIONS ──
  for (const section of sections) {
    if (section.key === 'HEADER') continue;
    const content = section.lines.map(l => l.trim()).filter(Boolean);
    if (!content.length) continue;

    const ku = section.key.toUpperCase();
    const isPersonal = ku.includes('PERSÖN') || ku.includes('PERSOEN') || ku.includes('PERSONAL');
    const isKandidaten = ku.includes('KANDIDATEN') || ku.includes('CANDIDATE ASSESSMENT') || ku.includes('CANDIDATE EVALUATION');
    const isExperience = ku.includes('BERUFSERFAHRUNG') || ku.includes('BERUFLICHER') || ku.includes('WORK EXPERIENCE') || ku.includes('PROFESSIONAL EXPERIENCE');

    // Section heading using berschrift2 style (exact navy, exact font)
    parts.push(makePara('berschrift2', section.key.toUpperCase()));
    parts.push(emptyPara());

    // PERSONAL DETAILS — two-column format using SPTBodytext66 for labels
    if (isPersonal) {
      for (const line of content) {
        if (line.includes(':')) {
          const idx = line.indexOf(':');
          const label = line.slice(0, idx).trim();
          const value = line.slice(idx + 1).trim();
          parts.push(makePara('SPTBodytext66', label));
          parts.push(makeNormalPara(value, { color: '262626', size: 20 }));
        } else {
          parts.push(makeNormalPara(line, { color: '262626', size: 20 }));
        }
      }
      parts.push(emptyPara());
      continue;
    }

    // KANDIDATENBEWERTUNG — sub-sections using Untertitel style
    if (isKandidaten) {
      let curSub = null;
      let subLines = [];
      const flush = () => {
        if (!curSub) return;
        parts.push(makePara('Untertitel', curSub.toUpperCase()));
        for (const l of subLines) {
          if (!l.trim()) continue;
          if (/^[-•]/.test(l)) {
            parts.push(makePara('Listenabsatz', l.replace(/^[-•]\s*/, '')));
          } else {
            parts.push(makeNormalPara(l, { justify: true, color: '262626', size: 22 }));
          }
        }
        parts.push(emptyPara());
        subLines = [];
      };
      for (const line of content) {
        if (isSubSection(line)) { flush(); curSub = line.trim(); }
        else subLines.push(line);
      }
      flush();
      continue;
    }

    // PROFESSIONAL EXPERIENCE — uses Amrop-header style for labels
    if (isExperience) {
      for (const line of content) {
        const isBullet = /^[-•]/.test(line);
        const isDateLine = /^\d{4}|^(Januar|Februar|März|April|Mai|Juni|Juli|August|September|Oktober|November|Dezember|January|February|March|April|May|June|July|August|September|October|November|December)/.test(line);
        const isCompanyInfo = /^\*/.test(line); // italic company description
        const isLabel = /^(Hauptverantwortlichkeiten|Key Achievements|Verantwortlichkeiten|Responsibilities|Main Responsibilities|Achievements|Company Information|Firmenbeschreibung):?$/i.test(line.trim());

        if (isDateLine && line.includes('|')) {
          // Career line: date | company | title
          const parts2 = line.split('|');
          parts.push(makePara('Amrop-header', parts2[0].trim()));
          if (parts2[1]) parts.push(makeNormalPara(parts2[1].trim().toUpperCase(), { bold: true, color: '262626', size: 22 }));
          if (parts2[2]) parts.push(makeNormalPara(parts2[2].trim(), { color: '262626', size: 22 }));
        } else if (isLabel) {
          parts.push(makePara('Amrop-header', line));
        } else if (isCompanyInfo) {
          parts.push(makePara('Listing1', line.replace(/^\*|\*$/g, '')));
        } else if (isBullet) {
          parts.push(makeNormalPara('– ' + line.replace(/^[-•]\s*/, ''), { color: '262626', size: 20 }));
        } else {
          parts.push(makeNormalPara(line, { color: '262626', size: 20 }));
        }
      }
      parts.push(emptyPara());
      continue;
    }

    // Default — normal paragraphs
    for (const line of content) {
      if (/^[-•]/.test(line)) {
        parts.push(makePara('Listenabsatz', line.replace(/^[-•]\s*/, '')));
      } else {
        parts.push(makeNormalPara(line, { justify: true, color: '262626', size: 20 }));
      }
    }
    parts.push(emptyPara());
  }

  // ── FOOTER ──
  parts.push(emptyPara());
  parts.push(makeNormalPara('Vorbereitet von: Dr. Sami Hamid  |  Managing Partner  |  Signium Austria', { bold: true, color: '102E66', size: 18 }));
  parts.push(makeNormalPara('t +43 664 4568862  |  sami.hamid@signium.com', { color: '595959', size: 17 }));

  return parts.join('\n');
}

// ── Handler ───────────────────────────────────────────────────────────────────

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method Not Allowed' });

  try {
    const { text, candidateName, position, client, datum } = req.body;
    if (!text) return res.status(400).json({ error: 'No text provided' });

    // Load the template
    const templatePath = path.join(process.cwd(), 'template.docx');
    const templateBuffer = fs.readFileSync(templatePath);
    const zip = await JSZip.loadAsync(templateBuffer);

    // Get the document.xml
    const docXmlRaw = await zip.file('word/document.xml').async('string');

    // Extract the body content wrapper
    const bodyStart = docXmlRaw.indexOf('<w:body>') + '<w:body>'.length;
    const bodyEnd = docXmlRaw.lastIndexOf('</w:body>');
    const beforeBody = docXmlRaw.substring(0, bodyStart);
    const afterBody = docXmlRaw.substring(bodyEnd); // includes </w:body></w:document>

    // Extract sectPr (page settings) from original - keep it
    const sectPrMatch = docXmlRaw.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/);
    const sectPr = sectPrMatch ? sectPrMatch[0] : '';

    // Build new body
    const newBodyContent = buildBodyXml(text, candidateName, position, client, datum);
    const newDocXml = beforeBody + '\n' + newBodyContent + '\n' + sectPr + '\n' + afterBody;

    // Update zip
    zip.file('word/document.xml', newDocXml);

    // Generate output
    const outputBuffer = await zip.generateAsync({
      type: 'nodebuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    });

    const safeName = (candidateName || 'Kandidat').replace(/\s+/g, '_');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${safeName}_Signium_Bericht.docx"`);
    res.send(outputBuffer);

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
}
