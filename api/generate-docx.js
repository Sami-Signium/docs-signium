const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

export const config = { api: { bodyParser: { sizeLimit: '10mb' } } };

function xe(str) {
  if (!str) return '';
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// Paragraph with named style
function sp(styleId, text, before, after, pageBreakBefore) {
  let ppr = `<w:pStyle w:val="${styleId}"/>`;
  if (pageBreakBefore) ppr += `<w:pageBreakBefore/>`;
  if (before !== undefined || after !== undefined) {
    const b = before !== undefined ? ` w:before="${before}"` : '';
    const a = after !== undefined ? ` w:after="${after}"` : '';
    ppr += `<w:spacing${b}${a}/>`;
  }
  if (!text && text !== 0) return `<w:p><w:pPr>${ppr}</w:pPr></w:p>`;
  return `<w:p><w:pPr>${ppr}</w:pPr><w:r><w:t xml:space="preserve">${xe(text)}</w:t></w:r></w:p>`;
}

// Normal paragraph (no named style)
function np(text, before, after, opts) {
  opts = opts || {};
  let ppr = '';
  if (before !== undefined || after !== undefined) {
    const b = before !== undefined ? ` w:before="${before}"` : '';
    const a = after !== undefined ? ` w:after="${after}"` : '';
    ppr += `<w:spacing${b}${a}/>`;
  }
  if (opts.jc) ppr += `<w:jc w:val="${opts.jc}"/>`;
  let rpr = '<w:rPr>';
  if (opts.bold) rpr += '<w:b/>';
  if (opts.italic) rpr += '<w:i/>';
  if (opts.color) rpr += `<w:color w:val="${opts.color}"/>`;
  // Always use 22 (11pt) for normal body text - matching template default
  const sz = opts.size || 22;
  rpr += `<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>`;
  rpr += '</w:rPr>';
  if (!text && text !== 0) return `<w:p><w:pPr>${ppr}</w:pPr></w:p>`;
  return `<w:p><w:pPr>${ppr}</w:pPr><w:r>${rpr}<w:t xml:space="preserve">${xe(text)}</w:t></w:r></w:p>`;
}

// Sections that need a page break before them
const PAGE_BREAK_SECTIONS = [
  'VERGÜTUNG', 'VERGUETUNG', 'COMPENSATION',
  'KARRIERE ZUSAMMENFASSUNG', 'CAREER SUMMARY',
  'KANDIDATENBEWERTUNG', 'FACHLICHES', 'BEWERTUNG', 'CANDIDATE',
  'BERUFSERFAHRUNG', 'BERUFLICHER', 'WORK EXPERIENCE', 'PROFESSIONAL EXPERIENCE'
];

const SECTION_KEYS = [
  'PERSÖNLICHE ANGABEN','PERSOENLICHE ANGABEN','PERSONAL DETAILS',
  'AUSBILDUNG UND QUALIFIKATIONEN','AUSBILDUNG','EDUCATION & QUALIFICATIONS','EDUCATION',
  'VERGÜTUNG UND VERFÜGBARKEIT','VERGUETUNG','COMPENSATION & AVAILABILITY','COMPENSATION',
  'KARRIERE ZUSAMMENFASSUNG','CAREER SUMMARY',
  'KANDIDATENBEWERTUNG','CANDIDATE ASSESSMENT','CANDIDATE EVALUATION',
  'FACHLICHES RESÜMEE','FACHLICHES RESUME','PROFESSIONAL SUMMARY',
  'BEWERTUNG','PERSONALITY','PERSÖNLICHKEIT',
  'BEWERBERMOTIVATION','MOTIVATION','KANDIDATENMOTIVATION',
  'BERUFSERFAHRUNG','BERUFLICHER WERDEGANG','WORK EXPERIENCE','PROFESSIONAL EXPERIENCE',
  'ANMERKUNGEN ZUM WERDEGANG'
];

const SUB_KEYS = [
  'FACHLICHES RESÜMEE','FACHLICHES RESUME','PROFESSIONAL SUMMARY',
  'BEWERTUNG','PERSONALITY','PERSÖNLICHKEIT',
  'BEWERBERMOTIVATION','MOTIVATION','KANDIDATENMOTIVATION'
];

function needsPageBreak(sectionKey) {
  const u = sectionKey.toUpperCase();
  return PAGE_BREAK_SECTIONS.some(k => u.startsWith(k));
}

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

function buildBodyXml(reportText, candidateName, position, client, datum) {
  const sections = parseReport(reportText);
  const parts = [];

  // ── COVER ──
  parts.push(sp('Titleheader', (candidateName || 'KANDIDAT').toUpperCase(), 120, 0));
  parts.push(sp('Coverdoctitle', 'VERTRAULICHER KANDIDATENBERICHT', 4080, 0));
  if (position) parts.push(sp('Coverdate', position.toUpperCase(), 720, 1000));
  if (client && client !== 'Vertraulich') parts.push(sp('Coverdate', client, 120, 120));
  parts.push(sp('Coverdate', datum || '', 120, 1000));

  // Confidentiality
  parts.push(np('Dieser Vertrauliche Bericht enthält zum Teil Informationen, die uns unter Zusicherung strengster Vertraulichkeit mitgeteilt wurden. Entsprechend unseren berufsethischen Prinzipien müssen wir Sie dazu verpflichten, nur einer begrenzten Auswahl von Personen, die sich direkt mit der Auswertung befassen, Einsicht in diese Berichte zu gewähren.', 120, undefined, { jc: 'both', italic: true, color: '595959', size: 18 }));
  parts.push(np('Der Inhalt muss auch jeglichen Drittpersonen gegenüber geheim gehalten werden. Es dürfen keinerlei Referenzen ohne Zustimmung des Kandidaten oder unsererseits eingeholt werden.', 120, undefined, { jc: 'both', italic: true, color: '595959', size: 18 }));
  parts.push(np('', 120));

  // ── SECTIONS ──
  for (const section of sections) {
    if (section.key === 'HEADER') continue;
    const content = section.lines.map(l => l.trim()).filter(Boolean);
    if (!content.length) continue;

    const ku = section.key.toUpperCase();
    const isPersonal = ku.includes('PERSÖN') || ku.includes('PERSOEN') || ku.includes('PERSONAL');
    const isKandidaten = ku.includes('KANDIDATEN') || ku.includes('CANDIDATE ASSESSMENT') || ku.includes('CANDIDATE EVALUATION');
    const isExperience = ku.includes('BERUFSERFAHRUNG') || ku.includes('BERUFLICHER') || ku.includes('WORK EXPERIENCE') || ku.includes('PROFESSIONAL EXPERIENCE');
    const isKarriere = ku.includes('KARRIERE') || ku.includes('CAREER SUMMARY');
    const pageBreak = needsPageBreak(section.key);

    // Section heading with optional page break
    parts.push(sp('berschrift2', section.key.toUpperCase(), 120, undefined, pageBreak));
    parts.push(np('', 120));

    // PERSONAL DETAILS
    if (isPersonal) {
      for (const line of content) {
        if (line.includes(':')) {
          const idx = line.indexOf(':');
          const label = line.slice(0, idx).trim();
          const value = line.slice(idx + 1).trim();
          parts.push(sp('SPTBodytext66', label));
          parts.push(np(value, undefined, undefined, { color: '262626' }));
        } else {
          parts.push(np(line, undefined, undefined, { color: '262626' }));
        }
      }
      parts.push(np('', 120));
      continue;
    }

    // KARRIERE ZUSAMMENFASSUNG
    if (isKarriere) {
      for (const line of content) {
        parts.push(np(line, 60, 120, { color: '262626', jc: 'both' }));
      }
      parts.push(np('', 120));
      continue;
    }

    // KANDIDATENBEWERTUNG with sub-sections
    if (isKandidaten) {
      let curSub = null;
      let subLines = [];
      const flush = () => {
        if (!curSub) return;
        parts.push(sp('Untertitel', curSub.toUpperCase(), 120, 120));
        for (const l of subLines) {
          if (!l.trim()) continue;
          if (/^[-•]/.test(l)) {
            parts.push(sp('Listenabsatz', l.replace(/^[-•]\s*/, '')));
          } else {
            parts.push(np(l, 120, undefined, { jc: 'both', color: '262626' }));
          }
        }
        parts.push(np('', 120));
        subLines = [];
      };
      for (const line of content) {
        if (isSubSection(line)) { flush(); curSub = line.trim(); }
        else subLines.push(line);
      }
      flush();
      continue;
    }

    // PROFESSIONAL EXPERIENCE — exact structure matching improved doc
    if (isExperience) {
      let i = 0;
      while (i < content.length) {
        const line = content[i];
        const isBullet = /^[-–•]/.test(line);
        const isCompany = /^\*/.test(line);
        const isLabel = /^(Hauptverantwortlichkeiten|Key Achievements|Verantwortlichkeiten|Responsibilities|Main Responsibilities|Achievements|Company Information|Firmenbeschreibung):?$/i.test(line.trim());
        const isDateHeader = /^\d{4}|^(Januar|Februar|März|April|Mai|Juni|Juli|August|September|Oktober|November|Dezember|January|February|March|April|May|June|July|August|September|October|November|December)/.test(line) && !isBullet;
        const hasCompanyInfo = line.includes(' - ') && isDateHeader;

        if (isDateHeader) {
          // Date + company on same line (e.g. "Oktober 2020 bis Januar 2026 - Constantia Flexibles")
          parts.push(sp('Amrop-header', line, 120, 120));
        } else if (isCompany) {
          parts.push(sp('Listing1', line.replace(/^\*|\*$/g, ''), 120));
        } else if (isLabel) {
          parts.push(sp('Amrop-header', line.replace(/:$/, '') + ':', 120, 0));
        } else if (isBullet) {
          parts.push(np(line.replace(/^[-–•]\s*/, ''), 60, 120, { color: '262626' }));
        } else if (line.trim()) {
          // Job title line — bold
          parts.push(np(line, 120, 120, { bold: true, color: '262626' }));
        }
        i++;
      }
      parts.push(np('', 120));
      continue;
    }

    // Default — education, compensation, etc.
    for (const line of content) {
      parts.push(np(line, 120, undefined, { color: '262626', jc: 'both' }));
    }
    parts.push(np('', 120));
  }

  // Footer
  parts.push(np('', 240));
  parts.push(np('Vorbereitet von: Dr. Sami Hamid  |  Managing Partner  |  Signium Austria', 120, 0, { bold: true, color: '102E66', size: 18 }));
  parts.push(np('t +43 664 4568862  |  sami.hamid@signium.com', 40, 0, { color: '595959', size: 17 }));

  return parts.join('\n');
}

function updateHeaders(zip, candidateName, position, client) {
  const headerFiles = ['word/header1.xml', 'word/header2.xml', 'word/header3.xml'];
  const promises = headerFiles.map(async (hf) => {
    const file = zip.file(hf);
    if (!file) return;
    let xml = await file.async('string');
    // Replace all occurrences of Quintin Stephen
    xml = xml.replace(/Quintin Stephen/g, xe(candidateName || ''));
    // Replace position
    xml = xml.replace(/Director of Identity &amp; Authentication/g, xe(position || ''));
    xml = xml.replace(/Director of Identity &amp;amp; Authentication/g, xe(position || ''));
    // Replace client/company
    xml = xml.replace(/Austriacard/g, xe(client && client !== 'Vertraulich' ? client : 'Confidential'));
    // Replace page number reference company
    xml = xml.replace(/AustriaCard Holdings[^<]*/g, xe(candidateName || ''));
    zip.file(hf, xml);
  });
  return Promise.all(promises);
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

    const templatePath = path.join(process.cwd(), 'template.docx');
    const templateBuffer = fs.readFileSync(templatePath);
    const zip = await JSZip.loadAsync(templateBuffer);

    // Update headers (fix Quintin Stephen issue)
    await updateHeaders(zip, candidateName, position, client);

    // Update document body
    const docXmlRaw = await zip.file('word/document.xml').async('string');
    const bodyStart = docXmlRaw.indexOf('<w:body>') + '<w:body>'.length;
    const bodyEnd = docXmlRaw.lastIndexOf('</w:body>');
    const beforeBody = docXmlRaw.substring(0, bodyStart);
    const afterBody = docXmlRaw.substring(bodyEnd);
    const sectPrMatch = docXmlRaw.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/);
    const sectPr = sectPrMatch ? sectPrMatch[0] : '';

    const newBodyContent = buildBodyXml(text, candidateName, position, client, datum);
    const newDocXml = beforeBody + '\n' + newBodyContent + '\n' + sectPr + '\n' + afterBody;
    zip.file('word/document.xml', newDocXml);

    const outputBuffer = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 6 } });
    const safeName = (candidateName || 'Kandidat').replace(/\s+/g, '_');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${safeName}_Signium_Bericht.docx"`);
    res.send(outputBuffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
}
