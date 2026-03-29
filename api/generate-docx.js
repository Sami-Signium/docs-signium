const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

export const config = { api: { bodyParser: { sizeLimit: '10mb' } } };

function xmlEscape(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// Build paragraph with exact style + optional spacing override
function p(styleId, text, spacingBefore, spacingAfter, extraOpts) {
  extraOpts = extraOpts || {};
  let pPr = `<w:pStyle w:val="${styleId}"/>`;
  if (spacingBefore !== null || spacingAfter !== null) {
    const b = spacingBefore !== null ? ` w:before="${spacingBefore}"` : '';
    const a = spacingAfter !== null ? ` w:after="${spacingAfter}"` : '';
    pPr += `<w:spacing${b}${a}/>`;
  }
  if (extraOpts.jc) pPr += `<w:jc w:val="${extraOpts.jc}"/>`;

  let rPr = '';
  if (extraOpts.bold || extraOpts.italic || extraOpts.color || extraOpts.size) {
    rPr = '<w:rPr>';
    if (extraOpts.bold) rPr += '<w:b/>';
    if (extraOpts.italic) rPr += '<w:i/>';
    if (extraOpts.color) rPr += `<w:color w:val="${extraOpts.color}"/>`;
    if (extraOpts.size) rPr += `<w:sz w:val="${extraOpts.size}"/><w:szCs w:val="${extraOpts.size}"/>`;
    rPr += '</w:rPr>';
  }

  if (!text && text !== 0) {
    return `<w:p><w:pPr>${pPr}</w:pPr></w:p>`;
  }
  return `<w:p><w:pPr>${pPr}</w:pPr><w:r>${rPr}<w:t xml:space="preserve">${xmlEscape(text)}</w:t></w:r></w:p>`;
}

// Normal paragraph (no style) with optional formatting
function np(text, spacingBefore, spacingAfter, opts) {
  opts = opts || {};
  let pPr = '';
  if (spacingBefore !== null || spacingAfter !== null) {
    const b = spacingBefore !== null ? ` w:before="${spacingBefore}"` : '';
    const a = spacingAfter !== null ? ` w:after="${spacingAfter}"` : '';
    pPr += `<w:spacing${b}${a}/>`;
  }
  if (opts.jc) pPr += `<w:jc w:val="${opts.jc}"/>`;

  let rPr = '<w:rPr>';
  if (opts.bold) rPr += '<w:b/>';
  if (opts.italic) rPr += '<w:i/>';
  if (opts.color) rPr += `<w:color w:val="${opts.color}"/>`;
  if (opts.size) rPr += `<w:sz w:val="${opts.size}"/><w:szCs w:val="${opts.size}"/>`;
  rPr += '</w:rPr>';

  if (!text && text !== 0) {
    return `<w:p><w:pPr>${pPr}</w:pPr></w:p>`;
  }
  return `<w:p><w:pPr>${pPr}</w:pPr><w:r>${rPr}<w:t xml:space="preserve">${xmlEscape(text)}</w:t></w:r></w:p>`;
}

const SECTION_KEYS = [
  'PERSÖNLICHE ANGABEN','PERSOENLICHE ANGABEN','PERSONAL DETAILS','PERSONAL DATA',
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
  // Titleheader: before=120, after=0
  parts.push(p('Titleheader', (candidateName || 'KANDIDAT').toUpperCase(), 120, 0));
  // Coverdoctitle: before=4080 (large gap, like original)
  parts.push(p('Coverdoctitle', 'VERTRAULICHER KANDIDATENBERICHT', 4080, 0));
  // Position as Coverdate: before=720, after=1000
  if (position) parts.push(p('Coverdate', position.toUpperCase(), 720, 1000));
  // Client
  if (client && client !== 'Vertraulich') parts.push(p('Coverdate', client, 120, 120));
  // Date
  parts.push(p('Coverdate', datum || '', 120, 1000));

  // Confidentiality paragraphs - before=120, justified
  parts.push(np('Dieser Vertrauliche Bericht enthält zum Teil Informationen, die uns unter Zusicherung strengster Vertraulichkeit mitgeteilt wurden. Entsprechend unseren berufsethischen Prinzipien müssen wir Sie dazu verpflichten, nur einer begrenzten Auswahl von Personen, die sich direkt mit der Auswertung befassen, Einsicht in diese Berichte zu gewähren.', 120, null, { jc: 'both', italic: true, color: '595959', size: 18 }));
  parts.push(np('Der Inhalt muss auch jeglichen Drittpersonen gegenüber geheim gehalten werden. Es dürfen keinerlei Referenzen ohne Zustimmung des Kandidaten oder unsererseits eingeholt werden.', 120, null, { jc: 'both', italic: true, color: '595959', size: 18 }));
  parts.push(np('', 120, null));

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

    // Section heading: berschrift2, before=120
    parts.push(p('berschrift2', section.key.toUpperCase(), 120, null));
    parts.push(np('', 120, null));

    // PERSONAL DETAILS
    if (isPersonal) {
      for (const line of content) {
        if (line.includes(':')) {
          const idx = line.indexOf(':');
          const label = line.slice(0, idx).trim();
          const value = line.slice(idx + 1).trim();
          parts.push(p('SPTBodytext66', label, null, null));
          parts.push(np(value, null, null, { color: '262626', size: 20 }));
        } else {
          parts.push(np(line, null, null, { color: '262626', size: 20 }));
        }
      }
      parts.push(np('', 120, null));
      continue;
    }

    // KARRIERE ZUSAMMENFASSUNG
    if (isKarriere) {
      for (const line of content) {
        parts.push(np(line, 60, 120, { color: '262626', size: 20, jc: 'both' }));
      }
      parts.push(np('', 120, null));
      continue;
    }

    // KANDIDATENBEWERTUNG with sub-sections
    if (isKandidaten) {
      let curSub = null;
      let subLines = [];
      const flush = () => {
        if (!curSub) return;
        // Untertitel: before=120, after=120
        parts.push(p('Untertitel', curSub.toUpperCase(), 120, 120));
        for (const l of subLines) {
          if (!l.trim()) continue;
          if (/^[-•]/.test(l)) {
            parts.push(p('Listenabsatz', l.replace(/^[-•]\s*/, ''), null, null));
          } else {
            parts.push(np(l, 120, null, { jc: 'both', color: '262626', size: 20 }));
          }
        }
        parts.push(np('', 120, null));
        subLines = [];
      };
      for (const line of content) {
        if (isSubSection(line)) { flush(); curSub = line.trim(); }
        else subLines.push(line);
      }
      flush();
      continue;
    }

    // PROFESSIONAL EXPERIENCE — exact structure from original
    if (isExperience) {
      let i = 0;
      while (i < content.length) {
        const line = content[i];
        // Date line followed by company name
        const isDateLine = /^\d{4}|^(Januar|Februar|März|April|Mai|Juni|Juli|August|September|Oktober|November|Dezember|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)/.test(line);

        if (isDateLine && !line.includes('|')) {
          // Amrop-header for date: before=0, after=120 (first) or before=120, after=120 (subsequent)
          const isFirst = parts.filter(x => x.includes('Amrop-header')).length === 0;
          parts.push(p('Amrop-header', line, isFirst ? 0 : 120, 120));
          // Next line is company + title
          if (i + 1 < content.length && !content[i+1].includes('Hauptverantwortlichkeiten') && !content[i+1].includes('Key Achievements') && !content[i+1].includes('*')) {
            i++;
            parts.push(np(content[i].toUpperCase(), null, 120, { bold: true, color: '262626', size: 22 }));
          }
        } else if (line.includes('|') && /\d{4}/.test(line)) {
          // Career line with pipe: split into date | company | title
          const pipeparts = line.split('|').map(x => x.trim());
          parts.push(p('Amrop-header', pipeparts[0], 120, 120));
          if (pipeparts[1]) parts.push(np(pipeparts[1].toUpperCase(), null, 120, { bold: true, color: '262626', size: 22 }));
          if (pipeparts[2]) parts.push(np(pipeparts[2], 120, 120, { color: '262626', size: 20 }));
        } else if (/^\*/.test(line)) {
          // Company description: Listing1, before=120
          parts.push(p('Listing1', line.replace(/^\*|\*$/g, ''), 120, null));
        } else if (/^(Hauptverantwortlichkeiten|Key Achievements|Verantwortlichkeiten|Responsibilities|Main Responsibilities|Achievements):?$/i.test(line.trim())) {
          // Label: Amrop-header, before=120, after=0
          parts.push(p('Amrop-header', line.replace(/:$/, '') + ':', 120, 0));
        } else if (/^[-–•]/.test(line)) {
          // Responsibility/achievement bullet: Normal, before=60, after=120
          parts.push(np(line.replace(/^[-–•]\s*/, ''), 60, 120, { color: '262626', size: 20 }));
        } else if (line.trim()) {
          parts.push(np(line, 120, 120, { color: '262626', size: 20 }));
        }
        i++;
      }
      parts.push(np('', 120, null));
      continue;
    }

    // Default
    for (const line of content) {
      parts.push(np(line, 120, null, { jc: 'both', color: '262626', size: 20 }));
    }
    parts.push(np('', 120, null));
  }

  // Footer
  parts.push(np('', 240, null));
  parts.push(np('Vorbereitet von: Dr. Sami Hamid  |  Managing Partner  |  Signium Austria', 120, 0, { bold: true, color: '102E66', size: 18 }));
  parts.push(np('t +43 664 4568862  |  sami.hamid@signium.com', 40, 0, { color: '595959', size: 17 }));

  return parts.join('\n');
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
