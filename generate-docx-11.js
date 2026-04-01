import { createRequire } from 'module';
const require = createRequire(import.meta.url);
import { readFileSync } from 'fs';
import { join } from 'path';
import JSZip from 'jszip';

export const config = { api: { bodyParser: { sizeLimit: '10mb' } } };

function xe(str) {
  if (!str) return '';
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function rpr(opts) {
  opts = opts || {};
  let x = '<w:rPr>';
  if (opts.major) {
    x += '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>';
  } else {
    x += '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/>';
  }
  if (opts.bold) x += '<w:b/>';
  if (opts.italic) x += '<w:i/>';
  const color = opts.color || '414042';
  x += `<w:color w:val="${color}"/>`;
  const sz = opts.sz || 22;
  x += `<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>`;
  x += '</w:rPr>';
  return x;
}

function run(text, opts) {
  return `<w:r>${rpr(opts)}<w:t xml:space="preserve">${xe(text)}</w:t></w:r>`;
}

function np(text, before, after, opts) {
  opts = opts || {};
  let ppr = '';
  if (before !== undefined || after !== undefined) {
    const b = before !== undefined ? ` w:before="${before}"` : '';
    const a = after !== undefined ? ` w:after="${after}"` : '';
    ppr += `<w:spacing${b}${a}/>`;
  }
  if (opts.jc) ppr += `<w:jc w:val="${opts.jc}"/>`;
  if (!text && text !== 0) return `<w:p><w:pPr>${ppr}</w:pPr></w:p>`;
  return `<w:p><w:pPr>${ppr}</w:pPr>${run(text, opts)}</w:p>`;
}

function personalRow(label, value) {
  const labelRpr = rpr({ sz: 22, color: '414042' });
  const valueRpr = rpr({ sz: 22, color: '262626' });
  // Pad label to fixed width with non-breaking spaces for alignment
  const padding = '\u00A0'.repeat(Math.max(1, 22 - label.length));
  return `<w:p>
    <w:pPr>
      <w:spacing w:before="80" w:after="80"/>
    </w:pPr>
    <w:r>${labelRpr}<w:t xml:space="preserve">${xe(label)}${xe(padding)}</w:t></w:r>
    <w:r>${valueRpr}<w:t xml:space="preserve">${xe(value)}</w:t></w:r>
  </w:p>`;
}

function hr() {
  return np('________________________________________________________________________________', 60, 60, { color: 'CCCCCC', sz: 16 });
}

function bullet(text) {
  const r = rpr({ sz: 24, color: '262626' });
  return `<w:p>
    <w:pPr>
      <w:pStyle w:val="Listenabsatz"/>
      <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
      <w:spacing w:before="60" w:after="120"/>
      ${r}
    </w:pPr>
    <w:r>${r}<w:t xml:space="preserve">${xe(text)}</w:t></w:r>
  </w:p>`;
}

function companyHeader(datePart, companyPart) {
  const dateRpr = rpr({ sz: 22, color: '414042' });
  const companyRpr = rpr({ sz: 24, color: '262626', bold: true });
  let ppr = `<w:pStyle w:val="Amrop-header"/>`;
  ppr += `<w:spacing w:before="120" w:after="120"/>`;
  ppr += rpr({ sz: 22, color: '414042' });
  if (companyPart) {
    return `<w:p><w:pPr>${ppr}</w:pPr>
      <w:r>${dateRpr}<w:t xml:space="preserve">${xe(datePart)}   </w:t></w:r>
      <w:r>${companyRpr}<w:t xml:space="preserve">${xe(companyPart)}</w:t></w:r>
    </w:p>`;
  }
  return `<w:p><w:pPr>${ppr}</w:pPr>
    <w:r>${dateRpr}<w:t xml:space="preserve">${xe(datePart)}</w:t></w:r>
  </w:p>`;
}

function sectionHead(text, pageBreak) {
  let ppr = `<w:pStyle w:val="berschrift2"/>`;
  if (pageBreak) ppr += `<w:pageBreakBefore/>`;
  ppr += `<w:spacing w:before="120"/>`;
  const r = rpr({ major: true, bold: true, sz: 28, color: '102E66' });
  ppr += r;
  return `<w:p><w:pPr>${ppr}</w:pPr><w:r>${r}<w:t xml:space="preserve">${xe(text)}</w:t></w:r></w:p>`;
}

// Sub-heading within a section (no page break)
function subHead(text) {
  const r = rpr({ major: true, bold: true, sz: 28, color: '102E66' });
  return `<w:p><w:pPr><w:spacing w:before="200" w:after="60"/>${r}</w:pPr><w:r>${r}<w:t xml:space="preserve">${xe(text)}</w:t></w:r></w:p>`;
}

const SECTION_KEYS = [
  'PERSOENLICHE ANGABEN','PERSONAL DETAILS',
  'PERSÖNLICHE ANGABEN',
  'AUSBILDUNG UND QUALIFIKATIONEN','AUSBILDUNG','EDUCATION & QUALIFICATIONS','EDUCATION',
  'VERGUETUNG UND VERFUEGBARKEIT','VERGUETUNG','VERGÜTUNG','COMPENSATION & AVAILABILITY','COMPENSATION',
  'KARRIERE ZUSAMMENFASSUNG','CAREER SUMMARY',
  'KANDIDATENBEWERTUNG','CANDIDATE ASSESSMENT','CANDIDATE EVALUATION',
  'FACHLICHES RESÜMEE','FACHLICHES RESUEMEE','PROFESSIONAL SUMMARY',
  'BEWERTUNG','PERSONALITY','PERSÖNLICHKEIT',
  'BEWERBERMOTIVATION','MOTIVATION','KANDIDATENMOTIVATION',
  'EMPFEHLUNG','RECOMMENDATION',
  'BERUFSERFAHRUNG','BERUFLICHER WERDEGANG','WORK EXPERIENCE','PROFESSIONAL EXPERIENCE',
  'ANMERKUNGEN ZUM WERDEGANG'
];

function needsPageBreak(key) {
  const u = key.toUpperCase();
  // Page 2: Personal details
  if (u.includes('PERS') && (u.includes('NLICH') || u.includes('ONAL'))) return true;
  // Page 3: Education
  if (u.includes('AUSBILDUNG') || u.includes('EDUCATION')) return true;
  // Page 4: Career summary
  if (u.includes('KARRIERE') || u.includes('CAREER SUMMARY')) return true;
  // Page 5: Candidate assessment section starts here
  if (u.includes('KANDIDATENBEWERTUNG') || u.includes('CANDIDATE ASSESSMENT') || u.includes('CANDIDATE EVALUATION')) return true;
  // FACHLICHES starts new page (in case KANDIDATENBEWERTUNG is skipped by AI)
  if (u.includes('FACHLICH') || u.includes('PROFESSIONAL SUMMARY')) return true;
  // Page 6: Professional experience
  if (u.includes('BERUFS') || u.includes('BERUFLICHER') || u.includes('WORK EXP') || u.includes('PROFESSIONAL EXP')) return true;
  return false;
}

function isSubSection(key) {
  const u = key.toUpperCase();
  // These are sub-sections within KANDIDATENBEWERTUNG — no page break
  return u.includes('BEWERTUNG') || u.includes('PERSONALITY') || u.includes('PERSÖNLICHKEIT') ||
         u.includes('BEWERBERMOTIVATION') || u.includes('MOTIVATION') ||
         u.includes('EMPFEHLUNG') || u.includes('RECOMMENDATION');
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

// Always render personal details section, even if empty
function renderPersonalSection(content) {
  const parts = [];
  const fields = ['Name','Geburtsdatum','Wohnort','Nationalität','Sprachen','Familienstand'];
  const parsed = {};
  for (const line of content) {
    if (line.includes(':')) {
      const idx = line.indexOf(':');
      parsed[line.slice(0, idx).trim()] = line.slice(idx + 1).trim();
    }
  }
  // Always show all standard fields
  for (const field of fields) {
    const value = parsed[field] || '';
    parts.push(personalRow(field, value));
  }
  // Show any extra fields from content
  for (const line of content) {
    if (line.includes(':')) {
      const idx = line.indexOf(':');
      const label = line.slice(0, idx).trim();
      if (!fields.includes(label)) {
        parts.push(personalRow(label, line.slice(idx + 1).trim()));
      }
    }
  }
  return parts;
}

function buildBodyXml(reportText, candidateName, position, client, datum) {
  const sections = parseReport(reportText);
  const parts = [];

  // Cover page
  parts.push(`<w:p><w:pPr><w:pStyle w:val="Titleheader"/><w:spacing w:before="120" w:after="0"/></w:pPr>
    ${run((candidateName || 'KANDIDAT').toUpperCase(), { major: true, bold: true, sz: 40, color: '808080' })}</w:p>`);
  parts.push(`<w:p><w:pPr><w:pStyle w:val="Coverdoctitle"/><w:spacing w:before="4080" w:after="0"/></w:pPr>
    ${run('VERTRAULICHER KANDIDATENBERICHT', { sz: 32, color: '102E66' })}</w:p>`);
  if (position) parts.push(`<w:p><w:pPr><w:pStyle w:val="Coverdate"/><w:spacing w:before="720" w:after="1000"/></w:pPr>
    ${run(position.toUpperCase(), { major: true, bold: true, sz: 28, color: '414042' })}</w:p>`);
  if (client && client !== 'Vertraulich') parts.push(`<w:p><w:pPr><w:pStyle w:val="Coverdate"/><w:spacing w:before="120" w:after="120"/></w:pPr>
    ${run(client, { major: true, bold: true, sz: 32, color: '414042' })}</w:p>`);
  parts.push(`<w:p><w:pPr><w:pStyle w:val="Coverdate"/><w:spacing w:before="120" w:after="1000"/></w:pPr>
    ${run(datum || '', { sz: 22, color: '414042' })}</w:p>`);
  parts.push(np('', 120));
  parts.push(np('Dieser Vertrauliche Bericht enthält zum Teil Informationen, die uns unter Zusicherung strengster Vertraulichkeit mitgeteilt wurden. Entsprechend unseren berufsethischen Prinzipien müssen wir Sie dazu verpflichten, nur einer begrenzten Auswahl von Personen Einsicht in diese Berichte zu gewähren.', 120, undefined, { italic: true, color: '595959', sz: 18, jc: 'both' }));
  parts.push(np('', 120));

  // Check if personal section exists in report
  const hasPersonal = sections.some(s => {
    const u = s.key.toUpperCase();
    return u.includes('PERS') && (u.includes('NLICH') || u.includes('ONAL'));
  });

  // If no personal section, add empty one
  if (!hasPersonal) {
    parts.push(sectionHead('PERSÖNLICHE ANGABEN', true));
    parts.push(hr());
    parts.push(...renderPersonalSection([]));
    parts.push(np('', 120));
  }

  for (const section of sections) {
    if (section.key === 'HEADER') continue;
    const content = section.lines.map(l => l.trim()).filter(Boolean);

    const ku = section.key.toUpperCase();
    const isPersonal = ku.includes('PERS') && (ku.includes('NLICH') || ku.includes('ONAL'));
    const isExperience = ku.includes('BERUFS') || ku.includes('BERUFLICHER') || ku.includes('WORK EXP') || ku.includes('PROFESSIONAL EXP');
    const isKarriere = ku.includes('KARRIERE') || ku.includes('CAREER SUMMARY');
    const isVergutung = ku.includes('VERG') || ku.includes('COMPENSATION');
    const isKandidatenBewertung = ku.includes('KANDIDATENBEWERTUNG') || ku.includes('CANDIDATE ASSESSMENT') || ku.includes('CANDIDATE EVALUATION');
    const isSubSec = isSubSection(section.key);
    const pageBreak = needsPageBreak(section.key);

    // Sub-sections get sub-heading, main sections get full heading
    if (isSubSec) {
      if (content.length === 0) continue; // skip empty sub-sections
      parts.push(subHead(section.key.toUpperCase()));
      parts.push(hr());
    } else {
      parts.push(sectionHead(section.key.toUpperCase(), pageBreak));
      parts.push(hr());
    }

    if (isPersonal) {
      parts.push(...renderPersonalSection(content));
      parts.push(np('', 120));
      continue;
    }

    if (isVergutung) {
      if (content.length === 0) { parts.push(np('', 120)); continue; }
      for (const line of content) {
        if (line.includes(':')) {
          const idx = line.indexOf(':');
          const label = line.slice(0, idx).trim();
          const value = line.slice(idx + 1).trim();
          const labelRpr = rpr({ bold: true, sz: 22, color: '414042' });
          const valueRpr = rpr({ sz: 22, color: '262626' });
          parts.push(`<w:p><w:pPr><w:spacing w:before="140" w:after="140"/></w:pPr>
            <w:r>${labelRpr}<w:t>${xe(label)}</w:t><w:tab/><w:tab/><w:tab/><w:tab/></w:r>
            <w:r>${valueRpr}<w:t xml:space="preserve">${xe(value)}</w:t></w:r>
          </w:p>`);
        } else { parts.push(np(line, 140, 140, { bold: true, sz: 22 })); }
      }
      parts.push(np('', 120));
      continue;
    }

    if (isKarriere) {
      if (content.length === 0) { parts.push(np('', 120)); continue; }
      for (const line of content) {
        const pipeParts = line.split('|').map(s => s.trim());
        if (pipeParts.length >= 2) {
          const dateRpr = rpr({ sz: 22, color: '414042' });
          const compRpr = rpr({ bold: true, sz: 22, color: '262626' });
          const titleRpr = rpr({ sz: 22, color: '262626' });
          const datePadded = pipeParts[0] + '\u00A0'.repeat(Math.max(1, 14 - pipeParts[0].length));
          parts.push(`<w:p><w:pPr>
            <w:spacing w:before="120" w:after="120"/>
          </w:pPr>
            <w:r>${dateRpr}<w:t xml:space="preserve">${xe(datePadded)}</w:t></w:r>
            <w:r>${compRpr}<w:t xml:space="preserve">${xe(pipeParts[1])}</w:t></w:r>
            ${pipeParts[2] ? `<w:r>${titleRpr}<w:t xml:space="preserve">  |  ${xe(pipeParts[2])}</w:t></w:r>` : ''}
          </w:p>`);
        } else { parts.push(np(line, 120, 120, { bold: true, sz: 22, color: '262626' })); }
      }
      parts.push(np('', 120));
      continue;
    }

    // KANDIDATENBEWERTUNG is just a title page — no content needed
    if (isKandidatenBewertung) {
      parts.push(np('', 120));
      continue;
    }

    // Sub-sections and other text sections
    if (content.length === 0) { parts.push(np('', 120)); continue; }

    if (isSubSec || isKandidatenBewertung) {
      for (const line of content) {
        if (!line.trim()) continue;
        if (/^[-•]/.test(line)) { parts.push(bullet(line.replace(/^[-•]\s*/, ''))); }
        else { parts.push(np(line, 160, 160, { jc: 'both', sz: 22, color: '262626' })); }
      }
      parts.push(np('', 120));
      continue;
    }

    if (isExperience) {
      let firstCompany = true;
      let i = 0;
      while (i < content.length) {
        const line = content[i];
        const isBullet = /^[-\u2013\u2022]/.test(line);
        const isCompanyDesc = /^\*/.test(line);
        // Date pattern - can be standalone date line
        const isDateOnly = /^(seit\s|ab\s|since\s)?\d{4}\s*[-–]\s*\d{4}$|^(seit|ab|since)\s+\d{4}$/.test(line.trim()) && !isBullet;
        // Full date+company on same line (e.g. "2019: Firma" or "2019-2021: Firma")
        const isDateWithCompany = /^(seit\s|ab\s)?\d{4}/.test(line) && (line.includes(': ') || (line.match(/\d{4}/) && line.length > 15 && !line.match(/^\d{4}\s*[-–]\s*\d{4}$/))) && !isBullet;
        // Company name line (all caps or title after date line)
        const nextLine = content[i + 1] || '';
        const isDateFollowedByCompany = isDateOnly && nextLine.length > 0 && !nextLine.startsWith('*') && !nextLine.startsWith('-');

        if (isDateWithCompany) {
          // Date and company on same line: "2019: Firma" or "seit 2021 - Firma"
          if (!firstCompany) parts.push(hr());
          firstCompany = false;
          let datePart = line, companyPart = '';
          const colonIdx = line.indexOf(': ');
          const dashIdx = line.indexOf(' - ');
          const dashIdx2 = line.indexOf(' – ');
          if (colonIdx > 4) { datePart = line.slice(0, colonIdx); companyPart = line.slice(colonIdx + 2); }
          else if (dashIdx > 4) { datePart = line.slice(0, dashIdx); companyPart = line.slice(dashIdx + 3); }
          else if (dashIdx2 > 4) { datePart = line.slice(0, dashIdx2); companyPart = line.slice(dashIdx2 + 3); }
          parts.push(companyHeader(datePart, companyPart));
          i++;
        } else if (isDateFollowedByCompany) {
          // Date on its own line, company name on next line
          if (!firstCompany) parts.push(hr());
          firstCompany = false;
          const companyName = nextLine.trim();
          parts.push(companyHeader(line, companyName));
          i += 2; // skip date line AND company name line
        } else if (isCompanyDesc) {
          const r = rpr({ sz: 22, color: '595959', italic: true });
          parts.push(`<w:p><w:pPr><w:pStyle w:val="Listing1"/><w:spacing w:before="60" w:after="60"/>${r}</w:pPr>
            <w:r>${r}<w:t xml:space="preserve">${xe(line.replace(/^\*|\*$/g, ''))}</w:t></w:r></w:p>`);
          parts.push(hr());
          i++;
        } else if (isBullet) {
          parts.push(bullet(line.replace(/^[-\u2013\u2022]\s*/, '')));
          i++;
        } else if (line.trim()) {
          // Job title or other text
          parts.push(np(line, 120, 80, { bold: true, sz: 24, color: '262626' }));
          i++;
        } else {
          i++;
        }
      }
      parts.push(np('', 120));
      continue;
    }

    for (const line of content) {
      parts.push(np(line, 120, 80, { sz: 22, color: '262626' }));
    }
    parts.push(np('', 120));
  }

  parts.push(np('', 240));
  parts.push(np('Vorbereitet von: Dr. Sami Hamid  |  Managing Partner  |  Signium Austria', 120, 0, { bold: true, color: '102E66', sz: 18 }));
  parts.push(np('t +43 664 4568862  |  sami.hamid@signium.com', 40, 0, { color: '595959', sz: 17 }));

  return parts.join('\n');
}

async function updateHeaders(zip, candidateName, position, client) {
  for (const hf of ['word/header1.xml','word/header2.xml','word/header3.xml']) {
    const file = zip.file(hf);
    if (!file) continue;
    let xml = await file.async('string');
    xml = xml.replace(/Quintin Stephen/g, xe(candidateName || ''));
    xml = xml.replace(/Director of Identity &amp; Authentication/g, xe(position || ''));
    xml = xml.replace(/Austriacard/g, xe(client && client !== 'Vertraulich' ? client : 'Confidential'));
    xml = xml.replace(/AustriaCard Holdings[^<]*/g, xe(candidateName || ''));
    zip.file(hf, xml);
  }
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

    const templatePath = join(process.cwd(), 'template.docx');
    const templateBuffer = readFileSync(templatePath);
    const zip = await JSZip.loadAsync(templateBuffer);

    await updateHeaders(zip, candidateName, position, client);

    const docXmlRaw = await zip.file('word/document.xml').async('string');
    const bodyStart = docXmlRaw.indexOf('<w:body>') + '<w:body>'.length;
    const bodyEnd = docXmlRaw.lastIndexOf('</w:body>');
    const sectPrMatch = docXmlRaw.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/);
    const sectPr = sectPrMatch ? sectPrMatch[0] : '';

    const newDocXml = docXmlRaw.substring(0, bodyStart) + '\n' +
      buildBodyXml(text, candidateName, position, client, datum) + '\n' +
      sectPr + '\n' + docXmlRaw.substring(bodyEnd);

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
