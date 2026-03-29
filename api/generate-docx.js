const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

export const config = { api: { bodyParser: { sizeLimit: '10mb' } } };

function xe(str) {
  if (!str) return '';
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// Theme font rPr — matches exactly what revised doc uses
// minorHAnsi = Calibri Light, majorHAnsi = Calibri
function rpr(opts) {
  opts = opts || {};
  let x = '<w:rPr>';
  // Font: use theme fonts exactly like revised doc
  if (opts.major) {
    x += '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>';
  } else {
    x += '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/>';
  }
  if (opts.bold) x += '<w:b/>';
  if (opts.italic) x += '<w:i/>';
  // Color: explicit dark gray unless specified
  const color = opts.color || '414042';
  x += `<w:color w:val="${color}"/>`;
  const sz = opts.sz || 22;
  x += `<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>`;
  x += '</w:rPr>';
  return x;
}

// Run element
function run(text, opts) {
  return `<w:r>${rpr(opts)}<w:t xml:space="preserve">${xe(text)}</w:t></w:r>`;
}

// Tab run
function tabrun(opts) {
  return `<w:r>${rpr(opts)}<w:tab/></w:r>`;
}

// Named style paragraph
function sp(styleId, text, before, after, opts) {
  opts = opts || {};
  let ppr = `<w:pStyle w:val="${styleId}"/>`;
  if (opts.pageBreak) ppr += `<w:pageBreakBefore/>`;
  if (before !== undefined || after !== undefined) {
    const b = before !== undefined ? ` w:before="${before}"` : '';
    const a = after !== undefined ? ` w:after="${after}"` : '';
    ppr += `<w:spacing${b}${a}/>`;
  }
  // pPr rPr to override style defaults
  ppr += rpr(opts).replace('<w:rPr>', '<w:rPr>'); // same rpr in pPr
  if (!text && text !== 0) return `<w:p><w:pPr>${ppr}</w:pPr></w:p>`;
  return `<w:p><w:pPr>${ppr}</w:pPr>${run(text, opts)}</w:p>`;
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
  if (!text && text !== 0) return `<w:p><w:pPr>${ppr}</w:pPr></w:p>`;
  return `<w:p><w:pPr>${ppr}</w:pPr>${run(text, opts)}</w:p>`;
}

// Personal details row: Label [4 tabs] Value — exact structure from revised doc
function personalRow(label, value) {
  const labelRpr = rpr({ sz: 22, color: '414042' });
  const valueRpr = rpr({ sz: 22, color: '262626' });
  return `<w:p>
    <w:pPr><w:pStyle w:val="SPTBodytext66"/>${rpr({ sz: 22 })}</w:pPr>
    <w:r>${labelRpr}<w:t>${xe(label)}</w:t><w:tab/><w:tab/><w:tab/><w:tab/></w:r>
    <w:r>${valueRpr}<w:t xml:space="preserve">${xe(value)}</w:t></w:r>
  </w:p>`;
}

// Horizontal rule — thin gray line
function hr() {
  return np('________________________________________________________________________________', 60, 60, { color: 'CCCCCC', sz: 16 });
}

// Listenabsatz bullet — exact structure from revised doc (with numPr)
function bullet(text, before, after) {
  const spacing = `<w:spacing w:before="${before||60}" w:after="${after||120}"/>`;
  const r = rpr({ sz: 24, color: '262626' });
  return `<w:p>
    <w:pPr>
      <w:pStyle w:val="Listenabsatz"/>
      <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
      ${spacing}
      ${rpr({ sz: 24, color: '262626' })}
    </w:pPr>
    <w:r>${r}<w:t xml:space="preserve">${xe(text)}</w:t></w:r>
  </w:p>`;
}

// Company header: "Date: Company" — date auto color, company bold larger
function companyHeader(datePart, companyPart, before, after) {
  const dateRpr = rpr({ sz: 22, color: '414042' });
  const companyRpr = rpr({ sz: 28, color: '262626', bold: true });
  let ppr = `<w:pStyle w:val="Amrop-header"/>`;
  ppr += `<w:spacing w:before="${before||120}" w:after="${after||120}"/>`;
  ppr += rpr({ sz: 22, color: '414042' }); // pPr rPr override petrol
  if (companyPart) {
    return `<w:p><w:pPr>${ppr}</w:pPr>
      <w:r>${dateRpr}<w:t xml:space="preserve">${xe(datePart)}: </w:t><w:tab/></w:r>
      <w:r>${companyRpr}<w:t xml:space="preserve">${xe(companyPart)}</w:t></w:r>
    </w:p>`;
  }
  return `<w:p><w:pPr>${ppr}</w:pPr>
    <w:r>${dateRpr}<w:t xml:space="preserve">${xe(datePart)}</w:t></w:r>
  </w:p>`;
}

// Section heading
function sectionHead(text, pageBreak) {
  let ppr = `<w:pStyle w:val="berschrift2"/>`;
  if (pageBreak) ppr += `<w:pageBreakBefore/>`;
  ppr += `<w:spacing w:before="120"/>`;
  const r = rpr({ major: true, bold: true, sz: 28, color: '102E66' });
  ppr += r; // override in pPr too
  return `<w:p><w:pPr>${ppr}</w:pPr><w:r>${r}<w:t xml:space="preserve">${xe(text)}</w:t></w:r></w:p>`;
}

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

function needsPageBreak(key) {
  const u = key.toUpperCase();
  if (u.includes('PERS') && (u.includes('NLICH') || u.includes('ONAL'))) return true;
  // AUSBILDUNG bleibt auf der gleichen Seite wie PERSÖNLICHE ANGABEN
  if (u.includes('VERG') || u.includes('COMPENSATION')) return true;
  if (u.includes('KARRIERE') || u.includes('CAREER')) return true;
  if (u.includes('FACHLICH') || u.includes('PROFESSIONAL SUMMARY')) return true;
  if (u.includes('BEWERTUNG') || u.includes('PERSONALITY')) return true;
  if (u.includes('BERUFS') || u.includes('BERUFLICHER') || u.includes('WORK EXP') || u.includes('PROFESSIONAL EXP')) return true;
  return false;
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

function buildBodyXml(reportText, candidateName, position, client, datum) {
  const sections = parseReport(reportText);
  const parts = [];

  // ── COVER ──
  parts.push(`<w:p><w:pPr><w:pStyle w:val="Titleheader"/><w:spacing w:before="120" w:after="0"/></w:pPr>
    ${run((candidateName || 'KANDIDAT').toUpperCase(), { major: true, bold: true, sz: 52, color: '414042' })}</w:p>`);

  parts.push(`<w:p><w:pPr><w:pStyle w:val="Coverdoctitle"/><w:spacing w:before="4080" w:after="0"/></w:pPr>
    ${run('VERTRAULICHER KANDIDATENBERICHT', { sz: 32, color: '102E66' })}</w:p>`);

  if (position) parts.push(`<w:p><w:pPr><w:pStyle w:val="Coverdate"/><w:spacing w:before="720" w:after="1000"/></w:pPr>
    ${run(position.toUpperCase(), { major: true, bold: true, sz: 28, color: '414042' })}</w:p>`);

  if (client && client !== 'Vertraulich') parts.push(`<w:p><w:pPr><w:pStyle w:val="Coverdate"/><w:spacing w:before="120" w:after="120"/></w:pPr>
    ${run(client, { major: true, bold: true, sz: 32, color: '414042' })}</w:p>`);

  parts.push(`<w:p><w:pPr><w:pStyle w:val="Coverdate"/><w:spacing w:before="120" w:after="1000"/></w:pPr>
    ${run(datum || '', { sz: 22, color: '414042' })}</w:p>`);

  parts.push(np('', 120));
  parts.push(np('Dieser Vertrauliche Bericht enthält zum Teil Informationen, die uns unter Zusicherung strengster Vertraulichkeit mitgeteilt wurden. Entsprechend unseren berufsethischen Prinzipien müssen wir Sie dazu verpflichten, nur einer begrenzten Auswahl von Personen, die sich direkt mit der Auswertung befassen, Einsicht in diese Berichte zu gewähren.', 120, undefined, { italic: true, color: '595959', sz: 18, jc: 'both' }));
  parts.push(np('Der Inhalt muss auch jeglichen Drittpersonen gegenüber geheim gehalten werden. Es dürfen keinerlei Referenzen ohne Zustimmung des Kandidaten oder unsererseits eingeholt werden.', 120, undefined, { italic: true, color: '595959', sz: 18, jc: 'both' }));
  parts.push(np('', 120));

  // ── SECTIONS ──
  for (const section of sections) {
    if (section.key === 'HEADER') continue;
    const content = section.lines.map(l => l.trim()).filter(Boolean);
    if (!content.length) continue;

    const ku = section.key.toUpperCase();
    const isPersonal = ku.includes('PERS') && (ku.includes('NLICH') || ku.includes('ONAL'));
    const isExperience = ku.includes('BERUFS') || ku.includes('BERUFLICHER') || ku.includes('WORK EXP') || ku.includes('PROFESSIONAL EXP');
    const isKarriere = ku.includes('KARRIERE') || ku.includes('CAREER');
    const isVergütung = ku.includes('VERG') || ku.includes('COMPENSATION');
    const isKandidaten = ku.includes('FACHLICH') || ku.includes('BEWERTUNG') || ku.includes('PERSONALITY') || ku.includes('KANDIDATEN') || ku.includes('MOTIVATION');
    const pageBreak = needsPageBreak(section.key);

    parts.push(sectionHead(section.key.toUpperCase(), pageBreak));
    parts.push(hr());

    // PERSONAL DETAILS
    if (isPersonal) {
      for (const line of content) {
        if (line.includes(':')) {
          const idx = line.indexOf(':');
          parts.push(personalRow(line.slice(0, idx).trim(), line.slice(idx + 1).trim()));
        } else {
          parts.push(personalRow(line, ''));
        }
      }
      parts.push(np('', 120));
      continue;
    }

    // VERGÜTUNG — Label [tab] Wert, bold label
    if (isVergütung) {
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
        } else {
          parts.push(np(line, 140, 140, { bold: true, sz: 22 }));
        }
      }
      parts.push(np('', 120));
      continue;
    }

    // KARRIERE — Datum [tab] Firma [tab] Titel
    if (isKarriere) {
      for (const line of content) {
        const pipeParts = line.split('|').map(s => s.trim());
        if (pipeParts.length >= 2) {
          const dateRpr = rpr({ sz: 22, color: '414042' });
          const compRpr = rpr({ bold: true, sz: 22, color: '262626' });
          const titleRpr = rpr({ sz: 22, color: '262626' });
          parts.push(`<w:p><w:pPr><w:spacing w:before="120" w:after="120"/></w:pPr>
            <w:r>${dateRpr}<w:t>${xe(pipeParts[0])}</w:t><w:tab/><w:tab/></w:r>
            <w:r>${compRpr}<w:t xml:space="preserve">${xe(pipeParts[1])}</w:t></w:r>
            ${pipeParts[2] ? `<w:r>${titleRpr}<w:t xml:space="preserve">  |  ${xe(pipeParts[2])}</w:t></w:r>` : ''}
          </w:p>`);
        } else {
          parts.push(np(line, 120, 120, { bold: true, sz: 22, color: '262626' }));
        }
      }
      parts.push(np('', 120));
      continue;
    }

    // FACHLICHES / BEWERTUNG
    if (isKandidaten) {
      for (const line of content) {
        if (!line.trim()) continue;
        if (/^[-•]/.test(line)) {
          parts.push(bullet(line.replace(/^[-•]\s*/, '')));
        } else {
          parts.push(np(line, 160, 160, { jc: 'both', sz: 22, color: '262626' }));
        }
      }
      parts.push(np('', 120));
      continue;
    }

    // PROFESSIONAL EXPERIENCE
    if (isExperience) {
      let firstCompany = true;
      let i = 0;
      while (i < content.length) {
        const line = content[i];
        const isBullet = /^[-–•]/.test(line);
        const isCompanyDesc = /^\*/.test(line);
        const isDateHeader = /^(Jan|Feb|Mär|Apr|Mai|Jun|Jul|Aug|Sep|Okt|Nov|Dez|Oct|Mar|Januar|Februar|März|April|Juni|Juli|August|September|Oktober|November|Dezember|January|February|March|April|May|June|July|August|September|October|November|December|\d{2}\/\d{4}|\d{4})/.test(line) && !isBullet;

        if (isDateHeader) {
          if (!firstCompany) parts.push(hr());
          firstCompany = false;
          // Split on colon or " - " to get date vs company
          let datePart = line, companyPart = '';
          const colonIdx = line.indexOf(': ');
          const dashIdx = line.indexOf(' - ');
          if (colonIdx > 4) { datePart = line.slice(0, colonIdx); companyPart = line.slice(colonIdx + 2); }
          else if (dashIdx > 4) { datePart = line.slice(0, dashIdx); companyPart = line.slice(dashIdx + 3); }
          parts.push(companyHeader(datePart, companyPart));
        } else if (isCompanyDesc) {
          // Listing1 — company description italic
          const r = rpr({ sz: 22, color: '595959', italic: true });
          parts.push(`<w:p><w:pPr><w:pStyle w:val="Listing1"/><w:spacing w:before="60" w:after="60"/>${r}</w:pPr>
            <w:r>${r}<w:t xml:space="preserve">${xe(line.replace(/^\*|\*$/g, ''))}</w:t></w:r></w:p>`);
          parts.push(hr());
        } else if (isBullet) {
          parts.push(bullet(line.replace(/^[-–•]\s*/, '')));
        } else if (line.trim()) {
          // Job title — bold, sz=24
          parts.push(np(line, 120, 80, { bold: true, sz: 24, color: '262626' }));
        }
        i++;
      }
      parts.push(np('', 120));
      continue;
    }

    // Default (education etc.)
    for (const line of content) {
      parts.push(np(line, 120, 80, { sz: 22, color: '262626' }));
    }
    parts.push(np('', 120));
  }

  // Footer
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
    xml = xml.replace(/Director of Identity &amp;amp; Authentication/g, xe(position || ''));
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

    const templatePath = path.join(process.cwd(), 'template.docx');
    const templateBuffer = fs.readFileSync(templatePath);
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
