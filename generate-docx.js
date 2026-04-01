import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import JSZip from 'jszip';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

export const config = { api: { bodyParser: { sizeLimit: '10mb' } } };

function xe(str) {
  if (!str) return '';
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function sp(styleId, text, before, after, opts) {
  opts = opts || {};
  let ppr = `<w:pStyle w:val="${styleId}"/>`;
  if (opts.pageBreak) ppr += `<w:pageBreakBefore/>`;
  if (before !== undefined || after !== undefined) {
    const b = before !== undefined ? ` w:before="${before}"` : '';
    const a = after !== undefined ? ` w:after="${after}"` : '';
    ppr += `<w:spacing${b}${a}/>`;
  }
  if (!text && text !== 0) return `<w:p><w:pPr>${ppr}</w:pPr></w:p>`;
  let rpr = `<w:rPr><w:color w:val="auto"/>`;
  if (opts.bold) rpr += '<w:b/>';
  if (opts.sz) rpr += `<w:sz w:val="${opts.sz}"/><w:szCs w:val="${opts.sz}"/>`;
  rpr += '</w:rPr>';
  return `<w:p><w:pPr>${ppr}</w:pPr><w:r>${rpr}<w:t xml:space="preserve">${xe(text)}</w:t></w:r></w:p>`;
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
  let rpr = `<w:rPr>`;
  if (opts.bold) rpr += '<w:b/>';
  if (opts.italic) rpr += '<w:i/>';
  const color = opts.color || 'auto';
  rpr += `<w:color w:val="${color}"/>`;
  const sz = opts.sz || 22;
  rpr += `<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>`;
  rpr += '</w:rPr>';
  if (!text && text !== 0) return `<w:p><w:pPr>${ppr}</w:pPr></w:p>`;
  return `<w:p><w:pPr>${ppr}</w:pPr><w:r>${rpr}<w:t xml:space="preserve">${xe(text)}</w:t></w:r></w:p>`;
}

function personalRow(label, value) {
  const isLong = value && value.length > 60;
  const indXml = isLong ? '<w:ind w:left="2200" w:hanging="2200"/>' : '';
  return `<w:p>
    <w:pPr>
      <w:tabs><w:tab w:val="left" w:pos="2200"/></w:tabs>
      <w:spacing w:before="80" w:after="80"/>
      ${indXml}
    </w:pPr>
    <w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/><w:color w:val="414042"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t>${xe(label)}</w:t></w:r>
    <w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/><w:color w:val="414042"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:tab/></w:r>
    <w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/><w:color w:val="262626"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t xml:space="preserve">${xe(value)}</w:t></w:r>
  </w:p>`;
}

function hr() {
  return np('________________________________________________________________________________', 60, 60, { color: 'AAAAAA', sz: 16 });
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

const NEW_PAGE = [
  'PERSÖNLICHE','PERSOENLICHE','PERSONAL D',
  'VERGÜTUNG','VERGUETUNG','COMPENSATION',
  'KARRIERE','CAREER SUMMARY',
  'FACHLICHES',
  'BERUFSERFAHRUNG','BERUFLICHER','WORK EXPERIENCE','PROFESSIONAL EXPERIENCE'
];

function needsPageBreak(key) {
  const u = key.toUpperCase();
  return NEW_PAGE.some(k => u.startsWith(k));
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

  parts.push(sp('Titleheader', (candidateName || 'KANDIDAT').toUpperCase(), 120, 0));
  parts.push(sp('Coverdoctitle', 'VERTRAULICHER KANDIDATENBERICHT', 4080, 0, { sz: 32 }));
  if (position) parts.push(sp('Coverdate', position.toUpperCase(), 720, 1000, { bold: true, sz: 28 }));
  if (client && client !== 'Vertraulich') parts.push(sp('Coverdate', client, 120, 120, { bold: true, sz: 32 }));
  parts.push(sp('Coverdate', datum || '', 120, 1000));
  parts.push(np('', 120));
  parts.push(np('Dieser Vertrauliche Bericht enthält zum Teil Informationen, die uns unter Zusicherung strengster Vertraulichkeit mitgeteilt wurden. Entsprechend unseren berufsethischen Prinzipien müssen wir Sie dazu verpflichten, nur einer begrenzten Auswahl von Personen Einsicht in diese Berichte zu gewähren.', 120, undefined, { italic: true, color: '595959', sz: 18, jc: 'both' }));
  parts.push(np('', 120));

  let vergütungInserted = false;

  for (const section of sections) {
    if (section.key === 'HEADER') continue;
    const content = section.lines.map(l => l.trim()).filter(Boolean);

    const ku = section.key.toUpperCase();
    const isPersonal   = ku.includes('PERSÖN') || ku.includes('PERSOEN') || ku.includes('PERSONAL');
    const isExperience = ku.includes('BERUFSERFAHRUNG') || ku.includes('BERUFLICHER') || ku.includes('WORK EXPERIENCE') || ku.includes('PROFESSIONAL EXPERIENCE');
    const isKarriere   = ku.includes('KARRIERE') || ku.includes('CAREER SUMMARY');
    const isVergütung  = ku.includes('VERGÜTUNG') || ku.includes('VERGUETUNG') || ku.includes('COMPENSATION');
    const isKandidaten = ku.includes('FACHLICHES') || ku.includes('BEWERTUNG') || ku.includes('PERSONALITY') || ku.includes('KANDIDATEN') || ku.includes('MOTIVATION');
    const pageBreak    = needsPageBreak(section.key);

    if ((isKarriere || isKandidaten || isExperience) && !vergütungInserted) {
      vergütungInserted = true;
      parts.push(sp('berschrift2', 'VERGÜTUNG UND VERFÜGBARKEIT', 120, undefined, { pageBreak: true, bold: true, sz: 28 }));
      parts.push(hr());
      for (const label of ['Aktuelles Fixgehalt','Aktueller Bonus','Gehaltsvorstellung','Kündigungsfrist','Verfügbarkeit','Reisebereitschaft']) {
        parts.push(personalRow(label, ''));
      }
      parts.push(np('', 120));
    }

    if (isVergütung) {
      vergütungInserted = true;
      parts.push(sp('berschrift2', 'VERGÜTUNG UND VERFÜGBARKEIT', 120, undefined, { pageBreak, bold: true, sz: 28 }));
      parts.push(hr());
      const provided = {};
      for (const line of content) {
        if (line.includes(':')) {
          const idx = line.indexOf(':');
          provided[line.slice(0, idx).trim()] = line.slice(idx + 1).trim();
        }
      }
      for (const label of ['Aktuelles Fixgehalt','Aktueller Bonus','Gehaltsvorstellung','Kündigungsfrist','Verfügbarkeit','Reisebereitschaft']) {
        parts.push(personalRow(label, provided[label] || ''));
      }
      parts.push(np('', 120));
      continue;
    }

    if (!content.length) continue;

    parts.push(sp('berschrift2', section.key.toUpperCase(), 120, undefined, { pageBreak, bold: true, sz: 28 }));
    parts.push(hr());

    if (isPersonal) {
      for (const line of content) {
        if (line.includes(':')) {
          const idx = line.indexOf(':');
          parts.push(personalRow(line.slice(0, idx).trim(), line.slice(idx + 1).trim()));
        } else {
          parts.push(np(line, 80, 80));
        }
      }
      parts.push(np('', 120));
      parts.push(np('', 120));
      parts.push(np('', 120));
      continue;
    }

    if (isKarriere) {
      for (const line of content) parts.push(np(line, 120, 160, { bold: true }));
      parts.push(np('', 120));
      continue;
    }

    if (isKandidaten) {
      for (const line of content) {
        if (!line.trim()) continue;
        if (/^[-•]/.test(line)) {
          parts.push(sp('Listenabsatz', line.replace(/^[-•]\s*/, ''), 60, 120, { sz: 24 }));
        } else {
          parts.push(np(line, 160, 160, { jc: 'both', sz: 22 }));
        }
      }
      parts.push(np('', 120));
      continue;
    }

    if (isExperience) {
      let firstCompany = true;
      for (let i = 0; i < content.length; i++) {
        const line = content[i];
        const isBullet = /^[-–•]/.test(line);
        const isCompanyDesc = /^\*/.test(line);
        const isCompanyHeader = /^(Januar|Februar|März|April|Mai|Juni|Juli|August|September|Oktober|November|Dezember|Jan\.|Feb\.|Mär\.|Apr\.|Jun\.|Jul\.|Aug\.|Sep\.|Okt\.|Nov\.|Dez\.|Oct\.|Sept?\.|Jan\s|Feb\s|\d{2}\/\d{4}|\d{4})/.test(line) && !isBullet;
        if (isCompanyHeader) {
          if (!firstCompany) parts.push(hr());
          firstCompany = false;
          const colonIdx = line.indexOf(':');
          const dashIdx = line.indexOf(' - ', 8);
          let datePart = line, companyPart = '';
          if (colonIdx > 0) { datePart = line.slice(0, colonIdx).trim(); companyPart = line.slice(colonIdx + 1).trim(); }
          else if (dashIdx > 0) { datePart = line.slice(0, dashIdx).trim(); companyPart = line.slice(dashIdx + 3).trim(); }
          if (companyPart) {
            parts.push(`<w:p><w:pPr><w:pStyle w:val="Amrop-header"/><w:spacing w:before="120" w:after="120"/></w:pPr><w:r><w:rPr><w:color w:val="auto"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t xml:space="preserve">${xe(datePart)}: </w:t></w:r><w:r><w:rPr><w:b/><w:color w:val="auto"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:t xml:space="preserve">${xe(companyPart)}</w:t></w:r></w:p>`);
          } else {
            parts.push(sp('Amrop-header', line, 120, 120, { sz: 22 }));
          }
        } else if (isCompanyDesc) {
          parts.push(sp('Listing1', line.replace(/^\*|\*$/g, ''), 120, undefined, { sz: 22 }));
          parts.push(hr());
        } else if (isBullet) {
          parts.push(sp('Listenabsatz', line.replace(/^[-–•]\s*/, ''), 60, 120, { sz: 24 }));
        } else if (line.trim()) {
          parts.push(np(line, 120, 120, { bold: true, sz: 24 }));
        }
      }
      parts.push(np('', 120));
      continue;
    }

    for (const line of content) parts.push(np(line, 120, 120, { jc: 'both' }));
    parts.push(np('', 120));
  }

  if (!vergütungInserted) {
    parts.push(sp('berschrift2', 'VERGÜTUNG UND VERFÜGBARKEIT', 120, undefined, { pageBreak: true, bold: true, sz: 28 }));
    parts.push(hr());
    for (const label of ['Aktuelles Fixgehalt','Aktueller Bonus','Gehaltsvorstellung','Kündigungsfrist','Verfügbarkeit','Reisebereitschaft']) {
      parts.push(personalRow(label, ''));
    }
    parts.push(np('', 120));
  }

  parts.push(np('', 240));
  parts.push(np('Vorbereitet von: Dr. Sami Hamid  |  Managing Partner  |  Signium Austria', 120, 0, { bold: true, color: '102E66', sz: 18 }));
  parts.push(np('t +43 664 4568862  |  sami.hamid@signium.com', 40, 0, { color: '595959', sz: 17 }));

  return parts.join('\n');
}

async function updateHeaders(zip, candidateName, position, client) {
  for (const hf of ['word/header1.xml','word/header2.xml','word/header3.xml','word/footer1.xml','word/footer2.xml','word/footer3.xml']) {
    const file = zip.file(hf);
    if (!file) continue;
    let xml = await file.async('string');
    xml = xml.replace(/Quintin Stephen/g, xe(candidateName || ''));
    xml = xml.replace(/Dr\. Sami Hamid(?!\s*\|)/g, xe(candidateName || ''));
    xml = xml.replace(/Director of Identity &amp; Authentication/g, xe(position || ''));
    xml = xml.replace(/Director of Identity &amp;amp; Authentication/g, xe(position || ''));
    xml = xml.replace(/Austriacard(?!\s*Holdings)/g, xe(client && client !== 'Vertraulich' ? client : 'Confidential'));
    xml = xml.replace(/AustriaCard Holdings[^<]*/g, xe(candidateName || ''));
    if (position) {
      xml = xml.replace(/(<w:t[^>]*>)Managing Partner(<\/w:t>)/g, `$1${xe(position)}$2`);
    }
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
    const newDocXml = docXmlRaw.substring(0, bodyStart) + '\n' + buildBodyXml(text, candidateName, position, client, datum) + '\n' + sectPr + '\n' + docXmlRaw.substring(bodyEnd);
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
