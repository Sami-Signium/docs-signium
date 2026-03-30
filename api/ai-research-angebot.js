// Netlify Function: ai-research-angebot
// Two modes: 
//   mode: "company" - web search + profile text
//   mode: "refine"  - refine freetext from keywords

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method Not Allowed' });

  let body;
  try { body = req.body; } 
  catch { return res.status(400).json({ error: 'Invalid JSON' }); }

  const { mode, company, positionTitle, keywords, language } = body;
  const lang = language || 'DE';
  const isDE = lang === 'DE';
  const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY;

  let systemPrompt, userPrompt, useWebSearch;

  if (mode === 'extract') {
    // Extract all fields from uploaded document
    const { fileBase64, mediaType, fileName, language: lang2 } = body;
    const isDE2 = (lang2 || 'DE') === 'DE';

    const extractSystem = isDE2
      ? `Du bist ein Executive Search Berater bei Signium Austria. Analysiere das hochgeladene Dokument (Job Profil, Briefing oder Anforderungsprofil) und extrahiere alle relevanten Informationen.
Antworte NUR mit einem JSON-Objekt, kein erklärender Text davor oder danach, keine Markdown-Backticks.
JSON-Struktur:
{
  "positionTitle": "Positionstitel",
  "clientCompany": "Firmenname",
  "clientContactName": "Vollständiger Name Ansprechpartner",
  "clientContactLastName": "Nur Nachname",
  "clientSalutation": "geehrte Frau / geehrter Herr",
  "clientAddress": "Straße und Hausnummer",
  "clientCity": "PLZ Ort",
  "clientEmail": "email@firma.at",
  "companyProfile": "Professioneller Fließtext über das Unternehmen, 200-300 Wörter",
  "positionDescription": "Professioneller Fließtext zur Position, 150-250 Wörter",
  "functionalTargets": ["Target 1", "Target 2", "Target 3"],
  "industryTargets": ["Branche 1", "Branche 2"],
  "geoTargets": ["Land 1", "Land 2"]
}
Felder die nicht im Dokument stehen: leerer String "" oder leeres Array [].
companyProfile und positionDescription immer als professionellen Fließtext formulieren.`
      : `You are an Executive Search consultant at Signium Austria. Analyse the uploaded document (job profile, briefing or requirements profile) and extract all relevant information.
Respond ONLY with a JSON object, no explanatory text before or after, no markdown backticks.
JSON structure:
{
  "positionTitle": "Position title",
  "clientCompany": "Company name",
  "clientContactName": "Full name of contact person",
  "clientContactLastName": "Last name only",
  "clientSalutation": "Ms. / Mr.",
  "clientAddress": "Street and number",
  "clientCity": "ZIP City",
  "clientEmail": "email@company.com",
  "companyProfile": "Professional prose about the company, 200-300 words",
  "positionDescription": "Professional prose about the position, 150-250 words",
  "functionalTargets": ["Target 1", "Target 2", "Target 3"],
  "industryTargets": ["Industry 1", "Industry 2"],
  "geoTargets": ["Country 1", "Country 2"]
}
Fields not found in the document: empty string "" or empty array [].
Always write companyProfile and positionDescription as professional prose.`;

    try {
      const isPdf = (mediaType || '').includes('pdf');
      const docContent = isPdf
        ? { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: fileBase64 } }
        : { type: 'text', text: `[Word document content — extract fields from: ${fileName}]\n\nNote: Parse the document structure and extract all available information.` };

      // For Word docs, we pass base64 as document type too
      const messageContent = isPdf
        ? [docContent, { type: 'text', text: isDE2 ? 'Analysiere dieses Dokument und extrahiere alle Felder als JSON.' : 'Analyse this document and extract all fields as JSON.' }]
        : [{ type: 'document', source: { type: 'base64', media_type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', data: fileBase64 } },
           { type: 'text', text: isDE2 ? 'Analysiere dieses Dokument und extrahiere alle Felder als JSON.' : 'Analyse this document and extract all fields as JSON.' }];

      const response = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': ANTHROPIC_API_KEY,
          'anthropic-version': '2023-06-01',
        },
        body: JSON.stringify({
          model: 'claude-sonnet-4-5',
          max_tokens: 2048,
          system: extractSystem,
          messages: [{ role: 'user', content: messageContent }],
        }),
      });

      if (!response.ok) throw new Error(`API ${response.status}`);
      const data = await response.json();
      const raw = data.content.filter(b => b.type === 'text').map(b => b.text).join('').trim();

      // Strip markdown fences if present
      const clean = raw.replace(/^```json\s*/i,'').replace(/^```\s*/,'').replace(/```\s*$/,'').trim();
      const extracted = JSON.parse(clean);

      return res.status(200).json({ extracted });
    } catch (err) {
      console.error('Extract error:', err);
      return res.status(500).json({ error: err.message });
    }
  }


    // Research company and write profile
    useWebSearch = true;
    systemPrompt = isDE
      ? `Du bist ein Executive Search Berater bei Signium Austria. Schreibe ein professionelles Unternehmensprofil für ein Angebot. 
Stil: sachlich, präzise, 3-5 Absätze, ca. 200-300 Wörter. Kein Marketing-Speak.
Inhalt: Gründung/Geschichte, Kerngeschäft, Größe (Mitarbeiter/Umsatz falls bekannt), Marktposition, Besonderheiten, Eigentümerstruktur.
Nur Fließtext, keine Aufzählungen, keine Überschriften. Auf Deutsch.`
      : `You are an Executive Search consultant at Signium Austria. Write a professional company profile for a proposal.
Style: factual, precise, 3-5 paragraphs, approx. 200-300 words. No marketing speak.
Content: founding/history, core business, size (employees/revenue if known), market position, ownership structure.
Plain prose only, no bullets, no headings. In English.`;
    userPrompt = isDE
      ? `Recherchiere das Unternehmen "${company}" und schreibe ein Unternehmensprofil für unser Executive Search Angebot.`
      : `Research the company "${company}" and write a company profile for our Executive Search proposal.`;

  } else if (mode === 'position') {
    // Refine position description from keywords
    useWebSearch = false;
    systemPrompt = isDE
      ? `Du bist ein Executive Search Berater bei Signium Austria. Schreibe eine professionelle Positionsbeschreibung für ein Angebot.
Stil: präzise, 2-3 Absätze, ca. 150-250 Wörter. Klarer Führungsauftrag sichtbar.
Inhalt: Kernaufgaben, Verantwortungsbereich, strategische Bedeutung, ggf. Reporting-Linie.
Nur Fließtext, keine Aufzählungen. Auf Deutsch.`
      : `You are an Executive Search consultant at Signium Austria. Write a professional position description for a proposal.
Style: precise, 2-3 paragraphs, approx. 150-250 words. Clear leadership mandate visible.
Content: core tasks, area of responsibility, strategic significance, reporting line if known.
Plain prose only, no bullets. In English.`;
    userPrompt = isDE
      ? `Unternehmen: ${company || '(nicht angegeben)'}
Position: ${positionTitle || ''}
Stichworte/Notizen des Beraters: ${keywords || '(keine weiteren Angaben)'}

Schreibe eine professionelle Positionsbeschreibung für unser Angebot.`
      : `Company: ${company || '(not specified)'}
Position: ${positionTitle || ''}
Consultant notes/keywords: ${keywords || '(none)'}

Write a professional position description for our proposal.`;

  } else {
    return res.status(400).json({ error: 'Unknown mode' });
  }

  try {
    const reqBody = {
      model: 'claude-sonnet-4-5',
      max_tokens: 1024,
      system: systemPrompt,
      messages: [{ role: 'user', content: userPrompt }],
    };

    if (useWebSearch) {
      reqBody.tools = [{ type: 'web_search_20250305', name: 'web_search' }];
    }

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01',
        'anthropic-beta': 'web-search-2025-03-05',
      },
      body: JSON.stringify(reqBody),
    });

    if (!response.ok) {
      const err = await response.text();
      throw new Error(`API error ${response.status}: ${err}`);
    }

    const data = await response.json();
    const text = data.content
      .filter(b => b.type === 'text')
      .map(b => b.text)
      .join('\n')
      .trim();

    return res.status(200).json({ text });

  } catch (err) {
    console.error('AI research error:', err);
    return res.status(500).json({ error: err.message });
  }
};

