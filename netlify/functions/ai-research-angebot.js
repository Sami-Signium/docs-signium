// Netlify Function: ai-research-angebot
// Two modes: 
//   mode: "company" - web search + profile text
//   mode: "refine"  - refine freetext from keywords

exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') return { statusCode: 405, body: 'Method Not Allowed' };

  let body;
  try { body = JSON.parse(event.body); } 
  catch { return { statusCode: 400, body: JSON.stringify({ error: 'Invalid JSON' }) }; }

  const { mode, company, positionTitle, keywords, language } = body;
  const lang = language || 'DE';
  const isDE = lang === 'DE';
  const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY;

  let systemPrompt, userPrompt, useWebSearch;

  if (mode === 'company') {
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
    return { statusCode: 400, body: JSON.stringify({ error: 'Unknown mode' }) };
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

    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ text }),
    };

  } catch (err) {
    console.error('AI research error:', err);
    return { statusCode: 500, body: JSON.stringify({ error: err.message }) };
  }
};
