const pdfParse = require('pdf-parse');
const mammoth = require('mammoth');

export const config = { api: { bodyParser: { sizeLimit: '10mb' } } };

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method Not Allowed' });

  try {
    const { fileBase64, fileName } = req.body;
    if (!fileBase64) return res.status(400).json({ error: 'No file data' });

    const buffer = Buffer.from(fileBase64, 'base64');
    const name = (fileName || '').toLowerCase();
    let text = '';

    if (name.endsWith('.pdf')) {
      const data = await pdfParse(buffer);
      text = data.text;
    } else if (name.endsWith('.docx') || name.endsWith('.doc')) {
      const result = await mammoth.extractRawText({ buffer });
      text = result.value;
    } else {
      text = buffer.toString('utf-8');
    }

    return res.status(200).json({ text: text.trim() });
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}
