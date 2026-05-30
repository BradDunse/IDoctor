const Anthropic = require('@anthropic-ai/sdk');
const {
  Document, Packer, Paragraph, TextRun,
  HeadingLevel, AlignmentType, LevelFormat,
  BorderStyle, WidthType
} = require('docx');

const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

function buildSystemPrompt(openingDensity, crossfadeSpeed) {
  return `You are The I Doctor, an expert writing editor. Your job is to rewrite first-person copy so it starts with the writer's authentic voice and gradually transitions to "you"-based language that speaks directly to the reader.

ZONE RULES:
- Zone 1 (first eighth of the piece): Keep self-references (I, me, my, mine, myself, I've, I'd, I'll, I'm, we, us, our) at a density matching level ${openingDensity} out of 10. Level 10 means keep nearly all self-references. Level 1 means very few.
- Zone 2 (from one-eighth to three-eighths): Crossfade speed is ${crossfadeSpeed} out of 10. Level 10 means transition abruptly. Level 1 means fade very slowly and gently.
- Zone 3 (beyond three-eighths): Write in "you"-based language. Only keep a self-reference if removing it would genuinely break the meaning or voice. These moments should be rare.

VOICE INTELLIGENCE RULES:
- When a sentence is the writer drawing on personal lived experience to make a point, preserve the "I" in Zone 1 and early Zone 2.
- When a sentence shifts to a universal lesson or implication for the reader, use "you" or "you would" framing.
- Never use flat "you had" or "you did" for experiences the reader hasn't lived. Use conditional "you would" instead.
- The seam between personal testimony and reader invitation is where the transition happens naturally.

FORMATTING RULES:
- Apply proper Markdown formatting throughout.
- Section headers get ## markdown heading format.
- Quoted speech from named people (like "Nick says...") gets italicized with *asterisks*.
- Key rules, punchy declarative lines, and core principles get **bolded**.
- Use bulleted or numbered lists only where the content genuinely calls for it, not by default.
- Keep formatting restrained. Bold and italic should mean something. If everything is emphasized, nothing is.
- Format for mobile reading: short paragraphs, generous white space, punchy sentences.

OUTPUT FORMAT:
Return a JSON object with exactly two fields:
{
  "rewritten": "the full rewritten text in Markdown",
  "selfRefCount": <number of self-references remaining>
}
Return ONLY the JSON object. No preamble, no explanation, no markdown code fences.`;
}

function parseMarkdownToDocx(markdown) {
  const lines = markdown.split('\n');
  const children = [];

  const bulletConfig = {
    reference: 'bullets',
    levels: [{
      level: 0,
      format: LevelFormat.BULLET,
      text: '\u2022',
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 720, hanging: 360 } } }
    }]
  };

  const numberConfig = {
    reference: 'numbers',
    levels: [{
      level: 0,
      format: LevelFormat.DECIMAL,
      text: '%1.',
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 720, hanging: 360 } } }
    }]
  };

  let numberRef = 0;

  function parseInline(text) {
    const runs = [];
    const regex = /(\*\*(.+?)\*\*)|(\*(.+?)\*)/g;
    let last = 0;
    let match;

    while ((match = regex.exec(text)) !== null) {
      if (match.index > last) {
        runs.push(new TextRun({ text: text.slice(last, match.index), font: 'Arial', size: 24 }));
      }
      if (match[1]) {
        runs.push(new TextRun({ text: match[2], bold: true, font: 'Arial', size: 24 }));
      } else if (match[3]) {
        runs.push(new TextRun({ text: match[4], italics: true, font: 'Arial', size: 24 }));
      }
      last = match.index + match[0].length;
    }

    if (last < text.length) {
      runs.push(new TextRun({ text: text.slice(last), font: 'Arial', size: 24 }));
    }

    return runs.length > 0 ? runs : [new TextRun({ text, font: 'Arial', size: 24 })];
  }

  for (const line of lines) {
    const trimmed = line.trim();

    if (!trimmed) {
      children.push(new Paragraph({ children: [new TextRun('')], spacing: { after: 120 } }));
      continue;
    }

    if (trimmed.startsWith('## ')) {
      children.push(new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun({ text: trimmed.slice(3), bold: true, font: 'Arial', size: 28 })],
        spacing: { before: 240, after: 120 }
      }));
    } else if (trimmed.startsWith('# ')) {
      children.push(new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: trimmed.slice(2), bold: true, font: 'Arial', size: 32 })],
        spacing: { before: 360, after: 180 }
      }));
    } else if (trimmed.startsWith('- ') || trimmed.startsWith('* ')) {
      children.push(new Paragraph({
        numbering: { reference: 'bullets', level: 0 },
        children: parseInline(trimmed.slice(2)),
        spacing: { after: 80 }
      }));
    } else if (/^\d+\.\s/.test(trimmed)) {
      numberRef++;
      children.push(new Paragraph({
        numbering: { reference: `numbers-${numberRef}`, level: 0 },
        children: parseInline(trimmed.replace(/^\d+\.\s/, '')),
        spacing: { after: 80 }
      }));
    } else {
      children.push(new Paragraph({
        children: parseInline(trimmed),
        spacing: { after: 160 }
      }));
    }
  }

  return { children, bulletConfig, numberConfig };
}

async function generateDocx(rewrittenText) {
  const { children, bulletConfig } = parseMarkdownToDocx(rewrittenText);

  const doc = new Document({
    styles: {
      default: {
        document: { run: { font: 'Arial', size: 24 } }
      },
      paragraphStyles: [
        {
          id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 32, bold: true, font: 'Arial', color: '000000' },
          paragraph: { spacing: { before: 360, after: 180 }, outlineLevel: 0 }
        },
        {
          id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 28, bold: true, font: 'Arial', color: '000000' },
          paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 }
        }
      ]
    },
    numbering: {
      config: [bulletConfig]
    },
    sections: [{
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
        }
      },
      children
    }]
  });

  return await Packer.toBuffer(doc);
}

module.exports = async function rewriteRoute(req, res) {
  const { text, email, openingDensity, crossfadeSpeed } = req.body;

  if (!text || !email) {
    return res.status(400).json({ error: 'Text and email are required.' });
  }

  if (!process.env.ANTHROPIC_API_KEY) {
    return res.status(500).json({ error: 'API key not configured.' });
  }

  try {
    const message = await client.messages.create({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 4096,
      system: buildSystemPrompt(openingDensity || 7, crossfadeSpeed || 5),
      messages: [{ role: 'user', content: text }]
    });

    const raw = message.content[0].text.trim();
    let parsed;
    try {
      parsed = JSON.parse(raw);
    } catch {
      const jsonMatch = raw.match(/\{[\s\S]*\}/);
      parsed = jsonMatch ? JSON.parse(jsonMatch[0]) : { rewritten: raw, selfRefCount: 0 };
    }

    const docxBuffer = await generateDocx(parsed.rewritten);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename="i-doctor-rewrite.docx"');
    res.setHeader('X-Self-Ref-Count', parsed.selfRefCount || 0);
    res.send(docxBuffer);

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Something went wrong. Please try again.' });
  }
};
