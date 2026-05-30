const Anthropic = require('@anthropic-ai/sdk');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, LevelFormat } = require('docx');

const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

function buildSystemPrompt(openingDensity, crossfadeSpeed) {
  var prompt = 'You are The I Doctor, an expert writing editor. Your job is to rewrite first-person copy so it starts with the writers authentic voice and gradually transitions to you-based language that speaks directly to the reader.\n\n';
  prompt += 'ZONE RULES:\n';
  prompt += '- Zone 1 (first eighth of the piece): Keep self-references (I, me, my, mine, myself, I\'ve, I\'d, I\'ll, I\'m, we, us, our) at a density matching level ' + openingDensity + ' out of 10. Level 10 means keep nearly all self-references. Level 1 means very few.\n';
  prompt += '- Zone 2 (from one-eighth to three-eighths): Crossfade speed is ' + crossfadeSpeed + ' out of 10. Level 10 means transition abruptly. Level 1 means fade very slowly and gently.\n';
  prompt += '- Zone 3 (beyond three-eighths): Write in you-based language. Only keep a self-reference if removing it would genuinely break the meaning or voice. These moments should be rare.\n\n';
  prompt += 'VOICE INTELLIGENCE RULES:\n';
  prompt += '- When a sentence is the writer drawing on personal lived experience to make a point, preserve the I in Zone 1 and early Zone 2.\n';
  prompt += '- When a sentence shifts to a universal lesson or implication for the reader, use you or you would framing.\n';
  prompt += '- Never use flat you had or you did for experiences the reader has not lived. Use conditional you would instead.\n';
  prompt += '- The seam between personal testimony and reader invitation is where the transition happens naturally.\n\n';
  prompt += 'FORMATTING RULES:\n';
  prompt += '- Apply proper Markdown formatting throughout.\n';
  prompt += '- Section headers get ## markdown heading format.\n';
  prompt += '- When italicizing quoted speech, place the quotation marks OUTSIDE the asterisks like this: "*spoken words*" not *"spoken words"*. This keeps punctuation separate from italic formatting.\n';
  prompt += '- Key rules, punchy declarative lines, and core principles get **bolded**.\n';
  prompt += '- Use bulleted or numbered lists only where the content genuinely calls for it.\n';
  prompt += '- Keep formatting restrained. Bold and italic should mean something.\n';
  prompt += '- Format for mobile reading: short paragraphs, generous white space, punchy sentences.\n\n';
  prompt += 'OUTPUT FORMAT:\n';
  prompt += 'Return a single valid JSON object and nothing else. No markdown fences, no explanation, no text before or after.\n';
  prompt += 'The JSON object must have exactly two fields:\n';
  prompt += '- rewritten: the full rewritten text as a string with \\n for line breaks\n';
  prompt += '- selfRefCount: a number counting remaining self-references\n';
  prompt += 'Start your response with { and end with }. Nothing else.';
  return prompt;
}

async function addToKit(email) {
  var kitApiKey = process.env.KIT_API_KEY;
  var formId = '6711021';
  if (!kitApiKey) return;
  try {
    await fetch('https://api.kit.com/v4/forms/' + formId + '/subscribers', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'X-Kit-Api-Key': kitApiKey },
      body: JSON.stringify({ email_address: email })
    });
  } catch (err) {
    console.error('Kit error:', err.message);
  }
}

function parseInline(text) {
  var runs = [];
  var regex = /(\*\*(.+?)\*\*)|(\*(.+?)\*)/g;
  var last = 0;
  var match;
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
  return runs.length > 0 ? runs : [new TextRun({ text: text, font: 'Arial', size: 24 })];
}

function parseMarkdownToDocx(markdown) {
  var lines = markdown.split('\n');
  var children = [];
  var bulletConfig = {
    reference: 'bullets',
    levels: [{ level: 0, format: LevelFormat.BULLET, text: '\u2022', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }]
  };
  for (var i = 0; i < lines.length; i++) {
    var trimmed = lines[i].trim();
    if (!trimmed) {
      children.push(new Paragraph({ children: [new TextRun('')], spacing: { after: 120 } }));
    } else if (trimmed.startsWith('## ')) {
      children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: trimmed.slice(3), bold: true, font: 'Arial', size: 28 })], spacing: { before: 240, after: 120 } }));
    } else if (trimmed.startsWith('# ')) {
      children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: trimmed.slice(2), bold: true, font: 'Arial', size: 32 })], spacing: { before: 360, after: 180 } }));
    } else if (trimmed.startsWith('- ') || trimmed.startsWith('* ')) {
      children.push(new Paragraph({ numbering: { reference: 'bullets', level: 0 }, children: parseInline(trimmed.slice(2)), spacing: { after: 80 } }));
    } else {
      children.push(new Paragraph({ children: parseInline(trimmed), spacing: { after: 160 } }));
    }
  }
  return { children: children, bulletConfig: bulletConfig };
}

async function generateDocx(rewrittenText) {
  var parsed = parseMarkdownToDocx(rewrittenText);
  var doc = new Document({
    styles: {
      default: { document: { run: { font: 'Arial', size: 24 } } },
      paragraphStyles: [
        { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 32, bold: true, font: 'Arial', color: '000000' }, paragraph: { spacing: { before: 360, after: 180 }, outlineLevel: 0 } },
        { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 28, bold: true, font: 'Arial', color: '000000' }, paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } }
      ]
    },
    numbering: { config: [parsed.bulletConfig] },
    sections: [{ properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } }, children: parsed.children }]
  });
  return await Packer.toBuffer(doc);
}

module.exports = async function rewriteRoute(req, res) {
  var text = req.body.text;
  var email = req.body.email;
  var openingDensity = req.body.openingDensity || 7;
  var crossfadeSpeed = req.body.crossfadeSpeed || 5;

  if (!text || !email) {
    return res.status(400).json({ error: 'Text and email are required.' });
  }
  if (!process.env.ANTHROPIC_API_KEY) {
    return res.status(500).json({ error: 'API key not configured.' });
  }

  try {
    await addToKit(email);

    var message = await client.messages.create({
      model: 'claude-sonnet-4-5',
      max_tokens: 4096,
      system: buildSystemPrompt(openingDensity, crossfadeSpeed),
      messages: [{ role: 'user', content: text }]
    });

    var raw = message.content[0].text.trim();

    var result;
    try {
      result = JSON.parse(raw);
    } catch (e) {
      try {
        var cleaned = raw.replace(/```json/g, '').replace(/```/g, '').trim();
        var jsonMatch = cleaned.match(/\{[\s\S]*\}/);
        result = jsonMatch ? JSON.parse(jsonMatch[0]) : { rewritten: raw, selfRefCount: 0 };
      } catch (e2) {
        result = { rewritten: raw, selfRefCount: 0 };
      }
    }

    var rewrittenText = result.rewritten || raw;
    var docxBuffer = await generateDocx(rewrittenText);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename="i-doctor-rewrite.docx"');
    res.setHeader('X-Self-Ref-Count', result.selfRefCount || 0);
    res.send(docxBuffer);

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Something went wrong. Please try again.' });
  }
};
