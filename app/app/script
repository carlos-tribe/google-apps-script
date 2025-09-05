/**
 * Tribe Proposal Generator — Google Docs version with Markdown Support
 * - Copies a Google Docs template and injects provided variables
 * - Parses markdown formatting for professional document output
 * - Returns a shareable URL as JSON
 *
 * Script Properties:
 *   TEMPLATE_ID      (required) Google Docs template file ID
 *   DEST_FOLDER_ID   (optional) Drive folder ID for new docs
 *   API_KEY          (optional) Shared secret; require in request body as { apiKey: "..." }
 *   SHARE_MODE       (optional) ONE OF: "ANYONE" | "DOMAIN" | "PRIVATE" (default: ANYONE)
 */

function doGet(e) {
  try {
    const props = PropertiesService.getScriptProperties();
    const info = {
      ok: true,
      version: '2.0',
      now: new Date().toISOString(),
      who: 'webapp-exec-as-owner',
      // quick visibility into config
      hasTemplate: !!props.getProperty('TEMPLATE_ID'),
      shareMode: (props.getProperty('SHARE_MODE') || 'ANYONE').toUpperCase(),
    };
    return jsonOut(info);
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    // --- Input guards ---
    if (!e || !e.postData || !e.postData.contents) {
      return jsonOut({ success: false, error: 'No post data' });
    }
    const data = JSON.parse(e.postData.contents || '{}');
    console.log('payload', JSON.stringify(data));

    // Simple health check via POST
    if (data && data.action === 'ping') {
      return jsonOut({ success: true, ping: true, time: new Date().toISOString() });
    }

    // --- Config ---
    const props = PropertiesService.getScriptProperties();
    const TEMPLATE_ID = props.getProperty('TEMPLATE_ID') || '';
    const DEST_FOLDER_ID = props.getProperty('DEST_FOLDER_ID') || '';
    const API_KEY = props.getProperty('API_KEY') || '';
    const SHARE_MODE = (props.getProperty('SHARE_MODE') || 'ANYONE').toUpperCase();

    if (!TEMPLATE_ID) return jsonOut({ success: false, error: 'Missing TEMPLATE_ID in Script properties' });
    if (API_KEY && data.apiKey !== API_KEY) return jsonOut({ success: false, error: 'unauthorized' });

    // Validate required company_name field (snake_case only)
    if (!data.company_name || typeof data.company_name !== 'string') {
      return jsonOut({ success: false, error: 'Missing: company_name' });
    }

    coerceStrings(data, [
      'company_name','project_name','proposal_date','context_and_opportunity',
      'objectives_scope','architecture_description','project_plan',
      'tribe_team','customer_team','tribe_responsibilities','customer_responsibilities',
      'acceptance_criteria','assumptions','technical_dependencies'
    ]);

    // --- Concurrency guard ---
    if (!lock.tryLock(5000)) {
      return jsonOut({ success: false, error: 'busy: try again' });
    }

    // --- Naming & copy ---
    const tz = Session.getScriptTimeZone() || 'UTC';
    const stamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    const baseName = `${data.company_name} Proposal - ${stamp}`;

    const template = DriveApp.getFileById(TEMPLATE_ID);
    let folder = null;
    if (DEST_FOLDER_ID) {
      try { folder = DriveApp.getFolderById(DEST_FOLDER_ID); }
      catch (err) { console.warn('Invalid DEST_FOLDER_ID'); }
    }
    const newFile = folder ? template.makeCopy(baseName, folder) : template.makeCopy(baseName);

    // --- Sharing ---
    if (SHARE_MODE !== 'PRIVATE') {
      const access = SHARE_MODE === 'DOMAIN'
        ? DriveApp.Access.DOMAIN_WITH_LINK
        : DriveApp.Access.ANYONE_WITH_LINK;
      newFile.setSharing(access, DriveApp.Permission.VIEW);
    }

    // --- Token replacement with markdown support ---
    const doc = DocumentApp.openById(newFile.getId());
    const body = doc.getBody();

    // Process each field with markdown parsing
    const markdownFields = {
      '{{company_name}}': fallback(data.company_name, '[Company Name]'),
      '{{project_name}}': fallback(data.project_name, '[Project Name]'),
      '{{proposal_date}}': fallback(data.proposal_date, Utilities.formatDate(new Date(), tz, 'MMM d, yyyy')),
      '{{context_and_opportunity}}': fallback(data.context_and_opportunity, ''),
      '{{objectives_scope}}': fallback(data.objectives_scope, ''),
      '{{architecture_description}}': fallback(data.architecture_description, ''),
      '{{project_plan}}': fallback(data.project_plan, ''),
      '{{tribe_team}}': fallback(data.tribe_team, ''),
      '{{customer_team}}': fallback(data.customer_team, ''),
      '{{tribe_responsibilities}}': fallback(data.tribe_responsibilities, ''),
      '{{customer_responsibilities}}': fallback(data.customer_responsibilities, ''),
      '{{acceptance_criteria}}': fallback(data.acceptance_criteria, ''),
      '{{assumptions}}': fallback(data.assumptions, ''),
      '{{technical_dependencies}}': fallback(data.technical_dependencies, '')
    };

    // Replace tokens with markdown-formatted content
    Object.keys(markdownFields).forEach(token => {
      replaceTokenWithMarkdown(body, token, markdownFields[token]);
    });

    // Key Challenges - handle as proper list
    if (Array.isArray(data.key_challenges)) {
      replaceTokenWithFormattedList(body, '{{key_challenges}}', data.key_challenges.map(String).filter(Boolean));
    } else if (data.key_challenges) {
      replaceTokenWithMarkdown(body, '{{key_challenges}}', data.key_challenges);
    }

    // Investment table
    if (Array.isArray(data.investment_table) && data.investment_table.length && Array.isArray(data.investment_table[0])) {
      replaceTokenWithTable(body, '{{investment_table}}', data.investment_table);
    } else if (data.investment_table) {
      replaceTokenWithMarkdown(body, '{{investment_table}}', data.investment_table);
    }

    doc.saveAndClose();

    const url = `https://docs.google.com/document/d/${newFile.getId()}/edit?usp=sharing`;
    return jsonOut({ success: true, url, docId: newFile.getId(), message: `Proposal generated successfully for ${data.company_name}` });

  } catch (error) {
    console.error(error);
    return jsonOut({ success: false, error: String(error), message: 'Failed to generate proposal' });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/* ----------------- Markdown Parsing Functions ------------------ */

function replaceTokenWithMarkdown(body, token, text) {
  const range = body.findText(escapeForFind(token));
  if (!range) return;
  
  const elem = range.getElement();
  const para = elem.getParent().asParagraph();
  const index = body.getChildIndex(para);
  para.removeFromParent();
  
  if (!text) {
    body.insertParagraph(index, '');
    return;
  }
  
  // Split into lines and process each
  const lines = String(text).split(/\r?\n/);
  let currentIndex = index;
  let inList = false;
  let listStartIndex = -1;
  
  lines.forEach((line) => {
    // Check if this is a list item
    const bulletMatch = line.match(/^\s*[-*•]\s+(.+)$/);
    const numberMatch = line.match(/^\s*(\d+)\.\s+(.+)$/);
    
    if (bulletMatch) {
      // Bullet list item
      if (!inList || inList !== 'bullet') {
        inList = 'bullet';
        listStartIndex = currentIndex;
      }
      const listItem = body.insertListItem(currentIndex, bulletMatch[1].trim());
      listItem.setGlyphType(DocumentApp.GlyphType.BULLET);
      parseInlineMarkdown(listItem, bulletMatch[1].trim());
      currentIndex++;
    } else if (numberMatch) {
      // Numbered list item
      if (!inList || inList !== 'number') {
        inList = 'number';
        listStartIndex = currentIndex;
      }
      const listItem = body.insertListItem(currentIndex, numberMatch[2].trim());
      listItem.setGlyphType(DocumentApp.GlyphType.NUMBER);
      parseInlineMarkdown(listItem, numberMatch[2].trim());
      currentIndex++;
    } else {
      // Not a list item
      inList = false;
      
      // Check for headers
      const h1Match = line.match(/^#\s+(.+)$/);
      const h2Match = line.match(/^##\s+(.+)$/);
      const h3Match = line.match(/^###\s+(.+)$/);
      
      let newPara;
      if (h3Match) {
        newPara = body.insertParagraph(currentIndex, '');
        newPara.setHeading(DocumentApp.ParagraphHeading.HEADING3);
        parseInlineMarkdown(newPara, h3Match[1].trim());
      } else if (h2Match) {
        newPara = body.insertParagraph(currentIndex, '');
        newPara.setHeading(DocumentApp.ParagraphHeading.HEADING2);
        parseInlineMarkdown(newPara, h2Match[1].trim());
      } else if (h1Match) {
        newPara = body.insertParagraph(currentIndex, '');
        newPara.setHeading(DocumentApp.ParagraphHeading.HEADING1);
        parseInlineMarkdown(newPara, h1Match[1].trim());
      } else {
        // Regular paragraph
        newPara = body.insertParagraph(currentIndex, '');
        parseInlineMarkdown(newPara, line);
      }
      currentIndex++;
    }
  });
}

function parseInlineMarkdown(element, text) {
  // Clear existing text
  element.clear();
  
  if (!text) return;
  
  let remaining = text;
  let position = 0;
  
  while (remaining.length > 0) {
    let matched = false;
    
    // Check for bold (**text**)
    const boldMatch = remaining.match(/^\*\*([^*]+)\*\*/);
    if (boldMatch) {
      element.appendText(boldMatch[1]).setBold(true);
      remaining = remaining.substring(boldMatch[0].length);
      matched = true;
    }
    
    // Check for italic (*text* or _text_)
    if (!matched) {
      const italicMatch = remaining.match(/^[*_]([^*_]+)[*_]/);
      if (italicMatch) {
        element.appendText(italicMatch[1]).setItalic(true);
        remaining = remaining.substring(italicMatch[0].length);
        matched = true;
      }
    }
    
    // Check for inline code (`code`)
    if (!matched) {
      const codeMatch = remaining.match(/^`([^`]+)`/);
      if (codeMatch) {
        element.appendText(codeMatch[1])
          .setFontFamily('Courier New')
          .setBackgroundColor('#f5f5f5');
        remaining = remaining.substring(codeMatch[0].length);
        matched = true;
      }
    }
    
    // Check for links [text](url)
    if (!matched) {
      const linkMatch = remaining.match(/^\[([^\]]+)\]\(([^)]+)\)/);
      if (linkMatch) {
        element.appendText(linkMatch[1]).setLinkUrl(linkMatch[2]);
        remaining = remaining.substring(linkMatch[0].length);
        matched = true;
      }
    }
    
    // If no markdown pattern matched, add the next character as plain text
    if (!matched) {
      // Find the next potential markdown character
      const nextSpecial = remaining.search(/[*_`\[]/);
      if (nextSpecial > 0) {
        element.appendText(remaining.substring(0, nextSpecial));
        remaining = remaining.substring(nextSpecial);
      } else {
        element.appendText(remaining);
        remaining = '';
      }
    }
  }
}

function replaceTokenWithFormattedList(body, token, items) {
  const range = body.findText(escapeForFind(token));
  if (!range) return;
  
  const elem = range.getElement();
  const para = elem.getParent().asParagraph();
  const insertIndex = body.getChildIndex(para);
  body.removeChild(para);
  
  items.forEach((item, i) => {
    const listItem = body.insertListItem(insertIndex + i, '');
    listItem.setGlyphType(DocumentApp.GlyphType.BULLET);
    parseInlineMarkdown(listItem, item);
  });
}

function replaceTokenWithTable(body, token, rows) {
  const range = body.findText(escapeForFind(token));
  if (!range) return;
  
  const elem = range.getElement();
  const para = elem.getParent().asParagraph();
  const insertIndex = body.getChildIndex(para);
  body.removeChild(para);
  
  const table = body.insertTable(insertIndex, rows.map(r => r.map(c => c == null ? '' : String(c))));
  table.setBorderWidth(1);
  
  // Format header row
  if (rows.length > 0) {
    const headerRow = table.getRow(0);
    for (let i = 0; i < headerRow.getNumCells(); i++) {
      const cell = headerRow.getCell(i);
      cell.setBackgroundColor('#f0f0f0');
      const text = cell.editAsText();
      text.setBold(true);
      // Parse any markdown in table cells
      const cellContent = rows[0][i];
      if (cellContent && typeof cellContent === 'string') {
        parseInlineMarkdownInText(text, cellContent);
      }
    }
  }
  
  // Parse markdown in data cells
  for (let r = 1; r < rows.length; r++) {
    const row = table.getRow(r);
    for (let c = 0; c < row.getNumCells(); c++) {
      const cell = row.getCell(c);
      const text = cell.editAsText();
      const cellContent = rows[r][c];
      if (cellContent && typeof cellContent === 'string') {
        parseInlineMarkdownInText(text, cellContent);
      }
    }
  }
}

function parseInlineMarkdownInText(textElement, content) {
  // For table cells, we need to work with Text objects differently
  textElement.setText('');
  
  let remaining = String(content);
  let currentIndex = 0;
  
  while (remaining.length > 0) {
    let matched = false;
    
    // Check for bold
    const boldMatch = remaining.match(/^\*\*([^*]+)\*\*/);
    if (boldMatch) {
      textElement.appendText(boldMatch[1]);
      textElement.setBold(currentIndex, currentIndex + boldMatch[1].length - 1, true);
      currentIndex += boldMatch[1].length;
      remaining = remaining.substring(boldMatch[0].length);
      matched = true;
    }
    
    // Check for italic
    if (!matched) {
      const italicMatch = remaining.match(/^[*_]([^*_]+)[*_]/);
      if (italicMatch) {
        textElement.appendText(italicMatch[1]);
        textElement.setItalic(currentIndex, currentIndex + italicMatch[1].length - 1, true);
        currentIndex += italicMatch[1].length;
        remaining = remaining.substring(italicMatch[0].length);
        matched = true;
      }
    }
    
    // If no pattern matched, add plain text
    if (!matched) {
      const nextSpecial = remaining.search(/[*_]/);
      if (nextSpecial > 0) {
        const plainText = remaining.substring(0, nextSpecial);
        textElement.appendText(plainText);
        currentIndex += plainText.length;
        remaining = remaining.substring(nextSpecial);
      } else {
        textElement.appendText(remaining);
        currentIndex += remaining.length;
        remaining = '';
      }
    }
  }
}

/* ----------------- Helper Functions ------------------ */

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function coerceStrings(obj, keys) {
  keys.forEach(k => {
    if (Object.prototype.hasOwnProperty.call(obj, k) && obj[k] != null && typeof obj[k] !== 'string') {
      obj[k] = String(obj[k]);
    }
  });
}

function fallback(v, dflt) {
  return (v === null || v === undefined || v === '') ? dflt : v;
}

function escapeForFind(token) {
  return token.replace(/([\\^$.*+?()[\]{}|\-])/g, '\\$1');
}

function sanitizeForDocs(value) {
  return String(value).replace(/\$/g, '\\$');
}
