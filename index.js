function saveTaxEmailsToDrive() {
  var taxLabel = 'GmailContactLabel/Tax';     // Gmail label to filter on
  var processedLabel = 'ProcessedTaxEmails';  // applied to threads after processing

  // Always target the current AU financial year
  var today = new Date();
  // Months are 0-based; July = 6
  var fyEndYear = (today.getMonth() >= 6) ? today.getFullYear() + 1 : today.getFullYear();

  // Current FY window: 1 Jul (fyEndYear-1) -> 30 Jun (fyEndYear), exclusive upper bound 1 Jul fyEndYear
  var fyStart = new Date(fyEndYear - 1, 6, 1);     // 1 Jul previous calendar year
  var fyEndExclusive = new Date(fyEndYear, 6, 1);  // 1 Jul fyEndYear (exclusive)
  var afterBoundary = addDays(fyStart, -1);        // 30 Jun before FY start
  var beforeBoundary = fyEndExclusive;             // 1 Jul after FY end

  // Folder path uses FY end year (e.g., .../Tax/2026)
  var driveFolderPath = 'Filing Cabinet/Tax/' + fyEndYear;

  // Gmail query: include 1 Jul → 30 Jun by using after: (30 Jun prior) and before: (1 Jul next)
  var query =
    '(' +
      'label:"' + taxLabel + '" ' +
      'after:' + formatQueryDate(afterBoundary) + ' ' +
      'before:' + formatQueryDate(beforeBoundary) + ' ' +
      'AND (subject:invoice OR subject:statement OR subject:receipt OR subject:subscription)' +
    ') ' +
    'AND NOT label:"' + processedLabel + '"';

  logOutput(3, 'Query: ' + query);

  var threads = GmailApp.search(query);
  logOutput(1, 'There are ' + threads.length + ' threads.');

  var rootFolder = getOrCreateFolder(driveFolderPath);
  var processedLabelObj = getOrCreateGmailLabel(processedLabel);

  threads.forEach(function(thread) {
    var messages = thread.getMessages();
    messages.forEach(function(message) {
      var sender = message.getFrom();
      var senderName = extractSenderName(sender);
      var senderFolder = getOrCreateFolder(driveFolderPath + '/' + senderName);

      // Include actual attachments; ignore inline images to avoid noise
      var attachments = message.getAttachments({ includeInlineImages: false, includeAttachments: true });
      var savedCount = 0;
      var fileCounter = 1;

      if (attachments.length > 0) logOutput(3, 'We have ' + attachments.length + ' attachments!');

      attachments.forEach(function(att) {
        var ct = String(att.getContentType()).toLowerCase();
        if (ct.indexOf('application/pdf') === 0) {
          var dateStr = formatFileDate(message.getDate());
          var newFileName = dateStr + '_' + fileCounter + '.pdf';
          senderFolder.createFile(att).setName(newFileName);
          savedCount++;
          fileCounter++;
        }
      });

      // If no PDFs saved, generate a PDF from the email body (plain text)
      if (savedCount === 0) {
        logOutput(3, 'No PDF attachments saved. Creating PDF from email HTML body.');
        var dateStr = formatFileDate(message.getDate());
        var emailHtml = message.getBody(); // prefer HTML
        var pdfBlob = generatePDFFromEmailHtml(emailHtml, dateStr + '_' + fileCounter);
        senderFolder.createFile(pdfBlob);
        fileCounter++;
      }
    });

    // Label the thread so we don't reprocess it
    thread.addLabel(processedLabelObj);
  });
}

/* Log Level
1 = Threads Processed
2 = Threads Processing
3 = Everything
*/
var logLevel = 1;
function logOutput(level, msg) {
  if (level <= logLevel) Logger.log(msg);
}

// Format dates for Gmail search (yyyy/MM/dd) using the script timezone
function formatQueryDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd');
}

// Format dates for filenames (yyyy-MM-dd) using the script timezone
function formatFileDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function addDays(d, days) {
  var n = new Date(d.getTime());
  n.setDate(n.getDate() + days);
  return n;
}

function getOrCreateFolder(path) {
  var folders = path.split('/');
  var parent = DriveApp.getRootFolder();
  folders.forEach(function(folderName) {
    var it = parent.getFoldersByName(folderName);
    parent = it.hasNext() ? it.next() : parent.createFolder(folderName);
  });
  return parent;
}

function extractSenderName(sender) {
  // Try "Name" <email@domain>
  var matchQuoted = sender.match(/"(.*?)"/);
  if (matchQuoted && matchQuoted[1]) return sanitizeName(matchQuoted[1]);

  // Try Name <email@domain>
  var matchAngle = sender.match(/^(.*?)\s*<[^>]+>/);
  if (matchAngle && matchAngle[1]) return sanitizeName(matchAngle[1]);

  // Fallback to raw
  return sanitizeName(sender);
}

function sanitizeName(name) {
  return String(name).replace(/[<>:"/\\|?*]+/g, '_').trim();
}

function getOrCreateGmailLabel(labelName) {
  var label = GmailApp.getUserLabelByName(labelName);
  return label ? label : GmailApp.createLabel(labelName);
}

function generatePDF(content, fileName) {
  var html =
    "<html><head><meta charset='UTF-8'></head><body>" +
    '<pre style="white-space: pre-wrap; font-family: monospace;">' +
    sanitizeHtml(content) +
    '</pre></body></html>';
  var blob = Utilities.newBlob(html, 'text/html', fileName + '.html').getAs('application/pdf');
  return blob.setName(fileName + '.pdf');
}

// Minimal HTML escaping for the plain text body → safe inside <pre>
function sanitizeHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function generatePDFFromHtml(htmlContent, fileName) {
  // Minimal print CSS to keep things readable and avoid clipping
  var wrapper =
    '<html><head><meta charset="UTF-8">' +
    '<style>' +
    '@page { size: A4; margin: 12mm; }' +
    'body { font-family: Arial, sans-serif; font-size: 12px; }' +
    'img { max-width: 100%; height: auto; }' +
    'table { border-collapse: collapse; width: 100%; }' +
    'th, td { border: 1px solid #ddd; padding: 4px; vertical-align: top; }' +
    'pre, code { white-space: pre-wrap; word-wrap: break-word; }' +
    '</style></head><body>' +
    htmlContent +
    '</body></html>';

  // HtmlService preserves much more formatting than Utilities.newBlob(...).getAs(...)
  var htmlOutput = HtmlService.createHtmlOutput(wrapper).setWidth(800).setHeight(1000);
  return htmlOutput.getAs('application/pdf').setName(fileName + '.pdf');
}

// Extract inner <body> HTML, remove outer <html>/<head>, strip scripts, and normalize a few quirks
function sanitizeEmailHtml(html) {
  if (!html) return '';
  // 1) Pull inner body if present
  var bodyMatch = html.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
  var inner = bodyMatch ? bodyMatch[1] : html;

  // 2) Drop any nested outer tags that can break HtmlService
  inner = inner.replace(/<\/?(?:html|head)\b[^>]*>/gi, '');

  // 3) Remove scripts entirely
  inner = inner.replace(/<script[\s\S]*?<\/script>/gi, '');

  // 4) Remove duplicate nested <tr><tr> (common in marketing HTML)
  inner = inner.replace(/<tr\b[^>]*>\s*<tr\b/gi, '<tr');

  // 5) Make sure images don’t overflow the page width when printed
  // (HtmlService respects inline styles; this is a light touch.)
  inner = inner.replace(/<img\b([^>]*)>/gi, function(_, attrs) {
    if (/style=/.test(attrs)) return '<img' + attrs + '>';
    return '<img' + attrs + ' style="max-width:100%;height:auto;">';
  });

  return inner;
}

// Try HtmlService first; if it throws, fall back to Blob->PDF conversion
function generatePDFFromEmailHtml(emailHtml, fileName) {
  var content = sanitizeEmailHtml(emailHtml) || wrapPlainTextAsHtml('(no body)');
  var wrapper =
    '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<style>' +
    '@page { size: A4; margin: 12mm; }' +
    'body { font-family: Arial, sans-serif; font-size: 12px; }' +
    'table { border-collapse: collapse; width: 100%; }' +
    'th, td { border: 1px solid #ddd; padding: 4px; vertical-align: top; }' +
    'pre, code { white-space: pre-wrap; word-wrap: break-word; }' +
    '</style></head><body>' + content + '</body></html>';

  try {
    var out = HtmlService.createHtmlOutput(wrapper).setWidth(800).setHeight(1000);
    return out.getAs('application/pdf').setName(fileName + '.pdf');
  } catch (err) {
    // Fallback converter is more tolerant of odd email HTML
    var blob = Utilities.newBlob(wrapper, 'text/html', fileName + '.html');
    return blob.getAs('application/pdf').setName(fileName + '.pdf');
  }
}

// Plain-text fallback (rarely used now)
function wrapPlainTextAsHtml(text) {
  return '<pre style="white-space:pre-wrap;word-wrap:break-word;">' +
         sanitizeHtml(text || '') +
         '</pre>';
}

