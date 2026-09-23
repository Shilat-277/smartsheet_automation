// Step 02e verifies whether the project contract attachment appears signed.
// run() respects DRY_RUN, finds a likely contract attachment, downloads/parses the PDF, writes the signed status, or records manual review.
// Helper functions find attachments, classify the first PDF line, update checklist rows, and build Smartsheet cell payloads.

const pdfParse = require('pdf-parse');
const config = require('../../config');
const { childLogger } = require('../utils/logger');
const runStateStore = require('../utils/runStateStore');
const { buildCell, cellValue, findColumnByTitle, findRowByPrimaryValue } = require('../utils/smartsheetSheet');

const MISSING_CONTRACT_MESSAGE = 'No contract found. No Letter of Agreement attachment was uploaded for this project.';

async function run(ctx) {
  const log = childLogger(ctx, 'step02e');

  if (isPatersonProject(ctx)) {
    const reason = 'Contract verification skipped for Paterson project';
    ctx.contract = { skipped: true, reason };
    log.info({ projectType: ctx.projectType, patersonProject: ctx.patersonProject }, reason);
    return ctx;
  }

  if (config.dryRun) {
    log.info('DRY_RUN enabled; contract verification marked for manual review');
    await markNeedsReview(ctx, 'DRY_RUN enabled');
    return ctx;
  }

  try {
    if (!ctx.projectNumber) {
      await markNeedsReview(ctx, 'Project number was not provided; skipping automatic contract attachment verification');
      return ctx;
    }

    const attachment = await findSignedContractAttachment(ctx);
    if (!attachment) {
      await markNeedsReview(ctx, MISSING_CONTRACT_MESSAGE, { missingContractAttachment: true });
      return ctx;
    }
    if (attachment.invalidFileType) {
      await markNeedsReview(ctx, invalidFileTypeMessage(attachment), {
        attachmentName: attachment.name || '',
        attachmentFileType: attachment.fileType || 'unsupported file',
        invalidFileType: true
      });
      return ctx;
    }

    const downloadUrl = await resolveAttachmentDownloadUrl(ctx, attachment);
    const graph = ctx.clients.graph;
    const buffer = await graph.download(downloadUrl);
    const parsed = await pdfParse(buffer, { max: 1 });
    const signedLine = findSignedLine(parsed.text || '');
    const firstLine = (parsed.text || '').split(/\r?\n/).find(Boolean) || '';
    const signed = Boolean(signedLine);
    const contract = {
      signed,
      attachmentId: attachment.id,
      attachmentName: attachment.name || '',
      attachmentUrl: downloadUrl,
      signedLine,
      firstLine
    };

    await writeSignedStatus(ctx, signed ? 'Yes' : '');
    ctx.contract = contract;
    log.info({ signed, attachmentId: attachment.id }, 'verified signed contract PDF first line');
    if (!signed) {
      addContractReviewAttachment(ctx, attachment, buffer);
      await markNeedsReview(ctx, 'Contract attachment was found but did not appear signed', contract);
    }
    return ctx;
  } catch (error) {
    log.warn({ err: error }, 'contract verification needs manual review');
    await markNeedsReview(ctx, error.message);
    return ctx;
  }
}

async function findSignedContractAttachment(ctx) {
  const smartsheet = ctx.clients.smartsheet;
  const { sheetId, rowId } = await resolveProjectRowContext(ctx);
  const response = await getRowAttachments(smartsheet, sheetId, rowId);
  const attachments = response.data.data || [];
  const likelyContractAttachments = attachments.filter(isLikelyContractAttachment);
  const pdfAttachments = attachments.filter(isPdfAttachment);
  const attachment = likelyContractAttachments.find(isPdfAttachment) || pdfAttachments[0];
  if (!attachment && likelyContractAttachments.length) {
    const invalidAttachment = likelyContractAttachments[0];
    return {
      ...invalidAttachment,
      sheetId,
      invalidFileType: true,
      fileType: describeAttachmentFileType(invalidAttachment)
    };
  }
  return attachment ? { ...attachment, sheetId } : null;
}

async function getRowAttachments(smartsheet, sheetId, rowId) {
  try {
    return await smartsheet.get(`/sheets/${sheetId}/rows/${rowId}/attachments`);
  } catch (error) {
    if (isNotFoundError(error)) {
      return { data: { data: [] } };
    }
    throw error;
  }
}

function isNotFoundError(error) {
  return error?.status === 404 || /^not found$/i.test(String(error?.message || '').trim());
}

function isLikelyContractAttachment(attachment) {
  return /letter|agreement|contract|loa/i.test(attachment.name || '');
}

function isPdfAttachment(attachment) {
  return /\.pdf$/i.test(attachment.name || '') || /\/pdf$/i.test(attachment.mimeType || attachment.contentType || '');
}

function describeAttachmentFileType(attachment) {
  const name = String(attachment.name || '').trim();
  const extension = name.match(/\.([^.]+)$/)?.[1]?.toLowerCase();
  const labels = {
    csv: 'spreadsheet file',
    gif: 'image file',
    heic: 'image file',
    jpeg: 'image file',
    jpg: 'image file',
    numbers: 'spreadsheet file',
    png: 'image file',
    tif: 'image file',
    tiff: 'image file',
    webp: 'image file',
    xls: 'Excel file',
    xlsm: 'Excel file',
    xlsx: 'Excel file'
  };
  return labels[extension] || attachment.mimeType || attachment.contentType || (extension ? `${extension.toUpperCase()} file` : 'unsupported file');
}

function invalidFileTypeMessage(attachment) {
  const fileType = attachment.fileType || describeAttachmentFileType(attachment);
  const fileName = attachment.name ? ` (${attachment.name})` : '';
  return `The contract attachment was not a valid file for automatic review. You uploaded ${articleFor(fileType)} ${fileType}${fileName}; please upload the contract as a PDF file.`;
}

function articleFor(value) {
  return /^[aeiou]/i.test(String(value || '')) ? 'an' : 'a';
}

function isPatersonProject(ctx) {
  if (/^(yes|y|true)$/i.test(String(ctx.patersonProject || '').trim())) {
    return true;
  }

  return /\bpat{1,2}erson\b/i.test(String(ctx.projectType || ctx.projectVertical || '').trim());
}

async function resolveAttachmentDownloadUrl(ctx, attachment) {
  if (attachment.url) {
    return attachment.url;
  }

  const sheetId = attachment.sheetId || normalizeSheetId(config.smartsheet.masterProjectListSheetId);
  const response = await ctx.clients.smartsheet.get(`/sheets/${sheetId}/attachments/${attachment.id}`);
  const url = response.data?.url;
  if (!url) {
    throw new Error(`Attachment ${attachment.id || attachment.name || 'unknown'} did not include a download URL`);
  }
  return url;
}

async function resolveProjectRowContext(ctx) {
  const sheetId = normalizeSheetId(ctx.masterProjectListSheetId || config.smartsheet.masterProjectListSheetId);

  if (!sheetId) {
    throw new Error('Master Project List sheet id was not provided');
  }

  if (ctx.masterProjectRowId) {
    return { sheetId, rowId: ctx.masterProjectRowId };
  }

  const sheet = (await ctx.clients.smartsheet.get(`/sheets/${sheetId}`)).data;
  const projectNumberColumn = findColumnByTitle(sheet, config.columns.masterProjectNumber, ['Project Number']);

  if (!projectNumberColumn) {
    throw new Error(`Project List is missing project number column: ${config.columns.masterProjectNumber}`);
  }

  const expectedProjectNumber = String(ctx.projectNumber).trim();
  const row = (sheet.rows || []).find((candidate) => {
    return String(cellValue(candidate, projectNumberColumn.id) || '').trim() === expectedProjectNumber;
  });

  if (!row) {
    throw new Error(`Project ${ctx.projectNumber} was not found on the Project List`);
  }

  ctx.masterProjectListSheetId = sheetId;
  ctx.masterProjectRowId = row.id;
  return { sheetId, rowId: row.id };
}

function normalizeSheetId(value) {
  const text = String(value || '').trim();
  const match = text.match(/\/sheets\/([^/?#]+)/i);
  return match?.[1] || text;
}

function classifySignedStatus(firstLine) {
  const signedPrefix = String(config.signedKeyword || '').trim().replace(/:\s*$/, '');
  const signedPattern = new RegExp(`^${escapeRegExp(signedPrefix)}:\\s*\\S+\\s*$`);
  return signedPattern.test(String(firstLine || '').trim());
}

function findSignedLine(text) {
  return String(text || '')
    .split(/\r?\n/)
    .map((line) => line.trim())
    .find(classifySignedStatus) || '';
}

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

async function writeSignedStatus(ctx, value) {
  const smartsheet = ctx.clients.smartsheet;
  const sheetId = ctx.sheetIds.gen009Checklist;
  const sheet = (await smartsheet.get(`/sheets/${sheetId}`)).data;
  const valueColumn = findColumnByTitle(sheet, config.columns.checklistValue, ['Status', 'Done']);
  const row = findRowByPrimaryValue(sheet, 'Verify letter of agreement was attached')
    || findRowByPrimaryValue(sheet, config.rows.signed, ['Signed LOA', 'Signed Letter of Agreement', 'Contract Signed']);

  if (!valueColumn || !row) {
    throw new Error('Checklist is missing signed-status target column or row');
  }

  await smartsheet.put(`/sheets/${sheetId}/rows`, [{ id: row.id, cells: [buildCell(valueColumn, checklistStatusValue(valueColumn, value))] }]);
  ctx.checklistRowMap.signed = row.id;
}

function checklistStatusValue(column, value) {
  const options = column.options || [];
  if (value === '') {
    return '';
  }
  if (options.includes(value)) {
    return value;
  }
  if (String(value).toLowerCase() === 'yes' && options.includes('Done')) {
    return 'Done';
  }
  return value;
}

function addContractReviewAttachment(ctx, attachment, buffer) {
  ctx.emailAttachments = ctx.emailAttachments || [];
  ctx.emailAttachments.push({
    name: attachment.name || `contract-${ctx.projectNumber || 'review'}.pdf`,
    contentType: 'application/pdf',
    contentBytes: buffer.toString('base64')
  });
}

async function markNeedsReview(ctx, reason, contract = {}) {
  ctx.stepStatus.step02e = 'needs_manual_review';
  ctx.contract = { ...contract, needsManualReview: true, reason };
  ctx.problems = ctx.problems || [];
  ctx.problems.push({ step: 'step02e', message: reason });
  await runStateStore.markStepNeedsManualReview(ctx, 'step02e', { reason });
}

module.exports = { classifySignedStatus, normalizeSheetId, run };
