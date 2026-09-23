// Step 04b creates the project OneDrive folders from configured template folders.
// run() copies CAD and Client Files templates into their configured destinations, then renames them for the project.
// Helper functions start Graph copy operations, poll asynchronous completion, and rename copied folders.

const path = require('path');
const config = require('../../config');
const { childLogger } = require('../utils/logger');
const { pollAsyncOperation } = require('../utils/pollAsyncOperation');
const { isRetryableHttpError } = require('../utils/retry');

async function run(ctx) {
  const log = childLogger(ctx, 'step04b');
  const finalName = `${ctx.projectName} - ${ctx.projectNumber}`;

  log.info({
    oneDriveUserId: config.graph.oneDriveUserId,
    finalName,
    cadTemplatePath: config.oneDrive.cadTemplatePath,
    cadDestinationPath: config.oneDrive.cadDestinationPath,
    clientTemplatePath: config.oneDrive.clientTemplatePath,
    clientDestinationPath: config.oneDrive.clientDestinationPath
  }, 'starting OneDrive folder creation');

  ctx.folderIds.oneDrive = ctx.folderIds.oneDrive || {};
  ctx.folderUrls = ctx.folderUrls || {};
  ctx.folderUrls.oneDrive = ctx.folderUrls.oneDrive || {};
  ctx.oneDriveFolderResults = ctx.oneDriveFolderResults || {};

  const failures = [];

  await copyRecordFolder({
    ctx,
    log,
    failures,
    key: 'cad',
    templatePath: config.oneDrive.cadTemplatePath,
    destinationPath: config.oneDrive.cadDestinationPath,
    finalName,
    label: 'CAD Files'
  });

  await copyRecordFolder({
    ctx,
    log,
    failures,
    key: 'client',
    templatePath: config.oneDrive.clientTemplatePath,
    destinationPath: config.oneDrive.clientDestinationPath,
    finalName,
    label: 'Client Files'
  });

  if (failures.length) {
    const message = failures.map((failure) => `${failure.label}: ${failure.message}`).join('; ');
    throw new Error(`OneDrive folder creation incomplete. ${message}`);
  }

  log.info({
    cadFolderId: ctx.folderIds.oneDrive.cad,
    cadFolderUrl: ctx.folderUrls.oneDrive.cad,
    clientFolderId: ctx.folderIds.oneDrive.client,
    clientFolderUrl: ctx.folderUrls.oneDrive.client
  }, 'created OneDrive project folders');
  return ctx;
}

async function copyRecordFolder({ ctx, log, failures, key, templatePath, destinationPath, finalName, label }) {
  try {
    const folder = await copyPollRename({ ctx, log, templatePath, destinationPath, finalName, label });
    ctx.folderIds.oneDrive[key] = folder.id;
    ctx.folderUrls.oneDrive[key] = folder.webUrl;
    ctx.oneDriveFolderResults[key] = { status: 'created', id: folder.id, webUrl: folder.webUrl };
  } catch (error) {
    const message = humanizeError(error);
    ctx.oneDriveFolderResults[key] = { status: 'failed', message };
    failures.push({ label, message });
    log.error({ err: error, label, templatePath, destinationPath, finalName }, 'OneDrive folder creation failed; continuing with remaining folders');
  }
}

async function copyPollRename({ ctx, log, templatePath, destinationPath, finalName, label }) {
  try {
    const graph = ctx.clients.oneDriveGraph || ctx.clients.graph;
    log.info({ label, templatePath, destinationPath }, 'resolving OneDrive template and destination folders');

    const template = (await graph.resolveDriveItemByPath(templatePath)).data;
    const destination = (await graph.resolveDriveItemByPath(destinationPath)).data;
    log.info({
      label,
      templateId: template.id,
      templateName: template.name,
      templateUrl: template.webUrl,
      destinationId: destination.id,
      destinationName: destination.name,
      destinationUrl: destination.webUrl
    }, 'resolved OneDrive template and destination folders');

    const temporaryName = `${path.basename(templatePath)} - ${ctx.runId}`.slice(0, 120);
    log.info({ label, templateId: template.id, destinationId: destination.id, temporaryName, finalName }, 'starting OneDrive template folder copy');

    const copyResponse = await graph.copyDriveItem({
      itemId: template.id,
      parentReferenceId: destination.id,
      name: temporaryName
    });
    const location = copyResponse.headers.get('location');

    if (!location) {
      throw new Error(`Graph copy for ${label} did not return a Location header`);
    }

    log.info({ label, status: copyResponse.status, temporaryName }, 'OneDrive copy accepted by Graph; polling for completion');

    const copied = await pollAsyncOperation({
      log,
      poll: () => pollGraphCopy(location, log),
      isComplete: (result) => ['completed', 'complete'].includes(String(result.status || '').toLowerCase()),
      intervalMs: 3000,
      timeoutMs: 300000
    });

    const copiedId = copied.resourceId || copied.resourceLocation?.split('/').pop();
    if (!copiedId) {
      throw new Error(`Graph copy for ${label} completed without a resource id`);
    }

    log.info({ label, copiedId, copyStatus: copied.status, temporaryName }, 'OneDrive copy completed; renaming copied folder');

    const renamed = (await graph.patchDriveItem(copiedId, { name: finalName })).data;
    log.info({
      label,
      folderId: renamed.id,
      folderName: renamed.name,
      folderUrl: renamed.webUrl,
      finalName,
      sourceTemplateUrl: template.webUrl,
      destinationUrl: destination.webUrl
    }, 'copied and renamed OneDrive template folder');
    return { id: renamed.id, webUrl: renamed.webUrl };
  } catch (error) {
    log.error({ err: error, label, templatePath, destinationPath, finalName }, 'OneDrive template folder copy failed');
    throw error;
  }
}

async function pollGraphCopy(location, log) {
  try {
    const response = await fetch(location);
    const text = await response.text();
    if (!response.ok) {
      const error = new Error(formatGraphCopyPollError(response, text));
      error.status = response.status;
      throw error;
    }
    return text ? JSON.parse(text) : {};
  } catch (error) {
    if (isRetryablePollError(error)) {
      if (log) {
        log.warn({ err: error, location }, 'OneDrive copy poll hit a transient error; will retry until copy timeout');
      }
      return { status: 'pending' };
    }

    throw error;
  }
}

function isRetryablePollError(error) {
  return isRetryableHttpError(error) || error.message === 'fetch failed';
}

function formatGraphCopyPollError(response, text) {
  const detail = summarizeResponseText(text);
  const statusText = response.statusText ? ` ${response.statusText}` : '';
  return detail
    ? `Graph copy poll failed: ${response.status}${statusText} - ${detail}`
    : `Graph copy poll failed: ${response.status}${statusText}`;
}

function summarizeResponseText(text) {
  const normalized = String(text || '')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  if (!normalized) {
    return '';
  }

  const technicalDetails = normalized.match(/Technical details:\s*([^]+)$/i)?.[1];
  if (technicalDetails) {
    return technicalDetails.slice(0, 180).trim();
  }

  return normalized.slice(0, 180).trim();
}

function humanizeError(error) {
  return error?.message || 'Unknown OneDrive error';
}

module.exports = { run };
