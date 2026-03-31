// ========================================
// DEDUPLICATION
// Check if a message has already been processed
// ========================================

/**
 * Check if a message ID is already in the Queue sheet.
 * Deduplicates by message_id only — label-agnostic so relabeling
 * a message does not reprocess it unless the queue row is removed.
 */
function isProcessed(messageId) {
  const { headers, rows } = getQueueData();
  if (rows.length === 0) return false;
  const col = headers.indexOf('message_id');
  return rows.some(row => row[col] === messageId);
}
