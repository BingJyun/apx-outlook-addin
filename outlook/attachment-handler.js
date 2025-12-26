/**
 * Attachment handler for Outlook Add-in.
 * Monitors attachment changes in Compose mode and automatically triggers APX Taskpane for large files.
 * 
 * This module listens for attachment additions/removals and checks for single attachments exceeding
 * the threshold defined in constants. If triggered, it removes the original attachment to prevent
 * double-upload and opens the Taskpane.
 */

(function () {
  "use strict";

  Office.initialize = function (_reason) {
    // Attach event handler for attachment changes in Compose mode
    Office.context.mailbox.item.addHandlerAsync(Office.EventType.AttachmentsChanged, onAttachmentsChanged);
  };

  /**
   * Event handler for attachment changes.
   * Checks if a single attachment exceeds the size threshold and triggers Taskpane opening.
   * 
   * @param {Office.AttachmentsChangedEventArgs} _event - The event arguments from Office.js (unused, marked with _).
   * @returns {Promise<void>}
   * @throws {Error} Propagates errors from Office.js APIs, handled by global error handler.
   */
  async function onAttachmentsChanged(_event) {
    const item = Office.context.mailbox.item;
    
    // Retrieve current attachments using async/await for clean code
    const attachments = await new Promise((resolve, reject) => {
      item.attachments.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(result.error);
        }
      });
    });
    
    // Only trigger for single large attachment as per PRD (supports single file)
    if (attachments.length === 1) {
      const attachment = attachments[0];
      if (attachment.size > window.constants.DEFAULTS.MAX_FILE_SIZE_BYTES) {
        // Remove the original attachment to avoid double-upload
        await new Promise((resolve, reject) => {
          item.removeAttachmentAsync(attachment.id, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              resolve();
            } else {
              reject(result.error);
            }
          });
        });
        
        // Open Taskpane as dialog (same as ribbon handler for consistency)
        const taskpaneUrl = `${window.location.protocol}//${window.location.host}/taskpane.html`;
        Office.context.ui.displayDialogAsync(taskpaneUrl, {
          height: window.constants.DEFAULTS.DIALOG_HEIGHT_PERCENT,
          width: window.constants.DEFAULTS.DIALOG_WIDTH_PERCENT,
          displayInIframe: true
        });
      }
    }
  }
})();