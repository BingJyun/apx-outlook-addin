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
})();