/**
 * Ribbon handler for Outlook Add-in.
 * Handles the manual trigger from Ribbon button to open APX Taskpane.
 * 
 * This module defines the function associated with the Ribbon button in the manifest.
 * It opens the Taskpane as a dialog with the specified dimensions to match PRD requirements.
 */

(function () {
  "use strict";

  // Office.initialize is included for consistency, though not strictly required for function commands
  Office.initialize = function (_reason) {
    // No specific initialization needed for ribbon handler
  };

  /**
   * Opens the APX Taskpane as a dialog when the Ribbon button is clicked.
   * Uses displayDialogAsync to present the Taskpane with width approximately 350px (35% of dialog width).
   * 
   * @param {Office.AddinCommands.Event} event - The event object passed by Office.js.
   * @returns {void}
   * @throws {Error} If dialog opening fails, but handled internally by Office.js.
   */
  function openAppxTaskpane(event) {
    // Construct Taskpane URL - in production, this should match the manifest's Taskpane URL
    const taskpaneUrl = `${window.location.protocol}//${window.location.host}/taskpane.html`;
    
    // Open as dialog with PRD-specified width (350px approximated as 35% width)
    Office.context.ui.displayDialogAsync(taskpaneUrl, {
      height: window.constants.DEFAULTS.DIALOG_HEIGHT_PERCENT, // Full height for usability
      width: window.constants.DEFAULTS.DIALOG_WIDTH_PERCENT,  // Approximate 350px width as per PRD
      displayInIframe: true
    });
    
    // Complete the event to inform Office.js that processing is done
    event.completed();
  }

  // Associate the function with the Ribbon button defined in manifest.xml
  Office.actions.associate("openAppxTaskpane", openAppxTaskpane);
})();