/*
 * Excel Add-in command functions
 */

/* global Office */

Office.onReady(() => {
  // Register functions
  Office.actions.associate("actionHandler", actionHandler);
});

/**
 * Shows a notification when add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function actionHandler(event: Office.AddinCommands.Event): void {
  // Your action code goes here

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

// Make sure functions are available in global scope
(window as any).actionHandler = actionHandler;

export {};
