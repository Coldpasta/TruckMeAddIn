
export function showTaskpane(event: Office.AddinCommands.Event)  {
  try {
    Office.addin.showAsTaskpane();
  } catch (e) {
    console.error("showTaskpane failed", e);
  } finally {
    event.completed();
  }
}

// associate for legacy manifest command behavior
try {
  Office.actions.associate("showTaskpane", showTaskpane);
} catch (e) {
  // at runtime this is fine in dev
  console.warn("Office.actions.associate may not be available in this context", e);
}