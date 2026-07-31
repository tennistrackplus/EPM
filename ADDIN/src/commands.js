Office.onReady(() => {
  // ready
});

function hidePane(event) {
  Office.addin.hideAsTaskpane();
  event.completed();
}

// Register the function so the manifest's ExecuteFunction action can find it
Office.actions.associate("hidePane", hidePane);
