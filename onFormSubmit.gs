function onFormSubmit(e) {
  Logger.log("Form submission detected. Calling runStockChecks for full stock processing.");
  // The full logic for adding new items and checking for missing ones is now in runStockChecks.
  runStockChecks();
}