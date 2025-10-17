/**
 * A simple Apps Script Webhook tester.
 * Sends a dummy JSON request to a specified Zapier Webhook URL.
 */
function testZapierWebhook() {
  // Replace with your Zapier Webhook URL from the Zapier setup page
  const ZAPIER_WEBHOOK_URL = 'https://hooks.zapier.com/hooks/catch/13989939/u4wwqbd/';

  // A dummy payload to send to Zapier
  const testPayload = {
    test_id: "APPSCRIPT_TEST_001",
    message: "Hello from Apps Script!",
    timestamp: new Date().toISOString()
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(testPayload)
  };

  try {
    const response = UrlFetchApp.fetch(ZAPIER_WEBHOOK_URL, options);
    const responseCode = response.getResponseCode();
    
    // Log the response from Zapier to verify success
    Logger.log('Webhook sent successfully.');
    Logger.log('Zapier Response Code: ' + responseCode);
    
    // Display a success message in a pop-up window
    SpreadsheetApp.getUi().alert('✅ Webhook sent successfully! Check Zapier to see the request.');
    
  } catch (e) {
    // If an error occurs, log it and display an error message
    Logger.log('Failed to send webhook: ' + e.message);
    SpreadsheetApp.getUi().alert('❌ Failed to send webhook. Please check the URL and permissions. Error: ' + e.message);
  }
}