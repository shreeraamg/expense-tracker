function getCategoryForDescription(description) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const keywordsTab = sheet.getSheetByName("Keywords");
  const data = keywordsTab.getDataRange().getValues();

  const lowerDesc = description.toLowerCase();

  for (let i = 1; i < data.length; i++) {
    const keyword = data[i][0].toString().toLowerCase();
    if (lowerDesc.includes(keyword)) {
      return data[i][1];
    }
  }

  return "OTHER EXPENSE";
}

function doPost(request) {
  try {
    const body = JSON.parse(request.postData.contents);

    const amount = body.amount;
    const description = body.description;
    const account = body.account;

    const category = getCategoryForDescription(description);

    const today = new Date();
    const date = today.getFullYear() + "-" +
                 String(today.getMonth() + 1).padStart(2, "0") + "-" +
                 String(today.getDate()).padStart(2, "0");

    const currency = "INR";
    const type = "OUT";
    const notes = "";

    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    const tabName = "Expenses_" + account;
    const expensesTab = sheet.getSheetByName(tabName);

    if (!expensesTab) {
      return ContentService
        .createTextOutput(JSON.stringify({ 
          status: "error", 
          message: "Tab not found: " + tabName 
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    expensesTab.appendRow([date, description, amount, currency, type, category, notes]);

    return ContentService
      .createTextOutput(JSON.stringify({ 
        status: "success", 
        category: category 
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ 
        status: "error", 
        message: error.message 
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function testDoPost() {
  const mockRequest = {
    postData: {
      contents: JSON.stringify({
        amount: 450,
        description: "Zomato dinner",
        account: "Axis_CC"
      })
    }
  };

  const result = doPost(mockRequest);
  Logger.log(result.getContent());
}
