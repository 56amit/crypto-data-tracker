function fetchCryptoData() {
  const apiUrl = "https://api.coingecko.com/api/v3/coins/markets?vs_currency=usd&order=market_cap_desc&per_page=15&page=1&meta_info=VR6";
  try {
    const response = UrlFetchApp.fetch(apiUrl);
    Logger.log(response.getContentText());
    const data = JSON.parse(response.getContentText());

    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    const currentSheet = sheet.getSheetByName("Current Prices");
    const historySheet = sheet.getSheetByName("Price History");

    currentSheet.getRange(2, 1, currentSheet.getLastRow(), 7).clearContent();
    const now = new Date();
    const currentData = [];
    const historyData = [];

    for (let i = 0; i < data.length; i++) {
      const coin = data[i];
      const row = [
        coin.id,
        coin.symbol.toUpperCase(),
        coin.name,
        coin.current_price,
        coin.market_cap,
        coin.price_change_percentage_24h,
        now.toLocaleString()
      ];
      currentData.push(row);

      const histRow = [
        now.toLocaleString(),
        coin.id,
        coin.symbol.toUpperCase(),
        coin.name,
        coin.current_price,
        coin.market_cap,
        coin.price_change_percentage_24h,
        "HIST" + Utilities.getUuid().substring(0, 5).toUpperCase()
      ];
      historyData.push(histRow);
    }

    currentSheet.getRange(2, 1, currentData.length, currentData[0].length).setValues(currentData);
    historySheet.getRange(historySheet.getLastRow() + 1, 1, historyData.length, historyData[0].length).setValues(historyData);

  } catch (e) {
    SpreadsheetApp.getUi().alert("API Error: " + e.message);
  }
}
