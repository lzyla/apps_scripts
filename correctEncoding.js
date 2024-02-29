function correctEncodingIssues() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getDataRange();
  var values = range.getValues();

  var correctedValues = values.map(row => {
    return row.map(cell => {
      if (typeof cell == 'string') {
        // Zbiorcze zamiany dla znalezionych problemów z kodowaniem
        return cell.replace(/Ä…/g, 'ą')
                   .replace(/Ä‡/g, 'ć')
                   .replace(/Ä™/g, 'ę')
                   .replace(/Ĺ‚/g, 'ł')
                   .replace(/Ĺ„/g, 'ń')
                   .replace(/Ăł/g, 'ó')
                   .replace(/Ĺ›/g, 'ś')
                   .replace(/ĹĽ/g, 'ż')
                   .replace(/Ĺş/g, 'ź')
                   .replace(/Ä„/g, 'Ą')
                   .replace(/Ä†/g, 'Ć')
                   .replace(/Ä˜/g, 'Ę')
                   .replace(/Ĺ/g, 'Ł')
                   .replace(/Ĺƒ/g, 'Ń')
                   .replace(/Ă“/g, 'Ó')
                   .replace(/Ĺš/g, 'Ś')
                   .replace(/Ĺť/g, 'Ż')
                   .replace(/Ĺą/g, 'Ź')
                   // Dodatkowe specyficzne zamiany
                   .replace(/nastÄ piła/g, 'nastąpiła')
                   .replace(/załÄ czeniu/g, 'załączeniu')
                   .replace(/UrzÄ d/g, 'Urząd')
                   // Dodatkowe poprawki na podstawie nowych przykładów
                   .replace(/sÄ /g, 'są ')
                   .replace(/nastÄpu/g, 'następujące ')
                   .replace(/udostÄp/g, 'udostępniane ');
      } else {
        // Zwraca komórkę bez zmian, jeśli nie jest typu string
        return cell;
      }
    });
  });

  // Aktualizacja wartości w arkuszu
  range.setValues(correctedValues);
}
