// ====================================================================================
// 📁 Helper: Highlight & Duplicate Topik - Resi Automation System
// ====================================================================================

// ✅ Highlight sel yang memiliki duplikat topik dalam baris yang sama
function highlightDuplicateTopik() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");

  const topicStartCol = 16; // Kolom P
  const topicEndCol = 30;   // Kolom AD
  const numRows = sheet.getLastRow() - 1;

  const range = sheet.getRange(2, topicStartCol, numRows, topicEndCol - topicStartCol + 1);
  const displayData = range.getDisplayValues();

  // 🔄 Reset background warna
  range.setBackground(null);

  for (let row = 0; row < displayData.length; row++) {
    const originalTopics = displayData[row];
    const normalizedTopics = originalTopics.map(item => item.trim().toLowerCase());
    const duplicates = findDuplicates(normalizedTopics);

    for (let i = 0; i < normalizedTopics.length; i++) {
      if (normalizedTopics[i] !== "" && duplicates.includes(normalizedTopics[i])) {
        sheet.getRange(row + 2, topicStartCol + i).setBackground("#f4cccc"); // row+2 karena data mulai baris ke-2
      }
    }
  }
}

// ✅ Utility: Temukan item duplikat dalam array
function findDuplicates(arr) {
  const counts = {};
  const duplicates = [];

  arr.forEach(item => {
    if (item && item !== "") {
      counts[item] = (counts[item] || 0) + 1;
      if (counts[item] === 2) {
        duplicates.push(item);
      }
    }
  });

  return duplicates;
}
