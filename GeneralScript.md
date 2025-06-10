```
function createGoogleFormWithQuizFromSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  function col(name) { return headers.indexOf(name); }

  // กำหนดชื่อฟอร์มจากบรรทัดแรกที่ไม่ใช่ header
  var formTitle = data[1][col('Question')] || 'New Quiz Form';
  var form = FormApp.create(formTitle);

  // เปิดเป็น Quiz
  form.setIsQuiz(true);

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[col('Type')] || !row[col('Question')]) continue;

    var type = (row[col('Type')] || '').trim();
    var question = row[col('Question')] || '';
    var desc = row[col('Desc')] || '';
    var img = row[col('Image')] || '';
    var required = String(row[col('Required')]).toUpperCase() === "TRUE";
    var other = String(row[col('Other')]).toUpperCase() === "TRUE";
    var points = parseInt(row[col('Points')], 10) || 0;
    var correctAns = row[col('Correct Answer')] || '';
    var correctFb = row[col('correct feedback')] || '';
    var correctUrl = row[col('correct url')] || '';
    var incorrectFb = row[col('incorrect feedback')] || '';
    var incorrectUrl = row[col('incorrect url')] || '';

    var item = null; // ใช้สำหรับ quiz item

    switch(type) {
      case 'Title':
        form.addTitleItem().setTitle(question);
        break;
      case 'Title and description':
        form.addTitleItem().setTitle(question).setHelpText(desc || "");
        break;
      case 'Section':
        form.addSectionHeaderItem().setTitle(question).setHelpText(desc || "");
        break;
      case 'Short answer':
        item = form.addTextItem().setTitle(question);
        if (desc) item.setHelpText(desc);
        item.setRequired(required);
        break;
      case 'Paragraph':
        item = form.addParagraphTextItem().setTitle(question);
        if (desc) item.setHelpText(desc);
        item.setRequired(required);
        break;
      case 'Multiple choice':
      case 'Checkboxes':
      case 'Dropdown':
        // Collect options (option start, option 2, ..., option end)
        var choices = [];
        for (let j = col('option start'); j <= col('option end'); j++) {
          if (row[j]) choices.push(row[j]);
        }
        // เพิ่ม Other ถ้าต้องการ (เฉพาะ multiple choice/checkboxes)
        if (other && (type === 'Multiple choice' || type === 'Checkboxes')) choices.push('Other');

        // Add item
        if (type === 'Multiple choice') {
          item = form.addMultipleChoiceItem().setTitle(question).setRequired(required);
        } else if (type === 'Checkboxes') {
          item = form.addCheckboxItem().setTitle(question).setRequired(required);
        } else if (type === 'Dropdown') {
          item = form.addListItem().setTitle(question).setRequired(required);
        }

        if (desc) item.setHelpText(desc);

        // Choices สำหรับแต่ละ type
        if (type === 'Multiple choice' || type === 'Checkboxes') {
          item.setChoices(choices.map(opt => item.createChoice(opt)));
        } else if (type === 'Dropdown') {
          item.setChoiceValues(choices);
        }

        // Quiz Logic
        if (points > 0 || correctAns) {
          // Set points
          item.setPoints(points);

          // Correct answer(s)
          if (type === 'Checkboxes') {
            // กรณี Checkbox เฉลยอาจคั่นด้วย comma หรือ | เช่น "Red,Blue"
            var correctSet = correctAns.split(/,|\|/).map(s => s.trim());
            item.setCorrectAnswer(correctSet);
          } else {
            // Multiple choice/Dropdown เฉลย 1 ตัว
            item.setCorrectAnswer(correctAns);
          }

          // Feedback
          var fbCorrect = FormApp.createFeedback().setText(correctFb || '');
          if (correctUrl) fbCorrect.addLink(correctUrl, 'เพิ่มเติม');

          var fbIncorrect = FormApp.createFeedback().setText(incorrectFb || '');
          if (incorrectUrl) fbIncorrect.addLink(incorrectUrl, 'เพิ่มเติม');

          item.setFeedbackForCorrect(fbCorrect);
          item.setFeedbackForIncorrect(fbIncorrect);
        }
        break;
      case 'Date':
        item = form.addDateItem().setTitle(question).setRequired(required);
        if (desc) item.setHelpText(desc);
        break;
      case 'Time':
        item = form.addTimeItem().setTitle(question).setRequired(required);
        if (desc) item.setHelpText(desc);
        break;
      case 'Date and time':
        item = form.addDateTimeItem().setTitle(question).setRequired(required);
        if (desc) item.setHelpText(desc);
        break;
      case 'Linear scale':
        // เพิ่ม logic กำหนด scale (อาจเพิ่ม column ใน sheet เพื่อกำหนด min/max label)
        item = form.addScaleItem()
          .setTitle(question)
          .setBounds(1, 5);
        if (desc) item.setHelpText(desc);
        item.setRequired(required);
        break;
      // Image & Grid types ยังไม่รองรับใน script นี้เพราะข้อจำกัด API/ความซับซ้อน
      default:
        Logger.log('Unknown/unsupported type: ' + type + ' for question: ' + question);
    }
  }
}
```
