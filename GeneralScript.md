```
function createGoogleFormWithQuizFromSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  function col(name) { return headers.findIndex(h => h.trim().toLowerCase() === name.trim().toLowerCase()); }

  // ฟอร์มจะชื่อ 'ไม่มีหัวข้อกำหนด' ถ้า cell ว่าง
  var formTitle = (data[1][col('Question')] && String(data[1][col('Question')]).trim()) ? data[1][col('Question')] : 'ไม่มีหัวข้อกำหนด';
  var form = FormApp.create(formTitle);
  form.setIsQuiz(true);

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[col('Type')]) continue; // ไม่สน row ว่าง

    var type = String(row[col('Type')]).trim().toLowerCase();
    var question = row[col('Question')] || '';
    var desc = row[col('Desc')] || '';
    var required = true; // กำหนด default เป็น required ทั้งหมด (ถ้าอยากกำหนดเพิ่มใน sheet ค่อยปรับเพิ่ม)
    var correctAns = row[col('Correct Answer')] || '';
    var item = null;
    var choices = [];

    switch(type) {
      case 'title':
        form.addTitleItem().setTitle(question || 'ไม่มีหัวข้อกำหนด');
        break;

      case 'title and description':
        form.addSectionHeaderItem().setTitle(question || 'ไม่มีหัวข้อกำหนด').setHelpText(desc || "");
        break;

      case 'section':
        form.addPageBreakItem().setTitle(question || '').setHelpText(desc || "");
        break;

      case 'video':
        // ใช้ column video แทน image
        var videoUrl = row[col('video')];
        if (videoUrl && /^https?:\/\/(www\.)?(youtube\.com|youtu\.be)\//i.test(videoUrl)) {
          try {
            form.addVideoItem()
              .setVideoUrl(videoUrl)
              .setTitle(question || '')
              .setHelpText(desc || "");
          } catch (e) {
            Logger.log('Error adding video: ' + videoUrl + ', ' + e);
          }
        } else {
          Logger.log('Invalid or missing YouTube video URL for: ' + question);
        }
        break;

      case 'short answer':
        item = form.addTextItem().setTitle(question);
        if (desc) item.setHelpText(desc);
        item.setRequired(required);
        break;

      case 'paragraph':
        item = form.addParagraphTextItem().setTitle(question);
        if (desc) item.setHelpText(desc);
        item.setRequired(required);
        break;

      case 'multiple choice':
      case 'checkboxes':
      case 'dropdown':
        var startOpt = col('option start');
        var endOpt = col('option end');
        for (let j = startOpt; j <= endOpt; j++) {
          if (row[j]) choices.push(String(row[j]));
        }
        if (type === 'multiple choice') {
          item = form.addMultipleChoiceItem().setTitle(question).setRequired(required);
          item.setChoices(choices.map(opt => item.createChoice(opt)));
        } else if (type === 'checkboxes') {
          item = form.addCheckboxItem().setTitle(question).setRequired(required);
          item.setChoices(choices.map(opt => item.createChoice(opt)));
        } else if (type === 'dropdown') {
          item = form.addListItem().setTitle(question).setRequired(required);
          item.setChoiceValues(choices);
        }
        // Quiz Logic: set correct answer ถ้ามี
        try {
          if (item && choices.length > 0 && correctAns) {
            if (type === 'checkboxes') {
              var correctSet = correctAns.split(/,|\|/).map(s => s.trim()).filter(Boolean);
              correctSet = correctSet.filter(c => choices.includes(c));
              if (correctSet.length > 0) item.setCorrectAnswer(correctSet);
            } else {
              if (choices.includes(correctAns)) item.setCorrectAnswer(correctAns);
            }
          }
        } catch (e) {
          Logger.log('Quiz/Points error at row ' + (i + 1) + ': ' + e);
        }
        break;

      case 'date':
        item = form.addDateItem().setTitle(question).setRequired(required);
        if (desc) item.setHelpText(desc);
        break;

      case 'time':
        item = form.addTimeItem().setTitle(question).setRequired(required);
        if (desc) item.setHelpText(desc);
        break;

      case 'date and time':
        item = form.addDateTimeItem().setTitle(question).setRequired(required);
        if (desc) item.setHelpText(desc);
        break;

      case 'linear scale':
        item = form.addScaleItem().setTitle(question).setBounds(1, 5).setRequired(required);
        if (desc) item.setHelpText(desc);
        break;

      default:
        Logger.log('Unknown/unsupported type: ' + type + ' for question: ' + question);
    }
  }
}
```
```
function createGoogleFormWithQuizFromSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  function col(name) { return headers.findIndex(h => h.trim().toLowerCase() === name.trim().toLowerCase()); }

  // ตั้งชื่อฟอร์มจากแถวแรก type Title หรือ Title and description ถ้าไม่มีใช้ 'ไม่มีหัวข้อกำหนด'
  var formTitle = 'ไม่มีหัวข้อกำหนด';
  var descTitle = '';
  for (var k = 1; k < data.length; k++) {
    var firstType = String(data[k][col('Type')]).trim().toLowerCase();
    if (firstType === 'title' || firstType === 'title and description') {
      formTitle = data[k][col('Question')] || 'ไม่มีหัวข้อกำหนด';
      descTitle = data[k][col('Desc')] || '';
      break;
    }
  }
  var form = FormApp.create(formTitle);
  form.setIsQuiz(true);

  var titleUsed = false; // ใช้สำหรับป้องกัน Title ซ้ำ

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[col('Type')]) continue;

    var type = String(row[col('Type')]).trim().toLowerCase();
    var question = row[col('Question')] || '';
    var desc = row[col('Desc')] || '';
    var required = true;
    var correctAns = row[col('Correct Answer')] || '';
    var item = null;
    var choices = [];

    // ** ห้ามแสดง Title ซ้ำ **
    if (type === 'title' || type === 'title and description') {
      // ข้าม row ที่ถูกใช้เป็นชื่อฟอร์มไปแล้ว
      if (!titleUsed && (formTitle === question || formTitle === 'ไม่มีหัวข้อกำหนด')) {
        // ถ้า Title and description และมี desc ให้ขึ้น Section Header (description ใหญ่) ด้านบนฟอร์ม
        if (type === 'title and description' && descTitle) {
          form.addSectionHeaderItem().setTitle('').setHelpText(descTitle);
        }
        titleUsed = true;
        continue;
      } else {
        // ถ้า title อื่น ๆ ที่เหลือ ให้ใช้เป็น section header
        form.addSectionHeaderItem().setTitle(question || '').setHelpText(desc || "");
        continue;
      }
    }

    switch(type) {
      case 'section':
        form.addPageBreakItem().setTitle(question || '').setHelpText(desc || "");
        break;

      case 'video':
        var videoUrl = row[col('video')];
        if (videoUrl && /^https?:\/\/(www\.)?(youtube\.com|youtu\.be)\//i.test(videoUrl)) {
          try {
            form.addVideoItem()
              .setVideoUrl(videoUrl)
              .setTitle(question || '')
              .setHelpText(desc || "");
          } catch (e) {
            Logger.log('Error adding video: ' + videoUrl + ', ' + e);
          }
        } else {
          Logger.log('Invalid or missing YouTube video URL for: ' + question);
        }
        break;

      case 'short answer':
        item = form.addTextItem().setTitle(question);
        if (desc) item.setHelpText(desc);
        item.setRequired(required);
        break;

      case 'paragraph':
        item = form.addParagraphTextItem().setTitle(question);
        if (desc) item.setHelpText(desc);
        item.setRequired(required);
        break;

      case 'multiple choice':
      case 'checkboxes':
      case 'dropdown':
        var startOpt = col('option start');
        var endOpt = col('option end');
        for (let j = startOpt; j <= endOpt; j++) {
          if (row[j]) choices.push(String(row[j]));
        }
        if (type === 'multiple choice') {
          item = form.addMultipleChoiceItem().setTitle(question).setRequired(required);
          item.setChoices(choices.map(opt => item.createChoice(opt)));
        } else if (type === 'checkboxes') {
          item = form.addCheckboxItem().setTitle(question).setRequired(required);
          item.setChoices(choices.map(opt => item.createChoice(opt)));
        } else if (type === 'dropdown') {
          item = form.addListItem().setTitle(question).setRequired(required);
          item.setChoiceValues(choices);
        }
        // Quiz Logic: set correct answer ถ้ามี
        try {
          if (item && choices.length > 0 && correctAns) {
            if (type === 'checkboxes') {
              var correctSet = correctAns.split(/,|\|/).map(s => s.trim()).filter(Boolean);
              correctSet = correctSet.filter(c => choices.includes(c));
              if (correctSet.length > 0) item.setCorrectAnswer(correctSet);
            } else {
              if (choices.includes(correctAns)) item.setCorrectAnswer(correctAns);
            }
          }
        } catch (e) {
          Logger.log('Quiz/Points error at row ' + (i + 1) + ': ' + e);
        }
        break;

      case 'date':
        item = form.addDateItem().setTitle(question).setRequired(required);
        if (desc) item.setHelpText(desc);
        break;

      case 'time':
        item = form.addTimeItem().setTitle(question).setRequired(required);
        if (desc) item.setHelpText(desc);
        break;

      case 'date and time':
        item = form.addDateTimeItem().setTitle(question).setRequired(required);
        if (desc) item.setHelpText(desc);
        break;

      case 'linear scale':
        item = form.addScaleItem().setTitle(question).setBounds(1, 5).setRequired(required);
        if (desc) item.setHelpText(desc);
        break;

      default:
        Logger.log('Unknown/unsupported type: ' + type + ' for question: ' + question);
    }
  }
}
```
