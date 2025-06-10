# 1
```
function createFormFromSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const rows = data.slice(1);

  // สร้างฟอร์มใหม่
  const form = FormApp.create("ฟอร์มจากชีท");
  let firstTitleSet = false; // flag สำหรับ setTitle เฉพาะรอบแรก

  rows.forEach(row => {
    const get = key => row[header.indexOf(key)] || '';

    const question = get('Question');
    const type = get('Type');
    const opt1 = get('option start');
    const opt2 = get('option 2');
    const opt3 = get('option 3');
    const opt4 = get('option end');
    const correct = get('Correct Answer');

    if (!question && !type) return;

    switch (type) {
      case 'Short answer':
        form.addTextItem().setTitle(question);
        break;

      case 'Multiple choice':
        const options = [opt1, opt2, opt3, opt4].filter(Boolean);
        form.addMultipleChoiceItem()
          .setTitle(question)
          .setChoiceValues(options);
        break;

      case 'Section':
        form.addSectionHeaderItem().setTitle(question);
        break;

      case 'Title and description':
        if (!firstTitleSet) {
          form.setTitle(question); // ใช้ setTitle เฉพาะอันแรกเท่านั้น
          firstTitleSet = true;
        } else {
          form.addSectionHeaderItem().setTitle(question);
        }
        // ถ้ามีคำอธิบายเพิ่ม เติม description ได้ (เช่น opt1)
        if (opt1) {
          form.addSectionHeaderItem().setTitle(opt1);
        }
        break;

      case 'Date and time':
        form.addDateTimeItem().setTitle(question);
        break;

      case 'Video':
        var url = opt1 || question;
        if (url && url.match(/(youtu\.be|youtube\.com)/)) {
          form.addVideoItem()
            .setVideoUrl(url)
            .setTitle(question || "Video");
        } else if (url) {
          form.addParagraphTextItem()
            .setTitle((question || "ดูวิดีโอ") + ' (คลิกลิงก์ด้านล่าง)')
            .setHelpText(url);
        }
        break;

      default:
        break;
    }
  });
}
```
# ผลที่ได้
https://docs.google.com/forms/d/1-B7B9tpe31-P4ZRs6-nIMHXRuKUcC4uVgt8dvApDK78/edit

# 2
```
function createFormFromSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const rows = data.slice(1);

  // สร้างฟอร์มใหม่
  const form = FormApp.create("ฟอร์มจากชีท");

  rows.forEach(row => {
    const get = key => header.indexOf(key) >= 0 ? row[header.indexOf(key)] : "";

    const question = get('Question');
    let type = get('Type');
    if (!type) return;
    type = type.trim().toLowerCase(); // make lowercase for switch
    const required = (get('Required') + '').toLowerCase() === "true";
    const opt1 = get('Option 1');
    const opt2 = get('Option 2');
    const opt3 = get('Option 3');
    const opt4 = get('Option 4');
    const other = (get('Other') + '').toLowerCase() === "true";
    const desc = get('Desc') || "";

    if (!question || !type) return;

    switch (type) {
      case 'short answer':
        var item = form.addTextItem().setTitle(question);
        if (desc) item.setHelpText(desc);
        item.setRequired(required);
        break;

      case 'paragraph':
        var item = form.addParagraphTextItem().setTitle(question);
        if (desc) item.setHelpText(desc);
        item.setRequired(required);
        break;

      case 'multiple choice':
        var choices = [opt1, opt2, opt3, opt4].filter(Boolean);
        var item = form.addMultipleChoiceItem().setTitle(question).setChoiceValues(choices);
        if (other) item.showOtherOption(true);
        if (desc) item.setHelpText(desc);
        item.setRequired(required);
        break;

      case 'checkbox':
      case 'checkboxes':
        var choices = [opt1, opt2, opt3, opt4].filter(Boolean);
        var item = form.addCheckboxItem().setTitle(question).setChoiceValues(choices);
        if (other) item.showOtherOption(true);
        if (desc) item.setHelpText(desc);
        item.setRequired(required);
        break;

      case 'dropdown':
        var choices = [opt1, opt2, opt3, opt4].filter(Boolean);
        // เพิ่ม "อื่นๆ" เป็นตัวเลือกใน dropdown ด้วย ถ้า other = true
        if (other) choices.push("อื่นๆ");
        var item = form.addListItem().setTitle(question).setChoiceValues(choices);
        if (desc) item.setHelpText(desc);
        item.setRequired(required);
        break;

      case 'date':
        var item = form.addDateItem().setTitle(question);
        if (desc) item.setHelpText(desc);
        item.setRequired(required);
        break;

      default:
        // ไม่ตรง type ใดเลย
        break;
    }
  });
}

```
# ผลที่ได้
https://docs.google.com/forms/d/19R7JEgBcz3jSM5-Fh0rKLypei1a5e9Tedz5CwbpDE0Q/preview
