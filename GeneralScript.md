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

```
