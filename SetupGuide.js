// ========================================
// SETUP GUIDE
// Creates an in-sheet onboarding guide for new users
// ========================================

function createSetupGuideSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = '📋 Setup Guide';

  // Remove existing if re-running setup
  const existing = ss.getSheetByName(sheetName);
  if (existing) ss.deleteSheet(existing);

  const sheet = ss.insertSheet(sheetName, 0); // First tab

  // --- TITLE ---
  sheet.getRange(1, 1, 1, 4).merge()
    .setValue('📧 Gmail Email Router — Setup Guide')
    .setFontSize(16)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 50);

  // --- SUBTITLE ---
  sheet.getRange(2, 1, 1, 4).merge()
    .setValue('Follow the steps below to get your email router running. Check each box when done.')
    .setFontSize(11)
    .setBackground('#e8f0fe')
    .setFontColor('#1a73e8')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(2, 30);

  // --- COLUMN HEADERS ---
  sheet.getRange(3, 1, 1, 4)
    .setValues([['✓ Done', 'Step', 'What To Do', 'Tips & Notes']])
    .setFontWeight('bold')
    .setBackground('#34a853')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  sheet.setRowHeight(3, 30);

  // --- STEPS ---
  const steps = [
    {
      done: true,
      step: 'Step 1\n✅ Complete!',
      what: 'Run Initial Setup\n\nThis step is already done — the Label Config, Queue, and Label Rules sheets have been created for you.',
      tips: 'If you ever need to re-run setup:\nGmail Tools → 🚀 Initial Setup'
    },
    {
      done: false,
      step: 'Step 2\n⚡ Required',
      what: 'Activate Labels to Monitor\n\n1. Click the "Label Config" sheet tab\n2. Find the labels you want to track\n3. Check the "active" checkbox for each one\n4. Make sure "capture_to_queue" is also checked',
      tips: 'Start with 1–2 labels to test before enabling everything.\n\nAll labels start inactive by default.\n\nPer-label options:\n• mark_read — mark emails as read after processing\n• archive — archive emails after processing'
    },
    {
      done: false,
      step: 'Step 3\n⚡ Required',
      what: 'Configure Webhook\n\n1. Copy your webhook URL from n8n (or any endpoint)\n2. Go to Gmail Tools → Webhook → Configure\n3. Paste the URL and optionally add an auth header',
      tips: 'Skip this step if you only want to log emails without sending to a webhook.\n\nSupports custom auth headers (X-API-Key, Bearer token, etc.).\n\nThe URL is stored securely in Script Properties.'
    },
    {
      done: false,
      step: 'Step 4\n🔀 Optional',
      what: 'Route Labels to n8n\n\n1. Go to the "Label Config" sheet\n2. For labels you want sent to your webhook, check "send_to_n8n"\n3. Fill in route_key for each label (used in webhook payload)',
      tips: '"capture_to_queue" = log it\n"send_to_n8n" = log it AND send to webhook\n\nYou can mix and match per label.'
    },
    {
      done: false,
      step: 'Step 5\n📎 Optional',
      what: 'Save PDF Attachments to Drive\n\n1. Create a folder in Google Drive\n2. Open the folder and copy the ID from the URL\n   (the long string after /folders/)\n3. Paste it in the "drive_folder_id" column in Label Config',
      tips: 'Only needed if emails have PDF attachments.\n\nExample Drive URL:\ndrive.google.com/drive/folders/\n1AbCdEfGhIjKlMnOpQrStUvWxYz'
    },
    {
      done: false,
      step: 'Step 6\n🚀 Required',
      what: 'Start the Automation\n\n1. Go to Gmail Tools → ⚙️ Setup Triggers\n2. Click OK to confirm\n\nThis starts:\n• Auto-label + email processing every 1 minute\n• Daily label sync at 3 AM\n• Weekly queue cleanup on Sundays at 2 AM',
      tips: 'Only do this once.\n\nTriggers keep running even after you close the spreadsheet.\n\nThe history tracker initializes on the first trigger run.'
    },
    {
      done: false,
      step: 'Step 7\n🧪 Recommended',
      what: 'Test It\n\n1. Gmail Tools → Webhook → Test Webhook (verify connectivity)\n2. Gmail Tools → ▶️ Run Processing Now\n3. Check the "Queue" sheet to see results\n\nFor pre-existing emails:\nGmail Tools → 📥 Backfill Existing Labels',
      tips: 'If nothing appears, check:\n• The label is marked "active" in Label Config\n• You clicked "Setup Triggers" in Step 6\n• Try "Backfill Existing Labels" for older emails'
    },
    {
      done: false,
      step: 'Step 8\n🏷️ Optional',
      what: 'Auto-Label Rules\n\n1. Open the "Label Rules" sheet\n2. Add rules with Gmail search queries\n3. Each rule auto-applies a label to matching emails\n4. Use Gmail Tools → Auto Label → Dry Run to preview',
      tips: 'Auto-labeling runs before email processing every minute.\n\nRules use Gmail search syntax:\n• from:sender@example.com\n• subject:"invoice"\n• has:attachment filename:pdf'
    },
    {
      done: false,
      step: 'Step 9\n🎉 You\'re done!',
      what: 'Monitor Your System\n\n• Gmail Tools → 📊 View Status — see a summary\n• Queue sheet — full processing history\n• Email alerts if Gmail labels are renamed or deleted\n• Daily sync runs automatically at 3 AM',
      tips: 'Need to add a new label later?\nGmail Tools → Label Management → Discover New Labels\n\nSomething stuck?\nGmail Tools → Maintenance → Retry Errors\n\nForgot where you were?\nGmail Tools → 🧭 Resume Setup Wizard'
    }
  ];

  for (let i = 0; i < steps.length; i++) {
    const row = i + 4;
    const step = steps[i];

    sheet.getRange(row, 1).insertCheckboxes().setValue(step.done);
    sheet.getRange(row, 2).setValue(step.step);
    sheet.getRange(row, 3).setValue(step.what);
    sheet.getRange(row, 4).setValue(step.tips);

    const bg = step.done ? '#e6f4ea' : (i % 2 === 0 ? '#f8f9fa' : '#ffffff');
    sheet.getRange(row, 1, 1, 4).setBackground(bg);

    if (step.done) {
      sheet.getRange(row, 2, 1, 3).setFontColor('#34a853');
    }

    sheet.setRowHeight(row, 110);
  }

  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(3, 380);
  sheet.setColumnWidth(4, 280);

  sheet.getRange(4, 2, steps.length, 3)
    .setWrap(true)
    .setVerticalAlignment('top');
  sheet.getRange(4, 1, steps.length, 1)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.setFrozenRows(3);
  ss.setActiveSheet(sheet);

  Logger.log('✅ Setup Guide sheet created');
  return sheet;
}
