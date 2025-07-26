// গ্লোবাল কনস্ট্যান্টস - আপনার স্প্রেডশীট আইডি এবং শীটের নাম এখানে দিন
const SPREADSHEET_ID = '1-cd-EPaVdsr9ay4bFvFK_M0eg4dklQ9k5xSb7UtNTCc'; // আপনার স্প্রেডশীট আইডি এখানে দিন
const MAIN_DATA_TAB_NAME = 'মীমাংসা রেজিস্টার'; // আপনার মূল ডেটা শীটের নাম (ফর্ম থেকে ডেটা এন্ট্রি হয়)
const MIIMANGSHA_REGISTER_TAB_NAME = 'মীমাংসা রেজিস্টার'; // আপনার মীমাংসা রেজিস্টার শীটের নাম (বর্তমান মাসের আপত্তি ডেটা থাকে)


// প্রতিটি মন্ত্রণালয়ের জন্য টেমপ্লেট শীটের নাম
const TEMPLATE_SHEET_NAMES = {
  'আর্থিক প্রতিষ্ঠান বিভাগ': 'পাক্ষিক রিটার্ণ টেমপ্লেটঃ আর্থিক প্রতিষ্ঠান',
  'বস্ত্র ও পাট': 'পাক্ষিক রিটার্ণ টেমপ্লেটঃ বস্ত্র ও পাট', // যদি এই মন্ত্রণালয়ের জন্য টেমপ্লেট থাকে, তবে নাম দিন
  'শিল্প+বাণিজ্য+বিমান': 'পাক্ষিক রিটার্ণ টেমপ্লেটঃ শিল্প+বাণিজ্য+বিমান' // যদি এই মন্ত্রণালয়ের জন্য টেমপ্লেট থাকে, তবে তবে নাম দিন
};


// মীমাংসা রেজিস্টার শীটে ডেটা পড়ার জন্য কলাম ম্যাপিং (0-ইনডেক্সড)
// ***গুরুত্বপূর্ণ: নিচের কলাম ইনডেক্সগুলো আপনার 'মীমাংসা রেজিস্টার' শীটের প্রকৃত বিন্যাস অনুযায়ী সেট করুন।***
const MIIMANGSHA_REGISTER_MAPPING = {
  ministryNameCol: 1, // B কলাম (0-indexed)
  institutionNameCol: 2, // C কলাম (0-indexed)
  branchNameCol: 3, // D কলাম (0-indexed) - এটি রিটার্নে ব্যবহৃত হবে না, তবে ডেটা পড়ার জন্য রাখা হয়েছে।
  dispositionTypeCol: 9, // J কলাম (0-indexed) - নিষ্পত্তির ধরণ
  involvedAmountCol: 10, // K কলাম (0-indexed) - জড়িত টাকা
  jariPatraDateCol: 20, // U কলাম (0-indexed) - জারিপত্র তাং
};


// টেমপ্লেট কলাম ম্যাপিং (নতুন রিটার্ন শীটের জন্য) - ডেটা কোথায় লেখা হবে
// কলাম ইনডেক্সগুলো 1-based (A=1, B=2, C=3, ...)
const TEMPLATE_COL_MAP = {
  // টেবিল ১ (আপত্তি সম্পর্কিত ডেটা)
  table1: {
    prevMonthObjectionCount: 3,    // C (index 3) - পূর্ববর্তী মাস পর্যন্ত আপত্তি সংখ্যা
    prevMonthObjectionAmount: 4,   // D (index 4) - পূর্ববর্তী মাস পর্যন্ত আপত্তি টাকা

    currentMonthObjectionCount: 5, // E (index 5) - বর্তমান মাসে উত্থাপিত আপত্তি সংখ্যা (এইগুলো 0 থাকবে)
    currentMonthObjectionAmount: 6, // F (index 6) - বর্তমান মাসে উত্থাপিত আপত্তি টাকা (এইগুলো 0 থাকবে)

    totalObjectionCount: 7,        // G (index 7) - মোট আপত্তি সংখ্যা
    totalObjectionAmount: 8,       // H (index 8) - মোট আপত্তি টাকা

    prevMonthSettledCount: 9,      // I (index 9) - পূর্ববর্তী মাস পর্যন্ত মীমাংসিত সংখ্যা
    prevMonthSettledAmount: 10     // J (index 10) - পূর্ববর্তী মাস পর্যন্ত মীমাংসিত টাকা
  },
  // টেবিল ২ (মীমাংসিত ও অন্যান্য ডেটা)
  table2: {
    currentMonthSettledCount: 12,  // L (index 12) - বর্তমান মাসে মীমাংসিত সংখ্যা
    currentMonthSettledAmount: 13, // M (index 13) - বর্তমান মাসে মীমাংসিত টাকা
    totalSettledCount: 14,         // N (index 14) - মোট মীমাংসিত সংখ্যা
    totalSettledAmount: 15,        // O (index 15) - মোট মীমাংসিত টাকা
    unsettledCount: 16,            // P (index 16) - অমীমাংসিত সংখ্যা
    unsettledAmount: 17,           // Q (index 17) - অমীমাংসিত টাকা
    broadsheetReplyCount: 18,      // R (index 18) - ব্রডশিট জবাবের সংখ্যা
    comments: 19                   // S (index 19) - মন্তব্য কলামের জন্য, যদি থাকে
  }
};


// টেমপ্লেট রো ম্যাপিং (নতুন রিটার্ন শীটের জন্য) - নির্দিষ্ট সারি কোথায় পাওয়া যাবে
// সব 0-indexed
const TEMPLATE_ROW_MAP = {
  periodInfo: 3, // রো ৪ (0-indexed 3) - সময়কাল লেখার জন্য
  ministryNameStart: 8, // রো ৯ (0-indexed 8) - প্রথম প্রতিষ্ঠানের ডেটা যেখানে শুরু হয় (মন্ত্রণালয়ের নাম)
  table1HeaderStart: 5, // রো ৬ (0-indexed 5) - টেবিল ১ এর হেডার
  table2HeaderStart: 26, // রো ২৭ (0-indexed 26) - টেবিল ২ এর হেডার
  ministryTotalRows: { // মন্ত্রণালয়ের নাম -> মোট সারির 0-indexed রো
    'আর্থিক প্রতিষ্ঠান বিভাগ': 23 // আর্থিক প্রতিষ্ঠান বিভাগের 'মোট' রো ২৪ (0-indexed 23)
  },
  anotherTotalRow: 43, // রো ৪৪ (0-indexed 43) এ দ্বিতীয় যোগফলের জন্য
  institutions: { // প্রতিষ্ঠানের নাম -> 0-indexed রো (আর্থিক প্রতিষ্ঠান বিভাগের টেমপ্লেটের জন্য)
    'সোনালী ব্যাংক পিএলসি': 8, // রো ৯ (0-indexed 8)
    'জনতা ব্যাংক পিএলসি': 9, // রো ১০ (0-indexed 9)
    'অগ্রণী ব্যাংক পিএলসি': 10, // রো ১১ (0-indexed 10)
    'বাংলাদেশ কৃষি ব্যাংক': 11, // রো ১২ (0-indexed 11)
    'রুপালী ব্যাংক পিএলসি': 12, // রো ১৩ (0-indexed 12)
    'বাংলাাদেশ ব্যাংক': 13, // রো ১৪ (0-indexed 13)
    'বাংলাদেশ ডেভেলপমেন্ট ব্যাংক পিএলসি': 14, // রো ১৫ (0-indexed 14)
    'গৃহনির্মাণ ঋণদান সংস্থা': 15, // রো ১৬ (0-indexed 15)
    'কর্মসংস্থান ব্যাংক': 16, // রো ১৭ (0-indexed 16)
    'বেসিক ব্যাংক': 17, // রো ১৮ (0-indexed 17)
    'আনসার ভিডিপি উন্নয়ন ব্যাংক': 18, // রো ১৯ (0-indexed 18)
    'ইনভেস্টমেন্ট কর্পোরেশন অব বাংলাদেশ': 19, // রো ২০ (0-indexed 19)
    'সাধারণ বীমা কর্পোরেশন': 20, // রো ২১ (0-indexed 20)
    'জীবন বীমা কর্পোরেশন': 21, // রো ২২ (0-indexed 21)
    'প্রবাসী কল্যান ব্যাংক': 22  // রো ২৩ (0-indexed 22)
  }
};

/**
 * সংখ্যাকে বাংলা অঙ্কে রূপান্তর করে এবং একক ডিজিটের ক্ষেত্রে বাম পাশে শূন্য যোগ করে।
 * @param {number} num - যে সংখ্যাকে ফরম্যাট করতে হবে।
 * @returns {string} ফরম্যাট করা বাংলা সংখ্যা স্ট্রিং।
 */
function formatCountForDisplay(num) {
    const bengaliDigits = ['০', '১', '২', '৩', '৪', '৫', '৬', '৭', '৮', '৯'];
    let formattedNum = String(num);

    if (num < 10 && num >= 0) { // নিশ্চিত করুন এটি একক সংখ্যা এবং ঋণাত্মক নয়
        formattedNum = '0' + formattedNum; // বাম পাশে শূন্য যোগ করুন
    }
    // বাংলা সংখ্যায় রূপান্তর করুন
    return formattedNum.split('').map(digit => {
        const parsedDigit = parseInt(digit);
        // যদি এটি একটি সংখ্যা হয়, তবে এটি রূপান্তর করুন, অন্যথায় এটি যেমন আছে তেমনই রাখুন (যেমন, অগ্রণী '0' এর জন্য)
        return isNaN(parsedDigit) ? digit : bengaliDigits[parsedDigit];
    }).join('');
}


/**
 * মীমাংসা রেজিস্টার ব্লক টোটাল এবং গ্র্যান্ড টোটাল সেট করে।
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - যে শীটে ডেটা সেট করা হবে।
 * @param {number} startRow - ডেটা ব্লকের শুরুর সারি (1-based)।
 * @param {number} numRows - ডেটা ব্লকের সারির সংখ্যা।
 * @param {Object} blockTotals - ব্লকের মোট যোগফলগুলোর অবজেক্ট।
 */
function applyFormatting(sheet, startRow, numRows, blockTotals) {
  Logger.log(`applyFormatting called for startRow: ${startRow}, numRows: ${numRows}`);

  const lastColumn = 22; // V কলাম পর্যন্ত, মোট 22টি কলাম (0-ভিত্তিক সূচকে 21)

  // Header row (row 1) in main sheet
  sheet.getRange(1, 1, 1, lastColumn).setVerticalAlignment('middle');
  sheet.getRange(1, 1, 1, lastColumn).setHorizontalAlignment('center');
  sheet.getRange(1, 1, 1, lastColumn).setFontSize(12); // Header font size

  const entireBlockRange = sheet.getRange(startRow, 1, numRows, lastColumn);
  entireBlockRange.setBorder(true, true, true, true, true, true);
  entireBlockRange.setWrap(true);
  entireBlockRange.setHorizontalAlignment('center'); // পুরো ব্লকের জন্য ডিফল্ট কেন্দ্র সারিবদ্ধতা
  entireBlockRange.setVerticalAlignment('middle');
  entireBlockRange.setFontSize(12); // Set font size for the entire data block

  // Specific column alignments overriding the entireBlockRange default
  sheet.getRange(startRow, 1, numRows, 1).setHorizontalAlignment('center'); // A কলাম: ক্রমিক নং (এটি ডিফল্ট কেন্দ্র সারিবদ্ধতা দ্বারা প্রভাবিত হবে)
  sheet.getRange(startRow, 2, numRows, 1).setHorizontalAlignment('center');   // B কলাম: মন্ত্রণালয়ের নাম - কেন্দ্র সারিবদ্ধ থাকবে
  sheet.getRange(startRow, 3, numRows, 1).setHorizontalAlignment('left');    // C কলাম: প্রতিষ্ঠানের নাম - বাম সারিবদ্ধ থাকবে
  sheet.getRange(startRow, 4, numRows, 1).setHorizontalAlignment('left');    // D কলাম: শাখার নাম - বাম সারিবদ্ধ থাকবে

  sheet.getRange(startRow, 5, numRows, 4).setHorizontalAlignment('center'); // E থেকে H কলাম: নিরীক্ষা বছর থেকে আপত্তির ধরণ
  sheet.getRange(startRow, 5, numRows, 4).setVerticalAlignment('middle');

  // I কলাম: অনুচ্ছেদ নং - এখন কেন্দ্র সারিবদ্ধ হবে
  sheet.getRange(startRow, 9, numRows, 1).setHorizontalAlignment('center'); // I কলাম: অনুচ্ছেদ নং
  sheet.getRange(startRow, 9, numRows, 1).setVerticalAlignment('middle');
  sheet.getRange(startRow, 9, numRows, 1).setNumberFormat('@'); // অনুচ্ছেদ নং এর জন্য

  // J কলাম: নিষ্পত্তির ধরণ - কেন্দ্র সারিবদ্ধ
  sheet.getRange(startRow, 10, numRows, 1).setHorizontalAlignment('center');
  sheet.getRange(startRow, 10, numRows, 1).setVerticalAlignment('middle');

  sheet.getRange(startRow, 11, numRows, 9).setHorizontalAlignment('center'); // K থেকে S কলাম: জড়িত টাকা থেকে মোট সমন্বয়
  sheet.getRange(startRow, 11, numRows, 9).setVerticalAlignment('middle');


  // জারিপত্র নং, জারিপত্র তাং, অতিরিক্ত তথ্য কলামের ফরম্যাটিং
  sheet.getRange(startRow, 20, numRows, 1).setNumberFormat('@'); // T কলাম: জারিপত্র নং
  sheet.getRange(startRow, 21, numRows, 1).setNumberFormat('yyyy-mm-dd'); // U কলাম: জারিপত্র তাং
  sheet.getRange(startRow, 22, numRows, 1).setNumberFormat('@'); // V কলাম: অতিরিক্ত তথ্য

  if (numRows > 1) {
    // এখানে মার্জিং লজিকটি পুনর্বিবেচনা করা প্রয়োজন যদি A থেকে G কলাম মার্জ করা হয়।
    // বর্তমানে, ফর্ম ডেটা এন্ট্রির জন্য প্রতিটি সারি আলাদাভাবে ডেটা নেয়, তাই উল্লম্ব মার্জিং এখানে সাধারণত প্রযোজ্য নয়।
    // যদি আপনি চান যে, একই জারিপত্র নং এর জন্য একাধিক অনুচ্ছেদ থাকলে সেগুলো মার্জ হবে, তবে সেই লজিক এখানে যোগ করা যেতে পারে।
    // আপাতত, আমি পূর্বের মার্জিং লজিকগুলো কমেন্ট করে রাখলাম, কারণ এটি ডেটা এন্ট্রি শীটের জন্য প্রযোজ্য নাও হতে পারে।
    // যদি আপনার ফর্ম এন্ট্রি শীটে একই জারিপত্র নং এর একাধিক অনুচ্ছেদ থাকে এবং আপনি সেগুলোকে মার্জ করতে চান, তবে আমাকে জানান।

    // A কলাম: ক্রমিক নং
    sheet.getRange(startRow, 1, numRows, 1).setHorizontalAlignment('center');
    sheet.getRange(startRow, 1, numRows, 1).setVerticalAlignment('middle');

    // জারিপত্র নং কলাম (T) মার্জ করা
    const jareeNoRange = sheet.getRange(startRow, 20, numRows, 1);
    jareeNoRange.mergeVertically();
    jareeNoRange.setVerticalAlignment('middle');
    jareeNoRange.setHorizontalAlignment('center');
    jareeNoRange.setBorder(true, true, true, true, true, true);

    // জারিপত্র তাং কলাম (U) মার্জ করা
    const jareeDateRange = sheet.getRange(startRow, 21, numRows, 1);
    jareeDateRange.mergeVertically();
    jareeDateRange.setVerticalAlignment('middle');
    jareeDateRange.setHorizontalAlignment('center');
    jareeDateRange.setBorder(true, true, true, true, true, true);

    // অতিরিক্ত তথ্য কলাম (V) মার্জ করা
    const additionalInfoRange = sheet.getRange(startRow, 22, numRows, 1);
    additionalInfoRange.mergeVertically();
    additionalInfoRange.setVerticalAlignment('middle');
    additionalInfoRange.setHorizontalAlignment('center');
    additionalInfoRange.setBorder(true, true, true, true, true, true);

  } else {
    sheet.getRange(startRow, 20, numRows, 1).setBorder(true, true, true, true, true, true);
    sheet.getRange(startRow, 21, numRows, 1).setBorder(true, true, true, true, true, true);
    sheet.getRange(startRow, 22, numRows, 1).setBorder(true, true, true, true, true, true);
  }

  const grandTotalRow = startRow + numRows; // ডেটা সারির ঠিক পরে হবে

  // Grand total row যোগ করুন
  sheet.insertRowsAfter(grandTotalRow - 1, 1); // ডেটা শেষ হওয়ার পরের সারিতে একটি নতুন সারি যোগ করবে


  const labelRange = sheet.getRange(grandTotalRow, 1, 1, 8); // A থেকে H কলাম পর্যন্ত মার্জ
  labelRange.mergeAcross();
  labelRange.setValue('মোট মীমাংসিত অনুচ্ছেদ এর সংখ্যা ও মোট টাকার পরিমাণ যথাক্রমে = ');
  labelRange.setHorizontalAlignment('center');
  labelRange.setVerticalAlignment('middle');
  labelRange.setFontWeight('bold');
  labelRange.setBackground('#c5e1a5');
  labelRange.setBorder(true, true, true, true, true, true);
  labelRange.setFontSize(12); // Set font size for label

  const paragraphCountCell = sheet.getRange(grandTotalRow, 9); // I কলাম
  paragraphCountCell.setValue(`${formatCountForDisplay(blockTotals.totalFullParagraphs)} টি`); // **পরিবর্তন এখানে**
  paragraphCountCell.setFontWeight('bold');
  paragraphCountCell.setBackground('#c5e1a5');
  paragraphCountCell.setHorizontalAlignment('center');
  paragraphCountCell.setVerticalAlignment('middle');
  paragraphCountCell.setBorder(true, true, true, true, true, true);
  paragraphCountCell.setFontSize(12); // Set font size for paragraph count


  const numericColumns = [
    { colIndex: 11, header: 'জড়িত টাকা', value: blockTotals.totalInvolvedAmount }, // K কলাম
    { colIndex: 12, header: 'ভ্যাট আদায়', value: blockTotals.totalVatCollected }, // L কলাম
    { colIndex: 13, header: 'ভ্যাট সমন্বয়', value: blockTotals.totalVatAdjusted }, // M কলাম
    { colIndex: 14, header: 'আয়কর আদায়', value: blockTotals.totalIncomeTaxCollected }, // N কলাম
    { colIndex: 15, header: 'আয়কর সমন্বয়', value: blockTotals.totalIncomeTaxAdjusted }, // O কলাম
    { colIndex: 16, header: 'অন্যান্য আদায়', value: blockTotals.totalOtherCollected }, // P কলাম
    { colIndex: 17, header: 'অন্যান্য সমন্বয়', value: blockTotals.totalOtherAdjusted }, // Q কলাম
    { colIndex: 18, header: 'মোট আদায়', value: blockTotals.totalCollectedAmount }, // R কলাম
    { colIndex: 19, header: 'মোট সমন্বয়', value: blockTotals.totalAdjustedAmount } // S কলাম
    // কলাম 20, 21, 22 (জারিপত্র নং, জারিপত্র তাং, অতিরিক্ত তথ্য) এখানে মোট যোগ করা হবে না
  ];

  numericColumns.forEach(col => {
    const cell = sheet.getRange(grandTotalRow, col.colIndex);
    cell.setValue(col.value);
    cell.setFontWeight('bold');
    cell.setBackground('#c5e1a5');
    cell.setHorizontalAlignment('center');
    cell.setVerticalAlignment('middle');
    cell.setBorder(true, true, true, true, true, true);
    cell.setFontSize(12); // Set font size for numeric total cells
  });

  // জারিপত্র নং (T), জারিপত্র তাং (U), অতিরিক্ত তথ্য (V) কলামগুলোর জন্য গ্র্যান্ড টোটাল রোতে খালি সেল সেট করা
  sheet.getRange(grandTotalRow, 20).setValue('');
  sheet.getRange(grandTotalRow, 21).setValue('');
  sheet.getRange(grandTotalRow, 22).setValue(''); // কলাম V

  sheet.getRange(grandTotalRow, 1, 1, lastColumn).setBorder(true, true, true, true, true, true);

  // গ্র্যান্ড টোটাল রো-এর নিচে একটি খালি সারি যোগ করা
  sheet.insertRowsAfter(grandTotalRow, 1);
  sheet.getRange(grandTotalRow + 1, 1, 1, lastColumn).setBackground('#ffffff');
  sheet.getRange(grandTotalRow + 1, 1, 1, lastColumn).clearFormat();
  sheet.getRange(grandTotalRow + 1, 1, 1, lastColumn).setHorizontalAlignment('center');
  sheet.getRange(grandTotalRow + 1, 1, 1, lastColumn).setVerticalAlignment('middle');
  sheet.getRange(grandTotalRow + 1, 1, 1, lastColumn).setFontSize(12); // Set font size for the empty row below grand total

  SpreadsheetApp.flush();
}


/**
 * Returns the URL of the Google Spreadsheet associated with SPREADSHEET_ID.
 * @returns {string} The URL of the spreadsheet.
 */
function getSpreadsheetUrl() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  // মীমাংসা রেজিস্টার লোড না হলে ত্রুটি হ্যান্ডলিং
  if (!ss) {
    Logger.log(`Error: Spreadsheet with ID '${SPREADSHEET_ID}' could not be opened in getSpreadsheetUrl.`);
    throw new Error(`মীমাংসা রেজিস্টার খুলতে ব্যর্থ। আইডি '${SPREADSHEET_ID}' সঠিক কিনা এবং আপনার Apps Script প্রোজেক্টের এর উপর প্রয়োজনীয় অনুমতি আছে কিনা নিশ্চিত করুন।`);
  }
  return ss.getUrl();
}


/**
 * একটি তারিখের জন্য পাক্ষিক সময়কাল গণনা করে।
 * @param {Date} date - যে তারিখের জন্য পাক্ষিক সময়কাল গণনা করা হবে।
 * @returns {Object} {startDate: Date, endDate: Date, startDateBengla: string, endDateBengla: string}
 */
function getFortnightlyPeriod(date) {
  const year = date.getFullYear();
  const month = date.getMonth(); // 0-indexed (জানুয়ারি = 0)
  const timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();


  let startDate;
  let endDate;


  // সবসময় বর্তমান মাসের ১৬ তারিখ থেকে পরবর্তী মাসের ১৫ তারিখ পর্যন্ত সময়কাল
  startDate = new Date(year, month, 16); // বর্তমান মাসের ১৬ তারিখ
  endDate = new Date(year, month + 1, 15); // পরবর্তী মাসের ১৫ তারিখ


  // বাংলা সংখ্যায় রূপান্তর করার জন্য একটি হেল্পার ফাংশন
  function convertToBengaliNumerals(num) {
    const bengaliDigits = ['০', '১', '২', '৩', '৪', '৫', '৬', '৭', '৮', '৯'];
    return String(num).split('').map(digit => bengaliDigits[parseInt(digit)]).join('');
  }


  // তারিখগুলোকে dd-mm-yyyy ফরম্যাটে ইংরেজিতে ফরম্যাট করুন
  let formattedStartDateEnglish = Utilities.formatDate(startDate, timeZone, "dd-MM-yyyy");
  let formattedEndDateEnglish = Utilities.formatDate(endDate, timeZone, "dd-MM-yyyy");


  // সংখ্যাগুলোকে বাংলাতে রূপান্তর করুন
  let startDateBengali = formattedStartDateEnglish.split('-').map(part => {
    return part.split('').map(digit => {
      return (digit >= '0' && digit <= '9') ? convertToBengaliNumerals(parseInt(digit)) : digit;
    }).join('');
  }).join('-');


  let endDateBengali = formattedEndDateEnglish.split('-').map(part => {
    return part.split('').map(digit => {
      return (digit >= '0' && digit <= '9') ? convertToBengaliNumerals(parseInt(digit)) : digit;
    }).join('');
  }).join('-');


  return {
    startDate: startDate,
    endDate: endDate,
    startDateBengla: startDateBengali,
    endDateBengla: endDateBengali
  };
}


/**
 * ওয়েব অ্যাপ হিসেবে স্থাপন (deploy) করার সময় HTML ফরমটি পরিবেশন করে।
 */
function doGet() {
  return HtmlService.createTemplateFromFile('meemangsa-register-form-html').evaluate();
}


/**
 * HTML ফরম থেকে প্রাপ্ত ডেটা প্রক্রিয়া করে Google Sheet-এর 'মূল ডেটা' শীটে যোগ করে।
 * @param {Object} formData - ফরম থেকে প্রাপ্ত ডেটা।
 * @returns {Object} - সফল বা ব্যর্থতার বার্তা।
 */
function processFormData(formData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    if (!ss) {
      throw new Error(`মীমাংসা রেজিস্টার খুলতে ব্যর্থ। আইডি '${SPREADSHEET_ID}' সঠিক কিনা নিশ্চিত করুন।`);
    }

    const sheet = ss.getSheetByName(MAIN_DATA_TAB_NAME);
    if (!sheet) {
      throw new Error(`শীট "${MAIN_DATA_TAB_NAME}" খুঁজে পাওয়া যায়নি।`);
    }

    // Determine the next serial number based on the current fortnight
    const lastRowBeforeAppend = sheet.getLastRow(); // ডেটা যুক্ত করার আগে শেষ সারি
    const currentFortnight = getFortnightlyPeriod(new Date()); // বর্তমান পাক্ষিকের তারিখগুলো নিন
    const timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();

    let serialNo = 1; // বর্তমান পাক্ষিকের জন্য ডিফল্ট ক্রমিক নং 1

    // শীটের সমস্ত ডেটা পড়ুন
    const allData = sheet.getDataRange().getValues();
    let maxSerialForCurrentFortnight = 0;

    // হেডার সারি বাদ দিতে সারি 2 (ইনডেক্স 1) থেকে শুরু করুন
    for (let i = 1; i < allData.length; i++) {
      const row = allData[i];
      const rowSerialNo = row[0]; // A কলাম (0-ইনডেক্সড) হলো ক্রমিক নং
      const jareeDateValue = row[MIIMANGSHA_REGISTER_MAPPING.jariPatraDateCol]; // U কলাম (0-ইনডেক্সড) হলো জারিপত্র তাং

      // নিশ্চিত করুন যে জারিপত্র তাং একটি বৈধ তারিখ অবজেক্ট
      if (jareeDateValue instanceof Date) {
        const entryDate = jareeDateValue;

        // এন্ট্রি ডেট বর্তমান পাক্ষিকের মধ্যে পড়ে কিনা তা পরীক্ষা করুন
        // টাইমজোন সমস্যা এড়াতে তারিখ স্ট্রিং তুলনা করা হচ্ছে
        const formattedEntryDate = Utilities.formatDate(entryDate, timeZone, "yyyy-MM-dd");
        const formattedFortnightStartDate = Utilities.formatDate(currentFortnight.startDate, timeZone, "yyyy-MM-dd");
        const formattedFortnightEndDate = Utilities.formatDate(currentFortnight.endDate, timeZone, "yyyy-MM-dd");

        if (formattedEntryDate >= formattedFortnightStartDate && formattedEntryDate <= formattedFortnightEndDate) {
          if (typeof rowSerialNo === 'number' && !isNaN(rowSerialNo)) {
            if (rowSerialNo > maxSerialForCurrentFortnight) {
              maxSerialForCurrentFortnight = rowSerialNo;
            }
          }
        }
      }
    }

    if (maxSerialForCurrentFortnight > 0) {
      serialNo = maxSerialForCurrentFortnight + 1; // বর্তমান পাক্ষিকের শেষ ক্রমিক নং এর পর থেকে শুরু করুন
    }
    // যদি maxSerialForCurrentFortnight 0 হয় (মানে এই পাক্ষিকের জন্য কোনো এন্ট্রি নেই), তাহলে serialNo 1 থাকবে।


    const rowsToAppend = [];
    let totalFullParagraphs = 0; // মোট অনুচ্ছেদের সংখ্যা গণনা করার জন্য

    let currentBlockTotalInvolvedAmount = 0;
    let currentBlockTotalVatCollected = 0;
    let currentBlockTotalVatAdjusted = 0;
    let currentBlockTotalIncomeTaxCollected = 0;
    let currentBlockTotalIncomeTaxAdjusted = 0;
    let currentBlockTotalOtherCollected = 0;
    let currentBlockTotalOtherAdjusted = 0;
    let currentBlockTotalCollectedAmount = 0;
    let currentBlockTotalAdjustedAmount = 0;


    formData.paragraphs.forEach((paragraph, index) => {
      const totalCollectedAmount = (typeof paragraph.vatCollected === 'number' ? paragraph.vatCollected : 0) +
                                   (typeof paragraph.incomeTaxCollected === 'number' ? paragraph.incomeTaxCollected : 0) +
                                   (typeof paragraph.otherCollected === 'number' ? paragraph.otherCollected : 0);
      const totalAdjustedAmount = (typeof paragraph.vatAdjusted === 'number' ? paragraph.vatAdjusted : 0) +
                                   (typeof paragraph.incomeTaxAdjusted === 'number' ? paragraph.incomeTaxAdjusted : 0) +
                                   (typeof paragraph.otherAdjusted === 'number' ? paragraph.otherAdjusted : 0);


      // Ensure jareeDate is a valid Date object if it exists
      let jareeDate = '';
      if (formData.jareeDate) {
        try {
          // Parse date string in YYYY-MM-DD format
          const [year, month, day] = formData.jareeDate.split('-').map(Number);
          jareeDate = new Date(year, month - 1, day); // month - 1 because Date month is 0-indexed
        } catch (e) {
          Logger.log(`Invalid jareeDate: ${formData.jareeDate}. Error: ${e.message}`);
          jareeDate = ''; // Set to empty if invalid
        }
      }

      const row = [
        serialNo + index, // ক্রমিক নং (A)
        formData.ministryName, // মন্ত্রণালয়ের নাম (B)
        formData.institutionName, // প্রতিষ্ঠানের নাম (C)
        formData.branchName, // শাখার নাম (D) - নতুন যোগ করা হয়েছে
        formData.auditYear, // নিরীক্ষা বছর (E)
        formData.letterNoDate, // পত্র নং ও তাং (F)
        formData.workPaperNoDate, // কার্যপত্র নং ও তাং (G)
        formData.objectionType, // আপত্তির ধরণ (H)
        paragraph.paragraphNo, // অনুচ্ছেদ নং (I)
        paragraph.paragraphType, // নিষ্পত্তির ধরণ (J)
        paragraph.involvedAmount, // জড়িত টাকা (K)
        paragraph.vatCollected, // ভ্যাট আদায় (L)
        paragraph.vatAdjusted, // ভ্যাট সমন্বয় (M)
        paragraph.incomeTaxCollected, // আয়কর আদায় (N)
        paragraph.incomeTaxAdjusted, // আয়কর সমন্বয় (O)
        paragraph.otherCollected, // অন্যান্য আদায় (P)
        paragraph.otherAdjusted, // অন্যান্য সমন্বয় (Q)
        totalCollectedAmount, // মোট আদায় (R)
        totalAdjustedAmount, // মোট সমন্বয় (S)
        formData.jareeNo, // জারিপত্র নং (T)
        jareeDate, // জারিপত্র তাং (U)
        formData.additionalInfo // অতিরিক্ত তথ্য (V)
      ];
      rowsToAppend.push(row);

      // যদি নিষ্পত্তির ধরণ 'পূর্ণাঙ্গ' হয়, তাহলে মোট পূর্ণাঙ্গ অনুচ্ছেদের সংখ্যা বাড়ান
      if (paragraph.paragraphType === 'পূর্ণাঙ্গ') {
        totalFullParagraphs++;
      }

      // ব্লকের মোট যোগফল গণনা - টাকার অংক পূর্ণাঙ্গ বা আংশিক উভয় নিষ্পত্তির জন্যই হিসাব করা হবে
      currentBlockTotalInvolvedAmount += typeof paragraph.involvedAmount === 'number' ? paragraph.involvedAmount : 0;
      currentBlockTotalVatCollected += typeof paragraph.vatCollected === 'number' ? paragraph.vatCollected : 0;
      currentBlockTotalVatAdjusted += typeof paragraph.vatAdjusted === 'number' ? paragraph.vatAdjusted : 0;
      currentBlockTotalIncomeTaxCollected += typeof paragraph.incomeTaxCollected === 'number' ? paragraph.incomeTaxCollected : 0;
      currentBlockTotalIncomeTaxAdjusted += typeof paragraph.incomeTaxAdjusted === 'number' ? paragraph.incomeTaxAdjusted : 0;
      currentBlockTotalOtherCollected += typeof paragraph.otherCollected === 'number' ? paragraph.otherCollected : 0;
      currentBlockTotalOtherAdjusted += typeof paragraph.otherAdjusted === 'number' ? paragraph.otherAdjusted : 0;
      currentBlockTotalCollectedAmount += totalCollectedAmount;
      currentBlockTotalAdjustedAmount += totalAdjustedAmount;
    });

    // ডেটা শীটে যোগ করুন
    sheet.getRange(lastRowBeforeAppend + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);

    // ডেটা জমা হওয়ার পর শেষ সারি (যেখান থেকে ফরম্যাটিং শুরু হবে)
    const newStartRow = lastRowBeforeAppend + 1; // নতুন যুক্ত হওয়া প্রথম সারির 1-based index

    // ব্লকের মোট যোগফলের জন্য একটি অবজেক্ট তৈরি করুন
    const blockTotals = {
      totalFullParagraphs: totalFullParagraphs,
      totalInvolvedAmount: currentBlockTotalInvolvedAmount,
      totalVatCollected: currentBlockTotalVatCollected,
      totalVatAdjusted: currentBlockTotalVatAdjusted,
      totalIncomeTaxCollected: currentBlockTotalIncomeTaxCollected,
      totalIncomeTaxAdjusted: currentBlockTotalIncomeTaxAdjusted,
      totalOtherCollected: currentBlockTotalOtherCollected,
      totalOtherAdjusted: currentBlockTotalOtherAdjusted,
      totalCollectedAmount: currentBlockTotalCollectedAmount,
      totalAdjustedAmount: currentBlockTotalAdjustedAmount
    };

    // applyFormatting ফাংশন কল করুন
    applyFormatting(sheet, newStartRow, rowsToAppend.length, blockTotals);


    return { success: true, message: 'ডেটা সফলভাবে জমা দেওয়া হয়েছে!' };
  } catch (e) {
    Logger.log(`Error processing form data: ${e.message}`);
    return { success: false, message: `ডেটা জমা দিতে ব্যর্থ: ${e.message}` };
  }
}

function getCurrentCyclePeriod() {
  const today = new Date();
  const startDateObj = new Date(today.getFullYear(), today.getMonth(), 16);
  const endDateObj = new Date(today.getFullYear(), today.getMonth() + 1, 15);

  // ইংরেজি সংখ্যাকে বাংলা সংখ্যায় রূপান্তরের জন্য একটি ম্যাপ
  const englishToBengaliDigits = {
    '0': '০', '1': '১', '2': '২', '3': '৩', '4': '৪',
    '5': '৫', '6': '৬', '7': '৭', '8': '৮', '9': '৯'
  };

  function convertToBengaliDigits(numberString) {
    return numberString.split('').map(digit => englishToBengaliDigits[digit] || digit).join('');
  }

  function formatDate(date) {
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = String(date.getFullYear());

    // সংখ্যাগুলোকে বাংলায় রূপান্তর
    const bengaliDay = convertToBengaliDigits(day);
    const bengaliMonth = convertToBengaliDigits(month);
    const bengaliYear = convertToBengaliDigits(year);

    return `${bengaliDay}/${bengaliMonth}/${bengaliYear}`;
  }

  const formattedStartDate = formatDate(startDateObj);
  const formattedEndDate = formatDate(endDateObj);

  // নতুন ফরম্যাটে স্ট্রিং তৈরি
  const result = `সময়কালঃ ${formattedStartDate} হতে ${formattedEndDate} খ্রিঃ তারিখ পর্যন্ত।`;

  Logger.log('getCurrentCyclePeriod result: ' + result);
  return result;
}
/**
 * পাক্ষিক রিটার্নের জন্য প্রয়োজনীয় ডেটা সংগ্রহ করে, মন্ত্রণালয় এবং সংস্থার ভিত্তিতে গ্রুপিং সহ।
 * @param {string} ministryType - যে মন্ত্রণালয়ের জন্য ডেটা সংগ্রহ করা হবে।
 * @returns {Object} {ministries: Object, grandTotals: Object}
 */
function getFortnightlyReturnData(ministryType) {
  Logger.log(`--- getFortnightlyReturnData started for ${ministryType} ---`);
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);


  if (!spreadsheet) {
    Logger.log(`Error: Spreadsheet with ID '${SPREADSHEET_ID}' could not be opened in getFortnightlyReturnData.`);
    throw new Error(`মীমাংসা রেজিস্টার খুলতে ব্যর্থ। আইডি '${SPREADSHEET_ID}' সঠিক কিনা এবং আপনার Apps Script প্রোজেক্টের এর উপর প্রয়োজনীয় অনুমতি আছে কিনা নিশ্চিত করুন।`);
  }


  const miimangshaRegisterSheet = spreadsheet.getSheetByName(MIIMANGSHA_REGISTER_TAB_NAME);
  if (!miimangshaRegisterSheet) {
    Logger.log(`Error: 'মীমাংসা রেজিস্টার' শীট খুঁজে পাওয়া যায়নি: ${MIIMANGSHA_REGISTER_TAB_NAME}`);
    throw new Error(`'মীমাংসা রেজিস্টার' শীট খুঁজে পাওয়া যায়নি: ${MIIMANGSHA_REGISTER_TAB_NAME}`);
  }


  const allMiimangshaRegisterData = miimangshaRegisterSheet.getDataRange().getValues();
  Logger.log(`Miimangsha Register rows: ${allMiimangshaRegisterData.length}`);


  // Initialize data structure to hold combined data
  const combinedMinistryInstitutionData = new Map(); // Map<MinistryName, Map<InstitutionName, { ...data }>>


  const templateSheetName = TEMPLATE_SHEET_NAMES[ministryType];
  if (!templateSheetName) {
    Logger.log(`Error: No template sheet name defined for ministry type: ${ministryType}`);
    throw new Error(`এই মন্ত্রণালয়ের জন্য কোনো টেমপ্লেট খুঁজে পাওয়া হয়নি: ${ministryType}`);
  }
  const templateSheet = spreadsheet.getSheetByName(templateSheetName);
  if (!templateSheet) {
    Logger.log(`Error: Template sheet named '${templateSheetName}' not found.`);
    throw new Error(`টেমপ্লেট শীট খুঁজে পাওয়া যায়নি: ${templateSheetName}`);
  }


  // 1. Populate initial data from the template itself (previous period's cumulative and broadsheet replies)
  // This loop now iterates through the institutions defined in TEMPLATE_ROW_MAP.institutions
  for (const institutionName in TEMPLATE_ROW_MAP.institutions) {
    const templateRowIndex = TEMPLATE_ROW_MAP.institutions[institutionName]; // 0-indexed row in the template
    const sheetRow = templateRowIndex + 1; // 1-based for sheet.getRange


    // Get previous month's data from Template Table 1
    // TEMPLATE_COL_MAP uses 1-based column indices directly, so no need for +1 here
    const prevObjectionCount = templateSheet.getRange(sheetRow, TEMPLATE_COL_MAP.table1.prevMonthObjectionCount).getValue() || 0;
    const prevObjectionAmount = templateSheet.getRange(sheetRow, TEMPLATE_COL_MAP.table1.prevMonthObjectionAmount).getValue() || 0;
    const prevSettledCount = templateSheet.getRange(sheetRow, TEMPLATE_COL_MAP.table1.prevMonthSettledCount).getValue() || 0;
    const prevSettledAmount = templateSheet.getRange(sheetRow, TEMPLATE_COL_MAP.table1.prevMonthSettledAmount).getValue() || 0;


    // Get broadsheet reply count from Template Table 2
    // Calculate table2InstitutionRow relative to table2HeaderStart
    const table2InstitutionRow = TEMPLATE_ROW_MAP.table2HeaderStart + (templateRowIndex - TEMPLATE_ROW_MAP.ministryNameStart);
    const broadsheetReplyCount = templateSheet.getRange(table2InstitutionRow + 1, TEMPLATE_COL_MAP.table2.broadsheetReplyCount).getValue() || 0;


    if (!combinedMinistryInstitutionData.has(ministryType)) {
      combinedMinistryInstitutionData.set(ministryType, new Map());
    }
    combinedMinistryInstitutionData.get(ministryType).set(institutionName, {
      prevMonthObjectionCount: typeof prevObjectionCount === 'number' ? prevObjectionCount : 0,
      prevMonthObjectionAmount: typeof prevObjectionAmount === 'number' ? prevObjectionAmount : 0,
      prevMonthSettledCount: typeof prevSettledCount === 'number' ? prevSettledCount : 0,
      prevMonthSettledAmount: typeof prevSettledAmount === 'number' ? prevSettledAmount : 0,
      currentMonthObjectionCount: 0, // Initialize to 0, will be filled from Miimangsha Register
      currentMonthObjectionAmount: 0, // Initialize to 0
      currentMonthSettledCount: 0, // Initialize to 0
      currentMonthSettledAmount: 0, // Initialize to 0
      broadsheetReplyCount: typeof broadsheetReplyCount === 'number' ? broadsheetReplyCount : 0
    });
    Logger.log(`Initialized from Template: ${ministryType} - ${institutionName}, PrevObjCount: ${prevObjectionCount}`);
  }


  // 2. Process data from Miimangsha Register (current month's new objections and settled amounts)
  // Get the current fortnight period to filter Miimangsha Register data
  const currentFortnight = getFortnightlyPeriod(new Date());
  const currentMonthStartDate = currentFortnight.startDate;
  const currentMonthEndDate = currentFortnight.endDate;
  const timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();


  // Start from the second row (index 1) to skip headers
  for (let i = 1; i < allMiimangshaRegisterData.length; i++) {
    const rowData = allMiimangshaRegisterData[i];

    const ministryName = rowData[MIIMANGSHA_REGISTER_MAPPING.ministryNameCol];
    const institutionName = rowData[MIIMANGSHA_REGISTER_MAPPING.institutionNameCol];
    const dispositionType = rowData[MIIMANGSHA_REGISTER_MAPPING.dispositionTypeCol];
    const involvedAmount = typeof rowData[MIIMANGSHA_REGISTER_MAPPING.involvedAmountCol] === 'number' ? rowData[MIIMANGSHA_REGISTER_MAPPING.involvedAmountCol] : 0;
    const jariPatraDateValue = rowData[MIIMANGSHA_REGISTER_MAPPING.jariPatraDateCol];

    // Ensure jariPatraDate is a valid Date object and falls within the current fortnight
    if (jariPatraDateValue instanceof Date) {
      const jariPatraDate = jariPatraDateValue;

      // Check if the date falls within the current fortnight period
      const formattedJariPatraDate = Utilities.formatDate(jariPatraDate, timeZone, "yyyy-MM-dd");
      const formattedFortnightStartDate = Utilities.formatDate(currentMonthStartDate, timeZone, "yyyy-MM-dd");
      const formattedFortnightEndDate = Utilities.formatDate(currentMonthEndDate, timeZone, "yyyy-MM-dd");

      if (formattedJariPatraDate >= formattedFortnightStartDate && formattedJariPatraDate <= formattedFortnightEndDate) {
        // Only process if the ministry and institution are relevant for this return
        if (ministryName === ministryType && combinedMinistryInstitutionData.has(ministryType) && combinedMinistryInstitutionData.get(ministryType).has(institutionName)) {
          const entry = combinedMinistryInstitutionData.get(ministryType).get(institutionName);

          // 'চলতি মাসে উত্থাপিত আপত্তি' সর্বদা 0 থাকবে
          entry.currentMonthObjectionCount = 0;
          entry.currentMonthObjectionAmount = 0;

          // 'পূর্ণাঙ্গ' হলে সংখ্যা গণনায় আসবে
          if (dispositionType === "পূর্ণাঙ্গ") {
            entry.currentMonthSettledCount += 1;
          }

          // 'আংশিক' অথবা 'পূর্ণাঙ্গ' উভয় হলেই টাকার পরিমাণ গণনায় আসবে
          if (dispositionType === "পূর্ণাঙ্গ" || dispositionType === "আংশিক") {
            entry.currentMonthSettledAmount += involvedAmount;
          }
          Logger.log(`Processed Miimangsha Register Entry: Institution: ${institutionName}, Disposition: ${dispositionType}, Amount: ${involvedAmount}, Date: ${jariPatraDate.toLocaleDateString()}, CurrentSettledCount: ${entry.currentMonthSettledCount}, CurrentSettledAmount: ${entry.currentMonthSettledAmount}`);
        } else {
          Logger.log(`Skipping Miimangsha Register Entry: Ministry '${ministryName}' or Institution '${institutionName}' not relevant for this return or not found in template mapping.`);
        }
      } else {
        Logger.log(`Skipping Miimangsha Register Entry: Date ${jariPatraDate.toLocaleDateString()} is outside current fortnight (${currentMonthStartDate.toLocaleDateString()} - ${currentMonthEndDate.toLocaleDateString()}).`);
      }
    } else {
      Logger.log(`Skipping Miimangsha Register Entry: Invalid Jari Patra Date value: ${jariPatraDateValue} in row ${i + 1}.`);
    }
  }


  // 3. Calculate final totals for each institution and grand totals
  const finalReturnData = {
    ministries: {},
    grandTotals: {
      prevMonthObjectionCount: 0,
      prevMonthObjectionAmount: 0,
      currentMonthObjectionCount: 0,
      currentMonthObjectionAmount: 0,
      totalObjectionCount: 0,
      totalObjectionAmount: 0,
      prevMonthSettledCount: 0,
      prevMonthSettledAmount: 0,
      currentMonthSettledCount: 0,
      currentMonthSettledAmount: 0,
      totalSettledCount: 0,
      totalSettledAmount: 0,
      unsettledCount: 0,
      unsettledAmount: 0,
      broadsheetReplyCount: 0
    }
  };


  // Ensure ministry entry exists
  if (!finalReturnData.ministries[ministryType]) {
    finalReturnData.ministries[ministryType] = {
      institutions: {},
      ministryTotals: {
        prevMonthObjectionCount: 0,
        prevMonthObjectionAmount: 0,
        currentMonthObjectionCount: 0,
        currentMonthObjectionAmount: 0,
        totalObjectionCount: 0,
        totalObjectionAmount: 0,
        prevMonthSettledCount: 0,
        prevMonthSettledAmount: 0,
        currentMonthSettledCount: 0,
        currentMonthSettledAmount: 0,
        totalSettledCount: 0,
        totalSettledAmount: 0,
        unsettledCount: 0,
        unsettledAmount: 0,
        broadsheetReplyCount: 0
      }
    };
  }


  // Iterate over institutions that were initialized from the template (and potentially updated from Miimangsha Register)
  // Ensure consistent order based on TEMPLATE_ROW_MAP.institutions
  const sortedInstitutions = Object.keys(TEMPLATE_ROW_MAP.institutions).sort((a, b) => {
    return TEMPLATE_ROW_MAP.institutions[a] - TEMPLATE_ROW_MAP.institutions[b];
  });


  sortedInstitutions.forEach(institutionName => {
    // Only process if the institution actually has data (either from template or miimangsha register)
    if (combinedMinistryInstitutionData.get(ministryType) && combinedMinistryInstitutionData.get(ministryType).has(institutionName)) {
      const data = combinedMinistryInstitutionData.get(ministryType).get(institutionName);


      // Calculate combined totals for each institution
      data.totalObjectionCount = data.prevMonthObjectionCount + data.currentMonthObjectionCount;
      data.totalObjectionAmount = data.prevMonthObjectionAmount + data.currentMonthObjectionAmount;
      data.totalSettledCount = data.prevMonthSettledCount + data.currentMonthSettledCount;
      data.totalSettledAmount = data.prevMonthSettledAmount + data.currentMonthSettledAmount;
      data.unsettledCount = data.totalObjectionCount - data.totalSettledCount;
      data.unsettledAmount = data.totalObjectionAmount - data.totalSettledAmount;


      // Ensure unsettled amounts are not negative
      if (data.unsettledCount < 0) data.unsettledCount = 0;
      if (data.unsettledAmount < 0) data.unsettledAmount = 0;


      finalReturnData.ministries[ministryType].institutions[institutionName] = data;


      // Accumulate ministry subtotals (these are the numbers that will be calculated by formulas in the sheet)
      // We are still calculating them in the script to provide a complete data structure for ministry totals,
      // but they will *not* be directly written to the sheet in the `generateReturn` function for sums.
      const ministryTotals = finalReturnData.ministries[ministryType].ministryTotals;
      ministryTotals.prevMonthObjectionCount += data.prevMonthObjectionCount;
      ministryTotals.prevMonthObjectionAmount += data.prevMonthObjectionAmount;
      ministryTotals.currentMonthObjectionCount += data.currentMonthObjectionCount;
      ministryTotals.currentMonthObjectionAmount += data.currentMonthObjectionAmount;
      ministryTotals.totalObjectionCount += data.totalObjectionCount;
      ministryTotals.totalObjectionAmount += data.totalObjectionAmount;
      ministryTotals.prevMonthSettledCount += data.prevMonthSettledCount;
      ministryTotals.prevMonthSettledAmount += data.prevMonthSettledAmount;
      ministryTotals.currentMonthSettledCount += data.currentMonthSettledCount;
      ministryTotals.currentMonthSettledAmount += data.currentMonthSettledAmount;
      ministryTotals.totalSettledCount += data.totalSettledCount;
      ministryTotals.totalSettledAmount += data.totalSettledAmount;
      ministryTotals.unsettledCount += data.unsettledCount;
      ministryTotals.unsettledAmount += data.unsettledAmount;
      ministryTotals.broadsheetReplyCount += data.broadsheetReplyCount;


      // Accumulate grand totals (if there were multiple ministries, this would sum across them)
      // For a single ministry return, grand totals will be same as ministry totals
      finalReturnData.grandTotals.prevMonthObjectionCount += data.prevMonthObjectionCount;
      finalReturnData.grandTotals.prevMonthObjectionAmount += data.prevMonthObjectionAmount;
      finalReturnData.grandTotals.currentMonthObjectionCount += data.currentMonthObjectionCount;
      finalReturnData.grandTotals.currentMonthObjectionAmount += data.currentMonthObjectionAmount;
      finalReturnData.grandTotals.totalObjectionCount += data.totalObjectionCount;
      finalReturnData.grandTotals.totalObjectionAmount += data.totalObjectionAmount;
      finalReturnData.grandTotals.prevMonthSettledCount += data.prevMonthSettledCount;
      finalReturnData.grandTotals.prevMonthSettledAmount += data.prevMonthSettledAmount;
      finalReturnData.grandTotals.currentMonthSettledCount += data.currentMonthSettledCount;
      finalReturnData.grandTotals.currentMonthSettledAmount += data.currentMonthSettledAmount;
      finalReturnData.grandTotals.totalSettledCount += data.totalSettledCount;
      finalReturnData.grandTotals.totalSettledAmount += data.totalSettledAmount;
      finalReturnData.grandTotals.unsettledCount += data.unsettledCount;
      finalReturnData.grandTotals.unsettledAmount += data.unsettledAmount;
      finalReturnData.broadsheetReplyCount += data.broadsheetReplyCount;
    }
  });


  Logger.log("--- getFortnightlyReturnData finished ---");
  return finalReturnData;
}


/**
 * নির্দিষ্ট রিটার্ণ পিরিয়ড এবং মন্ত্রণালয়ের জন্য পাক্ষিক রিটার্ণ তৈরি করে।
 * @param {string} returnPeriod - রিটার্ণ পিরিয়ড (e.g., 'পাক্ষিক', 'মাসিক')
 * @param {string} ministryType - মন্ত্রণালয়ের ধরণ (e.g., 'আর্থিক প্রতিষ্ঠান বিভাগ')
 * @returns {Object} - সফল বা ব্যর্থতার বার্তা এবং ফাইল URL।
 */
function generateReturn(returnPeriod, ministryType) {
  Logger.log(`generateReturn function started for Period: ${returnPeriod}, Ministry: ${ministryType}`);
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);


  if (!ss) {
    Logger.log(`Error: Spreadsheet with ID '${SPREADSHEET_ID}' could not be opened in generateReturn.`);
    return { success: false, message: `মীমাংসা রেজিস্টার খুলতে ব্যর্থ। আইডি '${SPREADSHEET_ID}' সঠিক কিনা এবং আপনার Apps Script প্রোজেক্টের এর উপর প্রয়োজনীয় অনুমতি আছে কিনা নিশ্চিত করুন।` };
  }


  const templateSheetName = TEMPLATE_SHEET_NAMES[ministryType];
  if (!templateSheetName) {
    Logger.log(`Error: No template sheet name defined for ministry type: ${ministryType}`);
    return { success: false, message: `এই মন্ত্রণালয়ের জন্য কোনো টেমপ্লেট খুঁজে পাওয়া যায়নি: ${ministryType}` };
  }


  Logger.log(`Accessing template sheet: '${templateSheetName}'`);
  const templateSheet = ss.getSheetByName(templateSheetName); // মূল টেমপ্লেট শীট
  if (!templateSheet) {
    Logger.log(`Error: Template sheet named '${templateSheetName}' not found.`);
    return { success: false, message: `টেমপ্লেট শীট খুঁজে পাওয়া যায়নি: ${templateSheetName}` };
  }


  // 1. Get processed data
  const returnData = getFortnightlyReturnData(ministryType);
  if (!returnData || Object.keys(returnData.ministries).length === 0) {
    Logger.log('No relevant data found for the return. Please check your data sheets.');
    return { success: false, message: 'রিটার্নের জন্য কোনো প্রাসঙ্গিক ডেটা খুঁজে পাওয়া হয়নি।' };
  }
  Logger.log('Data successfully processed for return generation.');


  // 2. Create a copy of the template sheet
  const currentFortnight = getFortnightlyPeriod(new Date());
  const newSheetName = `${ministryType} - ${returnPeriod} রিটার্ণ - ${currentFortnight.startDateBengla} থেকে ${currentFortnight.endDateBengla}`;
  Logger.log(`Creating new sheet named: ${newSheetName}`);
  const returnSheet = templateSheet.copyTo(ss).setName(newSheetName); // **পুনরায় চালু করা হয়েছে: নতুন কপি তৈরি করা হচ্ছে**
  Logger.log(`New sheet '${newSheetName}' created.`);


  // Set sharing permissions for the newly created sheet
  try {
    const newFile = DriveApp.getFileById(ss.getId()); // Get the spreadsheet file object
    newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    Logger.log(`Set sharing permissions for new sheet ID: ${returnSheet.getId()}`);
  } catch (e) {
    Logger.log(`Error setting sharing permissions for new sheet: ${e.message}`);
    // Continue even if setting permissions fails, but log the error.
  }


  // --- Start: Clear content and borders for columns L to S in rows 29 to 43 (if they exist) ---
  // Based on "অতিরিক্ত অংশ.png" and "বাড়তি অংশ.png", this area refers to rows 29-43 and columns L-S.
  // We are not deleting columns/rows, just clearing content and formatting.
  const startRowToClear = 29;
  const endRowToClear = 43;
  const startColToClear = 12; // L column
  const endColToClear = 19;   // S column


  Logger.log(`Clearing content and borders for range R${startRowToClear}C${startColToClear}:R${endRowToClear}C${endColToClear}`);


  if (startRowToClear <= returnSheet.getLastRow() && endColToClear <= returnSheet.getLastColumn()) {
    const rangeToClear = returnSheet.getRange(startRowToClear, startColToClear, endRowToClear - startRowToClear + 1, endColToClear - startColToClear + 1);
    rangeToClear.clearContent(); // Clear content
    rangeToClear.clearFormat();  // Clear formatting (including borders)
    Logger.log('Content and borders cleared for the specified range.');
  } else {
    Logger.log('Specified range for clearing content/borders is outside sheet dimensions. Skipping clear operation.');
  }
  // --- End: Clear content and borders ---


  // 3. Populate Fortnightly Period
  if (TEMPLATE_ROW_MAP.periodInfo !== -1) {
    returnSheet.getRange(TEMPLATE_ROW_MAP.periodInfo + 1, 2).setValue(`${currentFortnight.startDateBengla} থেকে ${currentFortnight.endDateBengla}`); // Column B, 1-based
    Logger.log(`Populated period: ${currentFortnight.startDateBengla} থেকে ${currentFortnight.endDateBengla}`);
  } else {
    Logger.log("Warning: Could not find 'সময়কাল:' row in the template to populate period information.");
  }


  // 4. Populate Ministry Name (merged cells A9:A23 for 'আর্থিক প্রতিষ্ঠান বিভাগ')
  if (TEMPLATE_ROW_MAP.ministryNameStart !== -1) {
    const ministryRange = returnSheet.getRange(TEMPLATE_ROW_MAP.ministryNameStart + 1, 1, 15, 1); // A9 to A23, 15 rows, 1 col
    // নতুন কপিতে কাজ করায় unmerge এর প্রয়োজন নেই, কারণ কপিটি টেমপ্লেটের ফরম্যাট নিয়েই তৈরি হবে।
    ministryRange.merge();
    ministryRange.setValue(ministryType);
    ministryRange.setVerticalAlignment('middle');
    ministryRange.setHorizontalAlignment('center');
    ministryRange.setBorder(true, true, true, true, true, true); // Add border
    Logger.log(`Populated ministry name '${ministryType}' in merged cells.`);
  }


  // 5. Populate Table 1 and Table 2 Data for institutions
  const institutionsData = returnData.ministries[ministryType].institutions;


  // Populate institution data rows first
  // Ensure consistent order based on TEMPLATE_ROW_MAP.institutions
  const sortedInstitutions = Object.keys(TEMPLATE_ROW_MAP.institutions).sort((a, b) => {
    return TEMPLATE_ROW_MAP.institutions[a] - TEMPLATE_ROW_MAP.institutions[b];
  });

  let serialCounter = 1; // Initialize serial number counter for the new return sheet

  for (const institutionName of sortedInstitutions) {
    const data = institutionsData[institutionName];
    const templateRow = TEMPLATE_ROW_MAP.institutions[institutionName]; // Use TEMPLATE_ROW_MAP.institutions for row


    if (templateRow !== undefined && templateRow !== -1) {
      const sheetRow = templateRow + 1; // Convert to 1-based index for sheet operations

      // Set the serial number in Column A (1-based index 1)
      returnSheet.getRange(sheetRow, 1).setValue(serialCounter++); // Set and increment serial number
      returnSheet.getRange(sheetRow, 1).setHorizontalAlignment('center'); // A কলাম: ক্রমিক নং - মাঝখানে
      returnSheet.getRange(sheetRow, 1).setVerticalAlignment('middle'); // উল্লম্বভাবে মাঝখানে

      // Set "সংস্থার নাম" (Column B) to left alignment
      returnSheet.getRange(sheetRow, 2).setHorizontalAlignment('left'); // B কলাম: সংস্থার নাম - বাম সারিবদ্ধ
      returnSheet.getRange(sheetRow, 2).setVerticalAlignment('middle'); // উল্লম্বভাবে মাঝখানে

      // Table 1 Data (C to J for individual institutions) - মাঝখানে সারিবদ্ধ
      returnSheet.getRange(sheetRow, TEMPLATE_COL_MAP.table1.prevMonthObjectionCount).setHorizontalAlignment('center');
      returnSheet.getRange(sheetRow, TEMPLATE_COL_MAP.table1.prevMonthObjectionAmount).setHorizontalAlignment('center');
      returnSheet.getRange(sheetRow, TEMPLATE_COL_MAP.table1.currentMonthObjectionCount).setHorizontalAlignment('center');
      returnSheet.getRange(sheetRow, TEMPLATE_COL_MAP.table1.currentMonthObjectionAmount).setHorizontalAlignment('center');
      returnSheet.getRange(sheetRow, TEMPLATE_COL_MAP.table1.totalObjectionCount).setHorizontalAlignment('center');
      returnSheet.getRange(sheetRow, TEMPLATE_COL_MAP.table1.totalObjectionAmount).setHorizontalAlignment('center');
      returnSheet.getRange(sheetRow, TEMPLATE_COL_MAP.table1.prevMonthSettledCount).setHorizontalAlignment('center');
      returnSheet.getRange(sheetRow, TEMPLATE_COL_MAP.table1.prevMonthSettledAmount).setHorizontalAlignment('center');

      // Table 2 Data (L to S for individual institutions) - মাঝখানে সারিবদ্ধ
      returnSheet.getRange(sheetRow, TEMPLATE_COL_MAP.table2.currentMonthSettledCount).setHorizontalAlignment('center');
      returnSheet.getRange(sheetRow, TEMPLATE_COL_MAP.table2.currentMonthSettledAmount).setHorizontalAlignment('center');
      returnSheet.getRange(sheetRow, TEMPLATE_COL_MAP.table2.totalSettledCount).setHorizontalAlignment('center');
      returnSheet.getRange(sheetRow, TEMPLATE_COL_MAP.table2.totalSettledAmount).setHorizontalAlignment('center');
      returnSheet.getRange(sheetRow, TEMPLATE_COL_MAP.table2.unsettledCount).setHorizontalAlignment('center');
      returnSheet.getRange(sheetRow, TEMPLATE_COL_MAP.table2.unsettledAmount).setHorizontalAlignment('center');
      returnSheet.getRange(sheetRow, TEMPLATE_COL_MAP.table2.broadsheetReplyCount).setHorizontalAlignment('center');
      returnSheet.getRange(sheetRow, TEMPLATE_COL_MAP.table2.comments).setHorizontalAlignment('center');

      // C থেকে S কলাম পর্যন্ত উল্লম্বভাবে মাঝখানে সারিবদ্ধ (একক রেঞ্জে সেট করা)
      returnSheet.getRange(sheetRow, 3, 1, 17).setVerticalAlignment('middle');


      // Write the data to the sheet
      returnSheet.getRange(sheetRow, TEMPLATE_COL_MAP.table1.prevMonthObjectionCount).setValue(data.prevMonthObjectionCount);
      returnSheet.getRange(sheetRow, TEMPLATE_COL_MAP.table1.prevMonthObjectionAmount).setValue(data.prevMonthObjectionAmount);
      returnSheet.getRange(sheetRow, TEMPLATE_COL_MAP.table1.currentMonthObjectionCount).setValue(data.currentMonthObjectionCount); // Should be 0
      returnSheet.getRange(sheetRow, TEMPLATE_COL_MAP.table1.currentMonthObjectionAmount).setValue(data.currentMonthObjectionAmount); // Should be 0
      returnSheet.getRange(sheetRow, TEMPLATE_COL_MAP.table1.totalObjectionCount).setValue(data.totalObjectionCount);
      returnSheet.getRange(sheetRow, TEMPLATE_COL_MAP.table1.totalObjectionAmount).setValue(data.totalObjectionAmount);
      returnSheet.getRange(sheetRow, TEMPLATE_COL_MAP.table1.prevMonthSettledCount).setValue(data.prevMonthSettledCount);
      returnSheet.getRange(sheetRow, TEMPLATE_COL_MAP.table1.prevMonthSettledAmount).setValue(data.prevMonthSettledAmount);

      Logger.log(`Populated data for institution ${institutionName} in row ${sheetRow} (Table 1 & 2)`);
    } else {
      Logger.log(`Warning: Could not find template row for institution: ${institutionName}`);
    }
  }


  // Populate Ministry Total (২৪ নং রো-তে ফর্মুলা ব্যবহার করে যোগফল)
  const ministryTotalRow = TEMPLATE_ROW_MAP.ministryTotalRows[ministryType]; // 0-indexed
  if (ministryTotalRow !== undefined && ministryTotalRow !== -1) {
    const sheetRow = ministryTotalRow + 1; // 1-based for sheet.getRange
    const dataStartRowFor24 = TEMPLATE_ROW_MAP.table1HeaderStart + 3; // 1-based start row for institution data (Row 9)
    const dataEndRowFor24 = sheetRow - 1; // The row just before the total row (Row 23)


    // C কলাম (index 3) থেকে J কলাম (index 10) পর্যন্ত রেঞ্জ পরিষ্কার করুন।
    // এই রেঞ্জটি (1-based indices for getRange) = (C to J) = (col 3 to col 10)
    const totalRange1 = returnSheet.getRange(sheetRow, 3, 1, 8); // 8 columns: C,D,E,F,G,H,I,J
    totalRange1.clearContent().clearFormat();


    // Table 1 Ministry Totals (C থেকে J পর্যন্ত) - ফর্মুলা বসানো হচ্ছে
    // TEMPLATE_COL_MAP is 1-based, getRange is 1-based.
    const colsToSumForRow24_1Based = [
      TEMPLATE_COL_MAP.table1.prevMonthObjectionCount,
      TEMPLATE_COL_MAP.table1.prevMonthObjectionAmount,
      TEMPLATE_COL_MAP.table1.currentMonthObjectionCount,
      TEMPLATE_COL_MAP.table1.currentMonthObjectionAmount,
      TEMPLATE_COL_MAP.table1.totalObjectionCount,
      TEMPLATE_COL_MAP.table1.totalObjectionAmount,
      TEMPLATE_COL_MAP.table1.prevMonthSettledCount,
      TEMPLATE_COL_MAP.table1.prevMonthSettledAmount
    ];


    for (const colIndex of colsToSumForRow24_1Based) {
      const colLetter = String.fromCharCode(64 + colIndex); // কলাম লেটার বের করুন (যেমন C, D, E...)
      const sumFormula = `=SUM(${colLetter}${dataStartRowFor24}:${colLetter}${dataEndRowFor24})`;
      returnSheet.getRange(sheetRow, colIndex).setFormula(sumFormula);
      returnSheet.getRange(sheetRow, colIndex).setHorizontalAlignment('center'); // নিশ্চিত করুন যোগফলও মাঝখানে থাকে
      returnSheet.getRange(sheetRow, colIndex).setVerticalAlignment('middle');
    }
    // ২৪ নং রো তে C কলাম থেকে J কলাম পর্যন্ত বর্ডার যোগ করা
    totalRange1.setBorder(true, true, true, true, true, true);
    Logger.log(`Populated Ministry Total formulas for ${ministryType} in row ${sheetRow} (C to J) with borders.`);


    // Table 2 Ministry Totals (L থেকে S পর্যন্ত) - এখানে কোন ফর্মুলা বসানো হবে না কারণ এই কলামগুলোর ডেটা/বর্ডার মুছে ফেলার কথা বলা হয়েছে।
    // তবে রো ২৪ এ K কলামে 1821 সংখ্যাটি আছে ("অতিরিক্ত অংশ.png"), তাই সেই সেলের বর্ডার নিশ্চিত করা হবে।
    // returnSheet.getRange(sheetRow, 11).setBorder(true, true, true, true, true, true); // Column K, Row 24 - এই লাইনটি পূর্ববর্তী অনুরোধে মুছে ফেলা হয়েছে।
    Logger.log('Confirmed border for K24.');
  } else {
    Logger.log(`Warning: Could not find Ministry Total row for ${ministryType} in template.`);
  }


  // Populate Another Total Row (৪৪ নং রো-তে ফর্মুলা ব্যবহার করে যোগফল)
  // নিশ্চিত করুন রো ৪৪ (0-indexed 43) এ ফর্মুলা বসানো হবে।
  const anotherTotalRow = TEMPLATE_ROW_MAP.anotherTotalRow; // 0-indexed
  if (anotherTotalRow !== undefined && anotherTotalRow !== -1) {
    const sheetRowFor44 = anotherTotalRow + 1; // 1-based for sheet.getRange


    // রো ৪৪ এর জন্য যোগফলের ডেটা রেঞ্জ।
    // এই ক্ষেত্রে, রো ২৯ থেকে ৪৩ পর্যন্ত প্রতিষ্ঠানের ডেটার যোগফল।
    const dataStartRowFor44 = TEMPLATE_ROW_MAP.table2HeaderStart + 2; // Row 29 (0-indexed 26 + 2 = 28, then +1 for 1-based = 29)
    const dataEndRowFor44 = sheetRowFor44 - 1; // Row 43


    // C কলাম (index 3) থেকে J কলাম (index 10) পর্যন্ত রেঞ্জ পরিষ্কার করুন।
    // এই রেঞ্জটি (1-based indices for getRange) = (C to J) = (col 3 to col 10)
    const totalRange44 = returnSheet.getRange(sheetRowFor44, 3, 1, 8); // 8 columns: C,D,E,F,G,H,I,J
    totalRange44.clearContent().clearFormat();


    // C থেকে J পর্যন্ত কলামে ফর্মুলা বসানো হচ্ছে (টেবিল ১ এর কলামগুলো)।
    // ছবিতে J পর্যন্ত বর্ডার এবং ডেটা দেখা যাচ্ছে।
    const colsToSumForRow44_1Based = [3, 4, 5, 6, 7, 8, 9, 10]; // C to J (1-based column indices)


    for (const colIndex of colsToSumForRow44_1Based) {
      const colLetter = String.fromCharCode(64 + colIndex); // কলাম লেটার বের করুন
      const sumFormula = `=SUM(${colLetter}${dataStartRowFor44}:${colLetter}${dataEndRowFor44})`;
      returnSheet.getRange(sheetRowFor44, colIndex).setFormula(sumFormula);
      returnSheet.getRange(sheetRowFor44, colIndex).setHorizontalAlignment('center'); // নিশ্চিত করুন যোগফলও মাঝখানে থাকে
      returnSheet.getRange(sheetRowFor44, colIndex).setVerticalAlignment('middle');
    }
    // রো ৪৪ তে C কলাম থেকে J কলাম পর্যন্ত বর্ডার যোগ করা
    totalRange44.setBorder(true, true, true, true, true, true);
    Logger.log(`Populated formulas for another total row in row ${sheetRowFor44} (C to J) with borders.`);
  } else {
    Logger.log("Warning: Could not find 'anotherTotalRow' in template to populate sums.");
  }


  // কলাম I এবং J এর ২৯ নং রো থেকে নিচ পযন্ত সেলগুলো বর্ডার দেওয়া নেই। যা দিতে হবে।
  // "সেল মিসিং.png" ছবিতে রো ২৯ থেকে ৪৩ পর্যন্ত I ও J কলামে বর্ডার নেই।
  // রো ২৯ থেকে ৪৩ পর্যন্ত C থেকে J কলামে বর্ডার যুক্ত করা হবে।
  const startRowForBorders_C_J = 29; // 1-based
  const endRowForBorders_C_J = 43; // 1-based
  const startColForBorders_C_J = 3; // C column
  const endColForBorders_C_J = 10; // J column


  if (startRowForBorders_C_J <= returnSheet.getLastRow() && endColForBorders_C_J <= returnSheet.getLastColumn()) {
    const rangeToBorder_C_J = returnSheet.getRange(startRowForBorders_C_J, startColForBorders_C_J, endRowForBorders_C_J - startRowForBorders_C_J + 1, endColForBorders_C_J - startColForBorders_C_J + 1);
    rangeToBorder_C_J.setBorder(true, true, true, true, true, true);
    Logger.log(`Added borders to columns C to J from row ${startRowForBorders_C_J} to ${endRowForBorders_C_J}.`);
  }


  SpreadsheetApp.flush();
  // সফলতার বার্তায় নতুন শীটের নাম উল্লেখ করা হয়েছে।
  return { success: true, message: `"${newSheetName}" সফলভাবে তৈরি করা হয়েছে!`, fileUrl: ss.getUrl() };
}

/**
 * Apps Script মেনুতে নতুন আইটেম তৈরি করে।
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('মীমাংসা রেজিস্টার')
    .addItem('ডেটা এন্ট্রি ফরম খুলুন', 'showForm')
    .addSubMenu(ui.createMenu('পাক্ষিক রিপোর্ট তৈরি করুন')
      .addItem('আর্থিক প্রতিষ্ঠান বিভাগ', 'generateFinancialInstitutionReport')
      .addItem('বস্ত্র ও পাট', 'generateTextileReport') // এই ফাংশন তৈরি করতে হবে
      .addItem('শিল্প+বাণিজ্য+বিমান', 'generateIndustryCommerceAviationReport') // এই ফাংশন তৈরি করতে হবে
    )
    .addToUi();
}

/**
 * ওয়েব অ্যাপ ফরম দেখায়।
 */
function showForm() {
  const html = HtmlService.createHtmlOutputFromFile('meemangsa-register-form-html')
    .setWidth(700)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'মীমাংসা রেজিস্টার ডেটা এন্ট্রি');
}


// --- Functions to call generateReturn for specific ministries ---
function generateFinancialInstitutionReport() {
  try {
    generateReturn('পাক্ষিক', 'আর্থিক প্রতিষ্ঠান বিভাগ'); // returnPeriod 'পাক্ষিক' হিসেবে পাস করা হয়েছে
    SpreadsheetApp.getUi().alert('পাক্ষিক রিপোর্ট তৈরি হয়েছে', 'আর্থিক প্রতিষ্ঠান বিভাগের রিপোর্ট সফলভাবে তৈরি করা হয়েছে!', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('ত্রুটি', 'রিপোর্ট তৈরিতে সমস্যা হয়েছে: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log(e.message);
  }
}

function generateTextileReport() {
  try {
    // এই ফাংশনটি 'বস্ত্র ও পাট' মন্ত্রণালয়ের জন্য generateReturn কল করবে।
    // আপনাকে TEMPLATE_SHEET_NAMES এবং TEMPLATE_ROW_MAP এ এই মন্ত্রণালয়ের টেমপ্লেট ও রো ম্যাপিং যোগ করতে হবে।
    generateReturn('পাক্ষিক', 'বস্ত্র ও পাট'); // returnPeriod 'পাক্ষিক' হিসেবে পাস করা হয়েছে
    SpreadsheetApp.getUi().alert('পাক্ষিক রিপোর্ট তৈরি হয়েছে', 'বস্ত্র ও পাট মন্ত্রণালয়ের রিপোর্ট সফলভাবে তৈরি করা হয়েছে!', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('ত্রুটি', 'রিপোর্ট তৈরিতে সমস্যা হয়েছে: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log(e.message);
  }
}

function generateIndustryCommerceAviationReport() {
  try {
    // এই ফাংশনটি 'শিল্প+বাণিজ্য+বিমান' মন্ত্রণালয়ের জন্য generateReturn কল করবে।
    // আপনাকে TEMPLATE_SHEET_NAMES এবং TEMPLATE_ROW_MAP এ এই মন্ত্রণালয়ের টেমপ্লেট ও রো ম্যাপিং যোগ করতে হবে।
    generateReturn('পাক্ষিক', 'শিল্প+বাণিজ্য+বিমান'); // returnPeriod 'পাক্ষিক' হিসেবে পাস করা হয়েছে
    SpreadsheetApp.getUi().alert('পাক্ষিক রিপোর্ট তৈরি হয়েছে', 'শিল্প+বাণিজ্য+বিমান মন্ত্রণালয়ের রিপোর্ট সফলভাবে তৈরি করা হয়েছে!', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('ত্রুটি', 'রিপোর্ট তৈরিতে সমস্যা হয়েছে: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log(e.message);
  }
}
