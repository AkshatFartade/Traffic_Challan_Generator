/**
 * strings.js — Centralized UI text dictionary for Traffic Notice Generator.
 * All user-facing text, error messages, and state labels are maintained here.
 */
const STRINGS = {
  header: {
    office: "कार्यालय अतिरिक्त पुलिस अधीक्षक यातायात / Office of Additional Superintendent of Police Traffic",
    titleHi: "यातायात नोटिस जनरेटर",
    titleEn: "Traffic Notice Generator",
  },
  steps: {
    upload: "1 Upload",
    check: "2 Check",
    generate: "3 Generate",
    download: "4 Download",
  },
  state1_empty: {
    heading: "Upload Challan Register",
    desc: "Select an official e-challan export file (.xlsx or .csv). The system detects the data block and verifies all records before generation.",
    dropPrompt: "Drop the Excel file here, or choose a file",
    dropHint: ".xlsx or .csv (up to 50 MB)",
    columnsNote: "Expected layout (columns B to J): Serial Number, E-Challan Number, Vehicle Number, Violation Date, Name, Address, Mobile Number, Fine Amount, Section.",
  },
  state2_file_chosen: {
    replace: "Replace file",
    remove: "Remove",
    inspectBtn: "Check File",
  },
  state3_checking: {
    title: "Inspecting file...",
    desc: "Scanning sheets, verifying column alignments, and checking preflight rules...",
  },
  state4_report: {
    heading: "Preflight Check Report",
    foundNotice: (count, sheet, start, end) => `Found ${count.toLocaleString()} notices on sheet ${sheet}, rows ${start} to ${end}.`,
    foundNoticeCsv: (count) => `Found ${count.toLocaleString()} notices in CSV data.`,
    previewTitle: "Record Preview (First 5 and Last 5 Rows)",
    blockingTitle: "Blocking Errors (Generation Disabled)",
    blockingHelp: "The file cannot be processed until these items are corrected in the Excel source.",
    warningsTitle: "Warnings for Supervisor Review",
    reviewCheckbox: "I have reviewed these warnings and approve proceeding with generation",
    downloadProblems: "Download list of flagged rows (.xlsx)",
    outcomeSingle: "This will create 1 PDF.",
    outcomeMulti: (batches, batchSize) => `This will create ${batches} PDFs inside one ZIP file (up to ${batchSize.toLocaleString()} notices each).`,
    outcomeOdd: "The last sheet will have one notice; its lower half stays blank.",
    btnStart: "Start Generation",
  },
  state5_generating: {
    heading: "Generating Notices",
    singleBatchNotice: "Converting to PDF. About 2 to 3 minutes for 1,000 notices on this computer.",
    batchProgress: (done, total) => `Batch ${done} of ${total}`,
    refreshNotice: "You can safely refresh this page. Generation runs in an isolated background process.",
    btnCancel: "Cancel Generation",
  },
  state6_verifying: {
    heading: "Verifying Layout & Positions",
    desc: "Checking that every notice is in the right place (Gate G1, G2, G3)...",
  },
  state7_done: {
    heading: "Generation Complete",
    verifiedSuccess: (count) => `${count.toLocaleString()} of ${count.toLocaleString()} notices verified against your Excel data.`,
    btnDownloadPdf: "Download Notices PDF",
    btnDownloadZip: "Download All Batches (ZIP)",
    previewLabel: "First Page Preview",
    retentionNotice: (minutes) => `Generated files are automatically deleted after ${minutes} minutes.`,
    individualBatches: "Individual Batch Downloads",
  },
  state8_failed: {
    heading: "Verification Failed",
    desc: "One or more safety verification gates failed. To protect official record integrity, PDF download has been disabled.",
  },
  errors: {
    network: "Connection lost. Please check if the server is running and refresh.",
    wrongType: "Unsupported file format. Please upload a modern Excel file (.xlsx) or CSV (.csv).",
    tooLarge: "File exceeds maximum upload size (50 MB).",
    sessionExpired: "Your session has expired. Please log in again.",
    cancelled: "Generation was cancelled.",
    general: "An error occurred while processing your request.",
  }
};
