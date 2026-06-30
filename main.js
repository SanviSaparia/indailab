// Initialize PDF.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.4.120/pdf.worker.min.js';

// Application State
const camsState = {
  pdfBytes: null,
  pdfFilename: '',
  isZipUpload: false,
  processedPdfBytes: null,
  processedExcelBytes: null
};

const kfintechState = {
  pdfBytes: null,
  pdfFilename: '',
  isZipUpload: false,
  processedPdfBytes: null,
  processedExcelBytes: null
};

const signState = {
  imageBytes: null,
  imageFilename: '',
  imageType: ''
};

const camsExcelState = {
  fileBytes: null,
  fileFilename: ''
};

const kfintechExcelState = {
  fileBytes: null,
  fileFilename: ''
};

const globalSignature = {
  imageBytes: null,
  imageFilename: '',
  imageType: ''
};

// Base64 <-> ArrayBuffer helper functions
function arrayBufferToBase64(buffer) {
  let binary = '';
  const bytes = new Uint8Array(buffer);
  const len = bytes.byteLength;
  for (let i = 0; i < len; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return window.btoa(binary);
}

function base64ToArrayBuffer(base64) {
  const binary_string = window.atob(base64);
  const len = binary_string.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) {
    bytes[i] = binary_string.charCodeAt(i);
  }
  return bytes.buffer;
}

function getMonthDigitFromText(text) {
  if (!text) return '06'; // default fallback
  const months = [
    { name: 'january', digit: '01' }, { name: 'february', digit: '02' }, { name: 'march', digit: '03' },
    { name: 'april', digit: '04' }, { name: 'may', digit: '05' }, { name: 'june', digit: '06' },
    { name: 'july', digit: '07' }, { name: 'august', digit: '08' }, { name: 'september', digit: '09' },
    { name: 'october', digit: '10' }, { name: 'november', digit: '11' }, { name: 'december', digit: '12' },
    { name: 'jan', digit: '01' }, { name: 'feb', digit: '02' }, { name: 'mar', digit: '03' },
    { name: 'apr', digit: '04' }, { name: 'jun', digit: '06' }, { name: 'jul', digit: '07' },
    { name: 'aug', digit: '08' }, { name: 'sep', digit: '09' }, { name: 'oct', digit: '10' },
    { name: 'nov', digit: '11' }, { name: 'dec', digit: '12' }
  ];
  
  const textLower = text.toLowerCase();
  for (const m of months) {
    if (textLower.includes(m.name)) {
      return m.digit;
    }
  }
  
  // If there's a digit form like 05/06/2026, let's extract the middle part as month
  const slashMatch = text.match(/([0-9]{2})\/([0-9]{2})\/([0-9]{4})/);
  if (slashMatch) {
    return slashMatch[2];
  }
  
  return '06'; // default fallback to June
}

function getYearRangeFromInvoiceNo(invoiceNo) {
  if (!invoiceNo) return '26-27'; // fallback default
  const match = invoiceNo.match(/(\d{2}-\d{2})/);
  if (match) {
    return match[1];
  }
  return '26-27'; // fallback default
}

async function extractPdfMetadata(pdfBytes) {
  try {
    const pdf = await pdfjsLib.getDocument({ data: pdfBytes.slice(0) }).promise;
    const numPages = pdf.numPages;
    let fullDocumentText = '';
    for (let pageNum = 1; pageNum <= numPages; pageNum++) {
      const page = await pdf.getPage(pageNum);
      const textContent = await page.getTextContent();
      const pageText = textContent.items.map(item => item.str || '').join(' ');
      fullDocumentText += pageText + '\n';
    }

    // GSTIN extraction
    const gstinRegex = /[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[0-9A-Z]{1}Z[0-9A-Z]{1}/gi;
    const gstins = (fullDocumentText.match(gstinRegex) || []).map(g => g.replace(/\s+/g, '').toUpperCase());

    // Invoice No extraction
    let invoiceNo = '';
    const invNoRegex = /Inv\s+serial\s+No\.\s*:\s*([^\s\r\n]+)/i;
    const matchInvNo = fullDocumentText.match(invNoRegex);
    if (matchInvNo) {
      invoiceNo = matchInvNo[1].trim();
    } else {
      const camsInvNoRegex = /Invoice\s+No\s*:\s*([^\s\r\n]+)/i;
      const matchCamsInvNo = fullDocumentText.match(camsInvNoRegex);
      if (matchCamsInvNo) {
        invoiceNo = matchCamsInvNo[1].trim();
      }
    }

    // Invoice Date extraction
    let invoiceDate = '';
    const dateRegex = /Date\s*:\s*([0-9]{2}\/[0-9]{2}\/[0-9]{4})/i;
    const matchDate = fullDocumentText.match(dateRegex);
    if (matchDate) {
      invoiceDate = matchDate[1].trim();
    } else {
      const camsDateRegex = /Invoice\s+Date\s*:\s*([A-Za-z]+)\s+([0-9]{1,2}),\s*([0-9]{4})/i;
      const matchCamsDate = fullDocumentText.match(camsDateRegex);
      if (matchCamsDate) {
        invoiceDate = matchCamsDate[0].replace(/Invoice\s+Date\s*:\s*/i, '').trim(); // e.g. "June 08, 2026"
      } else {
        const camsDateRegexFallback = /Invoice\s+Date\s*:\s*([^\n\r]+)/i;
        const matchCamsDateFallback = fullDocumentText.match(camsDateRegexFallback);
        if (matchCamsDateFallback) {
          invoiceDate = matchCamsDateFallback[1].trim();
        }
      }
    }

    return { gstins, invoiceNo, invoiceDate, fullText: fullDocumentText };
  } catch (error) {
    console.error("Failed to extract PDF metadata:", error);
    return { gstins: [], invoiceNo: '', invoiceDate: '', fullText: '' };
  }
}

function updateExcelReport(prefix, excelBytes, pdfMetaList) {
  try {
    const workbook = XLSX.read(excelBytes, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    for (let r = 1; r < jsonData.length; r++) {
      const row = jsonData[r];
      if (!row) continue;
      
      if (prefix === 'cams') {
        // CAMS matching: Column E (index 4) contains CAMS INVOICE NUMBER
        if (row.length <= 4) {
          while (row.length <= 4) row.push('');
        }
        const excelInvoiceNo = String(row[4] || '').trim();
        
        // Find a matched PDF where its fullText contains the Excel invoice number
        let matchedPdf = null;
        if (excelInvoiceNo) {
          matchedPdf = pdfMetaList.find(meta => {
            return meta.fullText && meta.fullText.includes(excelInvoiceNo);
          });
        }
        
        // Write the formatted broker invoice number in Column D (index 3) below the header
        const pdfMeta = matchedPdf || pdfMetaList[0];
        if (pdfMeta) {
          const monthDigit = getMonthDigitFromText(pdfMeta.invoiceDate || pdfMeta.fullText);
          const yearRange = getYearRangeFromInvoiceNo(pdfMeta.invoiceNo);
          const brokerInvoiceNo = 'CAMS/' + yearRange + '/000' + monthDigit;
          
          row[3] = brokerInvoiceNo;
          
          if (matchedPdf) {
            while (row.length <= 11) row.push('');
            row[11] = matchedPdf.filename.replace(/\.pdf$/i, '');
          }
        }
      } else if (prefix === 'kfintech') {
        // Kfintech matching: Column D (index 3) contains AMC GSTR Number
        if (row.length <= 3) continue;
        const excelGstr = String(row[3] || '').replace(/\s+/g, '').toUpperCase();
        if (!excelGstr) continue;
        
        const matchedPdf = pdfMetaList.find(meta => {
          return meta.gstins.some(g => g === excelGstr);
        });
        
        if (matchedPdf) {
          while (row.length <= 9) row.push('');
          row[7] = matchedPdf.invoiceNo;
          row[8] = matchedPdf.invoiceDate;
          row[9] = matchedPdf.filename.replace(/\.pdf$/i, '');
        }
      }
    }
    
    const updatedWorksheet = XLSX.utils.aoa_to_sheet(jsonData);
    workbook.Sheets[sheetName] = updatedWorksheet;
    return XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
  } catch (error) {
    console.error("Failed to update Excel report:", error);
    return excelBytes;
  }
}

function loadUserSignature(email) {
  // Populate profile info details dynamically
  const profileEmail = document.getElementById('profile-info-email');
  if (profileEmail) profileEmail.textContent = email;
  const profileMethod = document.getElementById('profile-info-method');
  if (profileMethod) {
    const userStorageVal = localStorage.getItem(`invoice_hub_user_${email.toLowerCase()}`);
    if (userStorageVal === 'google_auth') {
      profileMethod.textContent = 'Google Sign-In';
    } else {
      profileMethod.textContent = 'Email / Password';
    }
  }

  const sigDataJSON = localStorage.getItem(`invoice_hub_user_sig_${email.toLowerCase()}`);
  if (sigDataJSON) {
    try {
      const sigData = JSON.parse(sigDataJSON);
      if (sigData && sigData.imageBytesBase64) {
        // Automatically mark as onboarded if signature exists
        localStorage.setItem(`invoice_hub_user_onboarded_${email.toLowerCase()}`, 'true');
        const buffer = base64ToArrayBuffer(sigData.imageBytesBase64);
        
        // Populate global signature state
        globalSignature.imageBytes = buffer;
        globalSignature.imageFilename = sigData.imageFilename;
        globalSignature.imageType = sigData.imageType;
        
        // Also populate signState so they can see/change it
        signState.imageBytes = buffer;
        signState.imageFilename = sigData.imageFilename;
        signState.imageType = sigData.imageType;
        
        const sizeKb = (buffer.byteLength / 1024).toFixed(1);
        
        // Update UI pill in Sign Workspace to show the file name
        const pillSign = document.getElementById('pill-sign-img');
        if (pillSign) {
          pillSign.textContent = `✓ ${sigData.imageFilename} (${sizeKb} KB)`;
          pillSign.style.display = 'inline-flex';
        }

        // Update UI pill in Settings Workspace to show the file name
        const pillSettings = document.getElementById('pill-settings-sign-img');
        if (pillSettings) {
          pillSettings.textContent = `✓ ${sigData.imageFilename} (${sizeKb} KB)`;
          pillSettings.style.display = 'inline-flex';
        }
        
        // Update dropzones visibility in both workspaces
        const dropzoneSign = document.getElementById('dropzone-sign-img');
        if (dropzoneSign) dropzoneSign.style.display = 'none';
        const dropzoneSettings = document.getElementById('dropzone-settings-sign-img');
        if (dropzoneSettings) dropzoneSettings.style.display = 'none';
        
        // Update preview images in both workspaces
        const base64Str = sigData.imageBytesBase64;
        const mimeType = sigData.imageType || 'image/png';
        const dataUrl = `data:${mimeType};base64,${base64Str}`;
        const previewImgs = document.querySelectorAll('.sig-preview-img, .zoomed-image');
        previewImgs.forEach(img => {
          img.src = dataUrl;
        });

        // Enable Done / Upload buttons in both workspaces
        const btnDone = document.getElementById('btn-sign-done');
        if (btnDone) btnDone.removeAttribute('disabled');
        const btnSettingsDone = document.getElementById('btn-settings-sign-done');
        if (btnSettingsDone) btnSettingsDone.removeAttribute('disabled');
        
        // Recalculate input status for cams/kfintech workspaces
        checkInputs(camsState, 'btn-process-cams');
        checkInputs(kfintechState, 'btn-process-kfintech');
      }
    } catch (e) {
      console.error("Failed to parse user signature data:", e);
    }
  } else {
    // No signature saved for this user, clear global state and UI
    globalSignature.imageBytes = null;
    globalSignature.imageFilename = '';
    globalSignature.imageType = '';
    
    signState.imageBytes = null;
    signState.imageFilename = '';
    signState.imageType = '';
    
    const pillSign = document.getElementById('pill-sign-img');
    if (pillSign) {
      pillSign.style.display = 'none';
      pillSign.textContent = '';
    }

    const pillSettings = document.getElementById('pill-settings-sign-img');
    if (pillSettings) {
      pillSettings.style.display = 'none';
      pillSettings.textContent = '';
    }

    const dropzoneSign = document.getElementById('dropzone-sign-img');
    if (dropzoneSign) dropzoneSign.style.display = 'block';
    const dropzoneSettings = document.getElementById('dropzone-settings-sign-img');
    if (dropzoneSettings) dropzoneSettings.style.display = 'block';
    
    const previewImgs = document.querySelectorAll('.sig-preview-img, .zoomed-image');
    previewImgs.forEach(img => {
      img.src = './Sample Sign & Stamp.png';
    });

    const btnDone = document.getElementById('btn-sign-done');
    if (btnDone) btnDone.setAttribute('disabled', 'true');
    const btnSettingsDone = document.getElementById('btn-settings-sign-done');
    if (btnSettingsDone) btnSettingsDone.setAttribute('disabled', 'true');
    
    checkInputs(camsState, 'btn-process-cams');
    checkInputs(kfintechState, 'btn-process-kfintech');
  }
}

function loadUserExcelFiles(email) {
  // Clear CAMS Excel
  camsExcelState.fileBytes = null;
  camsExcelState.fileFilename = '';
  const camsPill = document.getElementById('pill-cams-excel');
  const camsDropzone = document.getElementById('dropzone-cams-excel');
  const camsDone = document.getElementById('btn-cams-excel-done');
  const camsPreviewName = document.getElementById('cams-excel-preview-name');
  const camsPreviewContainer = document.getElementById('cams-excel-preview-container');
  
  if (camsPill) {
    camsPill.style.display = 'none';
    camsPill.textContent = '';
  }
  if (camsDropzone) camsDropzone.style.display = 'block';
  if (camsDone) camsDone.setAttribute('disabled', 'true');
  if (camsPreviewName) camsPreviewName.textContent = 'CAMS_Sample.xlsx';
  if (camsPreviewContainer) camsPreviewContainer.style.display = 'none';

  // Clear Kfintech Excel
  kfintechExcelState.fileBytes = null;
  kfintechExcelState.fileFilename = '';
  const kfintechPill = document.getElementById('pill-kfintech-excel');
  const kfintechDropzone = document.getElementById('dropzone-kfintech-excel');
  const kfintechDone = document.getElementById('btn-kfintech-excel-done');
  const kfintechPreviewName = document.getElementById('kfintech-excel-preview-name');
  const kfintechPreviewContainer = document.getElementById('kfintech-excel-preview-container');

  if (kfintechPill) {
    kfintechPill.style.display = 'none';
    kfintechPill.textContent = '';
  }
  if (kfintechDropzone) kfintechDropzone.style.display = 'block';
  if (kfintechDone) kfintechDone.setAttribute('disabled', 'true');
  if (kfintechPreviewName) kfintechPreviewName.textContent = 'Kfintech_Sample.xlsx';
  if (kfintechPreviewContainer) kfintechPreviewContainer.style.display = 'none';
}

// ============================================================================
// 1. UI Navigation & Transitions
// ============================================================================

const cardsSection = document.querySelector('.dashboard-grid');

const cardCams = document.getElementById('card-cams');
const wsCams = document.getElementById('workspace-cams');
const btnCloseCams = document.getElementById('btn-close-cams-workspace');

const cardKfintech = document.getElementById('card-kfintech');
const wsKfintech = document.getElementById('workspace-kfintech');
const btnCloseKfintech = document.getElementById('btn-close-kfintech-workspace');

const cardSign = document.getElementById('card-sign');
const wsSign = document.getElementById('workspace-sign');
const btnCloseSign = document.getElementById('btn-close-sign-workspace');

function openWorkspace(ws, card) {
  // Clear all workspace active classes first to prevent overlap
  if (wsCams) wsCams.classList.remove('active');
  if (cardCams) cardCams.classList.remove('active');
  if (wsKfintech) wsKfintech.classList.remove('active');
  if (cardKfintech) cardKfintech.classList.remove('active');
  if (wsSign) wsSign.classList.remove('active');
  if (cardSign) cardSign.classList.remove('active');
  
  const wsCamsExcel = document.getElementById('workspace-cams-excel');
  const wsKfintechExcel = document.getElementById('workspace-kfintech-excel');
  if (wsCamsExcel) wsCamsExcel.classList.remove('active');
  if (wsKfintechExcel) wsKfintechExcel.classList.remove('active');
  
  const wsProfile = document.getElementById('workspace-profile');
  const wsSettings = document.getElementById('workspace-settings');
  if (wsProfile) wsProfile.classList.remove('active');
  if (wsSettings) wsSettings.classList.remove('active');

  cardsSection.style.display = 'none';
  ws.classList.add('active');
  if (card && card.classList) {
    card.classList.add('active');
  }
}

function navigateSubpage(view, hash) {
  const currentView = history.state && history.state.view;
  if (currentView && currentView !== 'dashboard') {
    history.replaceState({ view: view }, '', hash);
  } else {
    history.pushState({ view: view }, '', hash);
  }
}

function closeWorkspaces() {
  wsCams.classList.remove('active');
  cardCams.classList.remove('active');
  wsKfintech.classList.remove('active');
  cardKfintech.classList.remove('active');
  wsSign.classList.remove('active');
  cardSign.classList.remove('active');
  
  const wsCamsExcel = document.getElementById('workspace-cams-excel');
  const wsKfintechExcel = document.getElementById('workspace-kfintech-excel');
  if (wsCamsExcel) wsCamsExcel.classList.remove('active');
  if (wsKfintechExcel) wsKfintechExcel.classList.remove('active');
  
  const wsProfile = document.getElementById('workspace-profile');
  const wsSettings = document.getElementById('workspace-settings');
  if (wsProfile) wsProfile.classList.remove('active');
  if (wsSettings) wsSettings.classList.remove('active');
  
  // Enforce correct visibility of cardSign when dashboard is shown
  const sessionEmail = localStorage.getItem('invoice_hub_session');
  if (sessionEmail) {
    const isOnboarded = (localStorage.getItem(`invoice_hub_user_onboarded_${sessionEmail.toLowerCase()}`) === 'true') ||
                        (localStorage.getItem(`invoice_hub_user_sig_${sessionEmail.toLowerCase()}`) !== null);
    cardSign.style.display = isOnboarded ? 'none' : 'flex';
  } else {
    cardSign.style.display = 'flex';
  }
  
  updateVisitsDisplay();
  cardsSection.style.display = 'grid';
}

// CAMS Event Listeners for Workspace transitions
cardCams.addEventListener('click', () => {
  openWorkspace(wsCams, cardCams);
  
  const wsCamsExcel = document.getElementById('workspace-cams-excel');
  if (wsCamsExcel) wsCamsExcel.classList.add('active');
  
  // CAMS Excel Setup UI
  const hasCamsExcel = !!camsExcelState.fileBytes;
  const camsDrop = document.getElementById('dropzone-cams-excel');
  if (camsDrop) camsDrop.style.display = 'block';
  const pillCams = document.getElementById('pill-cams-excel');
  if (pillCams) pillCams.style.display = hasCamsExcel ? 'inline-flex' : 'none';
  const btnCamsDone = document.getElementById('btn-cams-excel-done');
  if (btnCamsDone) {
    if (hasCamsExcel) btnCamsDone.removeAttribute('disabled');
    else btnCamsDone.setAttribute('disabled', 'true');
  }
  const camsPreviewContainer = document.getElementById('cams-excel-preview-container');
  if (camsPreviewContainer) camsPreviewContainer.style.display = hasCamsExcel ? 'block' : 'none';
  
  if (!history.state || history.state.view !== 'cams') {
    navigateSubpage('cams', '#cams');
  }
});
btnCloseCams.addEventListener('click', (e) => {
  e.stopPropagation();
  goBackToDashboard();
});

// Keyboard Accessibility for CAMS
cardCams.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' || e.key === ' ') {
    e.preventDefault();
    cardCams.click();
  }
});

// Kfintech Event Listeners for Workspace transitions
cardKfintech.addEventListener('click', () => {
  openWorkspace(wsKfintech, cardKfintech);
  
  const wsKfintechExcel = document.getElementById('workspace-kfintech-excel');
  if (wsKfintechExcel) wsKfintechExcel.classList.add('active');
  
  // Kfintech Excel Setup UI
  const hasKfintechExcel = !!kfintechExcelState.fileBytes;
  const kfintechDrop = document.getElementById('dropzone-kfintech-excel');
  if (kfintechDrop) kfintechDrop.style.display = 'block';
  const pillKfintech = document.getElementById('pill-kfintech-excel');
  if (pillKfintech) pillKfintech.style.display = hasKfintechExcel ? 'inline-flex' : 'none';
  const btnKfintechDone = document.getElementById('btn-kfintech-excel-done');
  if (btnKfintechDone) {
    if (hasKfintechExcel) btnKfintechDone.removeAttribute('disabled');
    else btnKfintechDone.setAttribute('disabled', 'true');
  }
  const kfintechPreviewContainer = document.getElementById('kfintech-excel-preview-container');
  if (kfintechPreviewContainer) kfintechPreviewContainer.style.display = hasKfintechExcel ? 'block' : 'none';
  
  if (!history.state || history.state.view !== 'kfintech') {
    navigateSubpage('kfintech', '#kfintech');
  }
});
btnCloseKfintech.addEventListener('click', (e) => {
  e.stopPropagation();
  goBackToDashboard();
});

// Keyboard Accessibility for Kfintech
cardKfintech.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' || e.key === ' ') {
    e.preventDefault();
    cardKfintech.click();
  }
});

// Sign Event Listeners for Workspace transitions
cardSign.addEventListener('click', () => {
  openWorkspace(wsSign, cardSign);
  
  // If there's already an active signature, hide the dropzone and show the pill
  const hasSig = !!globalSignature.imageBytes;
  document.getElementById('dropzone-sign-img').style.display = hasSig ? 'none' : 'block';
  
  const pillSign = document.getElementById('pill-sign-img');
  if (pillSign) {
    pillSign.style.display = hasSig ? 'inline-flex' : 'none';
  }
  
  const previewImgs = document.querySelectorAll('.sig-preview-img, .zoomed-image');
  if (hasSig) {
    const base64Str = arrayBufferToBase64(globalSignature.imageBytes);
    const dataUrl = `data:${globalSignature.imageType || 'image/png'};base64,${base64Str}`;
    previewImgs.forEach(img => {
      img.src = dataUrl;
    });
  } else {
    previewImgs.forEach(img => {
      img.src = './Sample Sign & Stamp.png';
    });
  }

  const btnDone = document.getElementById('btn-sign-done');
  if (btnDone) {
    if (hasSig) {
      btnDone.removeAttribute('disabled');
    } else {
      btnDone.setAttribute('disabled', 'true');
    }
  }
  
  if (!history.state || history.state.view !== 'sign') {
    navigateSubpage('sign', '#sign');
  }
});
if (btnCloseSign) {
  btnCloseSign.addEventListener('click', (e) => {
    e.stopPropagation();
    goBackToDashboard();
  });
}

const btnCloseCamsExcel = document.getElementById('btn-close-cams-excel-workspace');
if (btnCloseCamsExcel) {
  btnCloseCamsExcel.addEventListener('click', (e) => {
    e.stopPropagation();
    goBackToDashboard();
  });
}

const btnCloseKfintechExcel = document.getElementById('btn-close-kfintech-excel-workspace');
if (btnCloseKfintechExcel) {
  btnCloseKfintechExcel.addEventListener('click', (e) => {
    e.stopPropagation();
    goBackToDashboard();
  });
}

// Keyboard Accessibility for Sign
cardSign.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' || e.key === ' ') {
    e.preventDefault();
    cardSign.click();
  }
});

// ============================================================================
// 2. Drag and Drop File Handlers
// ============================================================================

function setupDropzone(dropzoneId, inputId, pillId, fileType, callback) {
  const dropzone = document.getElementById(dropzoneId);
  const input = document.getElementById(inputId);
  const pill = document.getElementById(pillId);

  if (!dropzone || !input || !pill) return;

  // Trigger input click when clicking anywhere inside the dropzone
  dropzone.addEventListener('click', (e) => {
    if (e.target !== input) {
      input.click();
    }
  });

  // Drag over states
  dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.classList.add('dragover');
  });

  dropzone.addEventListener('dragleave', () => {
    dropzone.classList.remove('dragover');
  });

  dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.classList.remove('dragover');
    if (e.dataTransfer.files.length > 0) {
      input.files = e.dataTransfer.files;
      handleFile(input.files[0]);
    }
  });

  input.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
      handleFile(e.target.files[0]);
    }
  });

  input.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      let isFileLoaded = false;
      if (inputId === 'input-cams-pdf' && camsState.pdfBytes) isFileLoaded = true;
      if (inputId === 'input-kfintech-pdf' && kfintechState.pdfBytes) isFileLoaded = true;
      if (inputId === 'input-sign-img' && signState.imageBytes) isFileLoaded = true;
      if (inputId === 'input-cams-excel' && camsExcelState.fileBytes) isFileLoaded = true;
      if (inputId === 'input-kfintech-excel' && kfintechExcelState.fileBytes) isFileLoaded = true;

      if (isFileLoaded) {
        e.preventDefault(); // Prevent opening the file upload dialog
        
        // Trigger the corresponding action button
        if (inputId === 'input-cams-pdf') {
          const btn = document.getElementById('btn-process-cams');
          if (btn && !btn.disabled) btn.click();
        } else if (inputId === 'input-kfintech-pdf') {
          const btn = document.getElementById('btn-process-kfintech');
          if (btn && !btn.disabled) btn.click();
        } else if (inputId === 'input-sign-img') {
          const btn = document.getElementById('btn-sign-done');
          if (btn && !btn.disabled) btn.click();
        } else if (inputId === 'input-cams-excel') {
          const btn = document.getElementById('btn-cams-excel-done');
          if (btn && !btn.disabled) btn.click();
        } else if (inputId === 'input-kfintech-excel') {
          const btn = document.getElementById('btn-kfintech-excel-done');
          if (btn && !btn.disabled) btn.click();
        }
      }
    }
  });

  function handleFile(file) {
    // Validate file type
    if (fileType === 'pdf') {
      const nameLower = file.name.toLowerCase();
      if (!nameLower.endsWith('.pdf') && !nameLower.endsWith('.zip')) {
        alert('Please upload a PDF or ZIP file.');
        return;
      }
    }
    if (fileType === 'image' && !file.type.match('image/(png|jpeg)')) {
      alert('Please upload a PNG or JPEG image.');
      return;
    }
    if (fileType === 'excel') {
      const nameLower = file.name.toLowerCase();
      if (!nameLower.endsWith('.xlsx') && !nameLower.endsWith('.xls')) {
        alert('Please upload an Excel file (.xlsx or .xls).');
        return;
      }
    }

    const reader = new FileReader();
    reader.onload = function(evt) {
      callback(evt.target.result, file.name, file.type);
      
      // Update UI Pill
      pill.textContent = `✓ ${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
      pill.style.display = 'inline-flex';
      alert('Successfully uploaded.');
    };
    reader.readAsArrayBuffer(file);
  }
}

// Setup dropzones for CAMS workspace
setupDropzone('dropzone-cams-pdf', 'input-cams-pdf', 'pill-cams-pdf', 'pdf', (bytes, filename) => {
  camsState.pdfBytes = bytes;
  camsState.pdfFilename = filename;
  camsState.isZipUpload = filename.toLowerCase().endsWith('.zip');
  checkInputs(camsState, 'btn-process-cams');
});

// Setup dropzones for Kfintech workspace
setupDropzone('dropzone-kfintech-pdf', 'input-kfintech-pdf', 'pill-kfintech-pdf', 'pdf', (bytes, filename) => {
  kfintechState.pdfBytes = bytes;
  kfintechState.pdfFilename = filename;
  kfintechState.isZipUpload = filename.toLowerCase().endsWith('.zip');
  checkInputs(kfintechState, 'btn-process-kfintech');
});

// Setup dropzones for CAMS Excel workspace
setupDropzone('dropzone-cams-excel', 'input-cams-excel', 'pill-cams-excel', 'excel', (bytes, filename) => {
  camsExcelState.fileBytes = bytes;
  camsExcelState.fileFilename = filename;
  
  // Update corner preview name dynamically
  const camsPreviewName = document.getElementById('cams-excel-preview-name');
  if (camsPreviewName) camsPreviewName.textContent = filename;
  
  const camsPreviewContainer = document.getElementById('cams-excel-preview-container');
  if (camsPreviewContainer) camsPreviewContainer.style.display = 'block';

  // Enable Done button
  const btnDone = document.getElementById('btn-cams-excel-done');
  if (btnDone) btnDone.removeAttribute('disabled');
});

// Setup dropzones for Kfintech Excel workspace
setupDropzone('dropzone-kfintech-excel', 'input-kfintech-excel', 'pill-kfintech-excel', 'excel', (bytes, filename) => {
  kfintechExcelState.fileBytes = bytes;
  kfintechExcelState.fileFilename = filename;

  // Update corner preview name dynamically
  const kfintechPreviewName = document.getElementById('kfintech-excel-preview-name');
  if (kfintechPreviewName) kfintechPreviewName.textContent = filename;

  const kfintechPreviewContainer = document.getElementById('kfintech-excel-preview-container');
  if (kfintechPreviewContainer) kfintechPreviewContainer.style.display = 'block';

  // Enable Done button
  const btnDone = document.getElementById('btn-kfintech-excel-done');
  if (btnDone) btnDone.removeAttribute('disabled');
});

// Setup dropzones for Sign workspace (Only image dropzone is needed)
setupDropzone('dropzone-sign-img', 'input-sign-img', 'pill-sign-img', 'image', (bytes, filename, type) => {
  signState.imageBytes = bytes;
  signState.imageFilename = filename;
  signState.imageType = type;
  
  // Enable Done / Upload buttons in both
  const btnDone = document.getElementById('btn-sign-done');
  if (btnDone) btnDone.removeAttribute('disabled');
  const btnSettingsDone = document.getElementById('btn-settings-sign-done');
  if (btnSettingsDone) btnSettingsDone.removeAttribute('disabled');

  // Sync Settings pill
  const pillSettings = document.getElementById('pill-settings-sign-img');
  if (pillSettings) {
    pillSettings.textContent = `✓ ${filename} (${(bytes.byteLength / 1024).toFixed(1)} KB)`;
    pillSettings.style.display = 'inline-flex';
  }
});

// Setup dropzones for Settings Change Sign workspace
setupDropzone('dropzone-settings-sign-img', 'input-settings-sign-img', 'pill-settings-sign-img', 'image', (bytes, filename, type) => {
  signState.imageBytes = bytes;
  signState.imageFilename = filename;
  signState.imageType = type;
  
  // Enable Done / Upload buttons in both
  const btnDone = document.getElementById('btn-sign-done');
  if (btnDone) btnDone.removeAttribute('disabled');
  const btnSettingsDone = document.getElementById('btn-settings-sign-done');
  if (btnSettingsDone) btnSettingsDone.removeAttribute('disabled');

  // Sync Main pill
  const pillMain = document.getElementById('pill-sign-img');
  if (pillMain) {
    pillMain.textContent = `✓ ${filename} (${(bytes.byteLength / 1024).toFixed(1)} KB)`;
    pillMain.style.display = 'inline-flex';
  }
});

function checkInputs(targetState, buttonId) {
  const btnProcess = document.getElementById(buttonId);
  if (!btnProcess) return;
  if (targetState.pdfBytes && globalSignature.imageBytes) {
    btnProcess.removeAttribute('disabled');
  } else {
    btnProcess.setAttribute('disabled', 'true');
  }
}

// ============================================================================
// 3. Signature Placement Logic
// ============================================================================

// Helper function to sign a single PDF's ArrayBuffer and return the signed Uint8Array
async function signPdfDocument(pdfBytes, sigImageBytes, sigImageType, scaleFactor, xOffset, yOffset) {
  // 1. Scan PDF text to find coordinates of signature block elements using PDF.js
  const pdf = await pdfjsLib.getDocument({ data: pdfBytes.slice(0) }).promise;
  const numPages = pdf.numPages;
  const rawMatches = [];

  // First do a fast document-wide scan to detect conditions
  let fullDocumentText = '';
  for (let pageNum = 1; pageNum <= numPages; pageNum++) {
    const page = await pdf.getPage(pageNum);
    const textContent = await page.getTextContent();
    const pageText = textContent.items.map(item => item.str || '').join(' ');
    fullDocumentText += pageText + '\n';
  }

  const hasDesignationStatus = fullDocumentText.toLowerCase().includes('designation / status') || fullDocumentText.toLowerCase().includes('designation/status');
  const hasForSaparia = fullDocumentText.toLowerCase().includes('saparia global private limited') || fullDocumentText.toLowerCase().includes('saparia global');

  for (let pageNum = 1; pageNum <= numPages; pageNum++) {
    const page = await pdf.getPage(pageNum);
    const textContent = await page.getTextContent();
    
    textContent.items.forEach(item => {
      if (!item.str) return;
      const strLower = item.str.toLowerCase();
      let isMatch = false;

      const isSignatureKeyword = strLower.includes('signature');
      const isSignatoryKeyword = strLower.includes('signatory');
      const isAuthSignKeyword = strLower.includes('auth. sign') || strLower.includes('authorised signatory') || strLower.includes('authorized signatory');

      // Apply rules:
      if (hasDesignationStatus) {
        // Rule 1: designation status found -> only paste above "Signature", not "Name of the Signatory"
        if (isSignatureKeyword && !strLower.includes('name of the signatory') && !strLower.includes('name of signatory')) {
          isMatch = true;
        }
      } else {
        // Standard case: signature or signatory
        if (isSignatureKeyword || isSignatoryKeyword) {
          isMatch = true;
        }
      }

      // Rule 2: For SAPARIA GLOBAL PRIVATE LIMITED found -> paste above "Authorised Signatory" / "Auth. Sign."
      if (hasForSaparia) {
        if (isAuthSignKeyword) {
          isMatch = true;
        }
      }

      if (isMatch) {
        rawMatches.push({
          pageNum,
          x: item.transform[4],
          y: item.transform[5],
          width: item.width,
          height: item.height,
          text: item.str
        });
      }
    });
  }

  // Filter duplicates close on both axes, keeping the highest Y coordinate (upper text)
  const matches = [];
  rawMatches.forEach(match => {
    const duplicateIndex = matches.findIndex(m => 
      m.pageNum === match.pageNum &&
      Math.abs(m.x - match.x) < 50 &&
      Math.abs(m.y - match.y) < 100
    );

    if (duplicateIndex === -1) {
      matches.push(match);
    } else {
      if (match.y > matches[duplicateIndex].y) {
        matches[duplicateIndex] = match;
      }
    }
  });

  if (matches.length === 0) {
    // Default to first page, bottom right fallback
    matches.push({ pageNum: 1, x: 450, y: 80, width: 80, height: 15, text: 'Default Position' });
  }

  // 2. Load PDF into pdf-lib and paste the signature
  const pdfDoc = await PDFLib.PDFDocument.load(pdfBytes);
  const pages = pdfDoc.getPages();
  
  let embeddedImg;
  if (sigImageType === 'image/png') {
    embeddedImg = await pdfDoc.embedPng(sigImageBytes);
  } else {
    embeddedImg = await pdfDoc.embedJpg(sigImageBytes);
  }

  // Scale parameter from slider
  const baseMaxWidth = 120 * scaleFactor;
  const baseMaxHeight = 40 * scaleFactor;
  let sigWidth = embeddedImg.width;
  let sigHeight = embeddedImg.height;
  const aspect = sigWidth / sigHeight;

  if (sigWidth > baseMaxWidth) {
    sigWidth = baseMaxWidth;
    sigHeight = sigWidth / aspect;
  }
  if (sigHeight > baseMaxHeight) {
    sigHeight = baseMaxHeight;
    sigWidth = sigHeight * aspect;
  }

  // Embed image in coordinates
  matches.forEach(match => {
    const pageIndex = match.pageNum - 1;
    const page = pages[pageIndex];
    
    // Center horizontally over the word
    const posX = match.x + (match.width - sigWidth) / 2 + xOffset;
    // Place vertically above the baseline of the text
    const posY = match.y + match.height + yOffset;

    page.drawImage(embeddedImg, {
      x: posX,
      y: posY,
      width: sigWidth,
      height: sigHeight
    });
  });

  return await pdfDoc.save();
}

// Set up processing triggers for workspaces
function setupWorkspaceProcessor(prefix, targetState) {
  document.getElementById(`btn-process-${prefix}`).addEventListener('click', async () => {
    const feedback = document.getElementById(`feedback-${prefix}`);
    const progressBar = document.getElementById(`progress-${prefix}-fill`);
    const downloadPanel = document.getElementById(`download-${prefix}-panel`);
    const reportText = document.getElementById(`txt-${prefix}-report`);
    const btnDownload = document.getElementById(`btn-download-${prefix}-pdf`);
    
    feedback.classList.add('active');
    downloadPanel.style.display = 'none';
    progressBar.style.width = '0%';

    const scaleFactor = 1.0;
    const xOffset = 0;
    const yOffset = 2;

    const pdfMetaList = []; // Collect metadata for Excel matching

    try {
      if (targetState.isZipUpload) {
        if (prefix === 'kfintech') {
          // Kfintech Custom ZIP Flow: extract only exclusiveGST PDFs from sub-ZIPs and package flat
          progressBar.style.width = '10%';
          const mainZip = await JSZip.loadAsync(targetState.pdfBytes);
          progressBar.style.width = '20%';

          const zipFiles = [];
          mainZip.forEach((relativePath, fileEntry) => {
            if (relativePath.toLowerCase().endsWith('.zip')) {
              zipFiles.push(fileEntry);
            }
          });

          if (zipFiles.length === 0) {
            alert('No sub-ZIP folders found inside the uploaded ZIP folder.');
            feedback.classList.remove('active');
            return;
          }

          const outputZip = new JSZip();
          let processedCount = 0;
          let matchedPdfCount = 0;
          const totalFiles = zipFiles.length;

          for (const subZipEntry of zipFiles) {
            const subZipData = await subZipEntry.async('arraybuffer');
            const subZip = await JSZip.loadAsync(subZipData);
            
            // Find PDF files inside the sub-ZIP containing "exclusiveGST" in their name (case-insensitive)
            const pdfEntries = [];
            subZip.forEach((subPath, subEntry) => {
              const nameLower = subPath.toLowerCase();
              if (nameLower.endsWith('.pdf') && nameLower.includes('exclusivegst')) {
                pdfEntries.push(subEntry);
              }
            });

            // Sign each matching PDF and add to flat output ZIP
            for (const pdfEntry of pdfEntries) {
              const pdfBytes = await pdfEntry.async('arraybuffer');
              const signedPdfBytes = await signPdfDocument(pdfBytes, globalSignature.imageBytes, globalSignature.imageType, scaleFactor, xOffset, yOffset);
              
              // Prevent collisions in output ZIP
              let targetName = pdfEntry.name;
              if (outputZip.file(targetName)) {
                const subZipNameOnly = subZipEntry.name.replace(/\.zip$/i, '');
                targetName = `${subZipNameOnly}_${pdfEntry.name}`;
              }
              outputZip.file(targetName, signedPdfBytes);
              matchedPdfCount++;

              // Extract metadata for Excel report updating
              const meta = await extractPdfMetadata(pdfBytes);
              pdfMetaList.push({
                filename: targetName,
                gstins: meta.gstins,
                invoiceNo: meta.invoiceNo,
                invoiceDate: meta.invoiceDate,
                fullText: meta.fullText
              });
            }

            processedCount++;
            progressBar.style.width = `${20 + Math.floor((processedCount / totalFiles) * 70)}%`;
          }

          // Generate flat output ZIP
          targetState.processedPdfBytes = await outputZip.generateAsync({ type: 'uint8array' });
          
          progressBar.style.width = '100%';
          setTimeout(() => {
            // Process and update Excel report if template uploaded
            const excelTemplateBytes = kfintechExcelState.fileBytes;
            const btnDownloadExcel = document.getElementById('btn-download-kfintech-excel');
            if (excelTemplateBytes && pdfMetaList.length > 0) {
              targetState.processedExcelBytes = updateExcelReport('kfintech', excelTemplateBytes, pdfMetaList);
              if (btnDownloadExcel) btnDownloadExcel.style.display = 'inline-flex';
            } else {
              targetState.processedExcelBytes = null;
              if (btnDownloadExcel) btnDownloadExcel.style.display = 'none';
            }

            feedback.classList.remove('active');
            reportText.textContent = `Processed Kfintech ZIP folder. Extracted, signed, and packaged ${matchedPdfCount} matching "exclusiveGST" PDFs.`;
            btnDownload.querySelector('span').textContent = 'Download Signed ZIP';
            downloadPanel.style.display = 'flex';
          }, 500);
        } else {
          // ZIP Flow: Nested ZIPs extraction and re-packaging (CAMS)
          progressBar.style.width = '10%';
          const mainZip = await JSZip.loadAsync(targetState.pdfBytes);
          progressBar.style.width = '20%';

          const zipFiles = [];
          mainZip.forEach((relativePath, fileEntry) => {
            if (relativePath.toLowerCase().endsWith('.zip')) {
              zipFiles.push(fileEntry);
            }
          });

          if (zipFiles.length === 0) {
            alert('No sub-ZIP folders found inside the uploaded ZIP folder.');
            feedback.classList.remove('active');
            return;
          }

          let processedCount = 0;
          const totalFiles = zipFiles.length;

          for (const subZipEntry of zipFiles) {
            const subZipData = await subZipEntry.async('arraybuffer');
            const subZip = await JSZip.loadAsync(subZipData);
            
            // Find PDF files inside the sub-ZIP
            const pdfEntries = [];
            subZip.forEach((subPath, subEntry) => {
              if (subPath.toLowerCase().endsWith('.pdf')) {
                pdfEntries.push(subEntry);
              }
            });

            // Sign each PDF inside the sub-ZIP
            for (const pdfEntry of pdfEntries) {
              const pdfBytes = await pdfEntry.async('arraybuffer');
              const signedPdfBytes = await signPdfDocument(pdfBytes, globalSignature.imageBytes, globalSignature.imageType, scaleFactor, xOffset, yOffset);
              subZip.file(pdfEntry.name, signedPdfBytes);

              // Extract metadata for Excel report updating
              const meta = await extractPdfMetadata(pdfBytes);
              pdfMetaList.push({
                filename: pdfEntry.name,
                gstins: meta.gstins,
                invoiceNo: meta.invoiceNo,
                invoiceDate: meta.invoiceDate,
                fullText: meta.fullText
              });
            }

            // Recompile sub-ZIP ArrayBuffer
            const updatedSubZipBytes = await subZip.generateAsync({ type: 'arraybuffer' });
            mainZip.file(subZipEntry.name, updatedSubZipBytes);

            processedCount++;
            progressBar.style.width = `${20 + Math.floor((processedCount / totalFiles) * 70)}%`;
          }

          // Generate updated main ZIP bytes
          targetState.processedPdfBytes = await mainZip.generateAsync({ type: 'uint8array' });
          
          progressBar.style.width = '100%';
          setTimeout(() => {
            // Process and update Excel report if template uploaded
            const excelTemplateBytes = camsExcelState.fileBytes;
            const btnDownloadExcel = document.getElementById('btn-download-cams-excel');
            if (excelTemplateBytes && pdfMetaList.length > 0) {
              targetState.processedExcelBytes = updateExcelReport('cams', excelTemplateBytes, pdfMetaList);
              if (btnDownloadExcel) btnDownloadExcel.style.display = 'inline-flex';
            } else {
              targetState.processedExcelBytes = null;
              if (btnDownloadExcel) btnDownloadExcel.style.display = 'none';
            }

            feedback.classList.remove('active');
            reportText.textContent = `Processed parent ZIP folder. Extracted and signed PDFs inside all ${totalFiles} sub-ZIP folders.`;
            btnDownload.querySelector('span').textContent = 'Download Signed ZIP';
            downloadPanel.style.display = 'flex';
          }, 500);
        }

      } else {
        // Single PDF Flow
        progressBar.style.width = '20%';
        targetState.processedPdfBytes = await signPdfDocument(targetState.pdfBytes, globalSignature.imageBytes, globalSignature.imageType, scaleFactor, xOffset, yOffset);
        progressBar.style.width = '100%';
        
        // Extract metadata for Excel report updating
        const meta = await extractPdfMetadata(targetState.pdfBytes);
        pdfMetaList.push({
          filename: `signed_${targetState.pdfFilename}`,
          gstins: meta.gstins,
          invoiceNo: meta.invoiceNo,
          invoiceDate: meta.invoiceDate,
          fullText: meta.fullText
        });

        setTimeout(() => {
          // Process and update Excel report if template uploaded
          const excelTemplateBytes = prefix === 'cams' ? camsExcelState.fileBytes : kfintechExcelState.fileBytes;
          const btnDownloadExcel = document.getElementById(`btn-download-${prefix}-excel`);
          if (excelTemplateBytes && pdfMetaList.length > 0) {
            targetState.processedExcelBytes = updateExcelReport(prefix, excelTemplateBytes, pdfMetaList);
            if (btnDownloadExcel) btnDownloadExcel.style.display = 'inline-flex';
          } else {
            targetState.processedExcelBytes = null;
            if (btnDownloadExcel) btnDownloadExcel.style.display = 'none';
          }

          feedback.classList.remove('active');
          reportText.textContent = `Processed PDF. Signature embedded successfully above detected locations.`;
          btnDownload.querySelector('span').textContent = 'Download Signed PDF';
          downloadPanel.style.display = 'flex';
        }, 500);
      }

    } catch (error) {
      console.error(error);
      alert('An error occurred during processing: ' + error.message);
      feedback.classList.remove('active');
    }
  });

  // Trigger download for Signed file (PDF or ZIP depending on upload format)
  document.getElementById(`btn-download-${prefix}-pdf`).addEventListener('click', () => {
    if (!targetState.processedPdfBytes) return;
    const mimeType = targetState.isZipUpload ? 'application/zip' : 'application/pdf';
    
    let downloadName = 'signed_documents.zip';
    if (targetState.pdfFilename) {
      if (targetState.isZipUpload) {
        downloadName = `signed_${targetState.pdfFilename.replace(/\.zip$/i, '')}.zip`;
      } else {
        downloadName = `signed_${targetState.pdfFilename.replace(/\.pdf$/i, '')}.pdf`;
      }
    }

    const blob = new Blob([targetState.processedPdfBytes], { type: mimeType });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = downloadName;
    link.click();
  });
}

setupWorkspaceProcessor('cams', camsState);
setupWorkspaceProcessor('kfintech', kfintechState);

// Trigger download for updated Excel files
const btnDownloadCamsExcel = document.getElementById('btn-download-cams-excel');
if (btnDownloadCamsExcel) {
  btnDownloadCamsExcel.addEventListener('click', () => {
    if (!camsState.processedExcelBytes) return;
    const blob = new Blob([camsState.processedExcelBytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    const originalName = camsExcelState.fileFilename || 'CAMS_Report.xlsx';
    link.download = `updated_${originalName}`;
    link.click();
  });
}

const btnDownloadKfintechExcel = document.getElementById('btn-download-kfintech-excel');
if (btnDownloadKfintechExcel) {
  btnDownloadKfintechExcel.addEventListener('click', () => {
    if (!kfintechState.processedExcelBytes) return;
    const blob = new Blob([kfintechState.processedExcelBytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    const originalName = kfintechExcelState.fileFilename || 'Kfintech_Report.xlsx';
    link.download = `updated_${originalName}`;
    link.click();
  });
}

// Bind actions for Sign workspace buttons
function saveSignature() {
  if (!signState.imageBytes) return;

  // Propagate to global signature state
  globalSignature.imageBytes = signState.imageBytes;
  globalSignature.imageFilename = signState.imageFilename;
  globalSignature.imageType = signState.imageType;

  // Save the signature to localStorage for the currently logged in user
  const sessionEmail = localStorage.getItem('invoice_hub_session');
  if (sessionEmail) {
    const base64Str = arrayBufferToBase64(signState.imageBytes);
    const sigData = {
      imageBytesBase64: base64Str,
      imageFilename: signState.imageFilename,
      imageType: signState.imageType
    };
    localStorage.setItem(`invoice_hub_user_sig_${sessionEmail.toLowerCase()}`, JSON.stringify(sigData));
  }

  // Enable process buttons in CAMS/Kfintech if they already have PDFs uploaded
  checkInputs(camsState, 'btn-process-cams');
  checkInputs(kfintechState, 'btn-process-kfintech');

  // Hide dropzones in both workspaces
  const dzSign = document.getElementById('dropzone-sign-img');
  if (dzSign) dzSign.style.display = 'none';
  const dzSettings = document.getElementById('dropzone-settings-sign-img');
  if (dzSettings) dzSettings.style.display = 'none';

  const sizeKb = (signState.imageBytes.byteLength / 1024).toFixed(1);

  // Show pills in both workspaces
  const pillSign = document.getElementById('pill-sign-img');
  if (pillSign) {
    pillSign.textContent = `✓ ${signState.imageFilename} (${sizeKb} KB)`;
    pillSign.style.display = 'inline-flex';
  }
  const pillSettings = document.getElementById('pill-settings-sign-img');
  if (pillSettings) {
    pillSettings.textContent = `✓ ${signState.imageFilename} (${sizeKb} KB)`;
    pillSettings.style.display = 'inline-flex';
  }

  // Sync preview image source in both workspaces
  const base64Str = arrayBufferToBase64(signState.imageBytes);
  const dataUrl = `data:${signState.imageType || 'image/png'};base64,${base64Str}`;
  document.querySelectorAll('.sig-preview-img, .zoomed-image').forEach(img => {
    img.src = dataUrl;
  });
}

function changeSignature() {
  // Show dropzones in both workspaces
  const dzSign = document.getElementById('dropzone-sign-img');
  if (dzSign) dzSign.style.display = 'block';
  const dzSettings = document.getElementById('dropzone-settings-sign-img');
  if (dzSettings) dzSettings.style.display = 'block';

  // Hide pills in both workspaces
  const pillSign = document.getElementById('pill-sign-img');
  if (pillSign) {
    pillSign.style.display = 'none';
    pillSign.textContent = '';
  }
  const pillSettings = document.getElementById('pill-settings-sign-img');
  if (pillSettings) {
    pillSettings.style.display = 'none';
    pillSettings.textContent = '';
  }

  // Clear file inputs
  const inputSign = document.getElementById('input-sign-img');
  if (inputSign) inputSign.value = '';
  const inputSettings = document.getElementById('input-settings-sign-img');
  if (inputSettings) inputSettings.value = '';

  // Clear signState
  signState.imageBytes = null;
  signState.imageFilename = '';
  signState.imageType = '';

  // Reset preview images to default
  document.querySelectorAll('.sig-preview-img, .zoomed-image').forEach(img => {
    img.src = './Sample Sign & Stamp.png';
  });

  // Disable Upload buttons in both workspaces
  const btnDone = document.getElementById('btn-sign-done');
  if (btnDone) btnDone.setAttribute('disabled', 'true');
  const btnSettingsDone = document.getElementById('btn-settings-sign-done');
  if (btnSettingsDone) btnSettingsDone.setAttribute('disabled', 'true');
}

// Attach listeners for Main Sign workspace
document.getElementById('btn-sign-done').addEventListener('click', () => {
  saveSignature();
  
  const email = localStorage.getItem('invoice_hub_session');
  if (email) {
    localStorage.setItem(`invoice_hub_user_onboarded_${email.toLowerCase()}`, 'true');
  }
  
  cardSign.style.display = 'none';
  closeWorkspaces();
  history.replaceState({ view: 'dashboard' }, '', '#dashboard');
});
document.getElementById('btn-sign-change').addEventListener('click', (e) => {
  e.stopPropagation();
  changeSignature();
});

// Attach listeners for Settings Sign workspace
const btnSettingsDone = document.getElementById('btn-settings-sign-done');
if (btnSettingsDone) {
  btnSettingsDone.addEventListener('click', saveSignature);
}
const btnSettingsChange = document.getElementById('btn-settings-sign-change');
if (btnSettingsChange) {
  btnSettingsChange.addEventListener('click', (e) => {
    e.stopPropagation();
    changeSignature();
  });
}

// Bind actions for CAMS Excel workspace buttons
const btnCamsExcelDone = document.getElementById('btn-cams-excel-done');
if (btnCamsExcelDone) {
  btnCamsExcelDone.addEventListener('click', () => {
    if (!camsExcelState.fileBytes) return;

    // Show pill and preview in place instead of dashboard redirect
    const pillCams = document.getElementById('pill-cams-excel');
    if (pillCams) pillCams.style.display = 'inline-flex';
    const camsPreviewContainer = document.getElementById('cams-excel-preview-container');
    if (camsPreviewContainer) camsPreviewContainer.style.display = 'block';
  });
}

// Bind actions for Kfintech Excel workspace buttons
const btnKfintechExcelDone = document.getElementById('btn-kfintech-excel-done');
if (btnKfintechExcelDone) {
  btnKfintechExcelDone.addEventListener('click', () => {
    if (!kfintechExcelState.fileBytes) return;

    // Show pill and preview in place instead of dashboard redirect
    const pillKfintech = document.getElementById('pill-kfintech-excel');
    if (pillKfintech) pillKfintech.style.display = 'inline-flex';
    const kfintechPreviewContainer = document.getElementById('kfintech-excel-preview-container');
    if (kfintechPreviewContainer) kfintechPreviewContainer.style.display = 'block';
  });
}

// ============================================================================
// 5. Authentication Flow (Login & Signup)
// ============================================================================

const authContainer = document.getElementById('auth-container');
const appContainer = document.querySelector('.app-container');
const authOverlay = document.getElementById('auth-loading-overlay');
const overlayText = document.getElementById('loading-overlay-text');

// Form views
const viewLogin = document.getElementById('view-login');
const viewSignup = document.getElementById('view-signup');

// Nav buttons
const navBtnLogin = document.getElementById('nav-btn-login');
const navBtnSignup = document.getElementById('nav-btn-signup');

// Toggle forms via footer links
const linkRegister = document.getElementById('link-register');
const linkLoginFooter = document.getElementById('link-login-footer');

// User profile elements
const userDisplayEmail = document.getElementById('user-display-email');
const btnLogout = document.getElementById('btn-logout');
const linkForgetPassword = document.getElementById('link-forget-password');

// Default credentials
const DEFAULT_EMAIL = 'demo@example.com';
const DEFAULT_PASSWORD = 'password';

// Helper to show/hide loading spinner overlay
function showLoading(text, duration) {
  return new Promise((resolve) => {
    overlayText.textContent = text;
    authOverlay.style.display = 'flex';
    setTimeout(() => {
      authOverlay.style.display = 'none';
      resolve();
    }, duration);
  });
}

// State variables for signup verification flow
let generatedVerificationCode = null;
let verifiedEmail = '';

const signupStepEmail = document.getElementById('signup-step-email');
const signupStepCode = document.getElementById('signup-step-code');
const signupStepPassword = document.getElementById('signup-step-password');

const btnSignupVerifyEmail = document.getElementById('btn-signup-verify-email');
const btnSignupCodeSubmit = document.getElementById('btn-signup-code-submit');

function resetSignupSteps() {
  generatedVerificationCode = null;
  verifiedEmail = '';
  
  if (signupStepEmail) signupStepEmail.style.display = 'block';
  if (signupStepCode) signupStepCode.style.display = 'none';
  if (signupStepPassword) signupStepPassword.style.display = 'none';
  
  const emailInput = document.getElementById('signup-email');
  const codeInput = document.getElementById('signup-verification-code');
  const passwordInput = document.getElementById('signup-password');
  const confirmPasswordInput = document.getElementById('signup-confirm-password');
  
  if (emailInput) emailInput.value = '';
  if (codeInput) codeInput.value = '';
  if (passwordInput) passwordInput.value = '';
  if (confirmPasswordInput) confirmPasswordInput.value = '';
}

// Switch to Login View
function showLoginView() {
  resetSignupSteps();
  viewSignup.style.display = 'none';
  viewLogin.style.display = 'flex';
  navBtnSignup.classList.remove('active');
  navBtnSignup.classList.add('outline');
  navBtnLogin.classList.remove('outline');
  navBtnLogin.classList.add('active');
}

// Switch to Signup View
function showSignupView() {
  resetSignupSteps();
  viewLogin.style.display = 'none';
  viewSignup.style.display = 'flex';
  navBtnLogin.classList.remove('active');
  navBtnLogin.classList.add('outline');
  navBtnSignup.classList.remove('outline');
  navBtnSignup.classList.add('active');
}

// Log in session initialization
function doLoginSession(email) {
  localStorage.setItem('invoice_hub_session', email);
  if (!localStorage.getItem(`invoice_hub_user_${email.toLowerCase()}`)) {
    localStorage.setItem(`invoice_hub_user_${email.toLowerCase()}`, 'google_auth');
  }
  userDisplayEmail.textContent = email;
  authContainer.style.display = 'none';
  appContainer.style.display = 'flex';
  loadUserSignature(email);
  loadUserExcelFiles(email);
  updateVisitsDisplay();

  const isOnboarded = (localStorage.getItem(`invoice_hub_user_onboarded_${email.toLowerCase()}`) === 'true') ||
                      (localStorage.getItem(`invoice_hub_user_sig_${email.toLowerCase()}`) !== null);
  cardSign.style.display = isOnboarded ? 'none' : 'flex';

  if (isOnboarded) {
    history.replaceState({ view: 'dashboard' }, '', '#dashboard');
    closeWorkspaces();
  } else {
    history.replaceState({ view: 'dashboard' }, '', '#dashboard');
    history.pushState({ view: 'sign' }, '', '#sign');
    openWorkspace(wsSign, cardSign);
  }
}

// Check initial session state
function checkAuthSession() {
  const sessionEmail = localStorage.getItem('invoice_hub_session');
  if (sessionEmail) {
    userDisplayEmail.textContent = sessionEmail;
    authContainer.style.display = 'none';
    appContainer.style.display = 'flex';
    loadUserSignature(sessionEmail);
    loadUserExcelFiles(sessionEmail);
    updateVisitsDisplay();

    const isOnboarded = (localStorage.getItem(`invoice_hub_user_onboarded_${sessionEmail.toLowerCase()}`) === 'true') ||
                        (localStorage.getItem(`invoice_hub_user_sig_${sessionEmail.toLowerCase()}`) !== null);
    cardSign.style.display = isOnboarded ? 'none' : 'flex';

    if (isOnboarded) {
      history.replaceState({ view: 'dashboard' }, '', '#dashboard');
      closeWorkspaces();
    } else {
      history.replaceState({ view: 'dashboard' }, '', '#dashboard');
      history.pushState({ view: 'sign' }, '', '#sign');
      openWorkspace(wsSign, cardSign);
    }
  } else {
    appContainer.style.display = 'none';
    authContainer.style.display = 'flex';
    showLoginView();
  }
}

// Event Listeners for Nav Buttons
navBtnLogin.addEventListener('click', () => {
  if (!history.state || history.state.view !== 'login') {
    history.pushState({ view: 'login' }, '', '#login');
  }
  showLoginView();
});
navBtnSignup.addEventListener('click', () => {
  if (!history.state || history.state.view !== 'signup') {
    history.pushState({ view: 'signup' }, '', '#signup');
  }
  showSignupView();
});

// Event Listeners for Footer Links
linkRegister.addEventListener('click', (e) => {
  e.preventDefault();
  if (!history.state || history.state.view !== 'signup') {
    history.pushState({ view: 'signup' }, '', '#signup');
  }
  showSignupView();
});
linkLoginFooter.addEventListener('click', (e) => {
  e.preventDefault();
  if (!history.state || history.state.view !== 'login') {
    history.pushState({ view: 'login' }, '', '#login');
  }
  showLoginView();
});

// Mock Forget Password
linkForgetPassword.addEventListener('click', (e) => {
  e.preventDefault();
  alert('Mock feature: Password reset link has been sent to your email.');
});

// Google Account Chooser Mock Flow Integration
const openGoogleChooser = () => {
  const chooser = document.getElementById('google-chooser-modal');
  if (chooser) chooser.style.display = 'flex';
};

const btnGoogleLoginMock = document.getElementById('btn-google-login-mock');
if (btnGoogleLoginMock) {
  btnGoogleLoginMock.addEventListener('click', openGoogleChooser);
}

const btnGoogleSignupMock = document.getElementById('btn-google-signup-mock');
if (btnGoogleSignupMock) {
  btnGoogleSignupMock.addEventListener('click', openGoogleChooser);
}

// Modal elements
const googleChooserModal = document.getElementById('google-chooser-modal');
const googleCustomModal = document.getElementById('google-custom-modal');
const btnCancelGoogleChooser = document.getElementById('btn-cancel-google-chooser');
const btnAddGoogleAccount = document.getElementById('btn-add-google-account');
const btnBackGoogleChooser = document.getElementById('btn-back-google-chooser');
const formGoogleCustom = document.getElementById('form-google-custom');
const googleCustomEmailInput = document.getElementById('google-custom-email');

// Close Google chooser
if (btnCancelGoogleChooser) {
  btnCancelGoogleChooser.addEventListener('click', () => {
    googleChooserModal.style.display = 'none';
  });
}

// Show custom input modal
if (btnAddGoogleAccount) {
  btnAddGoogleAccount.addEventListener('click', () => {
    googleChooserModal.style.display = 'none';
    googleCustomModal.style.display = 'flex';
  });
}

// Go back from custom input to chooser modal
if (btnBackGoogleChooser) {
  btnBackGoogleChooser.addEventListener('click', () => {
    googleCustomModal.style.display = 'none';
    googleChooserModal.style.display = 'flex';
  });
}

// Handle account item clicks in chooser list
const googleAccountItems = document.querySelectorAll('.google-account-item[data-email]');
googleAccountItems.forEach(item => {
  item.addEventListener('click', async () => {
    const email = item.getAttribute('data-email');
    googleChooserModal.style.display = 'none';
    await showLoading('Signing in with Google...', 1000);
    doLoginSession(email);
  });
});

// Handle custom email submit
if (formGoogleCustom) {
  formGoogleCustom.addEventListener('submit', async (e) => {
    e.preventDefault();
    const email = googleCustomEmailInput.value.trim();
    if (!email) return;

    googleCustomModal.style.display = 'none';
    await showLoading('Signing in with Google...', 1200);
    doLoginSession(email);
    googleCustomEmailInput.value = '';
  });
}

// Email Login Form Submission
document.getElementById('form-login').addEventListener('submit', async (e) => {
  e.preventDefault();
  const emailInput = document.getElementById('login-email');
  const passwordInput = document.getElementById('login-password');
  const email = emailInput.value.trim();
  const password = passwordInput.value;

  if (!email || !password) {
    alert('Please enter your email and password.');
    return;
  }

  // Check against registered users in localStorage or default fallback
  const registeredPassword = localStorage.getItem(`invoice_hub_user_${email.toLowerCase()}`);
  const isValidDefault = (email.toLowerCase() === DEFAULT_EMAIL && password === DEFAULT_PASSWORD);
  const isValidRegistered = (registeredPassword !== null && password === registeredPassword);
  
  const isRegisteredEmail = (email.toLowerCase() === DEFAULT_EMAIL || registeredPassword !== null);

  if (isValidDefault || isValidRegistered) {
    await showLoading('Authenticating...', 1000);
    doLoginSession(email);
    // Clear forms
    emailInput.value = '';
    passwordInput.value = '';
  } else {
    if (isRegisteredEmail) {
      alert('Invalid email or password. Enter the right email or password.');
    } else {
      alert('Invalid email. Please signup.');
    }
  }
});



// Logout Event Listener
btnLogout.addEventListener('click', async () => {
  await showLoading('Logging out...', 600);
  localStorage.removeItem('invoice_hub_session');
  
  // Clear signature states on logout
  globalSignature.imageBytes = null;
  globalSignature.imageFilename = '';
  globalSignature.imageType = '';
  signState.imageBytes = null;
  signState.imageFilename = '';
  signState.imageType = '';

  // Clear Excel states on logout
  camsExcelState.fileBytes = null;
  camsExcelState.fileFilename = '';
  kfintechExcelState.fileBytes = null;
  kfintechExcelState.fileFilename = '';

  // Clear Excel UI pills and reset previews
  const camsPill = document.getElementById('pill-cams-excel');
  if (camsPill) camsPill.style.display = 'none';
  const kfintechPill = document.getElementById('pill-kfintech-excel');
  if (kfintechPill) kfintechPill.style.display = 'none';
  const camsPreviewName = document.getElementById('cams-excel-preview-name');
  if (camsPreviewName) camsPreviewName.textContent = 'CAMS_Sample.xlsx';
  const kfintechPreviewName = document.getElementById('kfintech-excel-preview-name');
  if (kfintechPreviewName) kfintechPreviewName.textContent = 'Kfintech_Sample.xlsx';

  resetSignupSteps();
  checkAuthSession();
});

// Run auth check on initialization
checkAuthSession();


// --- Extra: Appended Google Sign-In + EmailJS multi-step signup logic ---

// --- 1. GOOGLE SIGN-IN LOGIC ---
function decodeJwtResponse(token) {
    let base64Url = token.split('.')[1];
    let base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
    let jsonPayload = decodeURIComponent(window.atob(base64).split('').map(function(c) {
        return '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2);
    }).join(''));
    return JSON.parse(jsonPayload);
}

function handleCredentialResponse(response) {
    const responsePayload = decodeJwtResponse(response.credential);
    const email = responsePayload.email;
    localStorage.setItem('invoice_hub_session', email);
    if (!localStorage.getItem(`invoice_hub_user_${email.toLowerCase()}`)) {
      localStorage.setItem(`invoice_hub_user_${email.toLowerCase()}`, 'google_auth');
    }
    document.getElementById('user-display-email').innerText = email;
    document.getElementById('auth-container').style.display = 'none';
    const chooser = document.getElementById('google-chooser-modal');
    if (chooser) chooser.style.display = 'none';
    document.querySelector('.app-container').style.display = 'flex';
    loadUserSignature(email);
    loadUserExcelFiles(email);

    const isOnboarded = (localStorage.getItem(`invoice_hub_user_onboarded_${email.toLowerCase()}`) === 'true') ||
                        (localStorage.getItem(`invoice_hub_user_sig_${email.toLowerCase()}`) !== null);
    cardSign.style.display = isOnboarded ? 'none' : 'flex';

    if (isOnboarded) {
      history.replaceState({ view: 'dashboard' }, '', '#dashboard');
      closeWorkspaces();
    } else {
      history.replaceState({ view: 'dashboard' }, '', '#dashboard');
      history.pushState({ view: 'sign' }, '', '#sign');
      openWorkspace(wsSign, cardSign);
    }
}

function initGoogleSignIn() {
    if (window.google && google.accounts && google.accounts.id) {
        google.accounts.id.initialize({
            client_id: "1089142677625-f9gspg0hits2ovnt9i7km2h80gpcp5vv.apps.googleusercontent.com",
            callback: handleCredentialResponse
        });
        const loginDiv = document.getElementById("google-button-login");
        if(loginDiv) google.accounts.id.renderButton(loginDiv, { theme: "outline", size: "large", width: "300" });

        const signupDiv = document.getElementById("google-button-signup");
        if(signupDiv) google.accounts.id.renderButton(signupDiv, { theme: "outline", size: "large", width: "300" });
        
        google.accounts.id.prompt(); 
    }
}

if (document.readyState === 'complete') {
    initGoogleSignIn();
} else {
    window.addEventListener('load', initGoogleSignIn);
}

// --- 2. MULTI-STEP EMAILJS SIGNUP LOGIC ---

// EmailJS Credentials
const SERVICE_ID = "service_invoicehub";
const TEMPLATE_ID = "template_gdv3r9k";
const PUBLIC_KEY = "dCNuEyMUZ8Do45kVA";

// View Switching Logic (Login vs Signup tab)
const btnSignupNav = document.getElementById('nav-btn-signup');
const btnLoginNav = document.getElementById('nav-btn-login');

if (btnSignupNav && btnLoginNav) {
    btnSignupNav.addEventListener('click', function() {
        document.getElementById('view-login').style.display = 'none';
        document.getElementById('view-signup').style.display = 'block';
        this.classList.add('active');
        this.classList.remove('outline');
        btnLoginNav.classList.remove('active');
        btnLoginNav.classList.add('outline');
    });

    btnLoginNav.addEventListener('click', function() {
        document.getElementById('view-signup').style.display = 'none';
        document.getElementById('view-login').style.display = 'block';
        this.classList.add('active');
        this.classList.remove('outline');
        btnSignupNav.classList.remove('active');
        btnSignupNav.classList.add('outline');
    });
}

// Step 1: Send Verification Email
const btnVerifyEl = document.getElementById("btn-signup-verify-email");
if (btnVerifyEl) {
  btnVerifyEl.addEventListener("click", function(e) {
    e.preventDefault();
    
    const emailInput = document.getElementById("signup-email").value.trim();

    // STRICT VALIDATION: Must end exactly with @gmail.com
    const emailRegex = /^[a-zA-Z0-9._%+-]+@gmail\.com$/;
    
    if(!emailRegex.test(emailInput)) {
        alert("Invalid email. Please enter the correct email.");
        return; // This immediately stops the code, staying on Step 1
    }

    // Generate 6-digit OTP code using state variables from middle of the file
    const code = Math.floor(100000 + Math.random() * 900000).toString();
    generatedVerificationCode = code;
    verifiedEmail = emailInput;

    // Update button UI while sending
    const btn = document.getElementById("btn-signup-verify-email");
    const originalText = btn.innerHTML;
    btn.innerHTML = "<span>Sending...</span>";
    btn.disabled = true;

    // Send Email using EmailJS
    if (window.emailjs && emailjs.send) {
      emailjs.send(SERVICE_ID, TEMPLATE_ID, {
          email: emailInput,
          passcode: code
      }, PUBLIC_KEY)
      .then(function(response) {
          console.log("Email Sent Successfully!", response.status);
          
          alert(`A verification code has been sent to ${emailInput}. Please check your inbox and spam folder.`);

          // Move to Step 2 ONLY because the email sent successfully
          document.getElementById("signup-step-email").style.display = "none";
          document.getElementById("signup-step-code").style.display = "block";

          btn.innerHTML = originalText;
          btn.disabled = false;
      }, function(error) {
          console.log("Email Sending Failed...", error);
          alert("There was an issue sending the verification email. Please try again later.");
          
          btn.innerHTML = originalText;
          btn.disabled = false;
      });
    } else {
      // Fallback simulation if EmailJS not available (code logged to console, not shown in alert)
      console.warn('EmailJS not available — simulating verification email.');
      console.log(`[Simulation] A verification code has been generated. Your code: ${code}`);
      alert(`A verification code has been sent to ${emailInput}. Please check your inbox and spam folder.`);
      document.getElementById("signup-step-email").style.display = "none";
      document.getElementById("signup-step-code").style.display = "block";
      btn.innerHTML = originalText;
      btn.disabled = false;
    }
  });
}

// Step 2: Check Verification Code
const btnCodeSubmitEl = document.getElementById("btn-signup-code-submit");
if (btnCodeSubmitEl) {
  btnCodeSubmitEl.addEventListener("click", function(e) {
      e.preventDefault();
      
      const enteredCode = document.getElementById("signup-verification-code").value.trim();

      if(enteredCode === generatedVerificationCode) {
          // Code Matches! Move to Step 3
          document.getElementById("signup-step-code").style.display = "none";
          document.getElementById("signup-step-password").style.display = "block";
      } else {
          alert("Invalid verification code. Please check the code and try again.");
      }
  });
}

// Step 3: Final Form Submission
const formSignupEl = document.getElementById("form-signup");
if (formSignupEl) {
  formSignupEl.addEventListener("submit", async function(e) {
      e.preventDefault(); 

      // If Step 1 (Email) is currently active, route Enter key to validation/sending
      const stepEmail = document.getElementById("signup-step-email");
      if (stepEmail && stepEmail.style.display !== "none") {
          const btnVerify = document.getElementById("btn-signup-verify-email");
          if (btnVerify) btnVerify.click();
          return;
      }

      // If Step 2 (Code) is currently active, route Enter key to code validation
      const stepCode = document.getElementById("signup-step-code");
      if (stepCode && stepCode.style.display !== "none") {
          const btnCodeSubmit = document.getElementById("btn-signup-code-submit");
          if (btnCodeSubmit) btnCodeSubmit.click();
          return;
      }
      
      const pass = document.getElementById("signup-password").value;
      const confirmPass = document.getElementById("signup-confirm-password").value;

      if(pass !== confirmPass) {
          alert("Passwords do not match!");
          return;
      }
      
      if (pass.length < 6 || pass.length > 10) {
          alert("Invalid password. Create correct password.");
          return;
      }

      // Register user in localStorage using the verified email address
      localStorage.setItem(`invoice_hub_user_${verifiedEmail.toLowerCase()}`, pass);

      // Increment signup count
      let signupCount = parseInt(localStorage.getItem('invoice_hub_signup_count') || '1');
      localStorage.setItem('invoice_hub_signup_count', signupCount + 1);

      alert("Account Created Successfully!");
      updateVisitsDisplay();
      await showLoading('Creating Account...', 1200);
      doLoginSession(verifiedEmail);

      // Reset step views back to step 1
      resetSignupSteps();
  });
}

// --- 3. CORNER IMAGE ZOOM MODAL LOGIC ---
const sigPreviewContainer = document.querySelector('#workspace-sign .sig-preview-container');
const imageZoomModal = document.getElementById('image-zoom-modal');
const btnCloseImageZoom = document.getElementById('btn-close-image-zoom');

if (sigPreviewContainer && imageZoomModal) {
  // Click on the preview image opens the zoom modal
  sigPreviewContainer.addEventListener('click', (e) => {
    e.stopPropagation();
    imageZoomModal.style.display = 'flex';
  });

  // Click on the overlay (outside the image container) closes it
  imageZoomModal.addEventListener('click', () => {
    imageZoomModal.style.display = 'none';
  });

  if (btnCloseImageZoom) {
    btnCloseImageZoom.addEventListener('click', (e) => {
      e.stopPropagation();
      imageZoomModal.style.display = 'none';
    });
  }

  // Close zoom modal on pressing Escape key
  window.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && imageZoomModal.style.display === 'flex') {
      imageZoomModal.style.display = 'none';
    }
  });
}

// --- 4. EXCEL ZOOM MODAL LOGIC & SHEETJS PARSING ---

const camsDefaultExcelData = [
  ['Transaction ID', 'Investor Name', 'Scheme Name', 'Amount (INR)', 'Status'],
  ['TXN001', 'Abhishek Saparia', 'CAMS Equity Fund', '50,000.00', 'Success'],
  ['TXN002', 'Sanvi Saparia', 'CAMS Liquid Fund', '25,000.00', 'Success'],
  ['TXN003', 'System SGPL', 'CAMS Balanced Fund', '100,000.00', 'Success'],
  ['TXN004', 'Abhishek', 'CAMS Debt Fund', '10,000.00', 'Pending']
];

const kfintechDefaultExcelData = [
  ['Folio Number', 'Investor Name', 'Asset Class', 'Units', 'NAV (INR)', 'Current Value (INR)'],
  ['FOL8832', 'Abhishek Saparia', 'Equity - Growth', '1,250.45', '45.20', '56,520.34'],
  ['FOL9043', 'Sanvi Saparia', 'Hybrid - Growth', '850.12', '125.40', '106,605.05'],
  ['FOL7123', 'System SGPL', 'Debt - Direct', '500.00', '18.90', '9,450.00']
];

function renderMockExcelTable(dataArray, container) {
  let html = '<table><thead><tr>';
  dataArray[0].forEach(cell => {
    html += `<th>${cell}</th>`;
  });
  html += '</tr></thead><tbody>';
  
  for (let i = 1; i < dataArray.length; i++) {
    html += '<tr>';
    dataArray[i].forEach(cell => {
      html += `<td>${cell}</td>`;
    });
    html += '</tr>';
  }
  
  html += '</tbody></table>';
  container.innerHTML = html;
}

function renderExcelTableFromBytes(buffer, container) {
  try {
    const data = new Uint8Array(buffer);
    const workbook = XLSX.read(data, { type: 'array' });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    
    // Convert to HTML
    const htmlTable = XLSX.utils.sheet_to_html(worksheet);
    container.innerHTML = htmlTable;
  } catch (error) {
    console.error("SheetJS failed to parse Excel table:", error);
    container.innerHTML = `<div style="padding: 24px; text-align: center; color: var(--accent-orange); font-weight: 500;">Failed to parse Excel sheet contents: ${error.message}</div>`;
  }
}

const excelZoomModal = document.getElementById('excel-zoom-modal');
const excelTableContainer = document.getElementById('excel-table-container');
const btnCloseExcelZoom = document.getElementById('btn-close-excel-zoom');

function openExcelPreview(stateBytes, defaultMockData) {
  if (!excelZoomModal || !excelTableContainer) return;
  excelTableContainer.innerHTML = '';
  
  if (stateBytes) {
    renderExcelTableFromBytes(stateBytes, excelTableContainer);
  } else {
    renderMockExcelTable(defaultMockData, excelTableContainer);
  }
  excelZoomModal.style.display = 'flex';
}

const camsExcelPreviewBox = document.getElementById('cams-excel-preview-container');
if (camsExcelPreviewBox) {
  camsExcelPreviewBox.addEventListener('click', (e) => {
    e.stopPropagation();
    openExcelPreview(camsExcelState.fileBytes, camsDefaultExcelData);
  });
}

const kfintechExcelPreviewBox = document.getElementById('kfintech-excel-preview-container');
if (kfintechExcelPreviewBox) {
  kfintechExcelPreviewBox.addEventListener('click', (e) => {
    e.stopPropagation();
    openExcelPreview(kfintechExcelState.fileBytes, kfintechDefaultExcelData);
  });
}

if (excelZoomModal) {
  excelZoomModal.addEventListener('click', () => {
    excelZoomModal.style.display = 'none';
  });
  
  const excelContent = excelZoomModal.querySelector('.excel-zoom-content');
  if (excelContent) {
    excelContent.addEventListener('click', (e) => {
      e.stopPropagation(); // Prevent closing when clicking table itself
    });
  }
}

if (btnCloseExcelZoom) {
  btnCloseExcelZoom.addEventListener('click', (e) => {
    e.stopPropagation();
    excelZoomModal.style.display = 'none';
  });
}

window.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') {
    if (excelZoomModal && excelZoomModal.style.display === 'flex') {
      excelZoomModal.style.display = 'none';
    }
  }
});

// ============================================================================
// 6. Browser History & Trackpad Swiping Navigation
// ============================================================================

function goBackToDashboard() {
  if (history.state && history.state.view && history.state.view !== 'dashboard') {
    history.back();
  } else {
    history.pushState({ view: 'dashboard' }, '', '#dashboard');
    closeWorkspaces();
  }
}

// popstate event listener for browser history (back/forward and trackpad swipes)
window.addEventListener('popstate', (e) => {
  const state = e.state;
  const view = (state && state.view) ? state.view : 'dashboard';
  
  // Force new users to complete onboarding before they can access dashboard or swipe away
  const sessionEmail = localStorage.getItem('invoice_hub_session');
  if (sessionEmail) {
    const isOnboarded = localStorage.getItem(`invoice_hub_user_onboarded_${sessionEmail.toLowerCase()}`) === 'true';
    if (!isOnboarded) {
      history.replaceState({ view: 'dashboard' }, '', '#dashboard');
      history.pushState({ view: 'sign' }, '', '#sign');
      openWorkspace(wsSign, cardSign);
      return;
    }
  }
  
  if (view === 'dashboard') {
    closeWorkspaces();
  } else if (view === 'cams') {
    openWorkspace(wsCams, cardCams);
    const wsCamsExcel = document.getElementById('workspace-cams-excel');
    if (wsCamsExcel) wsCamsExcel.classList.add('active');
  } else if (view === 'kfintech') {
    openWorkspace(wsKfintech, cardKfintech);
    const wsKfintechExcel = document.getElementById('workspace-kfintech-excel');
    if (wsKfintechExcel) wsKfintechExcel.classList.add('active');
  } else if (view === 'sign') {
    openWorkspace(wsSign, cardSign);
  } else if (view === 'profile') {
    const wsProfile = document.getElementById('workspace-profile');
    if (wsProfile) openWorkspace(wsProfile);
  } else if (view === 'settings') {
    const wsSettings = document.getElementById('workspace-settings');
    if (wsSettings) openWorkspace(wsSettings);
  } else if (view === 'login') {
    authContainer.style.display = 'flex';
    appContainer.style.display = 'none';
    showLoginView();
    closeWorkspaces();
  } else if (view === 'signup') {
    authContainer.style.display = 'flex';
    appContainer.style.display = 'none';
    showSignupView();
    closeWorkspaces();
  }
});

// Initialize History state on load
const initialView = localStorage.getItem('invoice_hub_session') ? 'dashboard' : 'login';
history.replaceState({ view: initialView }, '', `#${initialView}`);

// Mousepad / Trackpad horizontal swipe detection
let wheelTimeout = null;
let accumulativeDeltaX = 0;
const SWIPE_THRESHOLD = 150; // px threshold for trackpad swipe detection

window.addEventListener('wheel', (e) => {
  // If vertical scrolling is dominant, ignore this gesture to avoid false positives
  if (Math.abs(e.deltaY) > Math.abs(e.deltaX)) {
    return;
  }
  
  accumulativeDeltaX += e.deltaX;
  
  if (wheelTimeout) {
    clearTimeout(wheelTimeout);
  }
  
  wheelTimeout = setTimeout(() => {
    if (Math.abs(accumulativeDeltaX) > SWIPE_THRESHOLD) {
      console.log("Trackpad swipe detected! DeltaX:", accumulativeDeltaX);
      history.back();
    }
    accumulativeDeltaX = 0;
  }, 100);
}, { passive: true });

// Touchscreen swipe gesture detection
let touchStartX = 0;
let touchStartY = 0;
const TOUCH_SWIPE_THRESHOLD = 80;

window.addEventListener('touchstart', (e) => {
  if (e.touches.length === 1) {
    touchStartX = e.touches[0].screenX;
    touchStartY = e.touches[0].screenY;
  }
}, { passive: true });

window.addEventListener('touchend', (e) => {
  if (e.changedTouches.length === 1) {
    const deltaX = e.changedTouches[0].screenX - touchStartX;
    const deltaY = e.changedTouches[0].screenY - touchStartY;
    
    if (Math.abs(deltaX) > Math.abs(deltaY) && Math.abs(deltaX) > TOUCH_SWIPE_THRESHOLD) {
      console.log("Touch swipe detected! DeltaX:", deltaX);
      history.back();
    }
  }
}, { passive: true });

// Toggle profile dropdown menu
const btnProfileDropdown = document.getElementById('btn-profile-dropdown');
const profileDropdownMenu = document.getElementById('profile-dropdown-menu');

if (btnProfileDropdown && profileDropdownMenu) {
  btnProfileDropdown.addEventListener('click', (e) => {
    e.stopPropagation();
    const isExpanded = btnProfileDropdown.getAttribute('aria-expanded') === 'true';
    btnProfileDropdown.setAttribute('aria-expanded', !isExpanded);
    profileDropdownMenu.classList.toggle('active');
  });

  // Close dropdown menu when clicking anywhere else
  window.addEventListener('click', () => {
    if (btnProfileDropdown) btnProfileDropdown.setAttribute('aria-expanded', 'false');
    if (profileDropdownMenu) profileDropdownMenu.classList.remove('active');
  });
}

// Dropdown menu action bindings
const btnDropdownProfile = document.getElementById('btn-dropdown-profile');
if (btnDropdownProfile) {
  btnDropdownProfile.addEventListener('click', () => {
    const wsProfile = document.getElementById('workspace-profile');
    if (wsProfile) {
      openWorkspace(wsProfile);
      if (!history.state || history.state.view !== 'profile') {
        navigateSubpage('profile', '#profile');
      }
    }
  });
}

const btnDropdownSettings = document.getElementById('btn-dropdown-settings');
if (btnDropdownSettings) {
  btnDropdownSettings.addEventListener('click', () => {
    const wsSettings = document.getElementById('workspace-settings');
    if (wsSettings) {
      openWorkspace(wsSettings);
      
      // Sync signature state and displays in settings card
      const hasSig = !!globalSignature.imageBytes;
      document.getElementById('dropzone-settings-sign-img').style.display = hasSig ? 'none' : 'block';
      const pillSettings = document.getElementById('pill-settings-sign-img');
      if (pillSettings) {
        pillSettings.style.display = hasSig ? 'inline-flex' : 'none';
      }
      
      if (!history.state || history.state.view !== 'settings') {
        navigateSubpage('settings', '#settings');
      }
    }
  });
}

// Close buttons for new workspaces
const btnCloseProfile = document.getElementById('btn-close-profile-workspace');
if (btnCloseProfile) {
  btnCloseProfile.addEventListener('click', (e) => {
    e.stopPropagation();
    goBackToDashboard();
  });
}

const btnCloseSettings = document.getElementById('btn-close-settings-workspace');
if (btnCloseSettings) {
  btnCloseSettings.addEventListener('click', (e) => {
    e.stopPropagation();
    goBackToDashboard();
  });
}

// Zoom click listener for settings signature preview
const settingsSigPreviewContainer = document.getElementById('settings-sig-preview-container');
if (settingsSigPreviewContainer && imageZoomModal) {
  settingsSigPreviewContainer.addEventListener('click', (e) => {
    e.stopPropagation();
    imageZoomModal.style.display = 'flex';
  });
}

// Logo redirection to dashboard
const logoLink = document.getElementById('logo-link');
if (logoLink) {
  logoLink.addEventListener('click', (e) => {
    e.preventDefault();
    history.pushState({ view: 'dashboard' }, '', '#dashboard');
    closeWorkspaces();
  });
}

// Video Tutorial Modal listener bindings
const btnVideoTutorial = document.getElementById('btn-video-tutorial');
const videoTutorialModal = document.getElementById('video-tutorial-modal');
const btnCloseVideoModal = document.getElementById('btn-close-video-modal');

if (btnVideoTutorial && videoTutorialModal) {
  btnVideoTutorial.addEventListener('click', (e) => {
    e.stopPropagation();
    videoTutorialModal.style.display = 'flex';
  });

  videoTutorialModal.addEventListener('click', () => {
    videoTutorialModal.style.display = 'none';
  });

  const videoCard = videoTutorialModal.querySelector('.video-modal-card');
  if (videoCard) {
    videoCard.addEventListener('click', (e) => {
      e.stopPropagation();
    });
  }

  if (btnCloseVideoModal) {
    btnCloseVideoModal.addEventListener('click', (e) => {
      e.stopPropagation();
      videoTutorialModal.style.display = 'none';
    });
  }
}

function initSignupCount() {
  if (!localStorage.getItem('invoice_hub_signup_count')) {
    let count = 1; // start with 1 for default demo@example.com account
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (key && key.startsWith('invoice_hub_user_') && !key.startsWith('invoice_hub_user_sig_') && !key.startsWith('invoice_hub_user_onboarded_')) {
        count++;
      }
    }
    localStorage.setItem('invoice_hub_signup_count', count);
  }
}

function updateVisitsDisplay() {
  const countEls = document.querySelectorAll('.visits-number');
  const count = localStorage.getItem('invoice_hub_signup_count') || '1';
  countEls.forEach(el => {
    el.textContent = count;
  });
}

// Initialize signup count and update UI
initSignupCount();
updateVisitsDisplay();

