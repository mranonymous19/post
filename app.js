console.log('ExcelJS loaded:', typeof ExcelJS);
console.log('Supabase loaded:', typeof supabase);

// const SUPABASE_URL = 'https://kowxagkvsnhbfqmfuzti.supabase.co';
// const SUPABASE_ANON_KEY = 'sb_publishable_r_iTeNJT1IHIeVzvHZAvHg_QP9hieXd';

const SUPABASE_URL = 'https://uflljrusgxggdfbsoxzc.supabase.co';
const SUPABASE_ANON_KEY = 'sb_publishable_odGXL4HY0DhX3HnmAi7Nxg_Lwr0xsL-';

const supabaseClient = supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

async function testConnection() {
  const { data, error } = await supabaseClient
    .from('orders')
    .select('shopify_order_id')
    .limit(1);

  console.log('Test query data:', data);
  console.log('Test query error:', error);
}

// testConnection(); // done verifying, disabled for now

// ---- Element references ----
let currentOrder = null; // holds the full fetched row for reuse at Submit time

const orderIdInput = document.getElementById('orderId');
const customerNameInput = document.getElementById('customerName');
const paymentTypeInput = document.getElementById('paymentType');
const orderPriceInput = document.getElementById('orderPrice');
const pincodeInput = document.getElementById('pincode');
const formError = document.getElementById('formError');
const barcodeInput = document.getElementById('barcode');
const weightInput = document.getElementById('weight');
const lengthInput = document.getElementById('length');
const breadthInput = document.getElementById('breadth');
const heightInput = document.getElementById('height');
const submitBtn = document.getElementById('submitBtn');
const manifestBody = document.getElementById('manifestBody');



// ---- Order ID lookup ----
function clearLookupFields() {
  customerNameInput.value = '';  
  paymentTypeInput.value = '';
  orderPriceInput.value = '';
  pincodeInput.value = '';
  currentOrder = null;
}

async function performLookup() {
  const rawInput = orderIdInput.value.trim();
  formError.textContent = '';

  if (!rawInput) {
    formError.textContent = 'Please enter an Order ID.';
    clearLookupFields();
    return;
  }

  const stripped = rawInput.replace(/^#/, '');
  const withHash = '#' + stripped;

  const { data, error } = await supabaseClient
    .from('orders')
    .select('*')
    .or(`order_name.eq.${stripped},order_name.eq.${withHash}`)
    .maybeSingle();

  if (error) {
    console.error('Lookup error:', error);
    formError.textContent = 'Lookup failed — check your connection and try again.';
    clearLookupFields();
    return;
  }

  if (!data) {
    formError.textContent = 'Order ID not found.';
    clearLookupFields();
    return;
  }

  currentOrder = data;
  customerNameInput.value = cleanValue(data.customer_name);   
  const invoiceUpper = (data.invoice_number || '').trim().toUpperCase();
  paymentTypeInput.value = invoiceUpper.startsWith('COD') ? 'COD' : invoiceUpper.startsWith('SHP') ? 'SHP' : '';
  orderPriceInput.value = data.amount_to_receive != null ? Number(data.amount_to_receive).toFixed(2) : '';
  pincodeInput.value = data.shipping_pincode || '';
}

orderIdInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') {
    e.preventDefault();
    performLookup();
  }
});

orderIdInput.addEventListener('input', () => {
  // if the operator edits the Order ID after a lookup, clear stale auto-filled data
  if (currentOrder) clearLookupFields();
});

// ---- Data cleaning ----
function cleanValue(value) {
  if (value === null || value === undefined) return '';
  let s = String(value).trim();
  if (s.toLowerCase() === 'null') return '';
  s = s.replace(/\bnull\b/gi, '');
  s = s.replace(/\s{2,}/g, ' ');
  s = s.replace(/^[,\s]+|[,\s]+$/g, '');
  return s.trim();
}

function cleanPhoneNumber(value) {
  if (!value) return '';
  const digitsOnly = String(value).replace(/\D/g, ''); // strip everything except digits
  return digitsOnly.length > 10 ? digitsOnly.slice(-10) : digitsOnly;
}

// ---- Address Line 1/2 rule ----
function buildAddressLines(rawAddress1, rawAddress2, rawCity) {
  let line1 = cleanValue(rawAddress1);
  const line2Base = cleanValue(rawAddress2) || cleanValue(rawCity);
  const overflow = [];

  while (line1.length >= 50) {
    const words = line1.split(' ');
    if (words.length === 1) break; // safety: single unbreakable word, stop
    overflow.unshift(words.pop());
    line1 = words.join(' ');
  }

  let line2 = overflow.length ? (overflow.join(' ') + ' ' + line2Base).trim() : line2Base;

  while (line2.length >= 50) {
    const words = line2.split(' ');
    if (words.length === 1) break;
    words.pop();
    line2 = words.join(' ');
  }

  return { line1, line2 };
}

// ---- Fixed values ----
const FIXED_VALUES = {
  shapeOfArticle: 'NROL',
  priorityFlag: true,
  deliveryInstruction: 'ND',
  instructionRts: 'RTS',
  senderName: 'AROVEHIC',
  senderCompany: 'ARIHANTH COMPLEX,3RD FLOOR',
  senderAddLine1: 'J C ROAD,A M ROAD,',
  senderAddLine2: 'BENGALURU-560002',
  senderCity: 'BENGALURU',
  senderState: 'KARNATAKA',
  senderPincode: '560002',
  senderEmailId: 'arovehic@gmail.com',
  senderMobileNo: '6360818919',
  altAddressFlag: false,
  pickupAddressFlag: false,
  dropOffPincode: '560001',
  dropoffPickupOfficeId: '21250001',
  ack: false,
  registration: false,
  otpBasedDelivery: false,
};

// ---- Manifest column definitions (order matches Excel template A–AV) ----
const MANIFEST_COLUMNS = [
  ['serialNumber', 'SERIAL NUMBER'], ['barcodeNo', 'BARCODE NO'],
  ['physicalWeight', 'PHYSICAL WEIGHT'], ['shapeOfArticle', 'SHAPE OF ARTICLE'], ['length', 'LENGTH'],
  ['breadthDiameter', 'BREADTH/DIAMETER'], ['height', 'HEIGHT'], ['priorityFlag', 'PRIORITY FLAG'],
  ['deliveryInstruction', 'DELIVERY INSTRUCTION'], ['instructionRts', 'INSTRUCTION RTS'],
  ['senderName', 'SENDER NAME'], ['senderCompany', 'SENDER COMPANY'], ['senderAddLine1', 'SENDER ADD LINE 1'],
  ['senderAddLine2', 'SENDER ADD LINE 2'], ['senderCity', 'SENDER CITY'], ['senderState', 'SENDER STATE'],
  ['senderPincode', 'SENDER PINCODE'], ['senderEmailId', 'SENDER EMAILID'], ['senderAltContact', 'SENDER ALT CONTACT'],
  ['senderKyc', 'SENDER KYC'], ['senderTaxReference', 'SENDER TAX REFERENCE'], ['receiverName', 'RECEIVER NAME'],
  ['receiverCompany', 'RECEIVER COMPANY'], ['receiverAddLine1', 'RECEIVER ADD LINE 1'],
  ['receiverAddLine2', 'RECEIVER ADD LINE 2'], ['receiverCity', 'RECEIVER CITY'], ['receiverState', 'RECEIVER STATE'],
  ['receiverPincode', 'RECEIVER PINCODE'], ['receiverEmailId', 'RECEIVER EMAILID'],
  ['receiverAltContact', 'RECEIVER ALT CONTACT'], ['receiverKyc', 'RECEIVER KYC'],
  ['receiverTaxReference', 'RECEIVER TAX REFERENCE'], ['altAddressFlag', 'ALT ADDRESS FLAG'],
  ['pickupAddressFlag', 'PICKUP ADDRESS FLAG'], ['dropOffPincode', 'DROP OFF PINCODE'],
  ['dropoffPickupOfficeId', 'DROPOFF/PICKUP OFFICE ID'], ['senderMobileNo', 'SENDER MOBILE NO'],
  ['receiverMobileNo', 'RECEIVER MOBILE NO'], ['prepaymentCode', 'PREPAYMENT CODE'],
  ['valueOfPrepayment', 'VALUE OF PREPAYMENT'], ['codrCod', 'CODR/COD'], ['valueForCodrCod', 'VALUE FOR CODR/COD'],
  ['insuranceType', 'INSURANCE TYPE'], ['valueOfInsurance', 'VALUE OF INSURANCE'], ['ack', 'ACK'],
  ['registration', 'REGISTRATION'], ['otpBasedDelivery', 'OTP BASED DELIVERY'], ['bulkReference', 'BULK REFERENCE'],
];

let manifestEntries = [];
let serialCounter = 1;

const STORAGE_KEY = 'courierManifest';

function saveManifestToStorage() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify({ serialCounter, entries: manifestEntries }));
}

function loadManifestFromStorage() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) return;
  try {
    const parsed = JSON.parse(raw);
    manifestEntries = parsed.entries || [];
    serialCounter = parsed.serialCounter || 1;
    manifestEntries.forEach(renderManifestRow);
  } catch (e) {
    console.error('Failed to load saved manifest:', e);
  }
}
// ---- Dimension minimums ----
const MIN_DIMENSIONS = { length: 14, breadth: 9, height: 1 };

// ---- Per-field error display helpers ----
const lengthErrorEl = document.getElementById('lengthError');
const breadthErrorEl = document.getElementById('breadthError');
const heightErrorEl = document.getElementById('heightError');

function showFieldError(el, message) {
  el.textContent = message;
  el.classList.remove('hidden');
}

function clearFieldError(el) {
  el.textContent = '';
  el.classList.add('hidden');
}

function clearAllDimensionErrors() {
  clearFieldError(lengthErrorEl);
  clearFieldError(breadthErrorEl);
  clearFieldError(heightErrorEl);
}

// ---- Submit validation ----
function validateManualFields() {
  clearAllDimensionErrors();

  if (!currentOrder) return 'Look up a valid Order ID first.';
  if (!barcodeInput.value.trim()) return 'Barcode No is required.';
  if (!weightInput.value || Number(weightInput.value) <= 0) return 'Physical Weight is required.';

  let hasDimensionError = false;

  if (!lengthInput.value || Number(lengthInput.value) < MIN_DIMENSIONS.length) {
    showFieldError(lengthErrorEl, `Length must be at least ${MIN_DIMENSIONS.length} cm.`);
    hasDimensionError = true;
  }
  if (!breadthInput.value || Number(breadthInput.value) < MIN_DIMENSIONS.breadth) {
    showFieldError(breadthErrorEl, `Breadth/Diameter must be at least ${MIN_DIMENSIONS.breadth} cm.`);
    hasDimensionError = true;
  }
  if (!heightInput.value || Number(heightInput.value) < MIN_DIMENSIONS.height) {
    showFieldError(heightErrorEl, `Height must be at least ${MIN_DIMENSIONS.height} cm.`);
    hasDimensionError = true;
  }

  if (hasDimensionError) return 'Please fix the highlighted dimension fields.';

  return null;
}

// ---- Row builder ----
function buildManifestRow() {
  const order = currentOrder;
  const isCOD = (order.invoice_number || '').trim().toUpperCase().startsWith('COD');
  const { line1, line2 } = buildAddressLines(order.shipping_address1, order.shipping_address2, order.shipping_city);

  return {
    orderId: order.shopify_order_id,
    orderName: orderIdInput.value.trim(),
    barcodeNo: barcodeInput.value.trim(),
    physicalWeight: Number(weightInput.value),
    shapeOfArticle: FIXED_VALUES.shapeOfArticle,
    length: Number(lengthInput.value),
    breadthDiameter: Number(breadthInput.value),
    height: Number(heightInput.value),
    priorityFlag: FIXED_VALUES.priorityFlag,
    deliveryInstruction: FIXED_VALUES.deliveryInstruction,
    instructionRts: FIXED_VALUES.instructionRts,
    senderName: FIXED_VALUES.senderName,
    senderCompany: FIXED_VALUES.senderCompany,
    senderAddLine1: FIXED_VALUES.senderAddLine1,
    senderAddLine2: FIXED_VALUES.senderAddLine2,
    senderCity: FIXED_VALUES.senderCity,
    senderState: FIXED_VALUES.senderState,
    senderPincode: FIXED_VALUES.senderPincode,
    senderEmailId: FIXED_VALUES.senderEmailId,
    senderAltContact: '', senderKyc: '', senderTaxReference: '',
    receiverName: cleanValue(order.customer_name),
    receiverCompany: '',
    receiverAddLine1: line1,
    receiverAddLine2: line2,
    receiverCity: cleanValue(order.shipping_city),
    receiverState: cleanValue(order.shipping_state),
    receiverPincode: order.shipping_pincode || '',
    receiverEmailId: '', receiverAltContact: '', receiverKyc: '', receiverTaxReference: '',
    altAddressFlag: FIXED_VALUES.altAddressFlag,
    pickupAddressFlag: FIXED_VALUES.pickupAddressFlag,
    dropOffPincode: FIXED_VALUES.dropOffPincode,
    dropoffPickupOfficeId: FIXED_VALUES.dropoffPickupOfficeId,
    senderMobileNo: FIXED_VALUES.senderMobileNo,
    receiverMobileNo: cleanPhoneNumber(order.customer_phone),
    prepaymentCode: '', valueOfPrepayment: '',
    codrCod: isCOD ? 'COD' : '',
    valueForCodrCod: isCOD ? (order.amount_to_receive != null ? Number(order.amount_to_receive.toFixed(2)) : '') : '',
    insuranceType: '', valueOfInsurance: '',
    ack: FIXED_VALUES.ack,
    registration: FIXED_VALUES.registration,
    otpBasedDelivery: FIXED_VALUES.otpBasedDelivery,
    bulkReference: '',
  };
}

// ---- Duplicate check, rendering, clearing ----


function renderManifestRow(row) {
  const tr = document.createElement('tr');
  MANIFEST_COLUMNS.forEach(([key]) => {
    const td = document.createElement('td');
    let val = row[key];
    if (typeof val === 'boolean') val = val ? 'TRUE' : 'FALSE';
    td.textContent = val === undefined || val === null || val === '' ? '' : val;
    tr.appendChild(td);
  });
  manifestBody.appendChild(tr);
}

function clearManualFields() {
  orderIdInput.value = '';
  barcodeInput.value = '';
  weightInput.value = '';
  lengthInput.value = '';
  breadthInput.value = '';
  heightInput.value = '';
  clearLookupFields();
  clearAllDimensionErrors();
}

// ---- Submit Entry ----
submitBtn.addEventListener('click', () => {
  formError.textContent = '';
  const validationError = validateManualFields();
  if (validationError) {
    formError.textContent = validationError;
    return;
  }

    if (isDuplicateOrderName(orderIdInput.value)) {
    formError.textContent = 'already this ordered id data is filled';
    return;
  }

  const row = buildManifestRow();
  row.serialNumber = serialCounter;
  manifestEntries.push(row);
  renderManifestRow(row);
  serialCounter++;
  saveManifestToStorage();
  clearManualFields();
  showToast(`Row ${row.serialNumber} added`);
});

function resetManifest() {
  manifestEntries = [];
  serialCounter = 1;
  manifestBody.innerHTML = '';
  localStorage.removeItem(STORAGE_KEY);
}

const downloadBtn = document.getElementById('downloadBtn');
const TEMPLATE_PATH = 'assets/bulkdomesticone_28042026_.xlsx';
const START_ROW = 2; // rows 2–18 are the preserved 17 existing rows

function getDateStamp() {
  const now = new Date();
  const dd = String(now.getDate()).padStart(2, '0');
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  const yyyy = now.getFullYear();
  return `${dd}${mm}${yyyy}`;
}

// ---- Save the manifest to Supabase (first entry per Order ID wins) ----
function toSupabaseRow(entry) {
  const text = (v) => (v === undefined || v === null || v === '' ? null : String(v));
  const num = (v) => (v === undefined || v === null || v === '' || Number.isNaN(Number(v)) ? null : Number(v));
  return {
    order_id: String(entry.orderName || '').trim(),
    barcode_no: text(entry.barcodeNo),
    physical_weight: num(entry.physicalWeight),
    length: num(entry.length),
    breadth_diameter: num(entry.breadthDiameter),
    height: num(entry.height),
    receiver_name: text(entry.receiverName),
    receiver_add_line_1: text(entry.receiverAddLine1),
    receiver_add_line_2: text(entry.receiverAddLine2),
    receiver_city: text(entry.receiverCity),
    receiver_state: text(entry.receiverState),
    receiver_pincode: text(entry.receiverPincode),
    receiver_mobile_no: text(entry.receiverMobileNo),
    codr_cod: text(entry.codrCod),
    value_for_codr_cod: num(entry.valueForCodrCod),
  };
}

async function saveManifestToSupabase() {
  const rows = manifestEntries.map(toSupabaseRow);
  const { data, error } = await supabaseClient.rpc('save_india_post_entries', { rows });
  if (error) throw error;
  return { saved: Number(data ?? 0), total: rows.length };
}

function showSaveMessage(message) {
  formError.textContent = message;
  formError.classList.replace('text-rose-600', 'text-emerald-600');
  setTimeout(() => {
    formError.classList.replace('text-emerald-600', 'text-rose-600');
    if (formError.textContent === message) formError.textContent = '';
  }, 6000);
}

async function downloadExcel() {
  formError.textContent = '';
  downloadBtn.disabled = true;
  try {
    const response = await fetch(TEMPLATE_PATH);
    if (!response.ok) throw new Error(`Template fetch failed: ${response.status}`);
    const buffer = await response.arrayBuffer();

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    const sheet = workbook.getWorksheet('ArticleDetails');
    if (!sheet) throw new Error('ArticleDetails sheet not found in template');

    manifestEntries.forEach((entry, i) => {
      const rowNum = START_ROW + i;
      MANIFEST_COLUMNS.forEach(([key], colIndex) => {
        sheet.getRow(rowNum).getCell(colIndex + 1).value = entry[key];
      });
    });

    const outBuffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([outBuffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `bulkdomesticone_${getDateStamp()}_.xlsx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);

    await saveAndClearAfterDownload();
  } catch (err) {
    console.error('Excel generation failed:', err);
    formError.textContent = 'Failed to generate Excel file. Please try again.';
  } finally {
    downloadBtn.disabled = false;
  }
}


downloadBtn.addEventListener('click', downloadExcel);

// ---- Auto-save to Supabase + clear after download ----
async function saveAndClearAfterDownload() {
  if (manifestEntries.length === 0) return; // nothing to save or clear

  try {
    const saveResult = await saveManifestToSupabase();
    resetManifest();
    const skipped = saveResult.total - saveResult.saved;
    showSaveMessage(
      skipped > 0
        ? `Saved ${saveResult.saved} new order(s) to Supabase. ${skipped} skipped (already saved). Active Manifest cleared.`
        : `Saved ${saveResult.saved} order(s) to Supabase. Active Manifest cleared.`
    );
    showToast(`Saved ${saveResult.saved} order(s) to Supabase`);
  } catch (err) {
    console.error('Supabase save failed:', err);
    formError.textContent = 'File downloaded, but saving to Supabase failed. Rows were NOT cleared — click Download Excel again to retry.';
    showToast('Saving to Supabase failed', 'error');
  }
}

// ---- Delete Row ----
const actionSelect = document.getElementById('actionSelect');
const deleteRowModal = document.getElementById('deleteRowModal');
const deleteOrderIdInput = document.getElementById('deleteOrderId');
const deleteRowError = document.getElementById('deleteRowError');
const deleteRowCloseBtn = document.getElementById('deleteRowCloseBtn');
const deleteRowCancelBtn = document.getElementById('deleteRowCancelBtn');
const deleteRowConfirmBtn = document.getElementById('deleteRowConfirmBtn');

const DELETE_FIELDS = [
  ['deleteSerialNumber', 'serialNumber'], ['deleteBarcode', 'barcodeNo'],
  ['deleteWeight', 'physicalWeight'], ['deleteLength', 'length'],
  ['deleteBreadth', 'breadthDiameter'], ['deleteHeight', 'height'],
  ['deleteReceiverName', 'receiverName'], ['deleteReceiverAddLine1', 'receiverAddLine1'],
  ['deleteReceiverAddLine2', 'receiverAddLine2'], ['deleteReceiverCity', 'receiverCity'],
  ['deleteReceiverState', 'receiverState'], ['deleteReceiverPincode', 'receiverPincode'],
  ['deleteReceiverMobileNo', 'receiverMobileNo'],
  ['deleteCodrCod', 'codrCod'], ['deleteValueForCodrCod', 'valueForCodrCod'],
];

let matchedDeleteEntry = null;

function clearDeleteDisplayFields() {
  DELETE_FIELDS.forEach(([elId]) => {
    document.getElementById(elId).value = '';
  });
}

function openDeleteModal() {
  deleteOrderIdInput.value = '';
  deleteRowError.textContent = '';
  clearDeleteDisplayFields();
  deleteRowConfirmBtn.disabled = true;
  matchedDeleteEntry = null;
  deleteRowModal.classList.remove('hidden');
  deleteOrderIdInput.focus();
}

function closeDeleteModal() {
  deleteRowModal.classList.add('hidden');
}

function performDeleteLookup() {
  const raw = deleteOrderIdInput.value.trim();
  deleteRowError.textContent = '';
  clearDeleteDisplayFields();
  deleteRowConfirmBtn.disabled = true;
  matchedDeleteEntry = null;

  if (!raw) {
    deleteRowError.textContent = 'Please enter an Order ID.';
    return;
  }

  const stripped = raw.replace(/^#/, '').toUpperCase();
  const found = manifestEntries.find((e) => {
    const entryId = (e.orderName || '').replace(/^#/, '').toUpperCase();
    return entryId === stripped;
  });

  if (!found) {
    deleteRowError.textContent = 'Order ID not found in Active Manifest.';
    return;
  }

  matchedDeleteEntry = found;
  DELETE_FIELDS.forEach(([elId, key]) => {
    let val = found[key];
    if (typeof val === 'boolean') val = val ? 'TRUE' : 'FALSE';
    document.getElementById(elId).value = val === undefined || val === null ? '' : val;
  });
  deleteRowConfirmBtn.disabled = false;
}

function rerenderManifestTable() {
  manifestBody.innerHTML = '';
  manifestEntries.forEach(renderManifestRow);
}

actionSelect.addEventListener('change', () => {
  if (actionSelect.value === 'addRow') {
    openAddModal();
  } else if (actionSelect.value === 'deleteRow') {
    openDeleteModal();
  } else if (actionSelect.value === 'editRow') {
    openEditModal();
  } else if (actionSelect.value === 'supabaseEdit') {
    openSupaEditModal();
  }
  actionSelect.value = ''; // always return dropdown to placeholder
});

deleteOrderIdInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') {
    e.preventDefault();
    performDeleteLookup();
  }
});

deleteRowCloseBtn.addEventListener('click', closeDeleteModal);
deleteRowCancelBtn.addEventListener('click', closeDeleteModal);

deleteRowConfirmBtn.addEventListener('click', () => {
  if (!matchedDeleteEntry) return;

  manifestEntries = manifestEntries.filter((e) => e !== matchedDeleteEntry);
  manifestEntries.forEach((e, i) => {
    e.serialNumber = i + 1;
  });
  serialCounter = manifestEntries.length + 1;

  rerenderManifestTable();
  saveManifestToStorage();
  closeDeleteModal();
  showToast('Row deleted');
});

// ---- Delete All Rows (with count verification) ----
const deleteAllOpenBtn = document.getElementById('deleteAllOpenBtn');
const deleteAllModal = document.getElementById('deleteAllModal');
const deleteAllCountInput = document.getElementById('deleteAllCount');
const deleteAllError = document.getElementById('deleteAllError');
const deleteAllCloseBtn = document.getElementById('deleteAllCloseBtn');
const deleteAllCancelBtn = document.getElementById('deleteAllCancelBtn');
const deleteAllConfirmBtn = document.getElementById('deleteAllConfirmBtn');

function openDeleteAllModal() {
  deleteRowError.textContent = '';
  if (manifestEntries.length === 0) {
    deleteRowError.textContent = 'The Active Manifest is already empty.';
    return;
  }
  deleteAllCountInput.value = '';
  deleteAllError.textContent = '';
  deleteAllModal.classList.remove('hidden');
  deleteAllCountInput.focus();
}

function closeDeleteAllModal() {
  deleteAllModal.classList.add('hidden');
}

function confirmDeleteAll() {
  deleteAllError.textContent = '';
  const raw = deleteAllCountInput.value.trim();

  if (raw === '') {
    deleteAllError.textContent = 'Enter the total number of rows to confirm.';
    return;
  }
  if (Number(raw) !== manifestEntries.length) {
    deleteAllError.textContent = 'That number does not match. Nothing was deleted.';
    return;
  }

  resetManifest();
  closeDeleteAllModal();
  closeDeleteModal();
  showToast('All rows deleted');
}

deleteAllOpenBtn.addEventListener('click', openDeleteAllModal);
deleteAllCloseBtn.addEventListener('click', closeDeleteAllModal);
deleteAllCancelBtn.addEventListener('click', closeDeleteAllModal);
deleteAllConfirmBtn.addEventListener('click', confirmDeleteAll);

// ---- Edit Row ----
const editRowModal = document.getElementById('editRowModal');
const editOrderIdInput = document.getElementById('editOrderId');
const editRowError = document.getElementById('editRowError');
const editRowCloseBtn = document.getElementById('editRowCloseBtn');
const editRowCancelBtn = document.getElementById('editRowCancelBtn');
const editRowSaveBtn = document.getElementById('editRowSaveBtn');
const editCodrCodSelect = document.getElementById('editCodrCod');
const editValueForCodrCodInput = document.getElementById('editValueForCodrCod');

const EDIT_FIELDS = [
  ['editSerialNumber', 'serialNumber'], ['editBarcode', 'barcodeNo'],
  ['editWeight', 'physicalWeight'], ['editLength', 'length'],
  ['editBreadth', 'breadthDiameter'], ['editHeight', 'height'],
  ['editReceiverName', 'receiverName'], ['editReceiverAddLine1', 'receiverAddLine1'],
  ['editReceiverAddLine2', 'receiverAddLine2'], ['editReceiverCity', 'receiverCity'],
  ['editReceiverState', 'receiverState'], ['editReceiverPincode', 'receiverPincode'],
  ['editReceiverMobileNo', 'receiverMobileNo'],
  ['editCodrCod', 'codrCod'], ['editValueForCodrCod', 'valueForCodrCod'],
];

let matchedEditEntry = null;

function clearEditFields() {
  EDIT_FIELDS.forEach(([elId]) => {
    document.getElementById(elId).value = '';
  });
  editValueForCodrCodInput.disabled = true;
}

function openEditModal() {
  editOrderIdInput.value = '';
  editRowError.textContent = '';
  clearEditFields();
  editRowSaveBtn.disabled = true;
  matchedEditEntry = null;
  editRowModal.classList.remove('hidden');
  editOrderIdInput.focus();
}

function closeEditModal() {
  editRowModal.classList.add('hidden');
}

function performEditLookup() {
  const raw = editOrderIdInput.value.trim();
  editRowError.textContent = '';
  clearEditFields();
  editRowSaveBtn.disabled = true;
  matchedEditEntry = null;

  if (!raw) {
    editRowError.textContent = 'Please enter an Order ID.';
    return;
  }

  const stripped = raw.replace(/^#/, '').toUpperCase();
  const found = manifestEntries.find((e) => {
    const entryId = (e.orderName || '').replace(/^#/, '').toUpperCase();
    return entryId === stripped;
  });

  if (!found) {
    editRowError.textContent = 'Order ID not found in Active Manifest.';
    return;
  }

  matchedEditEntry = found;
  EDIT_FIELDS.forEach(([elId, key]) => {
    const val = found[key];
    document.getElementById(elId).value = val === undefined || val === null ? '' : val;
  });
  editValueForCodrCodInput.disabled = editCodrCodSelect.value !== 'COD';
  editRowSaveBtn.disabled = false;
}

editCodrCodSelect.addEventListener('change', () => {
  if (editCodrCodSelect.value === 'COD') {
    editValueForCodrCodInput.disabled = false;
  } else {
    editValueForCodrCodInput.disabled = true;
    editValueForCodrCodInput.value = '';
  }
});

function saveEditedRow() {
  if (!matchedEditEntry) return;
  editRowError.textContent = '';

  const newSerial = Number(document.getElementById('editSerialNumber').value);
  const newBarcode = document.getElementById('editBarcode').value.trim();
  const newWeight = Number(document.getElementById('editWeight').value);
  const newLength = Number(document.getElementById('editLength').value);
  const newBreadth = Number(document.getElementById('editBreadth').value);
  const newHeight = Number(document.getElementById('editHeight').value);
  const newReceiverCity = cleanValue(document.getElementById('editReceiverCity').value);

  if (!newSerial || newSerial <= 0) { editRowError.textContent = 'Serial Number must be a positive number.'; return; }
  if (!newBarcode) { editRowError.textContent = 'Barcode No is required.'; return; }
  if (!newWeight || newWeight <= 0) { editRowError.textContent = 'Physical Weight is required.'; return; }
  if (!newLength || newLength < MIN_DIMENSIONS.length) { editRowError.textContent = `Length must be at least ${MIN_DIMENSIONS.length} cm.`; return; }
  if (!newBreadth || newBreadth < MIN_DIMENSIONS.breadth) { editRowError.textContent = `Breadth/Diameter must be at least ${MIN_DIMENSIONS.breadth} cm.`; return; }
  if (!newHeight || newHeight < MIN_DIMENSIONS.height) { editRowError.textContent = `Height must be at least ${MIN_DIMENSIONS.height} cm.`; return; }

  const { line1, line2 } = buildAddressLines(
    document.getElementById('editReceiverAddLine1').value,
    document.getElementById('editReceiverAddLine2').value,
    newReceiverCity
  );

  const isCOD = editCodrCodSelect.value === 'COD';

  matchedEditEntry.serialNumber = newSerial;
  matchedEditEntry.barcodeNo = newBarcode;
  matchedEditEntry.physicalWeight = newWeight;
  matchedEditEntry.length = newLength;
  matchedEditEntry.breadthDiameter = newBreadth;
  matchedEditEntry.height = newHeight;
  matchedEditEntry.receiverName = cleanValue(document.getElementById('editReceiverName').value);
  matchedEditEntry.receiverAddLine1 = line1;
  matchedEditEntry.receiverAddLine2 = line2;
  matchedEditEntry.receiverCity = newReceiverCity;
  matchedEditEntry.receiverState = cleanValue(document.getElementById('editReceiverState').value);
  matchedEditEntry.receiverPincode = document.getElementById('editReceiverPincode').value.trim();
  matchedEditEntry.receiverMobileNo = cleanPhoneNumber(document.getElementById('editReceiverMobileNo').value);
  matchedEditEntry.codrCod = isCOD ? 'COD' : '';
  matchedEditEntry.valueForCodrCod = isCOD
    ? (document.getElementById('editValueForCodrCod').value !== ''
        ? Number(Number(document.getElementById('editValueForCodrCod').value).toFixed(2))
        : '')
    : '';

  rerenderManifestTable();
  saveManifestToStorage();
  closeEditModal();
  showToast(`Row ${matchedEditEntry.serialNumber} updated`);
}

editOrderIdInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') {
    e.preventDefault();
    performEditLookup();
  }
});

editRowCloseBtn.addEventListener('click', closeEditModal);
editRowCancelBtn.addEventListener('click', closeEditModal);
editRowSaveBtn.addEventListener('click', saveEditedRow);

// ---- Add Row (manual) ----
const addRowModal = document.getElementById('addRowModal');
const addOrderIdInput = document.getElementById('addOrderId');
const addSerialNumberInput = document.getElementById('addSerialNumber');
const addRowError = document.getElementById('addRowError');
const addRowCloseBtn = document.getElementById('addRowCloseBtn');
const addRowCancelBtn = document.getElementById('addRowCancelBtn');
const addRowConfirmBtn = document.getElementById('addRowConfirmBtn');
const addCodrCodSelect = document.getElementById('addCodrCod');
const addValueForCodrCodInput = document.getElementById('addValueForCodrCod');

const ADD_TEXT_FIELDS = [
  'addBarcode', 'addWeight', 'addLength', 'addBreadth', 'addHeight',
  'addReceiverName', 'addReceiverAddLine1', 'addReceiverAddLine2', 'addReceiverCity',
  'addReceiverState', 'addReceiverPincode', 'addReceiverMobileNo',
];

function clearAddFields() {
  addOrderIdInput.value = '';
  ADD_TEXT_FIELDS.forEach((id) => { document.getElementById(id).value = ''; });
  addCodrCodSelect.value = '';
  addValueForCodrCodInput.value = '';
  addValueForCodrCodInput.disabled = true;
  addRowError.textContent = '';
}

function openAddModal() {
  clearAddFields();
  addSerialNumberInput.value = serialCounter;
  addRowModal.classList.remove('hidden');
  addOrderIdInput.focus();
}

function closeAddModal() {
  addRowModal.classList.add('hidden');
}

function normalizedOrderKey(raw) {
  return (raw || '').trim().replace(/^#/, '').toUpperCase();
}

function isDuplicateOrderName(raw) {
  const key = normalizedOrderKey(raw);
  return manifestEntries.some((e) => normalizedOrderKey(e.orderName) === key);
}

addCodrCodSelect.addEventListener('change', () => {
  if (addCodrCodSelect.value === 'COD') {
    addValueForCodrCodInput.disabled = false;
  } else {
    addValueForCodrCodInput.disabled = true;
    addValueForCodrCodInput.value = '';
  }
});

function saveAddedRow() {
  addRowError.textContent = '';

  const rawOrderId = addOrderIdInput.value.trim();
  const barcode = document.getElementById('addBarcode').value.trim();
  const weight = Number(document.getElementById('addWeight').value);
  const length = Number(document.getElementById('addLength').value);
  const breadth = Number(document.getElementById('addBreadth').value);
  const height = Number(document.getElementById('addHeight').value);

  if (!rawOrderId) { addRowError.textContent = 'Order ID is required.'; return; }
  if (isDuplicateOrderName(rawOrderId)) { addRowError.textContent = 'already this ordered id data is filled'; return; }
  if (!barcode) { addRowError.textContent = 'Barcode No is required.'; return; }
  if (!weight || weight <= 0) { addRowError.textContent = 'Physical Weight is required.'; return; }
  if (!length || length < MIN_DIMENSIONS.length) { addRowError.textContent = `Length must be at least ${MIN_DIMENSIONS.length} cm.`; return; }
  if (!breadth || breadth < MIN_DIMENSIONS.breadth) { addRowError.textContent = `Breadth/Diameter must be at least ${MIN_DIMENSIONS.breadth} cm.`; return; }
  if (!height || height < MIN_DIMENSIONS.height) { addRowError.textContent = `Height must be at least ${MIN_DIMENSIONS.height} cm.`; return; }

  const receiverCity = cleanValue(document.getElementById('addReceiverCity').value);
  const { line1, line2 } = buildAddressLines(
    document.getElementById('addReceiverAddLine1').value,
    document.getElementById('addReceiverAddLine2').value,
    receiverCity
  );

  const isCOD = addCodrCodSelect.value === 'COD';

  const row = {
    orderId: rawOrderId,
    orderName: rawOrderId,
    serialNumber: serialCounter,
    barcodeNo: barcode,
    physicalWeight: weight,
    shapeOfArticle: FIXED_VALUES.shapeOfArticle,
    length: length,
    breadthDiameter: breadth,
    height: height,
    priorityFlag: FIXED_VALUES.priorityFlag,
    deliveryInstruction: FIXED_VALUES.deliveryInstruction,
    instructionRts: FIXED_VALUES.instructionRts,
    senderName: FIXED_VALUES.senderName,
    senderCompany: FIXED_VALUES.senderCompany,
    senderAddLine1: FIXED_VALUES.senderAddLine1,
    senderAddLine2: FIXED_VALUES.senderAddLine2,
    senderCity: FIXED_VALUES.senderCity,
    senderState: FIXED_VALUES.senderState,
    senderPincode: FIXED_VALUES.senderPincode,
    senderEmailId: FIXED_VALUES.senderEmailId,
    senderAltContact: '', senderKyc: '', senderTaxReference: '',
    receiverName: cleanValue(document.getElementById('addReceiverName').value),
    receiverCompany: '',
    receiverAddLine1: line1,
    receiverAddLine2: line2,
    receiverCity: receiverCity,
    receiverState: cleanValue(document.getElementById('addReceiverState').value),
    receiverPincode: document.getElementById('addReceiverPincode').value.trim(),
    receiverEmailId: '',
    receiverAltContact: '', receiverKyc: '', receiverTaxReference: '',
    altAddressFlag: FIXED_VALUES.altAddressFlag,
    pickupAddressFlag: FIXED_VALUES.pickupAddressFlag,
    dropOffPincode: FIXED_VALUES.dropOffPincode,
    dropoffPickupOfficeId: FIXED_VALUES.dropoffPickupOfficeId,
    senderMobileNo: FIXED_VALUES.senderMobileNo,
    receiverMobileNo: cleanPhoneNumber(document.getElementById('addReceiverMobileNo').value),
    prepaymentCode: '', valueOfPrepayment: '',
    codrCod: isCOD ? 'COD' : '',
    valueForCodrCod: isCOD
      ? (addValueForCodrCodInput.value !== '' ? Number(Number(addValueForCodrCodInput.value).toFixed(2)) : '')
      : '',
    insuranceType: '', valueOfInsurance: '',
    ack: FIXED_VALUES.ack,
    registration: FIXED_VALUES.registration,
    otpBasedDelivery: FIXED_VALUES.otpBasedDelivery,
    bulkReference: '',
  };

  manifestEntries.push(row);
  renderManifestRow(row);
  serialCounter++;
  saveManifestToStorage();
  closeAddModal();
  showToast(`Row ${row.serialNumber} added`);
}

addOrderIdInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') e.preventDefault();
});

addRowCloseBtn.addEventListener('click', closeAddModal);
addRowCancelBtn.addEventListener('click', closeAddModal);
addRowConfirmBtn.addEventListener('click', saveAddedRow);

const copyPincodeBtn = document.getElementById('copyPincodeBtn');
const copiedNotice = document.getElementById('copiedNotice');


copyPincodeBtn.addEventListener('click', async () => {
  const value = pincodeInput.value.trim();
  if (!value) return; // nothing to copy

  try {
    await navigator.clipboard.writeText(value);
    copiedNotice.classList.add('show');
    setTimeout(() => copiedNotice.classList.remove('show'), 1500);
  } catch (err) {
    console.error('Clipboard copy failed:', err);
    formError.textContent = 'Could not copy — please copy manually.';
  }
});

// Stop Chrome offering to save/fill addresses on these fields
document.querySelectorAll('input, select').forEach((el) => {
  el.setAttribute('autocomplete', 'off');
});


// Stop Chrome offering to save/fill addresses on these fields
document.querySelectorAll('input, select').forEach((el) => {
  el.setAttribute('autocomplete', 'off');
});

// ---- Supabase Edit: change the barcode ("Transaction ID") of a saved india_post row ----
// Uses 2 Supabase functions (run supabase_edit_barcode.sql once):
//   get_india_post_barcode(p_order_id)                 -> current barcode
//   update_india_post_barcode(p_order_id, p_new_barcode) -> rows updated
const supaEditModal = document.getElementById('supaEditModal');
const supaOrderIdInput = document.getElementById('supaOrderId');
const supaCurrentBarcodeInput = document.getElementById('supaCurrentBarcode');
const supaNewBarcodeInput = document.getElementById('supaNewBarcode');
const supaEditError = document.getElementById('supaEditError');
const supaEditSaveBtn = document.getElementById('supaEditSaveBtn');

let supaMatchedOrderId = null;

function resetSupaEditFields() {
  supaCurrentBarcodeInput.value = '';
  supaNewBarcodeInput.value = '';
  supaNewBarcodeInput.disabled = true;
  supaEditSaveBtn.disabled = true;
  supaEditError.textContent = '';
  supaMatchedOrderId = null;
}

function openSupaEditModal() {
  supaOrderIdInput.value = '';
  resetSupaEditFields();
  supaEditModal.classList.remove('hidden');
  supaOrderIdInput.focus();
}

function closeSupaEditModal() {
  supaEditModal.classList.add('hidden');
}

async function performSupaLookup() {
  const raw = supaOrderIdInput.value.trim();
  resetSupaEditFields();
  if (!raw) {
    supaEditError.textContent = 'Please enter an Order ID.';
    return;
  }

  const { data, error } = await supabaseClient.rpc('get_india_post_barcode', { p_order_id: raw });
  if (error) {
    console.error('Supabase Edit lookup error:', error);
    supaEditError.textContent = `Lookup failed: ${error.message || 'check your connection and try again.'}`;
    return;
  }

  const row = Array.isArray(data) ? data[0] : data;
  if (!row) {
    supaEditError.textContent = 'Order ID not found in Supabase (india_post).';
    return;
  }

  supaMatchedOrderId = row.order_id;
  supaCurrentBarcodeInput.value = row.transaction_id || '';
  supaNewBarcodeInput.disabled = false;
  supaEditSaveBtn.disabled = false;
  supaNewBarcodeInput.focus();
}

async function saveSupaBarcode() {
  if (!supaMatchedOrderId) return;
  supaEditError.textContent = '';

  const newBarcode = supaNewBarcodeInput.value.trim();
  if (!newBarcode) {
    supaEditError.textContent = 'New Barcode is required.';
    return;
  }
  if (newBarcode === supaCurrentBarcodeInput.value.trim()) {
    supaEditError.textContent = 'New Barcode is the same as the current one.';
    return;
  }

  supaEditSaveBtn.disabled = true;
  const { data, error } = await supabaseClient.rpc('update_india_post_barcode', {
    p_order_id: supaMatchedOrderId,
    p_new_barcode: newBarcode,
  });
  supaEditSaveBtn.disabled = false;

  if (error) {
    console.error('Supabase Edit save error:', error);
    supaEditError.textContent = error.message || 'Could not save to Supabase. Please try again.';
    return;
  }
  if (!Number(data)) {
    supaEditError.textContent = 'Nothing was updated — the order was not found.';
    return;
  }

  closeSupaEditModal();
  showToast(`Barcode updated in Supabase for ${supaMatchedOrderId}`);
}

supaOrderIdInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') {
    e.preventDefault();
    performSupaLookup();
  }
});
supaOrderIdInput.addEventListener('input', () => {
  if (supaMatchedOrderId) resetSupaEditFields(); // Order ID changed: clear stale result
});
supaNewBarcodeInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') {
    e.preventDefault();
    saveSupaBarcode();
  }
});
document.getElementById('supaEditCloseBtn').addEventListener('click', closeSupaEditModal);
document.getElementById('supaEditCancelBtn').addEventListener('click', closeSupaEditModal);
supaEditSaveBtn.addEventListener('click', saveSupaBarcode);

// ---- UI helpers: toasts, today's date, row count, popup closing ----
function showToast(message, type = 'ok') {
  const stack = document.getElementById('toastStack');
  if (!stack) return;
  const icon = type === 'ok'
    ? '<path d="M20 6 9 17l-5-5"/>'
    : '<path d="M18 6 6 18M6 6l12 12"/>';
  const toast = document.createElement('div');
  toast.className = 'toast';
  toast.innerHTML =
    `<span class="toast-dot ${type === 'ok' ? 'ok' : 'err'}">` +
    `<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="3" stroke-linecap="round" stroke-linejoin="round">${icon}</svg>` +
    `</span><span></span>`;
  toast.lastChild.textContent = message;
  stack.appendChild(toast);
  setTimeout(() => {
    toast.classList.add('out');
    setTimeout(() => toast.remove(), 260);
  }, 2600);
}

(function setTodayPill() {
  const pill = document.getElementById('todayPill');
  if (!pill) return;
  pill.textContent = new Date().toLocaleDateString('en-IN', {
    weekday: 'short', day: 'numeric', month: 'short', year: 'numeric',
  });
})();

function updateRowCount() {
  const n = manifestBody.children.length;
  const label = `${n} ${n === 1 ? 'row' : 'rows'}`;
  const pill = document.getElementById('rowCountPill');
  const footer = document.getElementById('rowCountFooter');
  if (pill) pill.textContent = label;
  if (footer) footer.textContent = `Showing ${label}`;
}
new MutationObserver(updateRowCount).observe(manifestBody, { childList: true });

// Close popups with Esc or a click on the dark background (top-most popup first)
const POPUPS = [
  [deleteAllModal, closeDeleteAllModal],
  [deleteRowModal, closeDeleteModal],
  [editRowModal, closeEditModal],
  [addRowModal, closeAddModal],
  [supaEditModal, closeSupaEditModal],
];
POPUPS.forEach(([modalEl, closeFn]) => {
  modalEl.addEventListener('mousedown', (e) => {
    if (e.target === modalEl) closeFn();
  });
});
document.addEventListener('keydown', (e) => {
  if (e.key !== 'Escape') return;
  const open = POPUPS.find(([modalEl]) => !modalEl.classList.contains('hidden'));
  if (open) {
    e.preventDefault();
    open[1]();
  }
});

// ---- Action menu (styled dropdown that drives the hidden #actionSelect) ----
const actionMenuBtn = document.getElementById('actionMenuBtn');
const actionMenu = document.getElementById('actionMenu');
const actionMenuItems = Array.from(actionMenu.querySelectorAll('[data-action]'));

function openActionMenu() {
  actionMenu.classList.remove('hidden');
  actionMenuBtn.setAttribute('aria-expanded', 'true');
  actionMenuItems[0].focus();
}

function closeActionMenu() {
  actionMenu.classList.add('hidden');
  actionMenuBtn.setAttribute('aria-expanded', 'false');
}

actionMenuBtn.addEventListener('click', () => {
  if (actionMenu.classList.contains('hidden')) openActionMenu();
  else closeActionMenu();
});

actionMenuItems.forEach((item, i) => {
  item.addEventListener('click', () => {
    closeActionMenu();
    // Same path as before: set the select's value and fire its change handler
    actionSelect.value = item.dataset.action;
    actionSelect.dispatchEvent(new Event('change'));
  });
  item.addEventListener('keydown', (e) => {
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault();
      const step = e.key === 'ArrowDown' ? 1 : -1;
      actionMenuItems[(i + step + actionMenuItems.length) % actionMenuItems.length].focus();
    } else if (e.key === 'Escape') {
      e.stopPropagation();
      closeActionMenu();
      actionMenuBtn.focus();
    }
  });
});

document.addEventListener('mousedown', (e) => {
  if (!actionMenu.contains(e.target) && !actionMenuBtn.contains(e.target)) closeActionMenu();
});

// ---- Left/Right arrow keys move between entry fields ----
const ARROW_NAV_FIELDS = [orderIdInput, barcodeInput, weightInput, lengthInput, breadthInput, heightInput, submitBtn];

function caretAtEdge(el, direction) {
  // Number inputs and buttons have no caret position, so always allow the jump
  let start;
  let end;
  try {
    start = el.selectionStart;
    end = el.selectionEnd;
  } catch (e) {
    return true;
  }
  if (start === null || start === undefined) return true;
  if (start !== end) return false; // text is selected: let the arrow act normally
  return direction === 'right' ? end === el.value.length : start === 0;
}

function placeCaret(el, direction) {
  // Arriving with Right -> caret at end; arriving with Left -> caret at start
  try {
    const pos = direction === 'right' ? el.value.length : 0;
    el.setSelectionRange(pos, pos);
  } catch (e) {
    // number inputs and buttons don't support this; ignore
  }
}

ARROW_NAV_FIELDS.forEach((el, i) => {
  el.addEventListener('keydown', (e) => {
    if (e.key !== 'ArrowRight' && e.key !== 'ArrowLeft') return;
    if (e.shiftKey || e.ctrlKey || e.altKey || e.metaKey) return;

    const direction = e.key === 'ArrowRight' ? 'right' : 'left';
    if (!caretAtEdge(el, direction)) return;

    const next = ARROW_NAV_FIELDS[i + (direction === 'right' ? 1 : -1)];
    if (!next) return;

    e.preventDefault();

    // Leaving Order ID with Right: run the lookup if it hasn't been done yet
    if (el === orderIdInput && direction === 'right' && !currentOrder && orderIdInput.value.trim()) {
      performLookup();
    }

    next.focus();
    placeCaret(next, direction);
  });
});

loadManifestFromStorage();

// Remove leftover flag from the old reset popup (no longer used)
localStorage.removeItem('courierManifestPendingReset');