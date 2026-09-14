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
// ---- Submit validation ----
function validateManualFields() {
  if (!currentOrder) return 'Look up a valid Order ID first.';
  if (!barcodeInput.value.trim()) return 'Barcode No is required.';
  if (!weightInput.value || Number(weightInput.value) <= 0) return 'Physical Weight is required.';
  if (!lengthInput.value || Number(lengthInput.value) <= 0) return 'Length is required.';
  if (!breadthInput.value || Number(breadthInput.value) <= 0) return 'Breadth/Diameter is required.';
  if (!heightInput.value || Number(heightInput.value) <= 0) return 'Height is required.';
  return null;
}

// ---- Row builder ----
function buildManifestRow() {
  const order = currentOrder;
  const isCOD = (order.invoice_number || '').trim().toUpperCase().startsWith('COD');
  const { line1, line2 } = buildAddressLines(order.shipping_address1, order.shipping_address2, order.shipping_city);

  return {
    orderId: order.shopify_order_id,
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
    receiverMobileNo: order.customer_phone || '',
    prepaymentCode: '', valueOfPrepayment: '',
    codrCod: isCOD ? 'COD' : '',
    valueForCodrCod: isCOD ? Number(order.amount_to_receive.toFixed(2)) : '',
    insuranceType: '', valueOfInsurance: '',
    ack: FIXED_VALUES.ack,
    registration: FIXED_VALUES.registration,
    otpBasedDelivery: FIXED_VALUES.otpBasedDelivery,
    bulkReference: '',
  };
}

// ---- Duplicate check, rendering, clearing ----
function isDuplicate(orderId) {
  return manifestEntries.some((e) => e.orderId === orderId);
}

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
}

// ---- Submit Entry ----
submitBtn.addEventListener('click', () => {
  formError.textContent = '';
  const validationError = validateManualFields();
  if (validationError) {
    formError.textContent = validationError;
    return;
  }

  const orderId = currentOrder.shopify_order_id;
  if (isDuplicate(orderId)) {
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
});
const resetSelect = document.getElementById('resetSelect');

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

async function downloadExcel() {
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
  } catch (err) {
    console.error('Excel generation failed:', err);
    formError.textContent = 'Failed to generate Excel file. Please try again.';
  }
}

downloadBtn.addEventListener('click', downloadExcel);

resetSelect.addEventListener('change', () => {
  if (resetSelect.value === 'reset') {
    const confirmed = confirm('This will clear all entries in the Active Manifest. This cannot be undone. Continue?');
    if (confirmed) {
      resetManifest();
    }
    resetSelect.value = ''; // always return dropdown to placeholder
  }
});

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

loadManifestFromStorage();