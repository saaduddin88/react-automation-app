import { SADAD_SCREENS } from './sadadData';

export const IQAMA_NUMBER_LENGTH = 10; // Standard Saudi Iqama length — change here if your file differs.

export const IQAMA_EXPENSE_TYPES = [
  'تجديد اقامة ورخصة عمل',
  'خروج وعودة',
  'رخصة عمل',
  'نقل كفالة',
  'تجديد اقامة',
];

export const IQAMA_COLUMNS = ['Iqama Number', 'Amount', 'Expense Type', 'SAP Emp No.', 'Beneficiary', 'Own Account'];


export const iqamaRowActive = (it) =>
  [it.iqama, it.amount, it.sapNo].some((v) => String(v ?? '').trim() !== '');


export function validateIqamaItem(it) {
  const e = {};
  const active = iqamaRowActive(it);
  const iqama = String(it.iqama ?? '').trim();
  const amountRaw = String(it.amount ?? '').trim();
  const sapNo = String(it.sapNo ?? '').trim();
  const etype = String(it.expenseType ?? '').trim();
  const ownRaw = String(it.ownAccount ?? '').trim();

  if (active) {
    if (!iqama) e.iqama = 'Iqama Number is required.';
    else if (!/^\d+$/.test(iqama)) e.iqama = 'Digits only — no letters, spaces or symbols.';
    else if (iqama.length !== IQAMA_NUMBER_LENGTH) e.iqama = `Must be exactly ${IQAMA_NUMBER_LENGTH} digits.`;

    if (!amountRaw) e.amount = 'Amount is required.';
    else if (Number.isNaN(Number(amountRaw.replace(/,/g, '')))) e.amount = 'Amount must be a number.';

    if (!sapNo) e.sapNo = 'SAP Emp No. is required.';
    else if (!/^\d+$/.test(sapNo)) e.sapNo = 'Digits only — no decimals or symbols.';
    else if (sapNo.length > 5) e.sapNo = 'Up to 5 digits only.';

    if (!etype) e.expenseType = 'Expense Type is required.';
    else if (!IQAMA_EXPENSE_TYPES.includes(etype)) e.expenseType = 'Pick a value from the list.';
  } else if (etype && !IQAMA_EXPENSE_TYPES.includes(etype)) {
    e.expenseType = 'Pick a value from the list.';
  }

  // Own Account is optional, but if entered it must be a number and not exceed Amount.
  if (ownRaw) {
    if (Number.isNaN(Number(ownRaw.replace(/,/g, '')))) e.ownAccount = 'Must be a number.';
    else if (
      amountRaw !== '' &&
      !Number.isNaN(Number(amountRaw.replace(/,/g, ''))) &&
      Number(ownRaw.replace(/,/g, '')) > Number(amountRaw.replace(/,/g, ''))
    )
      e.ownAccount = 'Cannot exceed Amount.';
  }

  return e;
}


export const PAYMENT_COLUMNS = [
  'Bank reference',
  'Narrative',
  'Customer reference',
  'TRN type',
  'Value date',
  'Credit amount',
  'Debit amount',
  'Balance',
  'Time',
  'Post date',
];


export const isNumericValue = (s) => {
  const r = String(s).replace(/,/g, '').trim();
  return r !== '' && !Number.isNaN(Number(r));
};

export const isDigitsOnly = (s) => /^\d+$/.test(String(s).trim());

export const TIME_RE = /^([01]\d|2[0-3]):[0-5]\d$/;

export const DATE_RE = /^(\d{2})\/(\d{2})\/(\d{4})$/;

export const isValidDDMMYYYY = (s) => {
  const m = String(s).trim().match(DATE_RE);
  if (!m) return false;
  const d = +m[1];
  const mo = +m[2];
  const y = +m[3];
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return false;
  const dt = new Date(y, mo - 1, d);
  return dt.getFullYear() === y && dt.getMonth() === mo - 1 && dt.getDate() === d;
};

export const formatDateDDMMYYYY = (dt) => {
  const d = String(dt.getDate()).padStart(2, '0');
  const m = String(dt.getMonth() + 1).padStart(2, '0');
  return `${d}/${m}/${dt.getFullYear()}`;
};


export const paymentRowActive = (it) =>
  [it.bankRef, it.customerRef, it.valueDate, it.debitAmount, it.narrative].some(
    (v) => String(v ?? '').trim() !== ''
  );


export function validatePaymentItem(it) {
  const e = {};
  const active = paymentRowActive(it);
  const bankRef = String(it.bankRef ?? '').trim();
  const narrative = String(it.narrative ?? '').trim();
  const customerRef = String(it.customerRef ?? '').trim();
  const valueDate = String(it.valueDate ?? '').trim();
  const debit = String(it.debitAmount ?? '').trim();
  const credit = String(it.creditAmount ?? '').trim();
  const balance = String(it.balance ?? '').trim();
  const time = String(it.time ?? '').trim();
  const postDate = String(it.postDate ?? '').trim();

  if (active) {
    if (!bankRef) e.bankRef = 'Bank reference is required.';

    if (!narrative) e.narrative = 'Narrative is required.';

    if (!customerRef) e.customerRef = 'Customer reference is required.';
    else if (!isDigitsOnly(customerRef)) e.customerRef = 'Digits only.';

    if (!valueDate) e.valueDate = 'Value date is required.';
    else if (!isValidDDMMYYYY(valueDate)) e.valueDate = 'Use DD/MM/YYYY.';

    if (!debit) e.debitAmount = 'Debit amount is required.';
    else if (!isNumericValue(debit)) e.debitAmount = 'Must be a number.';
  }

  if (credit && !isNumericValue(credit)) e.creditAmount = 'Must be a number.';
  if (balance && !isNumericValue(balance)) e.balance = 'Must be a number.';
  if (time && !TIME_RE.test(time)) e.time = 'Use HH:MM (24h).';
  if (postDate && !isValidDDMMYYYY(postDate)) e.postDate = 'Use DD/MM/YYYY.';

  return e;
}

// 3-digit SADAD biller code -> screen key (e.g. '002' -> 'sec').

export const CODE_TO_KEY = SADAD_SCREENS.reduce((acc, s) => {
  if (s.code) acc[s.code] = s.key;
  return acc;
}, {});

// Identify which SADAD screen/category a raw invoice (bank statement) line belongs to.
// Returns a screen key (matching SADAD_SCREENS[].key) or null if it isn't a SADAD payment.

export function classifyInvoiceCategory(it) {
  const trn = String(it.trnType ?? '').trim();
  const nar = String(it.narrative ?? '');
  if (trn === 'TFR+') return null; // credit / refund / internal transfer — not a payment
  if (/BLR:094/.test(nar)) return 'vehicle-renewal';
  if (/BLR:090/.test(nar)) return 'iqama';
  if (nar.trim().toUpperCase().startsWith('NONREF')) return 'esal';
  const toks = nar.trim().split(/\s+/);
  const code = toks.length > 1 && /^\d{3}$/.test(toks[1]) ? toks[1] : null;
  if (code === '050') {
    // 050 is shared: 16-digit ref = traffic violation, else (15-digit permit no.) = permits.
    const ref = toks[0] || '';
    return ref.length >= 16 ? 'violations' : 'iqama';
  }
  return code ? CODE_TO_KEY[code] || null : null;
}

