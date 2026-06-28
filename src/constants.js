export const BUDGET_COLUMNS = [
  'Serial Number',
  'Fiscal Year',
  'Department',
  'Category',
  'Budget Amount',
];


export const ASSET_UPLOAD_COLUMNS = [
  'Serial Number',
  'Asset Name',
  'Asset Description',
  'Fiscal Year',
  'Department',
  'Asset Category',
  'Quantity',
];


export const ASSET_DETAIL_COLUMNS = [
  'Serial Number',
  'Asset Name',
  'Asset Description',
  'Fiscal Year',
  'Department',
  'Asset Category',
  'Quantity',
  'Amount',
];


export const ASSET_APPROVAL_STEPS = [
  'Requester',
  'Procurement',
  'Head of Department',
  'Director HR/Admin',
  'Finance Head',
  'VP Strategy',
  'Finance Accountant',
];


export const DEPARTMENTS = ['Admin', 'Strategy', 'IT'];

export const ASSET_CATEGORIES = [
  'Leasehold improvements',
  'Furniture and office equipment',
  'Motor vehicles',
  'Tools and equipment',
];


export const BUDGET_MATRIX = {
  Admin: {
    'Leasehold improvements': 1500000,
    'Furniture and office equipment': 1000000,
    'Motor vehicles': 2000000,
    'Tools and equipment': 800000,
  },
  Strategy: {
    'Leasehold improvements': 1200000,
    'Furniture and office equipment': 900000,
    'Motor vehicles': 1800000,
    'Tools and equipment': 700000,
  },
  IT: {
    'Leasehold improvements': 1000000,
    'Furniture and office equipment': 850000,
    'Motor vehicles': 1500000,
    'Tools and equipment': 1200000,
  },
};


export const REQUEST_HISTORY = [
  {
    id: 'REQ-1001',
    shortDescription: 'Office chair purchase',
    requestedDate: '2026-05-22',
    status: 'Pending for HOD approval',
    progress: 1,
    mode: 'active',
  },
  {
    id: 'REQ-1002',
    shortDescription: 'Printer toner supply',
    requestedDate: '2026-05-28',
    status: 'Pending for Director HR/Admin approval',
    progress: 2,
    mode: 'active',
  },
  {
    id: 'REQ-1003',
    shortDescription: 'Laptop upgrade batch',
    requestedDate: '2026-06-01',
    status: 'Pending for Finance head approval',
    progress: 3,
    mode: 'active',
  },
  {
    id: 'REQ-1004',
    shortDescription: 'Security camera replacement',
    requestedDate: '2026-06-04',
    status: 'Pending for VP Strategy approval',
    progress: 4,
    mode: 'active',
  },
  {
    id: 'REQ-1005',
    shortDescription: 'Reception desk renovation',
    requestedDate: '2026-06-07',
    status: 'Declined',
    progress: 5,
    mode: 'declined',
  },
  {
    id: 'REQ-1006',
    shortDescription: 'Vehicle tracker devices',
    requestedDate: '2026-06-10',
    status: 'Complete',
    progress: 6,
    mode: 'complete',
  },
];


export const BANK_ACCOUNTS = [
  { sNo: 1, houseBank: 'RJHI', companyCode: '2000', accountNumber: '206608010084566', accountId: 'S4566', availableAmount: 1250000 },
  { sNo: 2, houseBank: 'RJHI', companyCode: '1000', accountNumber: '206608010084574', accountId: 'S4574', availableAmount: 845000.5 },
  { sNo: 3, houseBank: 'SAIB', companyCode: '2000', accountNumber: '202227629001', accountId: 'S9001', availableAmount: 320000 },
  { sNo: 4, houseBank: 'SAIB', companyCode: '1000', accountNumber: '202227637001', accountId: 'S7001', availableAmount: 1785400.75 },
  { sNo: 5, houseBank: 'BSF', companyCode: '2000', accountNumber: '68745800133', accountId: 'S0133', availableAmount: 56000 },
  { sNo: 6, houseBank: 'BSF', companyCode: '1000', accountNumber: '68746400175', accountId: 'S0175', availableAmount: 990250 },
  { sNo: 7, houseBank: 'BSF', companyCode: '4000', accountNumber: 'R0492000137', accountId: 'S0137', availableAmount: 2400000 },
  { sNo: 8, houseBank: 'BSF', companyCode: '1000', accountNumber: '68746402212', accountId: 'S2212', availableAmount: 132500.25 },
  { sNo: 9, houseBank: 'BSF', companyCode: '1000', accountNumber: '68746402115', accountId: 'S2115', availableAmount: 78000 },
  { sNo: 10, houseBank: 'ANB', companyCode: '2000', accountNumber: '107007496040017', accountId: 'S4017', availableAmount: 3150000 },
  { sNo: 11, houseBank: 'ANB', companyCode: '1000', accountNumber: '107005383900017', accountId: 'S0017', availableAmount: 612000 },
  { sNo: 12, houseBank: 'RYAD', companyCode: '2000', accountNumber: '1021548729901', accountId: 'S2901', availableAmount: 1045000.5 },
  { sNo: 13, houseBank: 'RYAD', companyCode: '1000', accountNumber: '1021548059901', accountId: 'S5901', availableAmount: 287650 },
  { sNo: 14, houseBank: 'RYAD', companyCode: '1000', accountNumber: '1021548059942', accountId: 'S5942', availableAmount: 1500000 },
  { sNo: 15, houseBank: 'INMA', companyCode: '2000', accountNumber: '68200297140000', accountId: 'S0000', availableAmount: 92000 },
  { sNo: 16, houseBank: 'INMA', companyCode: '1000', accountNumber: '68200300004000', accountId: 'S4000', availableAmount: 740300.75 },
  { sNo: 17, houseBank: 'JZRH', companyCode: '2000', accountNumber: '7216300001', accountId: 'S0001', availableAmount: 2100000 },
  { sNo: 18, houseBank: 'JZRH', companyCode: '1000', accountNumber: '7216333001', accountId: 'S3001', availableAmount: 415000 },
  { sNo: 19, houseBank: 'JZRH', companyCode: '1000', accountNumber: '84295022350001', accountId: 'S5001', availableAmount: 68500 },
  { sNo: 20, houseBank: 'JZRH', companyCode: '1000', accountNumber: '84295022350002', accountId: 'S0002', availableAmount: 1990000 },
];

// Comma-separated money formatter for read-only amounts (e.g., 1,250,000.00).

export const TRANSFER_TO_BANKS = {
  RJHI: ['SA 16 8000 0106 6080 1684 6450', 'SA 16 8000 0106 6080 1684 6451'],
  SAB: ['SA 03 8000 0000 1234 5678 9012', 'SA 03 8000 0000 1234 5678 9013', 'SA 03 8000 0000 1234 5678 9014'],
  BSF: ['SA56 5500 0000 0132 0220 0520', 'SA56 5500 0000 0132 0220 0521', 'SA56 5500 0000 0132 0220 0522', 'SA56 5500 0000 0132 0220 0523'],
};


export const BANK_NAMES = ['RJHI', 'SAB', 'BSF'];


// Hardcoded cost centers for the Finance Accountant step (placeholder set).
export const COST_CENTERS = [
  'CC-1001 — Administration',
  'CC-1002 — Information Technology',
  'CC-1003 — Operations',
  'CC-1004 — Finance & Accounts',
  'CC-1005 — Strategy & Planning',
];

// Hardcoded company codes for the Finance Accountant step (placeholder set).
export const COMPANY_CODES = [
  '1000 — Ibrahim Shaker Co. Ltd.',
  '2000 — Shaker Trading',
  '3000 — Shaker Services',
];
