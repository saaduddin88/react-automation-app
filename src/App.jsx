import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Link, NavLink, Route, Routes, useLocation, useNavigate, useParams } from 'react-router-dom';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import {
  AlertTriangle,
  ArrowRight,
  Check,
  ChevronDown,
  CreditCard,
  Download,
  Edit3,
  FileSpreadsheet,
  LayoutDashboard,
  MapPinned,
  Search,
  ShieldCheck,
  Trash2,
  Upload,
} from 'lucide-react';
import { SADAD_GROUPS, SADAD_SCREENS, findScreen } from './sadadData';

const BUDGET_COLUMNS = [
  'Serial Number',
  'Fiscal Year',
  'Department',
  'Category',
  'Budget Amount',
];

const ASSET_UPLOAD_COLUMNS = [
  'Serial Number',
  'Asset Name',
  'Asset Description',
  'Fiscal Year',
  'Department',
  'Asset Category',
  'Quantity',
];

const ASSET_DETAIL_COLUMNS = [
  'Serial Number',
  'Asset Name',
  'Asset Description',
  'Fiscal Year',
  'Department',
  'Asset Category',
  'Quantity',
  'Amount',
];

const ASSET_APPROVAL_STEPS = [
  'Requester',
  'Procurement',
  'Head of Department',
  'Director HR/Admin',
  'Finance Head',
  'VP Strategy',
  'Finance Accountant',
];

const DEPARTMENTS = ['Admin', 'Strategy', 'IT'];
const ASSET_CATEGORIES = [
  'Leasehold improvements',
  'Furniture and office equipment',
  'Motor vehicles',
  'Tools and equipment',
];

const BUDGET_MATRIX = {
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

const REQUEST_HISTORY = [
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

const sampleLandingCards = [
  {
    title: 'Upload Budget',
    text: 'Download the budget template, upload budget lines and edit them any time.',
    to: '/budget-upload',
    icon: FileSpreadsheet,
  },
  {
    title: 'Asset Requisition',
    text: 'Download the template, upload requisition lines and submit your request.',
    to: '/asset-requisition',
    icon: FileSpreadsheet,
  },
  {
    title: 'Requisition Tracking',
    text: 'Open request history and inspect each approval timeline.',
    to: '/requisition-tracking',
    icon: ShieldCheck,
  },
  {
    title: 'Bank Transfers',
    text: 'Transfer funds between bank accounts and post to SAP.',
    to: '/bank-transfers',
    icon: FileSpreadsheet,
  },
  {
    title: 'SADAD Payments',
    text: 'Pay government, utility, telecom and commercial bills through one screen.',
    to: '/sadad-payments',
    icon: CreditCard,
  },
  {
    title: 'Reports',
    text: 'Reserved for budget utilization reports and category-wise analysis.',
    to: '/report-creation',
    icon: LayoutDashboard,
  },
];

const getCurrentFiscalYear = () => String(new Date().getFullYear());

// Local-date 'YYYY-MM-DD' (timezone-safe, so it doesn't roll over a day near midnight).
const getTodayISO = () => {
  const now = new Date();
  const local = new Date(now.getTime() - now.getTimezoneOffset() * 60000);
  return local.toISOString().split('T')[0];
};

const formatAmount = (value) => {
  const raw = String(value ?? '').replace(/,/g, '').trim();
  const parsed = parseFloat(raw);
  return Number.isNaN(parsed) ? '0' : parsed.toLocaleString();
};

const parseAmount = (value) => {
  const raw = String(value ?? '').replace(/,/g, '').trim();
  const parsed = parseFloat(raw);
  return Number.isNaN(parsed) ? 0 : parsed;
};

const getBudgetRequestTotal = (rows = []) => rows.reduce((sum, row) => sum + parseAmount(row['Budget Amount']), 0);

const summarizeField = (rows = [], field) => {
  return [...new Set(rows.map((row) => String(row[field] ?? '').trim()).filter(Boolean))].join(', ');
};

const normalizeExcelRow = (item) => {
  const normalized = {};
  Object.keys(item || {}).forEach((key) => {
    normalized[String(key).trim().toLowerCase()] = item[key];
  });
  return normalized;
};

const getModeForStatus = (status) => {
  if (status === 'Complete') return 'complete';
  if (status === 'Declined') return 'declined';
  return 'active';
};

const getProgressForStatus = (status) => {
  if (status === 'Pending for Procurement') return 2;
  if (status === 'Pending for HOD approval') return 3;
  if (status === 'Pending for Director HR/Admin approval') return 4;
  if (status === 'Pending for Finance head approval') return 5;
  if (status === 'Pending for VP Strategy approval') return 6;
  if (status === 'Complete') return 7;
  if (status === 'Declined') return 7;
  return 2;
};

const buildAssetApprovalTrail = (requestLike, status, actedAt = new Date().toLocaleString()) => {
  const request = requestLike || {};
  const trail = [
    {
      step: 'Requester',
      actor: 'Requester',
      action: 'Submitted',
      actedAt: request.submittedAt || actedAt,
    },
    {
      step: 'Procurement',
      actor: 'Procurement',
      action: 'Pending',
      actedAt: null,
    },
    {
      step: 'Head of Department',
      actor: 'Head of Department',
      action: 'Pending',
      actedAt: null,
    },
    {
      step: 'Director HR/Admin',
      actor: 'Director HR/Admin',
      action: 'Pending',
      actedAt: null,
    },
    {
      step: 'Finance Head',
      actor: 'Finance Head',
      action: 'Pending',
      actedAt: null,
    },
    {
      step: 'VP Strategy',
      actor: 'VP Strategy',
      action: 'Pending',
      actedAt: null,
    },
    {
      step: 'Finance Accountant',
      actor: 'Finance Accountant',
      action: 'Pending',
      actedAt: null,
    },
  ];

  const setStep = (stepName, action, timeValue = actedAt) => {
    const step = trail.find((item) => item.step === stepName);
    if (!step) return;
    step.action = action;
    step.actedAt = action === 'Pending' ? null : timeValue;
  };

  if (status === 'Pending for Procurement') return trail;

  if (status === 'Pending for HOD approval') {
    setStep('Procurement', 'Completed', request.procurementSavedAt || actedAt);
    return trail;
  }

  if (status === 'Pending for Director HR/Admin approval') {
    setStep('Procurement', 'Completed', request.procurementSavedAt || actedAt);
    setStep('Head of Department', 'Approved', request.hodApprovedAt || actedAt);
    return trail;
  }

  if (status === 'Pending for Finance head approval') {
    setStep('Procurement', 'Completed', request.procurementSavedAt || actedAt);
    setStep('Head of Department', 'Approved', request.hodApprovedAt || actedAt);
    setStep('Director HR/Admin', 'Approved', request.directorApprovedAt || actedAt);
    return trail;
  }

  if (status === 'Pending for VP Strategy approval') {
    setStep('Procurement', 'Completed', request.procurementSavedAt || actedAt);
    setStep('Head of Department', 'Approved', request.hodApprovedAt || actedAt);
    setStep('Director HR/Admin', 'Approved', request.directorApprovedAt || actedAt);
    setStep('Finance Head', 'Approved', request.financeHeadApprovedAt || actedAt);
    return trail;
  }

  if (status === 'Complete') {
    setStep('Procurement', 'Completed', request.procurementSavedAt || actedAt);
    setStep('Head of Department', 'Approved', request.hodApprovedAt || actedAt);
    setStep('Director HR/Admin', 'Approved', request.directorApprovedAt || actedAt);
    setStep('Finance Head', 'Approved', request.financeHeadApprovedAt || actedAt);
    setStep('VP Strategy', 'Approved', request.vpStrategyApprovedAt || actedAt);
    setStep('Finance Accountant', 'Completed', request.financeAccountantCompletedAt || actedAt);
    return trail;
  }

  if (status === 'Declined') {
    if (request.declinedAtStep === 'Procurement') {
      setStep('Procurement', 'Rejected', request.declinedAt || actedAt);
      return trail;
    }
    if (request.declinedAtStep === 'Head of Department') {
      setStep('Procurement', 'Completed', request.procurementSavedAt || actedAt);
      setStep('Head of Department', 'Rejected', request.declinedAt || actedAt);
      return trail;
    }
    if (request.declinedAtStep === 'Director HR/Admin') {
      setStep('Procurement', 'Completed', request.procurementSavedAt || actedAt);
      setStep('Head of Department', 'Approved', request.hodApprovedAt || actedAt);
      setStep('Director HR/Admin', 'Rejected', request.declinedAt || actedAt);
      return trail;
    }
    if (request.declinedAtStep === 'Finance Head') {
      setStep('Procurement', 'Completed', request.procurementSavedAt || actedAt);
      setStep('Head of Department', 'Approved', request.hodApprovedAt || actedAt);
      setStep('Director HR/Admin', 'Approved', request.directorApprovedAt || actedAt);
      setStep('Finance Head', 'Rejected', request.declinedAt || actedAt);
      return trail;
    }
    if (request.declinedAtStep === 'VP Strategy') {
      setStep('Procurement', 'Completed', request.procurementSavedAt || actedAt);
      setStep('Head of Department', 'Approved', request.hodApprovedAt || actedAt);
      setStep('Director HR/Admin', 'Approved', request.directorApprovedAt || actedAt);
      setStep('Finance Head', 'Approved', request.financeHeadApprovedAt || actedAt);
      setStep('VP Strategy', 'Rejected', request.declinedAt || actedAt);
      return trail;
    }
  }

  return trail;
};

function App() {
  const [rows, setRows] = useState([]);
  const [banner, setBanner] = useState('');
  const [budgetRows, setBudgetRows] = useState([]);
  const [budgetRequests, setBudgetRequests] = useState(() => {
    try {
      const saved = localStorage.getItem('budgetRequests');
      return saved ? JSON.parse(saved) : [];
    } catch (error) {
      console.error('Failed to parse budgetRequests from localStorage:', error);
      return [];
    }
  });
  const [editingBudgetRequestId, setEditingBudgetRequestId] = useState(null);
  const [assetRequests, setAssetRequests] = useState(() => {
    try {
      const saved = localStorage.getItem('assetRequests');
      return saved ? JSON.parse(saved) : [];
    } catch (error) {
      console.error('Failed to parse assetRequests from localStorage:', error);
      return [];
    }
  });

  useEffect(() => {
    localStorage.setItem('budgetRequests', JSON.stringify(budgetRequests));
  }, [budgetRequests]);

  useEffect(() => {
    localStorage.setItem('assetRequests', JSON.stringify(assetRequests));
  }, [assetRequests]);

  const location = useLocation();

  // Reset scroll to the top of the page whenever the route changes.
  useEffect(() => {
    window.scrollTo({ top: 0, left: 0, behavior: 'auto' });
  }, [location.pathname]);

  // Auto-dismiss the global banner so messages don't persist until refresh,
  // and scroll the user to the top so they actually see the message.
  useEffect(() => {
    if (!banner) return undefined;
    window.scrollTo({ top: 0, left: 0, behavior: 'smooth' });
    const timer = setTimeout(() => setBanner(''), 6000);
    return () => clearTimeout(timer);
  }, [banner]);

  return (
    <div className="app-shell">
      <Header />
      {banner ? <div className="global-banner">{banner}</div> : null}
      <main className="page-shell">
        <Routes>
          <Route path="/" element={<LandingPage />} />
          <Route
            path="/asset-requisition"
            element={
              <AssetRequisitionPage
                rows={rows}
                setRows={setRows}
                setBanner={setBanner}
                assetRequests={assetRequests}
                setAssetRequests={setAssetRequests}
              />
            }
          />
          <Route
            path="/asset-request/:id"
            element={
              <AssetRequestDetailsPage
                assetRequests={assetRequests}
                setAssetRequests={setAssetRequests}
                setBanner={setBanner}
              />
            }
          />
          <Route
            path="/budget-upload"
            element={
              <BudgetUploadPage
                budgetRows={budgetRows}
                setBudgetRows={setBudgetRows}
                budgetRequests={budgetRequests}
                setBudgetRequests={setBudgetRequests}
                editingBudgetRequestId={editingBudgetRequestId}
                setEditingBudgetRequestId={setEditingBudgetRequestId}
                setBanner={setBanner}
              />
            }
          />
          <Route
            path="/budget-request/:id"
            element={<BudgetRequestDetailsPage budgetRequests={budgetRequests} />}
          />
          <Route path="/requisition-tracking" element={<RequisitionTrackingPage assetRequests={assetRequests} />} />
		  <Route path="/bank-transfers" element={<BankTransfersPage setBanner={setBanner} />} />
          <Route path="/sadad-payments" element={<SADADPaymentsPage setBanner={setBanner} />} />
          <Route path="/sadad-payments/:screenKey" element={<SADADPaymentsPage setBanner={setBanner} />} />
          <Route path="/report-creation" element={<ReportsPage />} />
        </Routes>
      </main>
      <Footer />
    </div>
  );
}

function Header() {
  return (
    <header className="site-header">
      <div className="header-inner">
        <Link to="/" className="brand-block" aria-label="MuntQ home">
          <div className="brand-logo-wrap">
            <img src="/muntq-logo.png" alt="MuntQ logo" className="brand-logo" />
          </div>
          <div>
            <div className="brand-subtitle">Shaker Automation Portal</div>
          </div>
        </Link>

        <nav className="top-nav" aria-label="Primary navigation">
          <NavItem to="/" label="Home" />
          <NavDropdown
            label="Asset Management"
            items={[
              { to: '/asset-requisition', label: 'Asset Requisition' },
              { to: '/requisition-tracking', label: 'Requisition Tracking' },
            ]}
          />
          <NavDropdown
            label="Finance & Payments"
            items={[
              { to: '/budget-upload', label: 'Upload Budget' },
              { to: '/bank-transfers', label: 'Bank Transfers' },
              { to: '/sadad-payments', label: 'SADAD Payments' },
            ]}
          />
          <NavItem to="/report-creation" label="Reports" />
        </nav>
      </div>
    </header>
  );
}

function NavItem({ to, label }) {
  return (
    <NavLink
      to={to}
      className={({ isActive }) => (isActive ? 'nav-link active' : 'nav-link')}
    >
      {label}
    </NavLink>
  );
}

function NavDropdown({ label, items }) {
  const location = useLocation();
  const [open, setOpen] = useState(false);
  const isChildActive = items.some((item) => location.pathname === item.to);

  return (
    <div
      className="nav-dropdown"
      onMouseEnter={() => setOpen(true)}
      onMouseLeave={() => setOpen(false)}
    >
      <button
        type="button"
        className={isChildActive ? 'nav-link nav-dropdown-trigger active' : 'nav-link nav-dropdown-trigger'}
        aria-haspopup="true"
        aria-expanded={open}
        onClick={() => setOpen((current) => !current)}
      >
        {label}
        <ChevronDown size={15} className={open ? 'nav-dropdown-caret open' : 'nav-dropdown-caret'} />
      </button>
      <div className={open ? 'nav-dropdown-menu open' : 'nav-dropdown-menu'} role="menu">
        {items.map((item) => (
          <NavLink
            key={item.to}
            to={item.to}
            role="menuitem"
            onClick={() => setOpen(false)}
            className={({ isActive }) => (isActive ? 'nav-dropdown-item active' : 'nav-dropdown-item')}
          >
            {item.label}
          </NavLink>
        ))}
      </div>
    </div>
  );
}

function Footer() {
  return (
    <footer className="site-footer">
      <div className="footer-inner">
        <div>Built by QBS © {new Date().getFullYear()}</div>
        <div className="footer-note">An AI Initiative by Shaker ©</div>
      </div>
    </footer>
  );
}

function LandingPage() {
  return (
    <div className="page-grid">
      <section className="content-card landing-card-grid">
        {sampleLandingCards.map((item) => {
          const Icon = item.icon;
          return (
            <Link to={item.to} key={item.title} className="landing-card">
              <div className="landing-icon">
                <Icon size={22} />
              </div>
              <h3>{item.title}</h3>
              <p>{item.text}</p>
              <span className="card-link">
                Open <ArrowRight size={16} />
              </span>
            </Link>
          );
        })}
      </section>
    </div>
  );
}

function SADADPaymentsPage({ setBanner }) {
  const { screenKey } = useParams();
  const navigate = useNavigate();
  const screen = screenKey ? findScreen(screenKey) : null;

  return (
    <div className="page-grid">
      <section className="content-card">
        <div className="sadad-toolbar">
          <div className="sadad-toolbar-title">
            <h2>SADAD Payments</h2>
          </div>
          <div className="sadad-toolbar-select">
            <SadadSelect
              selectedKey={screenKey || ''}
              onSelect={(key) => navigate(`/sadad-payments/${key}`)}
            />
          </div>
        </div>

        {screen ? <SadadWorkspace screen={screen} setBanner={setBanner} /> : <SadadEmptyState />}
      </section>
    </div>
  );
}

function SadadSelect({ selectedKey, onSelect }) {
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState('');
  const ref = useRef(null);
  const selected = selectedKey ? findScreen(selectedKey) : null;

  useEffect(() => {
    function onDocClick(e) {
      if (ref.current && !ref.current.contains(e.target)) setOpen(false);
    }
    document.addEventListener('mousedown', onDocClick);
    return () => document.removeEventListener('mousedown', onDocClick);
  }, []);

  const q = query.trim().toLowerCase();
  const groups = SADAD_GROUPS.map((g) => ({
    group: g.group,
    screens: g.screens.filter(
      (s) =>
        !q ||
        s.nameAr.toLowerCase().includes(q) ||
        s.nameEn.toLowerCase().includes(q) ||
        (s.code && s.code.includes(q)) ||
        s.key.includes(q)
    ),
  })).filter((g) => g.screens.length > 0);

  return (
    <div className="sadad-select" ref={ref}>
      <div
        className={open ? 'sadad-field open' : 'sadad-field'}
        onClick={() => {
          if (!open) setOpen(true);
        }}
      >
        {open ? (
          <>
            <Search size={16} className="sadad-search-icon" />
            <input
              className="sadad-field-input"
              autoFocus
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Escape') {
                  setOpen(false);
                  setQuery('');
                }
              }}
              placeholder="Search by name or code…"
            />
            <span className="sadad-search-hint">{SADAD_SCREENS.length}</span>
          </>
        ) : selected ? (
          <span className="sadad-field-value">
            <img className="sadad-logo-thumb" src={selected.logo} alt="" />
            <span className="sadad-field-name">{selected.nameAr}</span>
            {selected.code ? <span className="sadad-code">{selected.code}</span> : null}
            <span className="sadad-field-en">· {selected.nameEn}</span>
          </span>
        ) : (
          <span className="sadad-placeholder">Select a payment screen…</span>
        )}
        <ChevronDown
          size={18}
          className="sadad-chevron"
          onClick={(e) => {
            e.stopPropagation();
            setOpen((o) => !o);
          }}
        />
      </div>

      {open ? (
        <div className="sadad-dropdown">
          <div className="sadad-dd-scroll">
            {groups.length === 0 ? (
              <div className="sadad-no-results">No screens match “{query}”.</div>
            ) : (
              groups.map((g) => (
                <div key={g.group}>
                  <div className="sadad-group-head">{g.group}</div>
                  {g.screens.map((s) => (
                    <button
                      type="button"
                      key={s.key}
                      className={s.key === selectedKey ? 'sadad-option selected' : 'sadad-option'}
                      onClick={() => {
                        onSelect(s.key);
                        setOpen(false);
                        setQuery('');
                      }}
                    >
                      <img className="sadad-logo-thumb" src={s.logo} alt="" />
                      <span className="sadad-opt-name">{s.nameAr}</span>
                      <span className="sadad-opt-en">· {s.nameEn}</span>
                      {s.code ? <span className="sadad-code sadad-opt-code">{s.code}</span> : null}
                      {s.key === selectedKey ? <Check size={16} className="sadad-opt-check" /> : null}
                    </button>
                  ))}
                </div>
              ))
            )}
          </div>
        </div>
      ) : null}
    </div>
  );
}

function SadadEmptyState() {
  return (
    <div className="sadad-empty">
      <div className="sadad-empty-circle">
        <Search size={24} />
      </div>
      <h3>No screen selected yet</h3>
      <p>
        Select a payment screen from the dropdown above to load it here. You can search by name or
        code (e.g. “GOSI”, “060”, “اجير”).
      </p>
    </div>
  );
}

function SadadWorkspace({ screen, setBanner }) {
  const [mode, setMode] = useState(screen.modes ? screen.modes[0] : null);

  useEffect(() => {
    setMode(screen.modes ? screen.modes[0] : null);
  }, [screen.key]);

  return (
    <div className="sadad-workspace">
      <div className="sadad-ws-head">
        <div className="sadad-ws-id">
          <img className="sadad-logo-lg" src={screen.logo} alt="" />
          <div className="sadad-ws-title">
            {screen.nameAr}
            <span className="sadad-ws-en">
              {' '}
              — {screen.nameEn}
              {screen.code ? ` · ${screen.code}` : ''}
            </span>
          </div>
        </div>
        {screen.modes ? (
          <div className="sadad-segmented">
            {screen.modes.map((m) => (
              <button
                type="button"
                key={m}
                className={m === mode ? 'sadad-seg active' : 'sadad-seg'}
                onClick={() => setMode(m)}
              >
                {m}
              </button>
            ))}
          </div>
        ) : null}
      </div>

      <div className="sadad-ws-body">
        {screen.key === 'iqama' && mode === 'Expense' ? (
          <IqamaExpenseScreen setBanner={setBanner} />
        ) : (screen.key === 'iqama' && mode === 'Payment') || screen.key === 'stc' ? (
          <IqamaPaymentScreen
            setBanner={setBanner}
            categoryKey={screen.key}
            categoryName={screen.nameEn}
          />
        ) : (
          <SadadGenericBody screen={screen} mode={mode} />
        )}
      </div>
    </div>
  );
}

function SadadGenericBody({ screen, mode }) {
  return (
    <>
      <div className="sadad-actions">
        <button type="button" className="secondary-button">
          <Download size={18} /> Download Template
        </button>
        <button type="button" className="primary-button">
          <Upload size={18} /> Upload File
        </button>
        <button type="button" className="secondary-button">+ Add Single Entry</button>
      </div>

      <div className="sadad-stub">
        The <strong>{screen.nameEn}</strong> screen will be built here
        {mode ? ` (${mode})` : ''}. Its fields, template and table columns will be specific to this
        screen.
      </div>
    </>
  );
}

const IQAMA_NUMBER_LENGTH = 10; // Standard Saudi Iqama length — change here if your file differs.
const IQAMA_EXPENSE_TYPES = [
  'تجديد اقامة ورخصة عمل',
  'خروج وعودة',
  'رخصة عمل',
  'نقل كفالة',
  'تجديد اقامة',
];
const IQAMA_COLUMNS = ['Iqama Number', 'Amount', 'Expense Type', 'SAP Emp No.', 'Beneficiary', 'Own Account'];

const iqamaRowActive = (it) =>
  [it.iqama, it.amount, it.sapNo].some((v) => String(v ?? '').trim() !== '');

function validateIqamaItem(it) {
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

function IqamaExpenseScreen({ setBanner }) {
  const [items, setItems] = useState([]);
  const [focusId, setFocusId] = useState(null);
  const newRowRef = useRef(null);

  useEffect(() => {
    if (focusId && newRowRef.current) {
      newRowRef.current.scrollIntoView({ behavior: 'smooth', block: 'center' });
      const firstInput = newRowRef.current.querySelector('input, select');
      if (firstInput) firstInput.focus();
      setFocusId(null);
    }
  }, [focusId]);

  const blankItem = () => ({
    id: crypto.randomUUID(),
    iqama: '',
    amount: '',
    expenseType: '',
    sapNo: '',
    beneficiary: '',
    ownAccount: '0.00',
  });

  const updateItem = (id, field, value) =>
    setItems((cur) => cur.map((it) => (it.id === id ? { ...it, [field]: value } : it)));
  const removeItem = (id) => setItems((cur) => cur.filter((it) => it.id !== id));
  const addItem = () => {
    const item = blankItem();
    setItems((cur) => [...cur, item]);
    setFocusId(item.id);
  };

  const handleAmountBlur = (id, field, value) => {
    const raw = String(value).replace(/,/g, '').trim();
    if (raw !== '' && !Number.isNaN(Number(raw))) updateItem(id, field, formatMoney(Number(raw)));
  };
  const handleAmountFocus = (id, field, value) =>
    updateItem(id, field, String(value).replace(/,/g, ''));

  const handleDownloadTemplate = async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Iqama Expense');
    const headerRow = worksheet.addRow(IQAMA_COLUMNS);
    headerRow.eachCell((cell) => {
      cell.font = { bold: true };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'EAF3EE' } };
      cell.border = {
        top: { style: 'thin', color: { argb: 'D0D7DE' } },
        left: { style: 'thin', color: { argb: 'D0D7DE' } },
        bottom: { style: 'thin', color: { argb: 'D0D7DE' } },
        right: { style: 'thin', color: { argb: 'D0D7DE' } },
      };
    });
    for (let i = 0; i < 50; i += 1) worksheet.addRow(['', '', '', '', '', '']);
    worksheet.dataValidations.add('C2:C500', {
      type: 'list',
      allowBlank: true,
      formulae: [`"${IQAMA_EXPENSE_TYPES.join(',')}"`],
      showErrorMessage: true,
      errorTitle: 'Invalid Expense Type',
      error: 'Please select an expense type from the dropdown list.',
    });
    [22, 16, 26, 12, 48, 16].forEach((w, i) => {
      worksheet.getColumn(i + 1).width = w;
    });
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'muntq-iqama-expense-template.xlsx';
    a.click();
    window.URL.revokeObjectURL(url);
    setBanner('Template downloaded. Fill it and upload it back anytime.');
  };

  const handleFileUpload = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
      if (!matrix.length) {
        setBanner('The uploaded file is empty.');
        event.target.value = '';
        return;
      }
      const header = (matrix[0] || []).map((h) => String(h ?? '').trim());
      const columnsMatch =
        header.length >= IQAMA_COLUMNS.length && IQAMA_COLUMNS.every((c, i) => header[i] === c);
      if (!columnsMatch) {
        const missing = IQAMA_COLUMNS.filter((c) => !header.includes(c));
        setBanner(
          missing.length
            ? `"${missing.join('", "')}" column(s) missing in the sheet. Required columns: ${IQAMA_COLUMNS.join(', ')}.`
            : `File columns don't match the template. Expected order: ${IQAMA_COLUMNS.join(', ')}.`
        );
        event.target.value = '';
        return;
      }
      const newItems = matrix
        .slice(1)
        .map((r) => ({
          id: crypto.randomUUID(),
          iqama: String(r[0] ?? '').trim(),
          amount: String(r[1] ?? '').trim(),
          expenseType: String(r[2] ?? '').trim(),
          sapNo: String(r[3] ?? '').trim(),
          beneficiary: String(r[4] ?? '').trim(),
          ownAccount: String(r[5] ?? '').trim(),
        }))
        .filter((it) =>
          [it.iqama, it.amount, it.expenseType, it.sapNo, it.beneficiary, it.ownAccount].some(
            (v) => v !== ''
          )
        )
        .map((it) => {
          const amtRaw = it.amount.replace(/,/g, '');
          if (amtRaw !== '' && !Number.isNaN(Number(amtRaw))) it.amount = formatMoney(Number(amtRaw));
          const ownRaw = it.ownAccount.replace(/,/g, '');
          if (ownRaw === '') it.ownAccount = '0.00';
          else if (!Number.isNaN(Number(ownRaw))) it.ownAccount = formatMoney(Number(ownRaw));
          return it;
        });
      if (!newItems.length) {
        setBanner('No line items were found in the uploaded file.');
        event.target.value = '';
        return;
      }
      setItems((cur) => [...cur, ...newItems]);
      const invalid = newItems.filter((it) => Object.keys(validateIqamaItem(it)).length > 0).length;
      setBanner(
        invalid
          ? `Uploaded ${newItems.length} line item(s). ${invalid} need attention — fix the highlighted fields below.`
          : `Uploaded ${newItems.length} line item(s) successfully.`
      );
      event.target.value = '';
    } catch (err) {
      setBanner('The file could not be read. Please use an .xlsx file based on the template.');
      event.target.value = '';
    }
  };

  const errorCount = items.filter((it) => Object.keys(validateIqamaItem(it)).length > 0).length;

  // Postable only when there is at least one line item, every row is filled (active),
  // and no row has any validation error.
  const canPost =
    items.length > 0 &&
    items.every((it) => iqamaRowActive(it) && Object.keys(validateIqamaItem(it)).length === 0);

  const handlePostToSAP = () => {
    if (!canPost) {
      setBanner('Cannot post to SAP. Fix the highlighted errors and remove any empty line items first.');
      return;
    }
    const count = items.length;
    setItems([]);
    setBanner(`Posted to SAP successfully! ${count} line item(s) posted. You can start a new upload.`);
  };

  return (
    <>
      <div className="sadad-actions">
        <button type="button" className="secondary-button" onClick={handleDownloadTemplate}>
          <Download size={18} /> Download Template
        </button>
        <label className="primary-button upload-button">
          <Upload size={18} /> Upload File
          <input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} />
        </label>
        <button type="button" className="secondary-button" onClick={addItem}>
          + Add Single Entry
        </button>
      </div>

      {items.length === 0 ? (
        <div className="sadad-stub">
          No line items yet. Download the template, upload a filled file, or add a single entry to
          begin.
        </div>
      ) : (
        <>
          <div className="sadad-table-meta">
            <span>{items.length} line item(s)</span>
            {errorCount > 0 ? (
              <span className="sadad-meta-error">{errorCount} with errors</span>
            ) : (
              <span className="sadad-meta-ok">All valid</span>
            )}
            <span className="sadad-meta-req"><span className="req-star">*</span> Mandatory</span>
          </div>
          <div className="sadad-table-wrap">
            <table className="data-table iqama-table">
              <thead>
                <tr>
                  <th>#</th>
                  <th>Iqama Number <span className="req-star">*</span></th>
                  <th>Amount <span className="req-star">*</span></th>
                  <th>Expense Type <span className="req-star">*</span></th>
                  <th>SAP Emp No. <span className="req-star">*</span></th>
                  <th>Beneficiary</th>
                  <th>Own Account</th>
                  <th aria-label="Remove" />
                </tr>
              </thead>
              <tbody>
                {items.map((it, idx) => {
                  const e = validateIqamaItem(it);
                  const typeInList = !it.expenseType || IQAMA_EXPENSE_TYPES.includes(it.expenseType);
                  return (
                    <tr key={it.id} ref={it.id === focusId ? newRowRef : undefined}>
                      <td>{idx + 1}</td>
                      <td className="cell-required">
                        <input
                          className={e.iqama ? 'cell-input input-error' : 'cell-input'}
                          value={it.iqama}
                          onChange={(ev) => updateItem(it.id, 'iqama', ev.target.value)}
                          placeholder={`${IQAMA_NUMBER_LENGTH} digits`}
                        />
                        {e.iqama ? <div className="cell-error">{e.iqama}</div> : null}
                      </td>
                      <td className="cell-required">
                        <input
                          className={e.amount ? 'cell-input input-error' : 'cell-input'}
                          value={it.amount}
                          onChange={(ev) => updateItem(it.id, 'amount', ev.target.value)}
                          onFocus={(ev) => handleAmountFocus(it.id, 'amount', ev.target.value)}
                          onBlur={(ev) => handleAmountBlur(it.id, 'amount', ev.target.value)}
                          placeholder="0.00"
                          inputMode="decimal"
                        />
                        {e.amount ? <div className="cell-error">{e.amount}</div> : null}
                      </td>
                      <td className="cell-required">
                        <select
                          className={e.expenseType ? 'cell-input input-error' : 'cell-input'}
                          value={typeInList ? it.expenseType : '__invalid__'}
                          onChange={(ev) =>
                            updateItem(
                              it.id,
                              'expenseType',
                              ev.target.value === '__invalid__' ? it.expenseType : ev.target.value
                            )
                          }
                        >
                          <option value="">—</option>
                          {IQAMA_EXPENSE_TYPES.map((t) => (
                            <option key={t} value={t}>
                              {t}
                            </option>
                          ))}
                          {!typeInList ? (
                            <option value="__invalid__">{it.expenseType} (invalid)</option>
                          ) : null}
                        </select>
                        {e.expenseType ? <div className="cell-error">{e.expenseType}</div> : null}
                      </td>
                      <td className="cell-required">
                        <input
                          className={e.sapNo ? 'cell-input input-error' : 'cell-input'}
                          value={it.sapNo}
                          onChange={(ev) => updateItem(it.id, 'sapNo', ev.target.value)}
                          placeholder="e.g. 10234"
                        />
                        {e.sapNo ? <div className="cell-error">{e.sapNo}</div> : null}
                      </td>
                      <td>
                        <input
                          className="cell-input"
                          value={it.beneficiary}
                          onChange={(ev) => updateItem(it.id, 'beneficiary', ev.target.value)}
                          placeholder="Optional"
                        />
                      </td>
                      <td>
                        <input
                          className={e.ownAccount ? 'cell-input input-error' : 'cell-input'}
                          value={it.ownAccount}
                          onChange={(ev) => updateItem(it.id, 'ownAccount', ev.target.value)}
                          onFocus={(ev) => handleAmountFocus(it.id, 'ownAccount', ev.target.value)}
                          onBlur={(ev) => handleAmountBlur(it.id, 'ownAccount', ev.target.value)}
                          placeholder="0.00"
                          inputMode="decimal"
                        />
                        {e.ownAccount ? <div className="cell-error">{e.ownAccount}</div> : null}
                      </td>
                      <td>
                        <button
                          type="button"
                          className="icon-button"
                          onClick={() => removeItem(it.id)}
                          aria-label="Remove line"
                        >
                          <Trash2 size={16} />
                        </button>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>

          <div className="submit-row">
            <button
              type="button"
              className="primary-button"
              onClick={handlePostToSAP}
              disabled={!canPost}
            >
              Post to SAP <ArrowRight size={18} />
            </button>
          </div>
        </>
      )}
    </>
  );
}

const PAYMENT_COLUMNS = [
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

const isNumericValue = (s) => {
  const r = String(s).replace(/,/g, '').trim();
  return r !== '' && !Number.isNaN(Number(r));
};
const isDigitsOnly = (s) => /^\d+$/.test(String(s).trim());
const TIME_RE = /^([01]\d|2[0-3]):[0-5]\d$/;
const DATE_RE = /^(\d{2})\/(\d{2})\/(\d{4})$/;
const isValidDDMMYYYY = (s) => {
  const m = String(s).trim().match(DATE_RE);
  if (!m) return false;
  const d = +m[1];
  const mo = +m[2];
  const y = +m[3];
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return false;
  const dt = new Date(y, mo - 1, d);
  return dt.getFullYear() === y && dt.getMonth() === mo - 1 && dt.getDate() === d;
};
const formatDateDDMMYYYY = (dt) => {
  const d = String(dt.getDate()).padStart(2, '0');
  const m = String(dt.getMonth() + 1).padStart(2, '0');
  return `${d}/${m}/${dt.getFullYear()}`;
};

const paymentRowActive = (it) =>
  [it.bankRef, it.customerRef, it.valueDate, it.debitAmount, it.narrative].some(
    (v) => String(v ?? '').trim() !== ''
  );

function validatePaymentItem(it) {
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
const CODE_TO_KEY = SADAD_SCREENS.reduce((acc, s) => {
  if (s.code) acc[s.code] = s.key;
  return acc;
}, {});

// Identify which SADAD screen/category a raw invoice (bank statement) line belongs to.
// Returns a screen key (matching SADAD_SCREENS[].key) or null if it isn't a SADAD payment.
function classifyInvoiceCategory(it) {
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

function IqamaPaymentScreen({ setBanner, categoryKey, categoryName }) {
  const [items, setItems] = useState([]);
  const [focusId, setFocusId] = useState(null);
  const newRowRef = useRef(null);

  useEffect(() => {
    if (focusId && newRowRef.current) {
      newRowRef.current.scrollIntoView({ behavior: 'smooth', block: 'center' });
      const firstInput = newRowRef.current.querySelector('input, select');
      if (firstInput) firstInput.focus();
      setFocusId(null);
    }
  }, [focusId]);

  const blankItem = () => ({
    id: crypto.randomUUID(),
    bankRef: '',
    narrative: '',
    customerRef: '',
    trnType: '',
    valueDate: '',
    creditAmount: '',
    debitAmount: '',
    balance: '',
    time: '',
    postDate: '',
  });

  const updateItem = (id, field, value) =>
    setItems((cur) => cur.map((it) => (it.id === id ? { ...it, [field]: value } : it)));
  const removeItem = (id) => setItems((cur) => cur.filter((it) => it.id !== id));
  const addItem = () => {
    const item = blankItem();
    setItems((cur) => [...cur, item]);
    setFocusId(item.id);
  };

  const handleAmtBlur = (id, field, value) => {
    const raw = String(value).replace(/,/g, '').trim();
    if (raw !== '' && !Number.isNaN(Number(raw))) updateItem(id, field, formatMoney(Number(raw)));
  };
  const handleAmtFocus = (id, field, value) =>
    updateItem(id, field, String(value).replace(/,/g, ''));

  const cellStr = (v) => String(v ?? '').trim();
  const cellAmt = (v) => {
    if (typeof v === 'number') return formatMoney(v);
    const s = String(v ?? '').trim();
    const raw = s.replace(/,/g, '');
    return s !== '' && !Number.isNaN(Number(raw)) ? formatMoney(Number(raw)) : s;
  };
  const cellDate = (v) => (v instanceof Date ? formatDateDDMMYYYY(v) : String(v ?? '').trim());

  const handleDownloadTemplate = async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Payment');
    const headerRow = worksheet.addRow(PAYMENT_COLUMNS);
    headerRow.eachCell((cell) => {
      cell.font = { bold: true };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'EAF3EE' } };
      cell.border = {
        top: { style: 'thin', color: { argb: 'D0D7DE' } },
        left: { style: 'thin', color: { argb: 'D0D7DE' } },
        bottom: { style: 'thin', color: { argb: 'D0D7DE' } },
        right: { style: 'thin', color: { argb: 'D0D7DE' } },
      };
    });
    for (let i = 0; i < 50; i += 1) worksheet.addRow(['', '', '', '', '', '', '', '', '', '']);
    [16, 40, 18, 14, 14, 15, 15, 15, 10, 14].forEach((w, i) => {
      worksheet.getColumn(i + 1).width = w;
    });
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `muntq-${categoryKey || 'sadad'}-payment-template.xlsx`;
    a.click();
    window.URL.revokeObjectURL(url);
    setBanner('Template downloaded. Fill it and upload it back anytime.');
  };

  const handleFileUpload = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
      if (!matrix.length) {
        setBanner('The uploaded file is empty.');
        event.target.value = '';
        return;
      }
      // Map columns by header name (tolerant to extra/blank columns in a raw bank statement).
      const header = (matrix[0] || []).map((h) => String(h ?? '').trim());
      const idx = (name) => header.indexOf(name);
      const at = (r, name) => {
        const i = idx(name);
        return i >= 0 ? r[i] : '';
      };
      // Essential columns to identify the category.
      const essential = ['Narrative', 'TRN type'];
      const missing = essential.filter((c) => idx(c) < 0);
      if (missing.length) {
        setBanner(
          `Can't read the invoice — missing column(s): ${missing.join(', ')}. The file needs at least Bank reference, Narrative and TRN type.`
        );
        event.target.value = '';
        return;
      }

      const allRows = matrix
        .slice(1)
        .map((r) => ({
          id: crypto.randomUUID(),
          bankRef: cellStr(at(r, 'Bank reference')),
          narrative: cellStr(at(r, 'Narrative')),
          customerRef: cellStr(at(r, 'Customer reference')),
          trnType: cellStr(at(r, 'TRN type')),
          valueDate: cellDate(at(r, 'Value date')),
          creditAmount: cellAmt(at(r, 'Credit amount')),
          debitAmount: cellAmt(at(r, 'Debit amount')),
          balance: cellAmt(at(r, 'Balance')),
          time: cellStr(at(r, 'Time')),
          postDate: cellDate(at(r, 'Post date')),
        }))
        .filter((it) =>
          [
            it.bankRef,
            it.narrative,
            it.customerRef,
            it.trnType,
            it.valueDate,
            it.creditAmount,
            it.debitAmount,
            it.balance,
            it.time,
            it.postDate,
          ].some((v) => v !== '')
        );

      if (!allRows.length) {
        setBanner('No line items were found in the uploaded file.');
        event.target.value = '';
        return;
      }

      // Keep only rows that belong to this screen's category.
      const matched = allRows.filter((it) => classifyInvoiceCategory(it) === categoryKey);

      // Drop exact duplicates (all displayed columns identical).
      const seen = new Set();
      const deduped = [];
      let dupCount = 0;
      matched.forEach((it) => {
        const key = [
          it.bankRef,
          it.narrative,
          it.customerRef,
          it.trnType,
          it.valueDate,
          it.creditAmount,
          it.debitAmount,
          it.balance,
          it.time,
          it.postDate,
        ].join('||');
        if (seen.has(key)) {
          dupCount += 1;
        } else {
          seen.add(key);
          deduped.push(it);
        }
      });

      const label = categoryName || 'this category';
      if (!deduped.length) {
        setItems([]);
        setBanner(
          `No "${label}" line items found in the file (scanned ${allRows.length} row(s)).`
        );
        event.target.value = '';
        return;
      }
      setItems(deduped);
      const dupNote = dupCount ? `, ${dupCount} duplicate(s) removed` : '';
      setBanner(
        `Identified ${deduped.length} "${label}" line item(s) out of ${allRows.length} row(s)${dupNote}. Review and post to SAP.`
      );
      event.target.value = '';
    } catch (err) {
      setBanner('The file could not be read. Please use an .xlsx file based on the template.');
      event.target.value = '';
    }
  };

  const errorCount = items.filter((it) => Object.keys(validatePaymentItem(it)).length > 0).length;
  const canPost =
    items.length > 0 &&
    items.every((it) => paymentRowActive(it) && Object.keys(validatePaymentItem(it)).length === 0);

  const sumField = (field) =>
    items.reduce((acc, it) => {
      const n = Number(String(it[field] ?? '').replace(/,/g, ''));
      return acc + (Number.isNaN(n) ? 0 : n);
    }, 0);
  const totalCredit = sumField('creditAmount');
  const totalDebit = sumField('debitAmount');

  const handlePostToSAP = () => {
    if (!canPost) {
      setBanner('Cannot post to SAP. Fix the highlighted errors and remove any empty line items first.');
      return;
    }
    const count = items.length;
    setItems([]);
    setBanner(`Posted to SAP successfully! ${count} line item(s) posted. You can start a new upload.`);
  };

  const amountCell = (it, field, err, required) => (
    <td className={required ? 'cell-required' : undefined}>
      <input
        className={err ? 'cell-input input-error' : 'cell-input'}
        value={it[field]}
        onChange={(ev) => updateItem(it.id, field, ev.target.value)}
        onFocus={(ev) => handleAmtFocus(it.id, field, ev.target.value)}
        onBlur={(ev) => handleAmtBlur(it.id, field, ev.target.value)}
        placeholder="0.00"
        inputMode="decimal"
      />
      {err ? <div className="cell-error">{err}</div> : null}
    </td>
  );

  const textCell = (it, field, err, placeholder, required) => (
    <td className={required ? 'cell-required' : undefined}>
      <input
        className={err ? 'cell-input input-error' : 'cell-input'}
        value={it[field]}
        onChange={(ev) => updateItem(it.id, field, ev.target.value)}
        placeholder={placeholder}
      />
      {err ? <div className="cell-error">{err}</div> : null}
    </td>
  );

  return (
    <>
      <div className="sadad-actions">
        <button type="button" className="secondary-button" onClick={handleDownloadTemplate}>
          <Download size={18} /> Download Template
        </button>
        <label className="primary-button upload-button">
          <Upload size={18} /> Upload File
          <input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} />
        </label>
        <button type="button" className="secondary-button" onClick={addItem}>
          + Add Single Entry
        </button>
      </div>

      {items.length === 0 ? (
        <div className="sadad-stub">
          No line items yet. Download the template, upload a filled file, or add a single entry to
          begin.
        </div>
      ) : (
        <>
          <div className="sadad-table-meta">
            <span>{items.length} line item(s)</span>
            {errorCount > 0 ? (
              <span className="sadad-meta-error">{errorCount} with errors</span>
            ) : (
              <span className="sadad-meta-ok">All valid</span>
            )}
            <span className="sadad-meta-req"><span className="req-star">*</span> Mandatory</span>
            <span className="sadad-totals">
              <span className="total-pill total-credit">Total Credit {formatMoney(totalCredit)}</span>
              <span className="total-pill total-debit">Total Debit {formatMoney(totalDebit)}</span>
            </span>
          </div>
          <div className="sadad-table-wrap">
            <table className="data-table iqama-table payment-table">
              <thead>
                <tr>
                  <th>#</th>
                  <th>Bank reference <span className="req-star">*</span></th>
                  <th>Narrative <span className="req-star">*</span></th>
                  <th>Customer reference <span className="req-star">*</span></th>
                  <th>TRN type</th>
                  <th>Value date <span className="req-star">*</span></th>
                  <th>Credit amount</th>
                  <th>Debit amount <span className="req-star">*</span></th>
                  <th>Balance</th>
                  <th>Time</th>
                  <th>Post date</th>
                  <th aria-label="Remove" />
                </tr>
              </thead>
              <tbody>
                {items.map((it, idx) => {
                  const e = validatePaymentItem(it);
                  return (
                    <tr key={it.id} ref={it.id === focusId ? newRowRef : undefined}>
                      <td>{idx + 1}</td>
                      {textCell(it, 'bankRef', e.bankRef, 'e.g. BTPE-00017', true)}
                      {textCell(it, 'narrative', e.narrative, 'Required', true)}
                      {textCell(it, 'customerRef', e.customerRef, 'Digits', true)}
                      {textCell(it, 'trnType', null, 'Optional')}
                      {textCell(it, 'valueDate', e.valueDate, 'DD/MM/YYYY', true)}
                      {amountCell(it, 'creditAmount', e.creditAmount)}
                      {amountCell(it, 'debitAmount', e.debitAmount, true)}
                      {amountCell(it, 'balance', e.balance)}
                      {textCell(it, 'time', e.time, 'HH:MM')}
                      {textCell(it, 'postDate', e.postDate, 'DD/MM/YYYY')}
                      <td>
                        <button
                          type="button"
                          className="icon-button"
                          onClick={() => removeItem(it.id)}
                          aria-label="Remove line"
                        >
                          <Trash2 size={16} />
                        </button>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>

          <div className="payment-footer">
            <div className="payment-footer-totals">
              <span>Total Credit <strong>{formatMoney(totalCredit)}</strong></span>
              <span>Total Debit <strong>{formatMoney(totalDebit)}</strong></span>
            </div>
            <button
              type="button"
              className="primary-button"
              onClick={handlePostToSAP}
              disabled={!canPost}
            >
              Post to SAP <ArrowRight size={18} />
            </button>
          </div>
        </>
      )}
    </>
  );
}

function AssetRequisitionPage({ rows, setRows, setBanner, assetRequests, setAssetRequests }) {
  const navigate = useNavigate();
  const [editingRowId, setEditingRowId] = useState(null);
  const [draftRows, setDraftRows] = useState({});
  const [uploadMessage, setUploadMessage] = useState('');
  const [selectedDepartment, setSelectedDepartment] = useState(DEPARTMENTS[0]);
  const [selectedCategory, setSelectedCategory] = useState(ASSET_CATEGORIES[0]);

  const fiscalYear = getCurrentFiscalYear();
  const totalBudget = BUDGET_MATRIX[selectedDepartment]?.[selectedCategory] ?? 0;

  const consumedBudget = useMemo(() => {
    return assetRequests.reduce((sum, request) => {
      if (String(request.fiscalYear) !== String(fiscalYear)) return sum;
      if (request.department !== selectedDepartment) return sum;
      if (request.category !== selectedCategory) return sum;
      if (request.status === 'Declined') return sum;
      return sum + parseAmount(request.totalAmount);
    }, 0);
  }, [assetRequests, fiscalYear, selectedDepartment, selectedCategory]);

  const availableBudget = totalBudget - consumedBudget;
  const currentRequestValue = rows.reduce((sum, row) => sum + parseAmount(row.Amount), 0);

  // Amount currently shown for a row (reflects in-progress inline edits).
  const amountForRow = (row) => parseAmount((draftRows[row.id] || row).Amount);

  // Per-category overview. On the requisition screen we always show every category
  // for the selected department so the budget card is visible before line items exist.
  const categoryBudgets = useMemo(() => {
    const dept = selectedDepartment;
    const fy = String(fiscalYear);

    return ASSET_CATEGORIES.map((category) => {
      const budget = (BUDGET_MATRIX[dept] && BUDGET_MATRIX[dept][category]) || 0;

      const thisRequisition = rows
        .filter((row) => (row['Asset Category'] || '') === category)
        .reduce((sum, row) => sum + amountForRow(row), 0);

      const consumedByOthers = assetRequests.reduce((sum, other) => {
        if (other.status === 'Declined') return sum;
        if ((other.department || '') !== dept) return sum;
        if (String(other.fiscalYear || '') !== fy) return sum;
        const otherCategoryTotal = (other.rows || []).reduce(
          (rowSum, row) => ((row['Asset Category'] || '') === category ? rowSum + parseAmount(row.Amount) : rowSum),
          0
        );
        return sum + otherCategoryTotal;
      }, 0);

      const available = budget - consumedByOthers;
      const remaining = available - thisRequisition;

      return { category, totalBudget: budget, thisRequisition, availableBudget: available, remaining };
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [rows, assetRequests, selectedDepartment, fiscalYear, draftRows]);

  const handleDownloadTemplate = async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Asset Requisition');

    const headerRow = worksheet.addRow(ASSET_UPLOAD_COLUMNS);
    headerRow.eachCell((cell) => {
      cell.font = { bold: true };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'EAF3EE' },
      };
      cell.border = {
        top: { style: 'thin', color: { argb: 'D0D7DE' } },
        left: { style: 'thin', color: { argb: 'D0D7DE' } },
        bottom: { style: 'thin', color: { argb: 'D0D7DE' } },
        right: { style: 'thin', color: { argb: 'D0D7DE' } },
      };
    });

    for (let i = 0; i < 50; i += 1) {
      worksheet.addRow([i + 1, '', '', fiscalYear, '', '', '']);
    }

    worksheet.dataValidations.add('E2:E200', {
      type: 'list',
      allowBlank: true,
      formulae: [`"${DEPARTMENTS.join(',')}"`],
      showErrorMessage: true,
      errorTitle: 'Invalid Department',
      error: 'Please select a department from the dropdown list.',
    });

    worksheet.dataValidations.add('F2:F200', {
      type: 'list',
      allowBlank: true,
      formulae: [`"${ASSET_CATEGORIES.join(',')}"`],
      showErrorMessage: true,
      errorTitle: 'Invalid Asset Category',
      error: 'Please select an asset category from the dropdown list.',
    });

    [18, 28, 32, 15, 18, 30, 12].forEach((width, index) => {
      worksheet.getColumn(index + 1).width = width;
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });

    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = 'muntq-asset-requisition-template.xlsx';
    anchor.click();
    window.URL.revokeObjectURL(url);

    setUploadMessage('Template downloaded successfully.');
    setBanner('Template ready. You can fill it later and upload the Excel file anytime.');
  };

  const sanitizeRow = (item, index = 0) => {
    const normalized = normalizeExcelRow(item);
    return {
      id: crypto.randomUUID(),
      'Serial Number': String(normalized['serial number'] || index + 1),
      'Asset Name': normalized['asset name'] ?? '',
      'Asset Description': normalized['asset description'] ?? '',
      'Fiscal Year': String(normalized['fiscal year'] || fiscalYear),
      'Department': normalized['department'] ?? '',
      'Asset Category': normalized['asset category'] ?? '',
      'Quantity': normalized['quantity'] ?? '',
      Amount: normalized['amount'] ?? '',
    };
  };

  const handleFileUpload = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const jsonRows = XLSX.utils.sheet_to_json(sheet, { defval: '' });

      const cleanedRows = jsonRows
        .map((row, index) => sanitizeRow(row, index))
        .filter((row) =>
          ['Asset Name', 'Asset Description', 'Department', 'Asset Category'].some(
            (column) => String(row[column] ?? '').trim() !== ''
          )
        );

      if (!cleanedRows.length) {
        setUploadMessage('No valid line items were found in the uploaded file.');
        return;
      }

      setRows((current) => [...current, ...cleanedRows]);
      setUploadMessage(`File uploaded successfully. ${cleanedRows.length} line item(s) added.`);
      setBanner('Upload successful. The new requisition lines were added below the existing items.');
      event.target.value = '';
    } catch (error) {
      setUploadMessage('The file could not be processed. Please use an Excel file with the required columns.');
    }
  };

  const startEdit = (row) => {
    setEditingRowId(row.id);
    setDraftRows((current) => ({ ...current, [row.id]: { ...row } }));
  };

  const updateDraftCell = (rowId, column, value) => {
    setDraftRows((current) => ({
      ...current,
      [rowId]: {
        ...(current[rowId] || {}),
        [column]: value,
      },
    }));
  };

  const saveEdit = (rowId) => {
    setRows((current) => current.map((row) => (row.id === rowId ? { ...draftRows[rowId] } : row)));
    setEditingRowId(null);
    setUploadMessage('Line item updated successfully.');
  };

  const deleteRow = (rowId) => {
    setRows((current) => current.filter((row) => row.id !== rowId));
    if (editingRowId === rowId) {
      setEditingRowId(null);
    }
    setUploadMessage('Line item deleted successfully.');
  };

  const submitRequisition = () => {
    if (!rows.length) return;

    const submittedAt = new Date().toLocaleString();
    const normalizedRows = rows.map((row, index) => ({
      ...row,
      'Serial Number': String(index + 1),
      'Fiscal Year': row['Fiscal Year'] || fiscalYear,
      'Department': row['Department'] || selectedDepartment,
      'Asset Category': row['Asset Category'] || selectedCategory,
      Amount: row.Amount ?? '',
    }));

    const newRequest = {
      id: `REQ-${Date.now()}`,
      submittedAt,
      fiscalYear,
      department: selectedDepartment,
      category: selectedCategory,
      status: 'Pending for Procurement',
      totalAmount: 0,
      rows: normalizedRows,
      procurementSavedAt: null,
      hodApprovedAt: null,
      directorApprovedAt: null,
      financeHeadApprovedAt: null,
      vpStrategyApprovedAt: null,
      financeAccountantCompletedAt: null,
      declinedAt: null,
      declinedAtStep: null,
      approvalTrail: buildAssetApprovalTrail({ submittedAt }, 'Pending for Procurement', submittedAt),
      progress: getProgressForStatus('Pending for Procurement'),
      mode: getModeForStatus('Pending for Procurement'),
    };

    setAssetRequests((current) => [...current, newRequest]);
    setBanner(`Asset request ${newRequest.id} submitted successfully.`);
    setRows([]);
    navigate('/requisition-tracking');
  };

  return (
    <div className="page-grid">
      <section className="content-card">
        <div className="section-heading-row">
          <div>
            <h2>Asset Requisition</h2>
            <p>Download the Excel guide if needed, upload one or multiple filled files and edit each row inline before submission.</p>
          </div>
          <div className="toolbar-row">
            <button className="secondary-button" onClick={handleDownloadTemplate}>
              <Download size={18} /> Download Template
            </button>
            <label className="primary-button upload-button">
              <Upload size={18} /> Upload Excel
              <input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} />
            </label>
          </div>
        </div>

        {uploadMessage ? <div className="success-message">{uploadMessage}</div> : null}

        <div className="budget-summary-row asset-summary-grid">
          <div className="budget-card">
            <div className="budget-label">Fiscal Year</div>
            <div className="budget-value">{fiscalYear}</div>
          </div>
          <div className="budget-card">
            <div className="budget-label">Department</div>
            <select
              className="cell-input summary-select"
              value={selectedDepartment}
              onChange={(event) => setSelectedDepartment(event.target.value)}
            >
              {DEPARTMENTS.map((department) => (
                <option key={department} value={department}>
                  {department}
                </option>
              ))}
            </select>
          </div>
          <div className="budget-card">
            <div className="budget-label">Category</div>
            <select
              className="cell-input summary-select"
              value={selectedCategory}
              onChange={(event) => setSelectedCategory(event.target.value)}
            >
              {ASSET_CATEGORIES.map((category) => (
                <option key={category} value={category}>
                  {category}
                </option>
              ))}
            </select>
          </div>
          <div className="budget-card">
            <div className="budget-label">Total Budget</div>
            <div className="budget-value">{formatAmount(totalBudget)}</div>
          </div>
          <div className="budget-card">
            <div className="budget-label">Consumed Budget</div>
            <div className="budget-value">{formatAmount(consumedBudget)}</div>
          </div>
          <div className="budget-card">
            <div className="budget-label">Available Budget</div>
            <div className={`budget-value ${availableBudget < 0 ? 'budget-value-danger' : ''}`}>
              {formatAmount(availableBudget)}
            </div>
          </div>
          <div className="budget-card">
            <div className="budget-label">Current Request Value</div>
            <div className="budget-value">{formatAmount(currentRequestValue)}</div>
          </div>
        </div>

        {rows.length > 0 &&
          (availableBudget < 0 ? (
            <div className="budget-message warning">
              <AlertTriangle size={16} />
              <span>Your request is exceeding the available budget, but you can still proceed by pressing the Submit Requisition button.</span>
            </div>
          ) : (
            <div className="budget-message ok">Your request is within the available budget.</div>
          ))}

        {categoryBudgets.length ? (
          <div className="category-budget-section">
            <h3 className="category-budget-title">Budget by Category</h3>
            <div className="category-budget-grid">
              {categoryBudgets.map((entry) => {
                const overBudget = entry.thisRequisition > entry.availableBudget;
                return (
                  <div className="category-budget-card" key={entry.category}>
                    <div className="category-budget-name">{entry.category}</div>
                    <div className="category-budget-row">
                      <span className="cb-label">Available Budget</span>
                      <span className="cb-value">{formatAmount(entry.availableBudget)}</span>
                    </div>
                    <div className="category-budget-row">
                      <span className="cb-label">This Requisition</span>
                      <span className={`cb-value ${overBudget ? 'danger' : ''}`}>
                        {formatAmount(entry.thisRequisition)}
                      </span>
                    </div>
                    <div className="category-budget-row cb-foot">
                      <span className="cb-label">Remaining</span>
                      <span className={`cb-value ${entry.remaining < 0 ? 'danger' : 'ok'}`}>
                        {formatAmount(entry.remaining)}
                      </span>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        ) : null}

        <div className="table-scroll table-space-top">
          <table className="data-table">
            <thead>
              <tr>
                {ASSET_UPLOAD_COLUMNS.map((column) => (
                  <th key={column}>{column}</th>
                ))}
                <th className="actions-col">&nbsp;</th>
              </tr>
            </thead>
            <tbody>
              {rows.length ? (
                rows.map((row) => {
                  const isEditing = editingRowId === row.id;
                  const source = isEditing ? draftRows[row.id] ?? row : row;

                  return (
                    <tr key={row.id}>
                      {ASSET_UPLOAD_COLUMNS.map((column) => (
                        <td key={`${row.id}-${column}`}>
                          {isEditing ? (
                            column === 'Department' ? (
                              <select
                                className="cell-input"
                                value={source[column] ?? ''}
                                onChange={(event) => updateDraftCell(row.id, column, event.target.value)}
                              >
                                <option value="">Select Department</option>
                                {DEPARTMENTS.map((department) => (
                                  <option key={department} value={department}>
                                    {department}
                                  </option>
                                ))}
                              </select>
                            ) : column === 'Asset Category' ? (
                              <select
                                className="cell-input"
                                value={source[column] ?? ''}
                                onChange={(event) => updateDraftCell(row.id, column, event.target.value)}
                              >
                                <option value="">Select Category</option>
                                {ASSET_CATEGORIES.map((category) => (
                                  <option key={category} value={category}>
                                    {category}
                                  </option>
                                ))}
                              </select>
                            ) : (
                              <input
                                className="cell-input"
                                value={source[column] ?? ''}
                                onChange={(event) => updateDraftCell(row.id, column, event.target.value)}
                              />
                            )
                          ) : (
                            <span>{source[column]}</span>
                          )}
                        </td>
                      ))}
                      <td className="row-actions">
                        <button
                          className="icon-button"
                          aria-label={isEditing ? 'Save row' : 'Edit row'}
                          onClick={() => (isEditing ? saveEdit(row.id) : startEdit(row))}
                        >
                          {isEditing ? <Check size={18} /> : <Edit3 size={18} />}
                        </button>
                        <button
                          className="icon-button danger"
                          aria-label="Delete row"
                          onClick={() => deleteRow(row.id)}
                        >
                          <Trash2 size={18} />
                        </button>
                      </td>
                    </tr>
                  );
                })
              ) : (
                <tr>
                  <td colSpan={ASSET_UPLOAD_COLUMNS.length + 1}>
                    <div className="empty-state-inline">No line items yet. Upload an Excel file to view your requisition lines here.</div>
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>

        <div className="submit-row">
          <button className="primary-button" onClick={submitRequisition} disabled={!rows.length}>
            Submit Requisition <ArrowRight size={18} />
          </button>
        </div>
      </section>
    </div>
  );
}

function AssetRequestDetailsPage({ assetRequests, setAssetRequests, setBanner }) {
  const { id } = useParams();
  const navigate = useNavigate();
  const [draftRows, setDraftRows] = useState({});
  const [isViewOnly, setIsViewOnly] = useState(true);

  const request = assetRequests.find((item) => item.id === id);

  useEffect(() => {
    if (!request) return;
    const initialDrafts = {};
    request.rows.forEach((row) => {
      initialDrafts[row.id] = { ...row };
    });
    setDraftRows(initialDrafts);
  }, [request]);

  if (!request) {
    return (
      <div className="page-grid">
        <section className="content-card">
          <h2>Asset Request Not Found</h2>
          <p>The selected asset request could not be found.</p>
        </section>
      </div>
    );
  }

  const isEditable = request.status === 'Pending for Procurement' || request.status === 'Pending for HOD approval';
  const canApproveByHOD = request.status === 'Pending for HOD approval';

  const updateAmount = (rowId, value) => {
    setDraftRows((current) => ({
      ...current,
      [rowId]: {
        ...(current[rowId] || {}),
        Amount: value,
      },
    }));
  };

  const saveAndSendToHOD = () => {
    const now = new Date().toLocaleString();
    const updatedRows = request.rows.map((row) => ({
      ...row,
      Amount: draftRows[row.id]?.Amount ?? row.Amount ?? '',
    }));
    const nextStatus = 'Pending for HOD approval';
    const totalAmount = updatedRows.reduce((sum, row) => sum + parseAmount(row.Amount), 0);

    setAssetRequests((current) =>
      current.map((item) =>
        item.id === request.id
          ? {
              ...item,
              rows: updatedRows,
              totalAmount,
              status: nextStatus,
              procurementSavedAt: now,
              approvalTrail: buildAssetApprovalTrail(
                {
                  ...item,
                  rows: updatedRows,
                  procurementSavedAt: now,
                },
                nextStatus,
                now
              ),
              progress: getProgressForStatus(nextStatus),
              mode: getModeForStatus(nextStatus),
            }
          : item
      )
    );

    setBanner(`Asset request ${request.id} saved and sent to HOD.`);
  };

  const approveByHOD = () => {
    const now = new Date().toLocaleString();
    const nextStatus = 'Pending for Director HR/Admin approval';

    setAssetRequests((current) =>
      current.map((item) =>
        item.id === request.id
          ? {
              ...item,
              status: nextStatus,
              hodApprovedAt: now,
              approvalTrail: buildAssetApprovalTrail(
                {
                  ...item,
                  hodApprovedAt: now,
                },
                nextStatus,
                now
              ),
              progress: getProgressForStatus(nextStatus),
              mode: getModeForStatus(nextStatus),
            }
          : item
      )
    );

    setBanner(`Asset request ${request.id} approved by HOD.`);
  };

  const declineByHOD = () => {
    const now = new Date().toLocaleString();

    setAssetRequests((current) =>
      current.map((item) =>
        item.id === request.id
          ? {
              ...item,
              status: 'Declined',
              declinedAt: now,
              declinedAtStep: 'Head of Department',
              approvalTrail: buildAssetApprovalTrail(
                {
                  ...item,
                  declinedAt: now,
                  declinedAtStep: 'Head of Department',
                },
                'Declined',
                now
              ),
              progress: getProgressForStatus('Declined'),
              mode: getModeForStatus('Declined'),
            }
          : item
      )
    );

    setBanner(`Asset request ${request.id} declined by HOD.`);
  };

  // Amount currently shown for a row (reflects in-progress edits).
  const amountForRow = (row) => parseAmount((draftRows[row.id] || row).Amount);

  // Per-category overview: available budget vs this requisition's draw, for the
  // categories that actually appear in this request's line items.
  const categoryBudgets = useMemo(() => {
    const dept = request.department;
    const fy = String(request.fiscalYear || '');

    const presentCategories = ASSET_CATEGORIES.filter((category) =>
      request.rows.some((row) => (row['Asset Category'] || '') === category)
    );

    return presentCategories.map((category) => {
      const totalBudget = (BUDGET_MATRIX[dept] && BUDGET_MATRIX[dept][category]) || 0;

      const thisRequisition = request.rows
        .filter((row) => (row['Asset Category'] || '') === category)
        .reduce((sum, row) => sum + amountForRow(row), 0);

      // What other (non-declined) requests in the same department + fiscal year already consume.
      const consumedByOthers = assetRequests.reduce((sum, other) => {
        if (other.id === request.id) return sum;
        if (other.status === 'Declined') return sum;
        if ((other.department || '') !== dept) return sum;
        if (String(other.fiscalYear || '') !== fy) return sum;
        const otherCategoryTotal = (other.rows || []).reduce(
          (rowSum, row) => ((row['Asset Category'] || '') === category ? rowSum + parseAmount(row.Amount) : rowSum),
          0
        );
        return sum + otherCategoryTotal;
      }, 0);

      const availableBudget = totalBudget - consumedByOthers;
      const remaining = availableBudget - thisRequisition;

      return { category, totalBudget, thisRequisition, availableBudget, remaining };
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [request, assetRequests, draftRows]);

  return (
    <div className="page-grid">
      {/* Card 1 — Summary: identity + actions + timeline + key facts + category budgets */}
      <section className={`content-card request-summary-card ${request.mode || ''}`}>
        <div className="section-heading-row compact">
          <div>
            <h2>Asset Request</h2>
            <p className="request-id-text">{request.id}</p>
          </div>
          <div className="toolbar-row">
            <button
              className="secondary-button"
              onClick={() => setIsViewOnly(!isViewOnly)}
              disabled={!isEditable}
              style={{ opacity: !isEditable ? 0.5 : 1, cursor: !isEditable ? 'not-allowed' : 'pointer' }}
            >
              <Edit3 size={18} /> {isViewOnly ? 'Edit Request' : 'View Only'}
            </button>
            <button className="secondary-button" onClick={() => navigate('/requisition-tracking')}>
              Back to Tracking
            </button>
          </div>
        </div>

        {/* Approval timeline — first glance at status */}
        <div className="summary-timeline timeline-card-compact">
          <ApprovalTimeline request={{ ...request, mode: request.mode }} />
        </div>

        {/* Key facts */}
        <div className="quick-details-grid summary-facts">
          <div className="quick-detail-item">
            <span className="quick-detail-label">Fiscal Year</span>
            <span className="quick-detail-value">{request.fiscalYear || getCurrentFiscalYear()}</span>
          </div>
          <div className="quick-detail-item">
            <span className="quick-detail-label">Department</span>
            <span className="quick-detail-value">{request.department || '—'}</span>
          </div>
          <div className="quick-detail-item">
            <span className="quick-detail-label">Status</span>
            <span className={`quick-detail-value status-badge ${badgeClassForStatus(request.mode)}`}>
              {request.status}
            </span>
          </div>
        </div>

        {/* Category budget overview */}
        {categoryBudgets.length ? (
          <div className="category-budget-section">
            <h3 className="category-budget-title">Budget by Category</h3>
            <div className="category-budget-grid">
              {categoryBudgets.map((entry) => {
                const overBudget = entry.thisRequisition > entry.availableBudget;
                return (
                  <div className="category-budget-card" key={entry.category}>
                    <div className="category-budget-name">{entry.category}</div>
                    <div className="category-budget-row">
                      <span className="cb-label">Available Budget</span>
                      <span className="cb-value">{formatAmount(entry.availableBudget)}</span>
                    </div>
                    <div className="category-budget-row">
                      <span className="cb-label">This Requisition</span>
                      <span className={`cb-value ${overBudget ? 'danger' : ''}`}>
                        {formatAmount(entry.thisRequisition)}
                      </span>
                    </div>
                    <div className="category-budget-row cb-foot">
                      <span className="cb-label">Remaining</span>
                      <span className={`cb-value ${entry.remaining < 0 ? 'danger' : 'ok'}`}>
                        {formatAmount(entry.remaining)}
                      </span>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        ) : null}
      </section>

      {/* Card 2 — Line Items (unchanged) */}
      <section className="content-card">
        <div className="section-heading-row compact">
          <div>
            <h2>Asset Line Items</h2>
          </div>
        </div>

        <div className="table-scroll">
          <table className="data-table">
            <thead>
              <tr>
                {ASSET_DETAIL_COLUMNS.map((column) => (
                  <th key={column}>{column}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {request.rows.map((row) => {
                const source = draftRows[row.id] || row;
                return (
                  <tr key={row.id}>
                    {ASSET_DETAIL_COLUMNS.map((column) => (
                      <td key={`${row.id}-${column}`}>
                        {column === 'Amount' ? (
                          isViewOnly || !isEditable ? (
                            <span>{formatAmount(source[column])}</span>
                          ) : (
                            <input
                              className="cell-input"
                              value={source[column] ?? ''}
                              onChange={(event) => updateAmount(row.id, event.target.value)}
                            />
                          )
                        ) : (
                          <span>{source[column]}</span>
                        )}
                      </td>
                    ))}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        {isEditable ? (
          <div className="submit-row">
            <button className="primary-button" onClick={saveAndSendToHOD}>
              Save Amounts <ArrowRight size={18} />
            </button>
            {canApproveByHOD ? (
              <>
                <button className="secondary-button" onClick={approveByHOD}>
                  Approve by HOD
                </button>
                <button className="secondary-button danger-button" onClick={declineByHOD}>
                  Decline by HOD
                </button>
              </>
            ) : null}
          </div>
        ) : null}
      </section>
    </div>
  );
}
const reindexBudgetRows = (rows) => {
  return rows.map((row, index) => ({
    ...row,
    'Serial Number': String(index + 1),
  }));
};

function BudgetUploadPage({
  budgetRows,
  setBudgetRows,
  budgetRequests,
  setBudgetRequests,
  editingBudgetRequestId,
  setEditingBudgetRequestId,
  setBanner,
}) {
  const navigate = useNavigate();
  const [editingRowId, setEditingRowId] = useState(null);
  const [draftRows, setDraftRows] = useState({});
  const [uploadMessage, setUploadMessage] = useState('');
  const currentYear = getCurrentFiscalYear();

  const handleDownloadTemplate = async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Budget Upload');

    const headerRow = worksheet.addRow(BUDGET_COLUMNS);
    headerRow.eachCell((cell) => {
      cell.font = { bold: true };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'EAF3EE' },
      };
      cell.border = {
        top: { style: 'thin', color: { argb: 'D0D7DE' } },
        left: { style: 'thin', color: { argb: 'D0D7DE' } },
        bottom: { style: 'thin', color: { argb: 'D0D7DE' } },
        right: { style: 'thin', color: { argb: 'D0D7DE' } },
      };
    });

    for (let i = 0; i < 50; i += 1) {
      worksheet.addRow([i + 1, currentYear, '', '', '']);
    }

    worksheet.dataValidations.add('C2:C200', {
      type: 'list',
      allowBlank: true,
      formulae: [`"${DEPARTMENTS.join(',')}"`],
      showErrorMessage: true,
      errorTitle: 'Invalid Department',
      error: 'Please select a department from the dropdown list.',
    });

    worksheet.dataValidations.add('D2:D200', {
      type: 'list',
      allowBlank: true,
      formulae: [`"${ASSET_CATEGORIES.join(',')}"`],
      showErrorMessage: true,
      errorTitle: 'Invalid Category',
      error: 'Please select an asset category from the dropdown list.',
    });

    [18, 15, 18, 30, 18].forEach((width, index) => {
      worksheet.getColumn(index + 1).width = width;
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });

    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = 'muntq-budget-upload-template.xlsx';
    anchor.click();
    window.URL.revokeObjectURL(url);

    setUploadMessage('Budget template downloaded successfully.');
    setBanner('Budget template ready. You can fill it later and upload the Excel file anytime.');
  };

  const sanitizeRow = (item, index = 0) => {
    const normalized = normalizeExcelRow(item);
    return {
      id: crypto.randomUUID(),
      'Serial Number': String(index + 1),
      'Fiscal Year': String(normalized['fiscal year'] || currentYear),
      'Department': normalized['department'] ?? '',
      'Category': normalized['category'] ?? '',
      'Budget Amount': normalized['budget amount'] ?? '',
    };
  };

  const handleFileUpload = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const jsonRows = XLSX.utils.sheet_to_json(sheet, { defval: '' });

      const cleanedRows = jsonRows
        .map((row, index) => sanitizeRow(row, index))
        .filter((row) =>
          ['Department', 'Category', 'Budget Amount'].some(
            (column) => String(row[column] ?? '').trim() !== ''
          )
        );

      if (!cleanedRows.length) {
        setUploadMessage('No valid budget rows were found in the uploaded file.');
        return;
      }

      setBudgetRows((current) => [...current, ...cleanedRows]);
      setUploadMessage(`Budget file uploaded successfully. ${cleanedRows.length} line item(s) added.`);
      setBanner('Budget upload successful. The new budget lines were added below the existing items.');
      event.target.value = '';
    } catch (error) {
      setUploadMessage('The file could not be processed. Please use an Excel file with the required columns.');
    }
  };

  const startEdit = (row) => {
    setEditingRowId(row.id);
    setDraftRows((current) => ({ ...current, [row.id]: { ...row } }));
  };

  const updateDraftCell = (rowId, column, value) => {
    setDraftRows((current) => ({
      ...current,
      [rowId]: {
        ...(current[rowId] || {}),
        [column]: value,
      },
    }));
  };

  const saveEdit = (rowId) => {
    setBudgetRows((current) => {
      const updated = current.map((row) =>
        row.id === rowId
          ? {
              ...draftRows[rowId],
              'Serial Number': row['Serial Number'],
            }
          : row
      );
      return reindexBudgetRows(updated);
    });

    setEditingRowId(null);
    setUploadMessage('Budget line updated successfully.');
  };

  const deleteRow = (rowId) => {
    setBudgetRows((current) => {
      const filtered = current.filter((row) => row.id !== rowId);
      return reindexBudgetRows(filtered);
    });

    if (editingRowId === rowId) {
      setEditingRowId(null);
    }

    setUploadMessage('Budget line deleted successfully.');
  };

  const addBudgetLineItem = () => {
    const newRow = {
      id: crypto.randomUUID(),
      'Serial Number': String(budgetRows.length + 1),
      'Fiscal Year': currentYear,
      'Department': '',
      'Category': '',
      'Budget Amount': '',
    };

    setBudgetRows((current) => [...current, newRow]);
    setEditingRowId(newRow.id);
    setDraftRows((current) => ({
      ...current,
      [newRow.id]: { ...newRow },
    }));
  };

  const submitBudget = () => {
    if (!budgetRows.length) return;

    const submittedAt = new Date().toLocaleString();

    if (editingBudgetRequestId) {
      const updatedRequestId = editingBudgetRequestId;

      setBudgetRequests((current) =>
        current.map((request) =>
          request.id === updatedRequestId
            ? {
                ...request,
                submittedAt,
                status: 'Submitted',
                rows: budgetRows.map((row) => ({ ...row })),
              }
            : request
        )
      );

      setBanner(`Budget request ${updatedRequestId} updated successfully.`);
      setEditingBudgetRequestId(null);
      setBudgetRows([]);
      return;
    }

    const newRequest = {
      id: `BUD-${Date.now()}`,
      submittedAt,
      status: 'Submitted',
      rows: budgetRows.map((row) => ({ ...row })),
    };

    setBudgetRequests((current) => [...current, newRequest]);
    setBanner(`Budget request ${newRequest.id} submitted successfully.`);
    setBudgetRows([]);
  };

  return (
    <div className="page-grid">
      <section className="content-card">
        <div className="section-heading-row">
          <div>
            <h2>Upload Budget</h2>
            <p>Download the budget template if needed, upload one or multiple filled files and edit each row inline at any time.</p>
          </div>
          <div className="toolbar-row">
            <button className="secondary-button" onClick={handleDownloadTemplate}>
              <Download size={18} /> Download Template
            </button>
            <label className="primary-button upload-button">
              <Upload size={18} /> Upload Excel
              <input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} />
            </label>
            <button className="secondary-button" onClick={addBudgetLineItem}>
              Add Line Item
            </button>
          </div>
        </div>

        {uploadMessage ? <div className="success-message">{uploadMessage}</div> : null}

        {editingBudgetRequestId ? (
          <div className="success-message">You are editing submitted budget request {editingBudgetRequestId}.</div>
        ) : null}

        <div className="table-scroll table-space-top">
          <table className="data-table">
            <thead>
              <tr>
                {BUDGET_COLUMNS.map((column) => (
                  <th key={column}>{column}</th>
                ))}
                <th className="actions-col">&nbsp;</th>
              </tr>
            </thead>
            <tbody>
              {budgetRows.length ? (
                budgetRows.map((row) => {
                  const isEditing = editingRowId === row.id;
                  const source = isEditing ? draftRows[row.id] ?? row : row;

                  return (
                    <tr key={row.id}>
                      {BUDGET_COLUMNS.map((column) => (
                        <td key={`${row.id}-${column}`}>
                          {column === 'Serial Number' ? (
                            <span>{source[column]}</span>
                          ) : isEditing ? (
                            column === 'Department' ? (
                              <select
                                className="cell-input"
                                value={source[column] ?? ''}
                                onChange={(event) => updateDraftCell(row.id, column, event.target.value)}
                              >
                                <option value="">Select Department</option>
                                {DEPARTMENTS.map((dept) => (
                                  <option key={dept} value={dept}>
                                    {dept}
                                  </option>
                                ))}
                              </select>
                            ) : column === 'Category' ? (
                              <select
                                className="cell-input"
                                value={source[column] ?? ''}
                                onChange={(event) => updateDraftCell(row.id, column, event.target.value)}
                              >
                                <option value="">Select Category</option>
                                {ASSET_CATEGORIES.map((cat) => (
                                  <option key={cat} value={cat}>
                                    {cat}
                                  </option>
                                ))}
                              </select>
                            ) : (
                              <input
                                className="cell-input"
                                value={source[column] ?? ''}
                                onChange={(event) => updateDraftCell(row.id, column, event.target.value)}
                              />
                            )
                          ) : column === 'Budget Amount' ? (
                            <span>{formatAmount(source[column])}</span>
                          ) : (
                            <span>{source[column]}</span>
                          )}
                        </td>
                      ))}

                      <td className="row-actions">
                        <button
                          className="icon-button"
                          aria-label={isEditing ? 'Save row' : 'Edit row'}
                          onClick={() => (isEditing ? saveEdit(row.id) : startEdit(row))}
                        >
                          {isEditing ? <Check size={18} /> : <Edit3 size={18} />}
                        </button>
                        <button
                          className="icon-button danger"
                          aria-label="Delete row"
                          onClick={() => deleteRow(row.id)}
                        >
                          <Trash2 size={18} />
                        </button>
                      </td>
                    </tr>
                  );
                })
              ) : (
                <tr>
                  <td colSpan={BUDGET_COLUMNS.length + 1}>
                    <div className="empty-state-inline">No budget lines yet. Upload an Excel file to view your budget rows here.</div>
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>

        <div className="submit-row">
          <button className="primary-button" onClick={submitBudget} disabled={!budgetRows.length}>
            {editingBudgetRequestId ? 'Update Budget' : 'Submit Budget'} <ArrowRight size={18} />
          </button>
        </div>
      </section>

      <section className="content-card" style={{ marginTop: '24px' }}>
        <div className="section-heading-row compact">
          <div>
            <h2>Submitted Budget History</h2>
            <p>Each submitted budget request can be reopened, edited, updated, or deleted.</p>
          </div>
        </div>

        <div className="table-scroll">
          <table className="data-table">
            <thead>
              <tr>
                <th>Request ID</th>
                <th>Submitted Date</th>
                <th>Fiscal Year</th>
                <th>Department</th>
                <th>Line Items</th>
                <th>Total Budget Amount</th>
                <th className="actions-col">&nbsp;</th>
              </tr>
            </thead>
            <tbody>
              {budgetRequests.length ? (
                budgetRequests.map((request) => (
                  <tr key={request.id}>
                    <td>
                      <Link className="inline-link-button" to={`/budget-request/${request.id}`}>{request.id}</Link>
                    </td>
                    <td>{request.submittedAt}</td>
                    <td>{summarizeField(request.rows, 'Fiscal Year')}</td>
                    <td>{summarizeField(request.rows, 'Department')}</td>
                    <td>{request.rows.length}</td>
                    <td>{formatAmount(getBudgetRequestTotal(request.rows))}</td>
                    <td className="row-actions">
                      <button
                        className="icon-button"
                        aria-label="Edit budget request"
                        onClick={() => {
                          setEditingBudgetRequestId(request.id);
                          setBudgetRows(request.rows.map((row) => ({ ...row })));
                          setBanner(`Editing budget request ${request.id}. Update the rows and click Update Budget.`);
                          navigate('/budget-upload');
                        }}
                      >
                        <Edit3 size={18} />
                      </button>
                      <button
                        className="icon-button danger"
                        aria-label="Delete budget request"
                        onClick={() => {
                          setBudgetRequests((current) => current.filter((item) => item.id !== request.id));

                          if (editingBudgetRequestId === request.id) {
                            setEditingBudgetRequestId(null);
                            setBudgetRows([]);
                          }

                          setBanner(`Budget request ${request.id} deleted successfully.`);
                        }}
                      >
                        <Trash2 size={18} />
                      </button>
                    </td>
                  </tr>
                ))
              ) : (
                <tr>
                  <td colSpan={7}>
                    <div className="empty-state-inline">No submitted budget requests yet.</div>
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </section>
    </div>
  );
}

function BudgetRequestDetailsPage({ budgetRequests }) {
  const { id } = useParams();
  const navigate = useNavigate();

  const request = budgetRequests.find((item) => item.id === id);

  if (!request) {
    return (
      <div className="page-grid">
        <section className="content-card">
          <h2>Budget Request Not Found</h2>
          <p>The selected budget request could not be found.</p>
        </section>
      </div>
    );
  }

  return (
    <div className="page-grid">
      <section className="content-card">
        <div className="section-heading-row compact">
          <div>
            <h2>Budget Request Details</h2>
            <p>{request.id}</p>
          </div>
          <div className="toolbar-row">
            <button className="secondary-button" onClick={() => navigate('/budget-upload')}>
              Back to Upload Budget
            </button>
          </div>
        </div>

        <div className="budget-summary-row">
          <div className="budget-card">
            <div className="budget-label">Request ID</div>
            <div className="budget-value">{request.id}</div>
          </div>
          <div className="budget-card">
            <div className="budget-label">Submitted Date</div>
            <div className="budget-value">{request.submittedAt}</div>
          </div>
        </div>
      </section>

      <section className="content-card">
        <div className="section-heading-row compact">
          <div>
            <h2>Budget Line Items</h2>
          </div>
        </div>

        <div className="table-scroll">
          <table className="data-table">
            <thead>
              <tr>
                {BUDGET_COLUMNS.map((column) => (
                  <th key={column}>{column}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {request.rows.map((row) => (
                <tr key={row.id}>
                  {BUDGET_COLUMNS.map((column) => (
                    <td key={`${row.id}-${column}`}>
                      {column === 'Budget Amount' ? formatAmount(row[column]) : row[column]}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>
    </div>
  );
}

function RequisitionTrackingPage({ assetRequests }) {
  const navigate = useNavigate();
  const [expandedRequestId, setExpandedRequestId] = useState(null);

  const toggleRequest = (requestId) => {
    setExpandedRequestId((current) => (current === requestId ? null : requestId));
  };

  const computeRequestAmount = (request) =>
    request.rows.reduce((sum, row) => sum + parseAmount(row.Amount), 0);

  return (
    <div className="page-grid">
      <section className="content-card">
        <div className="section-heading-row compact">
          <div>
            <h2>Request History</h2>
            <p>Click a request to show or hide its timeline.</p>
          </div>
        </div>

        <div className="table-scroll">
          <table className="data-table tracking-table">
            <thead>
              <tr>
                <th>Request ID</th>
                <th>Submitted Date</th>
                <th>Fiscal Year</th>
                <th>Department</th>
                <th>Category</th>
                <th>Amount</th>
                <th>Status</th>
                <th className="actions-col">&nbsp;</th>
              </tr>
            </thead>
            <tbody>
              {assetRequests.length ? (
                assetRequests.map((request) => {
                  const amount = request.totalAmount || computeRequestAmount(request);
                  const isExpanded = expandedRequestId === request.id;
                  const requestFiscalYear = request.fiscalYear || request.rows?.[0]?.['Fiscal Year'] || getCurrentFiscalYear();
                  const requestDepartment = request.department || summarizeField(request.rows || [], 'Department');
                  const requestCategory = request.category || summarizeField(request.rows || [], 'Asset Category');
                  const mode = getModeForStatus(request.status);
                  const isEditable = request.status === 'Pending for Procurement' || request.status === 'Pending for HOD approval';

                  return (
                    <React.Fragment key={request.id}>
                      <tr className={isExpanded ? 'selected-row' : ''}>
                        <td>
  <button 
    className="inline-link-button" 
    onClick={() => navigate(`/asset-request/${request.id}`)}
  >
    {request.id}
  </button>
</td>
                        <td>{request.submittedAt}</td>
                        <td>{requestFiscalYear}</td>
                        <td>{requestDepartment || '—'}</td>
                        <td>{requestCategory || '—'}</td>
                        <td>{formatAmount(amount || 0)}</td>
                        <td>
                          <span className={`status-badge ${badgeClassForStatus(mode)}`}>{request.status}</span>
                        </td>
                        <td className="row-actions">
                          <button
                            className="icon-button"
                            aria-label="Edit asset request"
                            disabled={!isEditable}
                            onClick={() => navigate(`/asset-request/${request.id}`)}
                          >
                            <Edit3 size={18} />
                          </button>
                        </td>
                      </tr>
                    </React.Fragment>
                  );
                })
              ) : (
                <tr>
                  <td colSpan={8}>
                    <div className="empty-state-inline">No asset requests yet. Submit a request from the Asset Requisition screen.</div>
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </section>
    </div>
  );
}

function badgeClassForStatus(mode) {
  if (mode === 'complete') return 'badge-complete';
  if (mode === 'declined') return 'badge-declined';
  return 'badge-pending';
}

function ApprovalTimeline({ request }) {
  // Display-only short labels so all 7 steps fit on one line; logic still uses item.step.
  const SHORT_STEP_LABELS = {
    'Requester': 'Requester',
    'Procurement': 'Procurement',
    'Head of Department': 'HOD',
    'Director HR/Admin': 'Director',
    'Finance Head': 'Finance Head',
    'VP Strategy': 'VP Strategy',
    'Finance Accountant': 'Accountant',
  };

  const lineClass =
    request.mode === 'declined'
      ? 'timeline-line declined'
      : request.mode === 'complete'
      ? 'timeline-line complete'
      : 'timeline-line';

  const trail = request.approvalTrail && request.approvalTrail.length
    ? request.approvalTrail
    : buildAssetApprovalTrail(request, request.status);

  return (
    <div className="timeline-shell">
      <div className="timeline-scroll">
        <div className="timeline-track seven-steps">
          <div className={lineClass} />
          {trail.map((item) => {
            const circleClasses = ['timeline-circle'];
            if (request.mode === 'declined') circleClasses.push('declined');
            if (request.mode === 'complete') circleClasses.push('complete');
            if (item.action !== 'Pending') circleClasses.push('filled');

            const labelClasses = ['timeline-label'];
            if (request.mode === 'declined') labelClasses.push('declined');
            if (request.mode === 'complete') labelClasses.push('complete');

            return (
              <div className="timeline-step" key={item.step}>
                <div className={circleClasses.join(' ')} aria-hidden="true" />
                <div className={labelClasses.join(' ')}>{SHORT_STEP_LABELS[item.step] || item.step}</div>
                <div className="timeline-meta">{item.action}</div>
                <div className="timeline-meta muted">{item.actedAt || '—'}</div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

function ReportsPage() {
  return (
    <div className="page-grid">
      <section className="content-card coming-soon-card">
        <div className="coming-soon-icon">
          <MapPinned size={28} />
        </div>
        <h2>Coming Soon</h2>
        <p>
          This page is intentionally prepared with the correct theme and structure so the report creation features can be added later without changing the overall experience.
        </p>
      </section>
    </div>
  );
}


const BANK_ACCOUNTS = [
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
const formatMoney = (value) =>
  Number(value || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

const TRANSFER_TO_BANKS = {
  RJHI: ['SA 16 8000 0106 6080 1684 6450', 'SA 16 8000 0106 6080 1684 6451'],
  SAB: ['SA 03 8000 0000 1234 5678 9012', 'SA 03 8000 0000 1234 5678 9013', 'SA 03 8000 0000 1234 5678 9014'],
  BSF: ['SA56 5500 0000 0132 0220 0520', 'SA56 5500 0000 0132 0220 0521', 'SA56 5500 0000 0132 0220 0522', 'SA56 5500 0000 0132 0220 0523'],
};

const BANK_NAMES = ['RJHI', 'SAB', 'BSF'];

function BankTransfersPage({ setBanner }) {
  const [transferAmounts, setTransferAmounts] = useState({});
  const [assignments, setAssignments] = useState({});
  const [assignmentsTo, setAssignmentsTo] = useState({});
  const [selectedBank, setSelectedBank] = useState('');
  const [selectedAccount, setSelectedAccount] = useState('');
  const [transferCompanyCode, setTransferCompanyCode] = useState('');
  const [postingDate, setPostingDate] = useState(getTodayISO());

  const availableAccounts = selectedBank ? TRANSFER_TO_BANKS[selectedBank] : [];

  // Display: add thousands separators to the integer part, keep decimals (and any
  // in-progress trailing dot) exactly as the user typed them. No forced rounding.
  const formatTransferAmount = (value) => {
    if (value === undefined || value === null || value === '') return '';
    const [intPart, decPart] = String(value).split('.');
    const withCommas = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return decPart !== undefined ? `${withCommas}.${decPart}` : withCommas;
  };

  // Store the raw string the user typed: digits + a single dot + up to 2 decimals.
  const handleTransferAmountChange = (sNo, value) => {
    let cleaned = value.replace(/[^0-9.]/g, '');
    const firstDot = cleaned.indexOf('.');
    if (firstDot !== -1) {
      // keep only the first dot, then limit to 2 decimal places
      const intPart = cleaned.slice(0, firstDot);
      const decPart = cleaned.slice(firstDot + 1).replace(/\./g, '').slice(0, 2);
      cleaned = `${intPart}.${decPart}`;
    }
    setTransferAmounts((current) => ({
      ...current,
      [sNo]: cleaned,
    }));
  };

  // On blur, pad a real value to 2 decimals (e.g. "99999" -> "99999.00").
  const handleTransferAmountBlur = (sNo) => {
    setTransferAmounts((current) => {
      const value = current[sNo];
      if (value === undefined || value === '' || value === '.') return current;
      const num = parseFloat(String(value));
      if (Number.isNaN(num)) return { ...current, [sNo]: '' };
      return { ...current, [sNo]: num.toFixed(2) };
    });
  };

  const handleAssignmentChange = (sNo, value) => {
    const validValue = value.replace(/[^a-zA-Z0-9-]/g, '');
    setAssignments((current) => ({
      ...current,
      [sNo]: validValue,
    }));
  };

  const handleAssignmentToChange = (sNo, value) => {
    const validValue = value.replace(/[^a-zA-Z0-9-]/g, '');
    setAssignmentsTo((current) => ({
      ...current,
      [sNo]: validValue,
    }));
  };

  // A row is postable when it has a transfer amount (> 0). Assignment is optional.
  const isRowPostable = (sNo) => {
    const amount = transferAmounts[sNo];
    return amount !== '' && amount !== undefined && !Number.isNaN(Number(amount)) && Number(amount) > 0;
  };

  const postableRows = BANK_ACCOUNTS.filter((account) => isRowPostable(account.sNo));

  // Numeric transfer value entered for a row (0 when empty/invalid).
  const transferValueFor = (sNo) => {
    const amount = transferAmounts[sNo];
    if (amount === '' || amount === undefined) return 0;
    const n = Number(amount);
    return Number.isNaN(n) ? 0 : n;
  };

  // A row is in error when its transfer amount exceeds that account's available amount.
  const isRowOverAvailable = (account) => transferValueFor(account.sNo) > account.availableAmount;

  const hasOverAvailableError = BANK_ACCOUNTS.some((account) => isRowOverAvailable(account));

  // Post to SAP is enabled when the header is complete, at least one row has an amount,
  // AND no row exceeds its available amount.
  const canPost =
    Boolean(selectedBank) &&
    Boolean(selectedAccount) &&
    Boolean(transferCompanyCode.trim()) &&
    Boolean(postingDate) &&
    postableRows.length > 0 &&
    !hasOverAvailableError;

  const handlePostToSAP = () => {
    if (hasOverAvailableError) {
      setBanner('One or more transfer amounts exceed the available amount. Reduce them before posting to SAP.');
      return;
    }
    if (!canPost) {
      setBanner(
        'Select House Bank, Account Number, Company Code, Posting Date and enter a Transfer Amount in at least one row.'
      );
      return;
    }

    // Every row with a transfer amount is posted; assignment is included when provided.
    const postedLines = postableRows.map((account) => ({
      sNo: account.sNo,
      houseBank: account.houseBank,
      accountNumber: account.accountNumber,
      accountId: account.accountId,
      transferAmount: transferAmounts[account.sNo],
      assignmentFrom: assignments[account.sNo] ?? '',
      assignmentTo: assignmentsTo[account.sNo] ?? '',
    }));

    console.log('Posting to SAP', {
      transferToBank: selectedBank,
      transferToAccount: selectedAccount,
      transferToCompanyCode: transferCompanyCode,
      postingDate,
      lines: postedLines,
    });

    // Show a single message via the global banner (it auto-dismisses after a few seconds).
    setBanner(`Your data has been posted to SAP successfully! ${postedLines.length} line item(s) posted.`);

    // Reset the form immediately after posting.
    setTransferAmounts({});
    setAssignments({});
    setAssignmentsTo({});
    setSelectedBank('');
    setSelectedAccount('');
    setTransferCompanyCode('');
    setPostingDate(getTodayISO());
  };

  const getTodayDate = () => getTodayISO();

  return (
    <div className="page-grid">
      <section className="content-card">
        <div className="section-heading-row">
          <div>
            <h2>Bank Transfers</h2>
            <p>Select transfer details and edit transfer amounts and assignments for each account.</p>
          </div>
        </div>

        {/* Transfer To Section */}
        <div className="transfer-section">
          <h3 className="transfer-section-title">Transfer To</h3>
          <div className="transfer-controls-grid">
            <div className="transfer-control">
              <label className="control-label">Company Code <span className="req-star">*</span></label>
              <input
                type="text"
                className="cell-input"
                value={transferCompanyCode}
                onChange={(e) => setTransferCompanyCode(e.target.value)}
                placeholder="e.g., 1000"
              />
            </div>

            <div className="transfer-control">
              <label className="control-label">House Bank <span className="req-star">*</span></label>
              <select
                className="cell-input"
                value={selectedBank}
                onChange={(e) => {
                  setSelectedBank(e.target.value);
                  setSelectedAccount('');
                }}
              >
                <option value="">Select Bank</option>
                {BANK_NAMES.map((bank) => (
                  <option key={bank} value={bank}>
                    {bank}
                  </option>
                ))}
              </select>
            </div>

            <div className="transfer-control">
              <label className="control-label">Account Number <span className="req-star">*</span></label>
              <select
                className="cell-input"
                value={selectedAccount}
                onChange={(e) => setSelectedAccount(e.target.value)}
                disabled={!selectedBank}
              >
                <option value="">Select Account</option>
                {availableAccounts.map((account) => (
                  <option key={account} value={account}>
                    {account}
                  </option>
                ))}
              </select>
            </div>

            <div className="transfer-control">
              <label className="control-label">Posting Date</label>
              <input
                type="date"
                className="cell-input"
                value={postingDate}
                onChange={(e) => setPostingDate(e.target.value)}
                max={getTodayDate()}
              />
            </div>
          </div>
        </div>

        {/* Bank Accounts Table */}
        <div className="table-scroll table-space-top">
          <table className="data-table transfers-table">
            <thead>
              <tr>
                <th>S. No</th>
                <th>House Bank</th>
                <th>Company Code</th>
                <th>Account Number (BANKN)</th>
                <th>Account ID</th>
                <th>Available Amount</th>
                <th>Transfer Amount</th>
                <th>Assignment (From)</th>
                <th>Assignment (To)</th>
              </tr>
            </thead>
            <tbody>
              {BANK_ACCOUNTS.map((account) => {
                const overAvailable = isRowOverAvailable(account);
                return (
                  <tr key={account.sNo}>
                    <td>{account.sNo}</td>
                    <td>{account.houseBank}</td>
                    <td>{account.companyCode}</td>
                    <td>{account.accountNumber}</td>
                    <td>{account.accountId}</td>
                    <td>{formatMoney(account.availableAmount)}</td>
                    <td>
                      <input
                        type="text"
                        inputMode="decimal"
                        className={overAvailable ? 'cell-input input-error' : 'cell-input'}
                        value={formatTransferAmount(transferAmounts[account.sNo])}
                        onChange={(e) => handleTransferAmountChange(account.sNo, e.target.value)}
                        onBlur={() => handleTransferAmountBlur(account.sNo)}
                        placeholder="0.00"
                        aria-invalid={overAvailable}
                      />
                      {overAvailable ? (
                        <div className="cell-error">
                          Exceeds available amount ({formatMoney(account.availableAmount)})
                        </div>
                      ) : null}
                    </td>
                    <td>
                      <input
                        type="text"
                        className="cell-input"
                        value={assignments[account.sNo] ?? ''}
                        onChange={(e) => handleAssignmentChange(account.sNo, e.target.value)}
                        placeholder="e.g., LKIO-003444"
                        maxLength="20"
                      />
                    </td>
                    <td>
                      <input
                        type="text"
                        className="cell-input"
                        value={assignmentsTo[account.sNo] ?? ''}
                        onChange={(e) => handleAssignmentToChange(account.sNo, e.target.value)}
                        placeholder="e.g., LKIO-003444"
                        maxLength="20"
                      />
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        {/* Post to SAP Button */}
        <div className="submit-row">
          <button className="primary-button" onClick={handlePostToSAP} disabled={!canPost}>
            Post to SAP <ArrowRight size={18} />
          </button>
        </div>
      </section>
    </div>
  );
}
export default App;
