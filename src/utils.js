export const getCurrentFiscalYear = () => String(new Date().getFullYear());

// Local-date 'YYYY-MM-DD' (timezone-safe, so it doesn't roll over a day near midnight).

export const getTodayISO = () => {
  const now = new Date();
  const local = new Date(now.getTime() - now.getTimezoneOffset() * 60000);
  return local.toISOString().split('T')[0];
};


export const formatAmount = (value) => {
  const raw = String(value ?? '').replace(/,/g, '').trim();
  const parsed = parseFloat(raw);
  return Number.isNaN(parsed) ? '0' : parsed.toLocaleString();
};


export const parseAmount = (value) => {
  const raw = String(value ?? '').replace(/,/g, '').trim();
  const parsed = parseFloat(raw);
  return Number.isNaN(parsed) ? 0 : parsed;
};


export const getBudgetRequestTotal = (rows = []) => rows.reduce((sum, row) => sum + parseAmount(row['Budget Amount']), 0);


export const summarizeField = (rows = [], field) => {
  return [...new Set(rows.map((row) => String(row[field] ?? '').trim()).filter(Boolean))].join(', ');
};


export const normalizeExcelRow = (item) => {
  const normalized = {};
  Object.keys(item || {}).forEach((key) => {
    normalized[String(key).trim().toLowerCase()] = item[key];
  });
  return normalized;
};


export const getModeForStatus = (status) => {
  if (status === 'Complete') return 'complete';
  if (status === 'Declined') return 'declined';
  return 'active';
};


export const getProgressForStatus = (status) => {
  if (status === 'Pending for Procurement') return 2;
  if (status === 'Pending for HOD approval') return 3;
  if (status === 'Pending for Director HR/Admin approval') return 4;
  if (status === 'Pending for Finance head approval') return 5;
  if (status === 'Pending for VP Strategy approval') return 6;
  if (status === 'Pending for Accountant approval') return 7;
  if (status === 'Complete') return 7;
  if (status === 'Declined') return 7;
  return 2;
};


export const buildAssetApprovalTrail = (requestLike, status, actedAt = new Date().toLocaleString()) => {
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

  if (status === 'Pending for Accountant approval') {
    setStep('Procurement', 'Completed', request.procurementSavedAt || actedAt);
    setStep('Head of Department', 'Approved', request.hodApprovedAt || actedAt);
    setStep('Director HR/Admin', 'Approved', request.directorApprovedAt || actedAt);
    setStep('Finance Head', 'Approved', request.financeHeadApprovedAt || actedAt);
    setStep('VP Strategy', 'Approved', request.vpStrategyApprovedAt || actedAt);
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


export const reindexBudgetRows = (rows) => {
  return rows.map((row, index) => ({
    ...row,
    'Serial Number': String(index + 1),
  }));
};


export function badgeClassForStatus(mode) {
  if (mode === 'complete') return 'badge-complete';
  if (mode === 'declined') return 'badge-declined';
  return 'badge-pending';
}


export const formatMoney = (value) =>
  Number(value || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

