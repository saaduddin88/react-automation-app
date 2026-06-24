import React, { useState, useEffect, useMemo } from 'react';
import { useNavigate, useParams } from 'react-router-dom';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { AlertTriangle, ArrowRight, Check, Download, Edit3, Trash2, Upload } from 'lucide-react';
import { ASSET_CATEGORIES, ASSET_DETAIL_COLUMNS, ASSET_UPLOAD_COLUMNS, BUDGET_MATRIX, DEPARTMENTS } from './constants';
import { badgeClassForStatus, buildAssetApprovalTrail, formatAmount, getCurrentFiscalYear, getModeForStatus, getProgressForStatus, normalizeExcelRow, parseAmount } from './utils';

export function AssetRequisitionPage({ rows, setRows, setBanner, assetRequests, setAssetRequests }) {
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


export function AssetRequestDetailsPage({ assetRequests, setAssetRequests, setBanner }) {
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

export function ApprovalTimeline({ request }) {
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

