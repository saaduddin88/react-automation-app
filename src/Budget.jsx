import React, { useState } from 'react';
import { Link, useNavigate, useParams } from 'react-router-dom';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { ArrowRight, Check, Download, Edit3, Trash2, Upload } from 'lucide-react';
import { ASSET_CATEGORIES, BUDGET_COLUMNS, DEPARTMENTS } from './constants';
import { formatAmount, getBudgetRequestTotal, getCurrentFiscalYear, normalizeExcelRow, reindexBudgetRows, summarizeField } from './utils';

export function BudgetUploadPage({
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


export function BudgetRequestDetailsPage({ budgetRequests }) {
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

