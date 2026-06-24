import React, { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { ArrowRight, Download, Trash2, Upload } from 'lucide-react';
import { IQAMA_COLUMNS, IQAMA_EXPENSE_TYPES, IQAMA_NUMBER_LENGTH, iqamaRowActive, validateIqamaItem } from './sadadValidation';
import { formatMoney, getTodayISO } from './utils';

export function IqamaExpenseScreen({ setBanner }) {
  const [items, setItems] = useState([]);
  const [postingDate, setPostingDate] = useState(getTodayISO());
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

  // Postable only when there is at least one line item, a posting date is set,
  // every row is filled (active), and no row has any validation error.
  const canPost =
    items.length > 0 &&
    postingDate.trim() !== '' &&
    items.every((it) => iqamaRowActive(it) && Object.keys(validateIqamaItem(it)).length === 0);

  const handlePostToSAP = () => {
    if (!postingDate.trim()) {
      setBanner('Please select a Posting Date before posting to SAP.');
      return;
    }
    if (!canPost) {
      setBanner('Cannot post to SAP. Fix the highlighted errors and remove any empty line items first.');
      return;
    }
    const count = items.length;
    setItems([]);
    setBanner(`Posted to SAP successfully! ${count} line item(s) posted with posting date ${postingDate}. You can start a new upload.`);
  };

  return (
    <>
      <div className="sadad-actions">
        <div className="posting-date-field">
          <label htmlFor="iqama-posting-date">
            Posting Date <span className="req-star">*</span>
          </label>
          <input
            id="iqama-posting-date"
            type="date"
            className={postingDate ? 'posting-date-input' : 'posting-date-input input-error'}
            value={postingDate}
            onChange={(ev) => setPostingDate(ev.target.value)}
          />
        </div>
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

