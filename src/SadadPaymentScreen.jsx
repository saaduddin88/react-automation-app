import React, { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { ArrowRight, Download, Trash2, Upload } from 'lucide-react';
import { PAYMENT_COLUMNS, classifyInvoiceCategory, formatDateDDMMYYYY, paymentRowActive, validatePaymentItem } from './sadadValidation';
import { formatMoney } from './utils';

export function IqamaPaymentScreen({ setBanner, categoryKey, categoryName }) {
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

