import React, { useState } from 'react';
import { ArrowRight } from 'lucide-react';
import { BANK_ACCOUNTS, BANK_NAMES, TRANSFER_TO_BANKS } from './constants';
import { formatMoney, getTodayISO } from './utils';

export function BankTransfersPage({ setBanner }) {
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
