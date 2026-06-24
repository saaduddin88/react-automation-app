import React, { useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { Edit3 } from 'lucide-react';
import { badgeClassForStatus, formatAmount, getCurrentFiscalYear, getModeForStatus, parseAmount, summarizeField } from './utils';

export function RequisitionTrackingPage({ assetRequests }) {
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

