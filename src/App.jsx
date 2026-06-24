import React, { useState, useEffect } from 'react';
import { Route, Routes, useLocation } from 'react-router-dom';
import { AssetRequestDetailsPage, AssetRequisitionPage } from './AssetRequisition';
import { BankTransfersPage } from './BankTransfers';
import { BudgetRequestDetailsPage, BudgetUploadPage } from './Budget';
import { LandingPage } from './LandingPage';
import { Footer, Header } from './Layout';
import { ReportsPage } from './Reports';
import { RequisitionTrackingPage } from './RequisitionTracking';
import { SADADPaymentsPage } from './SadadPayments';

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


export default App;

