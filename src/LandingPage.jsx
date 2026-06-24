import React from 'react';
import { Link } from 'react-router-dom';
import { ArrowRight, CreditCard, Download, FileSpreadsheet, LayoutDashboard, ShieldCheck, Upload } from 'lucide-react';

export const sampleLandingCards = [
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


export function LandingPage() {
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

