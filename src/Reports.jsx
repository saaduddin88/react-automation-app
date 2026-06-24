import React from 'react';
import { MapPinned } from 'lucide-react';

export function ReportsPage() {
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


