import React, { useState } from 'react';
import { Link, NavLink, useLocation } from 'react-router-dom';
import { ChevronDown, Upload } from 'lucide-react';

export function Header() {
  return (
    <header className="site-header">
      <div className="header-inner">
        <Link to="/" className="brand-block" aria-label="MuntQ home">
          <div className="brand-logo-wrap">
            <img src="/muntq-logo.png" alt="MuntQ logo" className="brand-logo" />
          </div>
          <div>
            <div className="brand-subtitle">Shaker Automation Portal</div>
          </div>
        </Link>

        <nav className="top-nav" aria-label="Primary navigation">
          <NavItem to="/" label="Home" />
          <NavDropdown
            label="Asset Management"
            items={[
              { to: '/asset-requisition', label: 'Asset Requisition' },
              { to: '/requisition-tracking', label: 'Requisition Tracking' },
            ]}
          />
          <NavDropdown
            label="Finance & Payments"
            items={[
              { to: '/budget-upload', label: 'Upload Budget' },
              { to: '/bank-transfers', label: 'Bank Transfers' },
              { to: '/sadad-payments', label: 'SADAD Payments' },
            ]}
          />
          <NavItem to="/report-creation" label="Reports" />
        </nav>
      </div>
    </header>
  );
}


export function NavItem({ to, label }) {
  return (
    <NavLink
      to={to}
      className={({ isActive }) => (isActive ? 'nav-link active' : 'nav-link')}
    >
      {label}
    </NavLink>
  );
}


export function NavDropdown({ label, items }) {
  const location = useLocation();
  const [open, setOpen] = useState(false);
  const isChildActive = items.some((item) => location.pathname === item.to);

  return (
    <div
      className="nav-dropdown"
      onMouseEnter={() => setOpen(true)}
      onMouseLeave={() => setOpen(false)}
    >
      <button
        type="button"
        className={isChildActive ? 'nav-link nav-dropdown-trigger active' : 'nav-link nav-dropdown-trigger'}
        aria-haspopup="true"
        aria-expanded={open}
        onClick={() => setOpen((current) => !current)}
      >
        {label}
        <ChevronDown size={15} className={open ? 'nav-dropdown-caret open' : 'nav-dropdown-caret'} />
      </button>
      <div className={open ? 'nav-dropdown-menu open' : 'nav-dropdown-menu'} role="menu">
        {items.map((item) => (
          <NavLink
            key={item.to}
            to={item.to}
            role="menuitem"
            onClick={() => setOpen(false)}
            className={({ isActive }) => (isActive ? 'nav-dropdown-item active' : 'nav-dropdown-item')}
          >
            {item.label}
          </NavLink>
        ))}
      </div>
    </div>
  );
}


export function Footer() {
  return (
    <footer className="site-footer">
      <div className="footer-inner">
        <div>Built by QBS © {new Date().getFullYear()}</div>
        <div className="footer-note">An AI Initiative by Shaker ©</div>
      </div>
    </footer>
  );
}

