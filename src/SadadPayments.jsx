import React, { useState, useEffect, useRef } from 'react';
import { useNavigate, useParams } from 'react-router-dom';
import { Check, ChevronDown, Download, Search, Upload } from 'lucide-react';
import { IqamaExpenseScreen } from './IqamaExpenseScreen';
import { IqamaPaymentScreen } from './SadadPaymentScreen';
import { SADAD_GROUPS, SADAD_SCREENS, findScreen } from './sadadData';

export function SADADPaymentsPage({ setBanner }) {
  const { screenKey } = useParams();
  const navigate = useNavigate();
  const screen = screenKey ? findScreen(screenKey) : null;

  return (
    <div className="page-grid">
      <section className="content-card">
        <div className="sadad-toolbar">
          <div className="sadad-toolbar-title">
            <h2>SADAD Payments</h2>
          </div>
          <div className="sadad-toolbar-select">
            <SadadSelect
              selectedKey={screenKey || ''}
              onSelect={(key) => navigate(`/sadad-payments/${key}`)}
            />
          </div>
        </div>

        {screen ? <SadadWorkspace screen={screen} setBanner={setBanner} /> : <SadadEmptyState />}
      </section>
    </div>
  );
}


export function SadadSelect({ selectedKey, onSelect }) {
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState('');
  const ref = useRef(null);
  const selected = selectedKey ? findScreen(selectedKey) : null;

  useEffect(() => {
    function onDocClick(e) {
      if (ref.current && !ref.current.contains(e.target)) setOpen(false);
    }
    document.addEventListener('mousedown', onDocClick);
    return () => document.removeEventListener('mousedown', onDocClick);
  }, []);

  const q = query.trim().toLowerCase();
  const groups = SADAD_GROUPS.map((g) => ({
    group: g.group,
    screens: g.screens.filter(
      (s) =>
        !q ||
        s.nameAr.toLowerCase().includes(q) ||
        s.nameEn.toLowerCase().includes(q) ||
        (s.code && s.code.includes(q)) ||
        s.key.includes(q)
    ),
  })).filter((g) => g.screens.length > 0);

  return (
    <div className="sadad-select" ref={ref}>
      <div
        className={open ? 'sadad-field open' : 'sadad-field'}
        onClick={() => {
          if (!open) setOpen(true);
        }}
      >
        {open ? (
          <>
            <Search size={16} className="sadad-search-icon" />
            <input
              className="sadad-field-input"
              autoFocus
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Escape') {
                  setOpen(false);
                  setQuery('');
                }
              }}
              placeholder="Search by name or code…"
            />
            <span className="sadad-search-hint">{SADAD_SCREENS.length}</span>
          </>
        ) : selected ? (
          <span className="sadad-field-value">
            <img className="sadad-logo-thumb" src={selected.logo} alt="" />
            <span className="sadad-field-name">{selected.nameAr}</span>
            {selected.code ? <span className="sadad-code">{selected.code}</span> : null}
            <span className="sadad-field-en">· {selected.nameEn}</span>
          </span>
        ) : (
          <span className="sadad-placeholder">Select a payment screen…</span>
        )}
        <ChevronDown
          size={18}
          className="sadad-chevron"
          onClick={(e) => {
            e.stopPropagation();
            setOpen((o) => !o);
          }}
        />
      </div>

      {open ? (
        <div className="sadad-dropdown">
          <div className="sadad-dd-scroll">
            {groups.length === 0 ? (
              <div className="sadad-no-results">No screens match “{query}”.</div>
            ) : (
              groups.map((g) => (
                <div key={g.group}>
                  <div className="sadad-group-head">{g.group}</div>
                  {g.screens.map((s) => (
                    <button
                      type="button"
                      key={s.key}
                      className={s.key === selectedKey ? 'sadad-option selected' : 'sadad-option'}
                      onClick={() => {
                        onSelect(s.key);
                        setOpen(false);
                        setQuery('');
                      }}
                    >
                      <img className="sadad-logo-thumb" src={s.logo} alt="" />
                      <span className="sadad-opt-name">{s.nameAr}</span>
                      <span className="sadad-opt-en">· {s.nameEn}</span>
                      {s.code ? <span className="sadad-code sadad-opt-code">{s.code}</span> : null}
                      {s.key === selectedKey ? <Check size={16} className="sadad-opt-check" /> : null}
                    </button>
                  ))}
                </div>
              ))
            )}
          </div>
        </div>
      ) : null}
    </div>
  );
}


export function SadadEmptyState() {
  return (
    <div className="sadad-empty">
      <div className="sadad-empty-circle">
        <Search size={24} />
      </div>
      <h3>No screen selected yet</h3>
      <p>
        Select a payment screen from the dropdown above to load it here. You can search by name or
        code (e.g. “GOSI”, “060”, “اجير”).
      </p>
    </div>
  );
}


export function SadadWorkspace({ screen, setBanner }) {
  const [mode, setMode] = useState(screen.modes ? screen.modes[0] : null);

  useEffect(() => {
    setMode(screen.modes ? screen.modes[0] : null);
  }, [screen.key]);

  return (
    <div className="sadad-workspace">
      <div className="sadad-ws-head">
        <div className="sadad-ws-id">
          <img className="sadad-logo-lg" src={screen.logo} alt="" />
          <div className="sadad-ws-title">
            {screen.nameAr}
            <span className="sadad-ws-en">
              {' '}
              — {screen.nameEn}
              {screen.code ? ` · ${screen.code}` : ''}
            </span>
          </div>
        </div>
        {screen.modes ? (
          <div className="sadad-segmented">
            {screen.modes.map((m) => (
              <button
                type="button"
                key={m}
                className={m === mode ? 'sadad-seg active' : 'sadad-seg'}
                onClick={() => setMode(m)}
              >
                {m}
              </button>
            ))}
          </div>
        ) : null}
      </div>

      <div className="sadad-ws-body">
        {screen.key === 'iqama' && mode === 'Expense' ? (
          <IqamaExpenseScreen setBanner={setBanner} />
        ) : (screen.key === 'iqama' && mode === 'Payment') || screen.key === 'stc' ? (
          <IqamaPaymentScreen
            setBanner={setBanner}
            categoryKey={screen.key}
            categoryName={screen.nameEn}
          />
        ) : (
          <SadadGenericBody screen={screen} mode={mode} />
        )}
      </div>
    </div>
  );
}


export function SadadGenericBody({ screen, mode }) {
  return (
    <>
      <div className="sadad-actions">
        <button type="button" className="secondary-button">
          <Download size={18} /> Download Template
        </button>
        <button type="button" className="primary-button">
          <Upload size={18} /> Upload File
        </button>
        <button type="button" className="secondary-button">+ Add Single Entry</button>
      </div>

      <div className="sadad-stub">
        The <strong>{screen.nameEn}</strong> screen will be built here
        {mode ? ` (${mode})` : ''}. Its fields, template and table columns will be specific to this
        screen.
      </div>
    </>
  );
}

