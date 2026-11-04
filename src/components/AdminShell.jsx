'use client';

import { useState } from 'react';
import Sidebar from '@/components/Sidebar';
import { Menu, X, Download, FileDown, Presentation } from 'lucide-react';
import mudirajLogo from '@/Public/mudiraj_logo.png';

export default function AdminShell({ admin, children }) {
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);

  return (
    <div className="admin-shell">
      {/* Mobile Overlay */}
      {isMobileMenuOpen && (
        <div 
          className="mobile-overlay" 
          onClick={() => setIsMobileMenuOpen(false)}
        />
      )}
      
      <Sidebar admin={admin} isOpen={isMobileMenuOpen} onClose={() => setIsMobileMenuOpen(false)} />

      <div className="admin-main">
        <header className="admin-topbar">
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
            <button 
              className="mobile-menu-btn"
              onClick={() => setIsMobileMenuOpen(true)}
              aria-label="Open Menu"
            >
              <Menu size={20} />
            </button>
            <img
              src={mudirajLogo.src || mudirajLogo}
              alt="Mudiraj Logo"
              style={{
                width: '32px',
                height: '32px',
                borderRadius: '50%',
                objectFit: 'cover',
                border: '1.5px solid var(--accent)',
                flexShrink: 0,
              }}
            />
            <span style={{ fontWeight: 700, fontSize: '15px', color: 'var(--primary-dark)' }} className="topbar-title">
              Mudiraj Community Management System <span className="hide-on-mobile">— Andhra Pradesh Admin Console</span>
            </span>
          </div>

          <div style={{ display: 'flex', alignItems: 'center', gap: '12px', fontSize: '13px', flexWrap: 'wrap' }}>
            {/* Download Admin Features PPT Button in Navbar */}
            <a
              href="/downloads/Mudiraj_Admin_Features_Guide.pptx"
              download="Mudiraj_Admin_Features_Guide.pptx"
              className="btn btn-sm"
              style={{
                display: 'inline-flex',
                alignItems: 'center',
                gap: '6px',
                backgroundColor: '#ffffff',
                color: 'var(--primary)',
                border: '1.5px solid #bfdbfe',
                borderRadius: '6px',
                fontWeight: 700,
                fontSize: '12px',
                padding: '6px 12px',
                boxShadow: '0 1px 2px rgba(0,0,0,0.05)',
                textDecoration: 'none',
                transition: 'all 0.15s',
                whiteSpace: 'nowrap'
              }}
              onMouseEnter={(e) => {
                e.currentTarget.style.backgroundColor = '#eff6ff';
                e.currentTarget.style.borderColor = 'var(--primary)';
              }}
              onMouseLeave={(e) => {
                e.currentTarget.style.backgroundColor = '#ffffff';
                e.currentTarget.style.borderColor = '#bfdbfe';
              }}
              title="Download Admin Features Guide Presentation (.pptx covering all admin modules)"
            >
              <Download size={14} color="var(--accent)" />
              <span>Download Admin Guide (PPT)</span>
            </a>

            <span className="badge badge-active hide-on-mobile">State: Andhra Pradesh</span>
          </div>
        </header>

        <main className="admin-workspace">{children}</main>
      </div>
    </div>
  );
}
