'use client';

import Link from 'next/link';
import { usePathname } from 'next/navigation';
import { ShieldCheck, UserPlus, ChevronDown, Presentation, Download, FileDown } from 'lucide-react';
import mudirajLogo from '@/Public/mudiraj_logo.png';

export default function Header() {
  const pathname = usePathname();

  return (
    <>
      <div className="top-banner">
        <div className="container top-banner-inner">
          <span>
            <strong>Andhra Pradesh Mudiraj Mahasabha</strong> — Unity, Empowerment &amp; Community Welfare
          </span>
          <div style={{ display: 'flex', alignItems: 'center', gap: '16px', flexWrap: 'wrap' }}>
            <a
              href="/downloads/Mudiraj_Website_User_Guide.pptx"
              download="Mudiraj_Website_User_Guide.pptx"
              style={{
                color: '#fef08a',
                background: 'rgba(253, 230, 138, 0.15)',
                border: '1px solid rgba(253, 230, 138, 0.35)',
                padding: '3px 8px',
                borderRadius: '4px',
                fontSize: '12px',
                fontWeight: 600,
                display: 'inline-flex',
                alignItems: 'center',
                gap: '5px',
                textDecoration: 'none',
                transition: 'all 0.2s',
              }}
              onMouseEnter={(e) => { e.currentTarget.style.background = 'rgba(253, 230, 138, 0.25)'; }}
              onMouseLeave={(e) => { e.currentTarget.style.background = 'rgba(253, 230, 138, 0.15)'; }}
              title="Download Full Website User Guide Presentation (.pptx)"
            >
              <FileDown size={13} /> Download Website Guide (PPT)
            </a>
            <span>State Jurisdiction: Andhra Pradesh</span>
            <Link
              href="/login"
              style={{
                color: '#fde68a',
                fontWeight: 600,
                display: 'inline-flex',
                alignItems: 'center',
                gap: '4px',
                padding: '4px 8px',
                borderRadius: '4px',
                border: '1px solid transparent',
                transition: 'all 0.2s'
              }}
              onMouseEnter={(e) => { e.currentTarget.style.border = '1px solid #fde68a'; e.currentTarget.style.background = 'rgba(253, 230, 138, 0.1)' }}
              onMouseLeave={(e) => { e.currentTarget.style.border = '1px solid transparent'; e.currentTarget.style.background = 'transparent' }}
            >
              <ShieldCheck size={14} /> Admin Login
            </Link>
          </div>
        </div>
      </div>

      <header className="site-header">
        <div className="container header-inner">
          <Link href="/" className="brand-logo">
            <img
              src={mudirajLogo.src || mudirajLogo}
              alt="Mudiraj Community Logo"
              className="brand-emblem"
              style={{ objectFit: 'cover', borderRadius: '50%', width: '48px', height: '48px' }}
            />
            <div>
              <div className="brand-title">Mudiraj Community</div>
              <div className="brand-subtitle">ANDHRA PRADESH STATE PORTAL</div>
            </div>
          </Link>

          <nav className="nav-links" aria-label="Main Navigation">
            <Link href="/" className={`nav-link ${pathname === '/' ? 'active' : ''}`}>
              Home
            </Link>
            
            <div className="dropdown">
              <div className={`nav-link ${pathname?.startsWith('/about') || pathname?.startsWith('/contact') ? 'active' : ''}`} style={{ display: 'flex', alignItems: 'center', gap: '4px', cursor: 'pointer' }}>
                About Us <ChevronDown size={14} />
              </div>
              <div className="dropdown-content">
                <Link href="/about">About Organization</Link>
                <Link href="/contact">Contact Us</Link>
              </div>
            </div>

            <div className="dropdown">
              <div className={`nav-link ${pathname?.startsWith('/organization') || pathname?.startsWith('/leadership') ? 'active' : ''}`} style={{ display: 'flex', alignItems: 'center', gap: '4px', cursor: 'pointer' }}>
                Structure <ChevronDown size={14} />
              </div>
              <div className="dropdown-content">
                <Link href="/organization">Organization Tree</Link>
                <Link href="/leadership">Leadership Directory</Link>
              </div>
            </div>

            <Link href="/membership" className={`nav-link ${pathname?.startsWith('/membership') ? 'active' : ''}`}>
              Membership Status
            </Link>

            {/* Download Website Guide PPT Dropdown */}
            <div className="dropdown">
              <div
                className="nav-link"
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: '6px',
                  cursor: 'pointer',
                  color: 'var(--primary)',
                  background: 'rgba(30, 58, 138, 0.06)',
                  border: '1px solid rgba(30, 58, 138, 0.18)',
                  borderRadius: '6px',
                  fontWeight: 600,
                }}
              >
                <Presentation size={15} color="var(--primary)" />
                <span>Website Guide (PPT)</span>
                <ChevronDown size={13} />
              </div>
              <div className="dropdown-content" style={{ minWidth: '320px', padding: '10px', right: 0 }}>
                <a
                  href="/downloads/Mudiraj_Website_User_Guide.pptx"
                  download="Mudiraj_Website_User_Guide.pptx"
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    gap: '10px',
                    fontWeight: 700,
                    color: '#1e3a8a',
                    background: '#eff6ff',
                    padding: '10px 12px',
                    borderRadius: '6px',
                    border: '1px solid #bfdbfe',
                    marginBottom: '8px',
                    textDecoration: 'none',
                  }}
                >
                  <Download size={18} color="#d97706" style={{ flexShrink: 0 }} />
                  <div>
                    <div style={{ fontSize: '13px', lineHeight: 1.2 }}>Download Website Guide (.pptx)</div>
                    <div style={{ fontSize: '11px', fontWeight: 500, color: '#64748b', marginTop: '2px' }}>
                      Visual User Manual &amp; Form Guide • 17 Slides (572 KB)
                    </div>
                  </div>
                </a>
                <div style={{ fontSize: '11px', fontWeight: 700, color: '#94a3b8', padding: '4px 6px', textTransform: 'uppercase', letterSpacing: '0.5px' }}>
                  Pages Covered in Presentation:
                </div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '2px' }}>
                  <Link href="/" style={{ fontSize: '12px', padding: '5px 8px', borderRadius: '4px' }}>• Home Page (State Central Portal)</Link>
                  <Link href="/about" style={{ fontSize: '12px', padding: '5px 8px', borderRadius: '4px' }}>• About Organization &amp; 30-Member Teams</Link>
                  <Link href="/organization" style={{ fontSize: '12px', padding: '5px 8px', borderRadius: '4px' }}>• 5-Tier Structure Explorer</Link>
                  <Link href="/leadership" style={{ fontSize: '12px', padding: '5px 8px', borderRadius: '4px' }}>• Leadership Directory (3 Wings)</Link>
                  <Link href="/membership" style={{ fontSize: '12px', padding: '5px 8px', borderRadius: '4px' }}>• Status Verification &amp; Lookup</Link>
                  <Link href="/apply" style={{ fontSize: '12px', padding: '5px 8px', borderRadius: '4px' }}>• Apply for Membership Form</Link>
                  <Link href="/id-card" style={{ fontSize: '12px', padding: '5px 8px', borderRadius: '4px' }}>• Digital Dual-Sided ID Card</Link>
                  <Link href="/contact" style={{ fontSize: '12px', padding: '5px 8px', borderRadius: '4px' }}>• Helpdesk &amp; District Centers</Link>
                </div>
              </div>
            </div>

            <Link href="/apply" className="btn btn-accent btn-sm" style={{ marginLeft: '4px' }}>
              <UserPlus size={15} /> Join Now
            </Link>
          </nav>
        </div>
      </header>
    </>
  );
}
