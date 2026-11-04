'use client';

import Link from 'next/link';
import { usePathname, useRouter } from 'next/navigation';
import {
  LayoutDashboard,
  Users,
  FileCheck2,
  Network,
  UsersRound,
  Award,
  BarChart3,
  Settings,
  LogOut,
  Globe,
  Plane,
  FileDown,
} from 'lucide-react';
import mudirajLogo from '@/Public/mudiraj_logo.png';

export default function Sidebar({ admin, isOpen, onClose }) {
  const pathname = usePathname();
  const router = useRouter();

  const navItems = [
    { href: '/admin', label: 'Dashboard', icon: LayoutDashboard, exact: true },
    { href: '/admin/members', label: 'Members', icon: Users },
    { href: '/admin/applications', label: 'Membership Applications', icon: FileCheck2 },
    { href: '/admin/organizations', label: 'Masters', icon: Network },
    { href: '/admin/teams', label: 'Teams', icon: UsersRound },
    { href: '/admin/nri-teams', label: 'NRI Teams', icon: Plane },
    { href: '/admin/leaders', label: 'Leaders', icon: Award },
    { href: '/admin/reports', label: 'Reports', icon: BarChart3 },
    { href: '/admin/settings', label: 'Settings', icon: Settings },
  ];

  async function handleLogout() {
    if (!window.confirm('Are you sure you want to log out of the Admin Dashboard?')) {
      return;
    }
    await fetch('/api/auth/logout', { method: 'POST' });
    router.push('/login');
    router.refresh();
  }

  const handleLinkClick = () => {
    if (onClose) onClose();
  };

  return (
    <aside className={`admin-sidebar ${isOpen ? 'open' : ''}`}>
      <div className="admin-sidebar-header">
        <Link href="/admin" onClick={handleLinkClick} style={{ display: 'flex', alignItems: 'center', gap: '10px', textDecoration: 'none' }}>
          <img
            src={mudirajLogo.src || mudirajLogo}
            alt="Mudiraj Community Logo"
            className="brand-emblem"
            style={{
              width: '42px',
              height: '42px',
              borderRadius: '50%',
              objectFit: 'cover',
              border: '2px solid #fbbf24',
              backgroundColor: '#ffffff',
            }}
          />
          <div>
            <div style={{ fontWeight: 700, fontSize: '15px', color: '#ffffff' }}>
              Mudiraj Admin
            </div>
            <div style={{ fontSize: '11px', color: '#fbbf24', fontWeight: 600 }}>
              ANDHRA PRADESH STATE
            </div>
          </div>
        </Link>
      </div>

      <nav className="admin-nav" aria-label="Admin Navigation">
        {navItems.map((item) => {
          const Icon = item.icon;
          const isActive = item.exact
            ? pathname === item.href
            : pathname === item.href || pathname?.startsWith(`${item.href}/`);
          return (
            <Link
              key={item.href}
              href={item.href}
              onClick={handleLinkClick}
              className={`admin-nav-item ${isActive ? 'active' : ''}`}
            >
              <Icon size={18} />
              <span>{item.label}</span>
            </Link>
          );
        })}

        <div style={{ marginTop: 'auto', paddingTop: '16px', borderTop: '1px solid rgba(255,255,255,0.1)' }}>
          <a
            href="/downloads/Mudiraj_Admin_Features_Guide.pptx"
            download="Mudiraj_Admin_Features_Guide.pptx"
            className="admin-nav-item"
            style={{ color: '#fbbf24', textDecoration: 'none' }}
            title="Download Admin Features Guide Presentation (.pptx)"
          >
            <FileDown size={18} />
            <span>Admin Guide (PPT)</span>
          </a>
          <Link href="/" onClick={handleLinkClick} className="admin-nav-item">
            <Globe size={18} />
            <span>Public Website</span>
          </Link>
          <button
            type="button"
            onClick={handleLogout}
            className="admin-nav-item"
            style={{
              width: '100%',
              background: 'transparent',
              border: 'none',
              cursor: 'pointer',
              color: '#fca5a5',
              textAlign: 'left',
              fontFamily: 'inherit',
            }}
          >
            <LogOut size={18} />
            <span>Logout</span>
          </button>
        </div>
      </nav>
    </aside>
  );
}
