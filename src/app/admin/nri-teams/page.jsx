'use client';

import { useState, useEffect, useCallback } from 'react';
import Link from 'next/link';
import TeamCard from '@/components/TeamCard';
import {
  Globe,
  Plane,
  UsersRound,
  UserPlus,
  PlusCircle,
  Building,
  Award,
  Search,
  CheckCircle2,
  AlertCircle,
  ChevronRight,
  ShieldCheck,
  X,
} from 'lucide-react';
import { DEFAULT_TEAM_POSITIONS } from '@/lib/validation';

export default function AdminNriTeamsPage() {
  const [teams, setTeams] = useState([]);
  const [chapters, setChapters] = useState([]);
  const [loading, setLoading] = useState(true);
  const [selectedChapter, setSelectedChapter] = useState('');
  const [selectedTeamType, setSelectedTeamType] = useState('');
  const [search, setSearch] = useState('');

  // Notifications
  const [errorMsg, setErrorMsg] = useState('');
  const [successMsg, setSuccessMsg] = useState('');

  // Modals
  const [directModalOpen, setDirectModalOpen] = useState(false);
  const [chapterModalOpen, setChapterModalOpen] = useState(false);
  const [submitting, setSubmitting] = useState(false);

  // Direct Leader Form State
  const [directForm, setDirectForm] = useState({
    teamId: '',
    positionCode: 'PRESIDENT',
    fullName: '',
    country: 'United States of America',
    city: '',
    mobile: '',
    email: '',
  });

  // New Chapter Form State
  const [chapterForm, setChapterForm] = useState({
    name: '',
    countryCode: '',
    description: '',
  });

  const loadNriTeams = useCallback(async () => {
    setLoading(true);
    try {
      const params = new URLSearchParams();
      if (selectedChapter) params.set('chapterId', selectedChapter);
      if (selectedTeamType) params.set('teamType', selectedTeamType);
      if (search) params.set('search', search);

      const res = await fetch(`/api/admin/nri-teams?${params.toString()}`);
      const data = await res.json();
      if (res.ok) {
        setTeams(data.teams || []);
        setChapters(data.chapters || []);
        if (!directForm.teamId && data.teams && data.teams.length > 0) {
          setDirectForm((prev) => ({ ...prev, teamId: data.teams[0].id }));
        }
      } else {
        setErrorMsg(data.error || 'Failed to load NRI teams.');
      }
    } catch (err) {
      setErrorMsg(err.message || 'Unable to fetch NRI data.');
    } finally {
      setLoading(false);
    }
  }, [selectedChapter, selectedTeamType, search]);

  useEffect(() => {
    loadNriTeams();
  }, [loadNriTeams]);

  function handleTeamUpdated(updatedTeam) {
    setTeams((prev) =>
      prev.map((t) => (t.id === updatedTeam.id ? updatedTeam : t))
    );
    setSuccessMsg(`Team "${updatedTeam.orgName} - ${updatedTeam.teamType}" updated successfully.`);
    setTimeout(() => setSuccessMsg(''), 4000);
  }

  async function handleDirectLeaderSubmit(e) {
    e.preventDefault();
    setErrorMsg('');
    setSuccessMsg('');
    setSubmitting(true);

    try {
      const res = await fetch('/api/admin/nri-teams', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'ASSIGN_DIRECT_LEADER',
          ...directForm,
        }),
      });
      const data = await res.json();

      if (!res.ok) {
        setErrorMsg(data.error || 'Failed to assign NRI leader.');
        return;
      }

      setSuccessMsg(data.message || 'NRI leader assigned successfully.');
      setDirectModalOpen(false);
      setDirectForm({
        teamId: teams[0]?.id || '',
        positionCode: 'PRESIDENT',
        fullName: '',
        country: 'United States of America',
        city: '',
        mobile: '',
        email: '',
      });
      loadNriTeams();
      setTimeout(() => setSuccessMsg(''), 5000);
    } catch (err) {
      setErrorMsg(err.message || 'Error assigning NRI leader.');
    } finally {
      setSubmitting(false);
    }
  }

  async function handleChapterSubmit(e) {
    e.preventDefault();
    setErrorMsg('');
    setSuccessMsg('');
    setSubmitting(true);

    try {
      const res = await fetch('/api/admin/nri-teams', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'CREATE_CHAPTER',
          ...chapterForm,
        }),
      });
      const data = await res.json();

      if (!res.ok) {
        setErrorMsg(data.error || 'Failed to create chapter.');
        return;
      }

      setSuccessMsg(data.message || 'New NRI chapter created with Main, Youth, and Mahila teams.');
      setChapterModalOpen(false);
      setChapterForm({ name: '', countryCode: '', description: '' });
      loadNriTeams();
      setTimeout(() => setSuccessMsg(''), 5000);
    } catch (err) {
      setErrorMsg(err.message || 'Error creating chapter.');
    } finally {
      setSubmitting(false);
    }
  }

  // Calculate summary statistics
  const totalLeaders = teams.reduce((acc, t) => acc + (t.leaders?.length || 0), 0);
  const totalCapacity = teams.reduce((acc, t) => acc + (t.totalPositionsCapacity || 0), 0);

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: '22px' }}>
      {/* Top Banner & Header */}
      <div
        style={{
          background: 'linear-gradient(135deg, #0f172a 0%, #1e293b 100%)',
          color: '#ffffff',
          padding: '24px',
          borderRadius: '12px',
          boxShadow: '0 4px 16px rgba(0,0,0,0.1)',
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center',
          flexWrap: 'wrap',
          gap: '16px',
        }}
      >
        <div style={{ maxWidth: '780px' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '8px', flexWrap: 'wrap' }}>
            <span
              className="badge"
              style={{
                backgroundColor: '#0284c7',
                color: '#fff',
                fontSize: '11px',
                fontWeight: 700,
                letterSpacing: '0.5px',
                padding: '3px 9px',
              }}
            >
              <Plane size={13} style={{ marginRight: '4px', verticalAlign: 'middle' }} />
              SEPARATE WING
            </span>
            <span
              className="badge"
              style={{
                backgroundColor: 'rgba(255,255,255,0.15)',
                color: '#e2e8f0',
                fontSize: '11px',
                padding: '3px 9px',
              }}
            >
              International Diaspora Jurisdiction
            </span>
          </div>

          <h1 style={{ fontSize: '24px', fontWeight: 800, color: '#ffffff', margin: 0, lineHeight: 1.2 }}>
            NRI Community Teams (Non-Resident Indian)
          </h1>
          <p style={{ fontSize: '13px', color: '#cbd5e1', marginTop: '6px', lineHeight: 1.5 }}>
            Dedicated autonomous administration for global Mudiraj diaspora committees.
            Under this NRI wing, <strong>Main Team</strong>, <strong>Youth Team</strong>, and{' '}
            <strong>Mahila Team</strong> operate independently. Each team provides fixed positions for{' '}
            <em>President / Chairman, Vice President, General Secretary, Secretary, Treasurer</em>, and{' '}
            <em>Executive Members</em>.
          </p>
        </div>

        <div style={{ display: 'flex', gap: '12px', flexWrap: 'wrap' }}>
          <button
            type="button"
            className="btn"
            onClick={() => setDirectModalOpen(true)}
            style={{
              backgroundColor: '#d97706',
              color: '#ffffff',
              border: '1px solid #b45309',
              fontWeight: 700,
              fontSize: '13.5px',
              padding: '9px 18px',
              borderRadius: '6px',
              display: 'inline-flex',
              alignItems: 'center',
              gap: '8px',
              cursor: 'pointer',
              boxShadow: '0 2px 6px rgba(217, 119, 6, 0.4)',
              transition: 'all 0.2s',
            }}
            onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = '#b45309')}
            onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = '#d97706')}
          >
            <UserPlus size={17} /> Direct Assign NRI Leader
          </button>

          <button
            type="button"
            className="btn"
            onClick={() => setChapterModalOpen(true)}
            style={{
              backgroundColor: '#059669',
              color: '#ffffff',
              border: '1px solid #047857',
              fontWeight: 700,
              fontSize: '13.5px',
              padding: '9px 18px',
              borderRadius: '6px',
              display: 'inline-flex',
              alignItems: 'center',
              gap: '8px',
              cursor: 'pointer',
              boxShadow: '0 2px 6px rgba(5, 150, 105, 0.4)',
              transition: 'all 0.2s',
            }}
            onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = '#047857')}
            onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = '#059669')}
          >
            <PlusCircle size={17} /> Add Country Chapter
          </button>
        </div>
      </div>

      {/* Alerts */}
      {errorMsg && (
        <div className="alert alert-error" style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
          <AlertCircle size={18} />
          <span>{errorMsg}</span>
          <button
            type="button"
            onClick={() => setErrorMsg('')}
            style={{ marginLeft: 'auto', background: 'none', border: 'none', cursor: 'pointer', color: 'inherit' }}
          >
            <X size={16} />
          </button>
        </div>
      )}
      {successMsg && (
        <div className="alert alert-success" style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
          <CheckCircle2 size={18} />
          <span>{successMsg}</span>
          <button
            type="button"
            onClick={() => setSuccessMsg('')}
            style={{ marginLeft: 'auto', background: 'none', border: 'none', cursor: 'pointer', color: 'inherit' }}
          >
            <X size={16} />
          </button>
        </div>
      )}

      {/* KPI Stats Bar */}
      <div className="grid-4">
        <div className="stat-card">
          <div>
            <div className="stat-label">NRI Teams Active</div>
            <div className="stat-value">{teams.length}</div>
            <div className="stat-subtext">Main, Youth &amp; Mahila Wings</div>
          </div>
          <div className="stat-icon" style={{ backgroundColor: '#e0f2fe', color: '#0369a1' }}>
            <Plane size={22} />
          </div>
        </div>

        <div className="stat-card">
          <div>
            <div className="stat-label">Assigned NRI Leaders</div>
            <div className="stat-value">{totalLeaders}</div>
            <div className="stat-subtext">Across All Active NRI Teams</div>
          </div>
          <div className="stat-icon" style={{ backgroundColor: '#dcfce7', color: '#15803d' }}>
            <Award size={22} />
          </div>
        </div>

        <div className="stat-card">
          <div>
            <div className="stat-label">Total Roster Capacity</div>
            <div className="stat-value">{totalCapacity}</div>
            <div className="stat-subtext">Across All 6 Positions</div>
          </div>
          <div className="stat-icon" style={{ backgroundColor: '#fef3c7', color: '#b45309' }}>
            <UsersRound size={22} />
          </div>
        </div>

        <div className="stat-card">
          <div>
            <div className="stat-label">Registered Chapters</div>
            <div className="stat-value">{chapters.length || 1}</div>
            <div className="stat-subtext">Global &amp; Country Chapters</div>
          </div>
          <div className="stat-icon" style={{ backgroundColor: '#f3e8ff', color: '#6b21a8' }}>
            <Globe size={22} />
          </div>
        </div>
      </div>

      {/* Positions Breakdown Explainer Card */}
      <div className="card" style={{ background: '#f8fafc', border: '1px solid #e2e8f0', padding: '16px 20px' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '8px' }}>
          <ShieldCheck size={18} color="var(--primary)" />
          <h3 style={{ fontSize: '14px', fontWeight: 700, margin: 0, color: 'var(--primary-dark)' }}>
            Configured Positions in each NRI Team (Main, Youth, Mahila)
          </h3>
        </div>
        <div
          style={{
            display: 'grid',
            gridTemplateColumns: 'repeat(auto-fit, minmax(170px, 1fr))',
            gap: '8px',
            fontSize: '12px',
          }}
        >
          {DEFAULT_TEAM_POSITIONS.map((pos) => (
            <div
              key={pos.code}
              style={{
                background: '#ffffff',
                border: '1px solid #cbd5e1',
                padding: '8px 12px',
                borderRadius: '6px',
                display: 'flex',
                justifyContent: 'space-between',
                alignItems: 'center',
              }}
            >
              <span style={{ fontWeight: 600, color: '#334155' }}>{pos.title}</span>
              <span
                style={{
                  fontWeight: 800,
                  background: '#f1f5f9',
                  color: 'var(--primary)',
                  padding: '2px 6px',
                  borderRadius: '4px',
                }}
              >
                {pos.code === 'EXECUTIVE_MEMBER' ? '14 (Var)' : `${pos.maxCount} Slot${pos.maxCount > 1 ? 's' : ''}`}
              </span>
            </div>
          ))}
        </div>
      </div>

      {/* Filter and Tab Controls */}
      <div
        className="card"
        style={{
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center',
          flexWrap: 'wrap',
          gap: '14px',
          padding: '16px 20px',
        }}
      >
        {/* Team Type Tabs */}
        <div style={{ display: 'flex', gap: '6px', flexWrap: 'wrap' }}>
          {[
            { label: 'All NRI Wings', value: '' },
            { label: 'Main Team', value: 'Main Team' },
            { label: 'Youth Team', value: 'Youth Team' },
            { label: 'Mahila Team', value: 'Mahila Team' },
          ].map((tab) => {
            const isActive = selectedTeamType === tab.value;
            return (
              <button
                key={tab.label}
                type="button"
                onClick={() => setSelectedTeamType(tab.value)}
                className={`btn btn-sm ${isActive ? 'btn-primary' : 'btn-outline'}`}
                style={{
                  borderRadius: '20px',
                  padding: '6px 14px',
                  fontSize: '13px',
                  fontWeight: isActive ? 700 : 500,
                }}
              >
                {tab.label}
              </button>
            );
          })}
        </div>

        {/* Chapter Filter & Search */}
        <div style={{ display: 'flex', gap: '10px', alignItems: 'center', flexWrap: 'wrap' }}>
          {chapters.length > 1 && (
            <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
              <span style={{ fontSize: '12px', fontWeight: 600, color: 'var(--text-secondary)' }}>Chapter:</span>
              <select
                className="form-control"
                value={selectedChapter}
                onChange={(e) => setSelectedChapter(e.target.value)}
                style={{ padding: '6px 10px', fontSize: '13px', minWidth: '180px' }}
              >
                <option value="">All NRI Chapters</option>
                {chapters.map((ch) => (
                  <option key={ch.id} value={ch.id}>
                    {ch.name}
                  </option>
                ))}
              </select>
            </div>
          )}

          <div style={{ position: 'relative', width: '220px' }}>
            <Search
              size={15}
              style={{
                position: 'absolute',
                left: '10px',
                top: '50%',
                transform: 'translateY(-50%)',
                color: 'var(--text-muted)',
              }}
            />
            <input
              type="text"
              className="form-control"
              placeholder="Search NRI teams..."
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              style={{ paddingLeft: '32px', fontSize: '13px', padding: '6px 10px 6px 32px' }}
            />
          </div>
        </div>
      </div>

      {/* Teams Render Grid / List */}
      {loading ? (
        <div className="card" style={{ padding: '40px', textAlign: 'center', color: 'var(--text-muted)' }}>
          <Plane size={36} className="animate-spin" style={{ margin: '0 auto 12px', opacity: 0.6 }} />
          <div style={{ fontSize: '14px', fontWeight: 600 }}>Loading NRI teams and leadership rosters...</div>
        </div>
      ) : teams.length === 0 ? (
        <div className="card" style={{ padding: '48px', textAlign: 'center', color: 'var(--text-muted)' }}>
          <Plane size={44} style={{ margin: '0 auto 12px', opacity: 0.4 }} />
          <h3 style={{ fontSize: '16px', fontWeight: 700, color: 'var(--text-main)', marginBottom: '6px' }}>
            No NRI Teams Found
          </h3>
          <p style={{ fontSize: '13px', maxWidth: '440px', margin: '0 auto 18px' }}>
            There are currently no NRI teams matching the selected chapter or filters.
          </p>
          <button
            type="button"
            className="btn btn-primary"
            onClick={() => {
              setSelectedChapter('');
              setSelectedTeamType('');
              setSearch('');
            }}
          >
            Reset Filters
          </button>
        </div>
      ) : (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '22px' }}>
          {teams.map((team) => (
            <TeamCard
              key={team.id}
              team={team}
              isAdmin={true}
              onTeamUpdated={handleTeamUpdated}
            />
          ))}
        </div>
      )}

      {/* MODAL 1: Direct Register & Assign NRI Leader */}
      {directModalOpen && (
        <div className="modal-backdrop" onClick={() => setDirectModalOpen(false)}>
          <div
            className="modal-panel"
            onClick={(e) => e.stopPropagation()}
            style={{
              maxWidth: '600px',
              width: '94%',
              maxHeight: 'calc(100vh - 40px)',
              margin: 'auto',
              display: 'flex',
              flexDirection: 'column',
              backgroundColor: '#ffffff',
              color: '#0f172a',
              borderRadius: '12px',
              border: '1px solid #cbd5e1',
              boxShadow: '0 25px 50px -12px rgba(0, 0, 0, 0.45)',
              overflow: 'hidden',
            }}
          >
            <div
              className="modal-header"
              style={{
                backgroundColor: '#ffffff',
                borderBottom: '1px solid #e2e8f0',
                padding: '18px 24px',
                flexShrink: 0,
              }}
            >
              <div>
                <h3 style={{ fontSize: '18px', fontWeight: 800, margin: 0, color: '#0f172a' }}>
                  Assign NRI Community Leader
                </h3>
                <p style={{ fontSize: '13px', color: '#475569', marginTop: '3px', margin: 0 }}>
                  Directly appoint an overseas Mudiraj leader to one of the 6 official positions.
                </p>
              </div>
              <button
                type="button"
                className="btn-icon"
                onClick={() => setDirectModalOpen(false)}
                aria-label="Close"
                style={{
                  background: '#f1f5f9',
                  border: '1px solid #cbd5e1',
                  borderRadius: '6px',
                  padding: '6px',
                  cursor: 'pointer',
                  color: '#334155',
                }}
              >
                <X size={18} />
              </button>
            </div>

            <form
              onSubmit={handleDirectLeaderSubmit}
              style={{
                margin: 0,
                backgroundColor: '#ffffff',
                display: 'flex',
                flexDirection: 'column',
                flex: 1,
                minHeight: 0,
                overflow: 'hidden',
              }}
            >
              <div
                className="modal-body"
                style={{
                  padding: '20px 24px',
                  display: 'flex',
                  flexDirection: 'column',
                  gap: '14px',
                  backgroundColor: '#ffffff',
                  overflowY: 'auto',
                  maxHeight: 'calc(100vh - 180px)',
                  flex: 1,
                }}
              >
                <div className="form-grid" style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))', gap: '14px' }}>
                  <div className="form-group">
                    <label className="form-label" style={{ color: '#0f172a', fontWeight: 700, fontSize: '13px', marginBottom: '6px', display: 'block' }}>
                      Target NRI Team *
                    </label>
                    <select
                      className="form-control"
                      value={directForm.teamId}
                      onChange={(e) => setDirectForm({ ...directForm, teamId: e.target.value })}
                      required
                      style={{
                        backgroundColor: '#ffffff',
                        color: '#0f172a',
                        border: '1.5px solid #cbd5e1',
                        borderRadius: '6px',
                        padding: '9px 12px',
                        fontSize: '13.5px',
                        width: '100%',
                      }}
                    >
                      {teams.map((t) => (
                        <option key={t.id} value={t.id}>
                          {t.orgName} — {t.teamType}
                        </option>
                      ))}
                    </select>
                  </div>

                  <div className="form-group">
                    <label className="form-label" style={{ color: '#0f172a', fontWeight: 700, fontSize: '13px', marginBottom: '6px', display: 'block' }}>
                      Position / Designation *
                    </label>
                    <select
                      className="form-control"
                      value={directForm.positionCode}
                      onChange={(e) => setDirectForm({ ...directForm, positionCode: e.target.value })}
                      required
                      style={{
                        backgroundColor: '#ffffff',
                        color: '#0f172a',
                        border: '1.5px solid #cbd5e1',
                        borderRadius: '6px',
                        padding: '9px 12px',
                        fontSize: '13.5px',
                        width: '100%',
                      }}
                    >
                      <option value="PRESIDENT">President / Chairman (1 Slot)</option>
                      <option value="VICE_PRESIDENT">Vice President (6 Slots)</option>
                      <option value="GENERAL_SECRETARY">General Secretary (2 Slots)</option>
                      <option value="SECRETARY">Secretary (6 Slots)</option>
                      <option value="TREASURER">Treasurer (1 Slot)</option>
                      <option value="EXECUTIVE_MEMBER">Executive Member (Variable Slots)</option>
                    </select>
                  </div>
                </div>

                <div className="form-grid" style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))', gap: '14px' }}>
                  <div className="form-group">
                    <label className="form-label" style={{ color: '#0f172a', fontWeight: 700, fontSize: '13px', marginBottom: '6px', display: 'block' }}>
                      Full Name *
                    </label>
                    <input
                      type="text"
                      className="form-control"
                      placeholder="e.g. Ramesh Mudiraj"
                      value={directForm.fullName}
                      onChange={(e) => setDirectForm({ ...directForm, fullName: e.target.value })}
                      required
                      style={{
                        backgroundColor: '#ffffff',
                        color: '#0f172a',
                        border: '1.5px solid #cbd5e1',
                        borderRadius: '6px',
                        padding: '9px 12px',
                        fontSize: '13.5px',
                        width: '100%',
                      }}
                    />
                  </div>
                  <div className="form-group">
                    <label className="form-label" style={{ color: '#0f172a', fontWeight: 700, fontSize: '13px', marginBottom: '6px', display: 'block' }}>
                      Country of Residence *
                    </label>
                    <input
                      type="text"
                      className="form-control"
                      placeholder="e.g. United States / UAE / UK"
                      value={directForm.country}
                      onChange={(e) => setDirectForm({ ...directForm, country: e.target.value })}
                      required
                      style={{
                        backgroundColor: '#ffffff',
                        color: '#0f172a',
                        border: '1.5px solid #cbd5e1',
                        borderRadius: '6px',
                        padding: '9px 12px',
                        fontSize: '13.5px',
                        width: '100%',
                      }}
                    />
                  </div>
                </div>

                <div className="form-grid" style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))', gap: '14px' }}>
                  <div className="form-group">
                    <label className="form-label" style={{ color: '#0f172a', fontWeight: 700, fontSize: '13px', marginBottom: '6px', display: 'block' }}>
                      City / State Abroad
                    </label>
                    <input
                      type="text"
                      className="form-control"
                      placeholder="e.g. Dallas, Texas / Dubai"
                      value={directForm.city}
                      onChange={(e) => setDirectForm({ ...directForm, city: e.target.value })}
                      style={{
                        backgroundColor: '#ffffff',
                        color: '#0f172a',
                        border: '1.5px solid #cbd5e1',
                        borderRadius: '6px',
                        padding: '9px 12px',
                        fontSize: '13.5px',
                        width: '100%',
                      }}
                    />
                  </div>
                  <div className="form-group">
                    <label className="form-label" style={{ color: '#0f172a', fontWeight: 700, fontSize: '13px', marginBottom: '6px', display: 'block' }}>
                      Contact / WhatsApp Number
                    </label>
                    <input
                      type="text"
                      className="form-control"
                      placeholder="e.g. +1 469 555 0192"
                      value={directForm.mobile}
                      onChange={(e) => setDirectForm({ ...directForm, mobile: e.target.value })}
                      style={{
                        backgroundColor: '#ffffff',
                        color: '#0f172a',
                        border: '1.5px solid #cbd5e1',
                        borderRadius: '6px',
                        padding: '9px 12px',
                        fontSize: '13.5px',
                        width: '100%',
                      }}
                    />
                  </div>
                </div>

                <div className="form-group">
                  <label className="form-label" style={{ color: '#0f172a', fontWeight: 700, fontSize: '13px', marginBottom: '6px', display: 'block' }}>
                    Email Address
                  </label>
                  <input
                    type="email"
                    className="form-control"
                    placeholder="e.g. ramesh.mudiraj@gmail.com"
                    value={directForm.email}
                    onChange={(e) => setDirectForm({ ...directForm, email: e.target.value })}
                    style={{
                      backgroundColor: '#ffffff',
                      color: '#0f172a',
                      border: '1.5px solid #cbd5e1',
                      borderRadius: '6px',
                      padding: '9px 12px',
                      fontSize: '13.5px',
                      width: '100%',
                    }}
                  />
                </div>
              </div>

              <div
                className="modal-footer"
                style={{
                  backgroundColor: '#f8fafc',
                  borderTop: '1px solid #e2e8f0',
                  padding: '14px 24px',
                  display: 'flex',
                  justifyContent: 'flex-end',
                  gap: '12px',
                  flexShrink: 0,
                  position: 'sticky',
                  bottom: 0,
                  zIndex: 20,
                }}
              >
                <button
                  type="button"
                  onClick={() => setDirectModalOpen(false)}
                  disabled={submitting}
                  style={{
                    backgroundColor: '#ffffff',
                    color: '#334155',
                    border: '1.5px solid #cbd5e1',
                    borderRadius: '6px',
                    padding: '9px 18px',
                    fontWeight: 600,
                    fontSize: '13.5px',
                    cursor: 'pointer',
                  }}
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  disabled={submitting}
                  style={{
                    backgroundColor: '#d97706',
                    color: '#ffffff',
                    border: '1px solid #b45309',
                    borderRadius: '6px',
                    padding: '9px 22px',
                    fontWeight: 700,
                    fontSize: '13.5px',
                    cursor: 'pointer',
                    boxShadow: '0 2px 4px rgba(217, 119, 6, 0.3)',
                  }}
                >
                  {submitting ? 'Assigning...' : 'Appoint & Assign Leader'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* MODAL 2: Add Country Chapter */}
      {chapterModalOpen && (
        <div className="modal-backdrop" onClick={() => setChapterModalOpen(false)}>
          <div
            className="modal-panel"
            onClick={(e) => e.stopPropagation()}
            style={{
              maxWidth: '540px',
              width: '94%',
              maxHeight: 'calc(100vh - 40px)',
              margin: 'auto',
              display: 'flex',
              flexDirection: 'column',
              backgroundColor: '#ffffff',
              color: '#0f172a',
              borderRadius: '12px',
              border: '1px solid #cbd5e1',
              boxShadow: '0 25px 50px -12px rgba(0, 0, 0, 0.45)',
              overflow: 'hidden',
            }}
          >
            <div
              className="modal-header"
              style={{
                backgroundColor: '#ffffff',
                borderBottom: '1px solid #e2e8f0',
                padding: '18px 24px',
                flexShrink: 0,
              }}
            >
              <div>
                <h3 style={{ fontSize: '18px', fontWeight: 800, margin: 0, color: '#0f172a' }}>
                  Add Country / Regional NRI Chapter
                </h3>
                <p style={{ fontSize: '13px', color: '#475569', marginTop: '3px', margin: 0 }}>
                  Automatically provisions Main, Youth, and Mahila teams with full leadership rosters.
                </p>
              </div>
              <button
                type="button"
                className="btn-icon"
                onClick={() => setChapterModalOpen(false)}
                aria-label="Close"
                style={{
                  background: '#f1f5f9',
                  border: '1px solid #cbd5e1',
                  borderRadius: '6px',
                  padding: '6px',
                  cursor: 'pointer',
                  color: '#334155',
                }}
              >
                <X size={18} />
              </button>
            </div>

            <form
              onSubmit={handleChapterSubmit}
              style={{
                margin: 0,
                backgroundColor: '#ffffff',
                display: 'flex',
                flexDirection: 'column',
                flex: 1,
                minHeight: 0,
                overflow: 'hidden',
              }}
            >
              <div
                className="modal-body"
                style={{
                  padding: '20px 24px',
                  display: 'flex',
                  flexDirection: 'column',
                  gap: '14px',
                  backgroundColor: '#ffffff',
                  overflowY: 'auto',
                  maxHeight: 'calc(100vh - 180px)',
                  flex: 1,
                }}
              >
                <div className="form-group">
                  <label className="form-label" style={{ color: '#0f172a', fontWeight: 700, fontSize: '13px', marginBottom: '6px', display: 'block' }}>
                    Country / Region Name *
                  </label>
                  <input
                    type="text"
                    className="form-control"
                    placeholder="e.g. United Kingdom (UK) / UAE / Australia"
                    value={chapterForm.name}
                    onChange={(e) => setChapterForm({ ...chapterForm, name: e.target.value })}
                    required
                    style={{
                      backgroundColor: '#ffffff',
                      color: '#0f172a',
                      border: '1.5px solid #cbd5e1',
                      borderRadius: '6px',
                      padding: '9px 12px',
                      fontSize: '13.5px',
                      width: '100%',
                    }}
                  />
                </div>

                <div className="form-group">
                  <label className="form-label" style={{ color: '#0f172a', fontWeight: 700, fontSize: '13px', marginBottom: '6px', display: 'block' }}>
                    Short Code (Optional)
                  </label>
                  <input
                    type="text"
                    className="form-control"
                    placeholder="e.g. UK / UAE / AUS"
                    value={chapterForm.countryCode}
                    onChange={(e) => setChapterForm({ ...chapterForm, countryCode: e.target.value.toUpperCase() })}
                    maxLength={6}
                    style={{
                      backgroundColor: '#ffffff',
                      color: '#0f172a',
                      border: '1.5px solid #cbd5e1',
                      borderRadius: '6px',
                      padding: '9px 12px',
                      fontSize: '13.5px',
                      width: '100%',
                    }}
                  />
                </div>

                <div className="form-group">
                  <label className="form-label" style={{ color: '#0f172a', fontWeight: 700, fontSize: '13px', marginBottom: '6px', display: 'block' }}>
                    Description / Remarks (Optional)
                  </label>
                  <textarea
                    className="form-control"
                    rows={3}
                    placeholder="Overseas Mudiraj Community Chapter operations..."
                    value={chapterForm.description}
                    onChange={(e) => setChapterForm({ ...chapterForm, description: e.target.value })}
                    style={{
                      backgroundColor: '#ffffff',
                      color: '#0f172a',
                      border: '1.5px solid #cbd5e1',
                      borderRadius: '6px',
                      padding: '9px 12px',
                      fontSize: '13.5px',
                      width: '100%',
                      resize: 'vertical',
                    }}
                  />
                </div>
              </div>

              <div
                className="modal-footer"
                style={{
                  backgroundColor: '#f8fafc',
                  borderTop: '1px solid #e2e8f0',
                  padding: '14px 24px',
                  display: 'flex',
                  justifyContent: 'flex-end',
                  gap: '12px',
                  flexShrink: 0,
                  position: 'sticky',
                  bottom: 0,
                  zIndex: 20,
                }}
              >
                <button
                  type="button"
                  onClick={() => setChapterModalOpen(false)}
                  disabled={submitting}
                  style={{
                    backgroundColor: '#ffffff',
                    color: '#334155',
                    border: '1.5px solid #cbd5e1',
                    borderRadius: '6px',
                    padding: '9px 18px',
                    fontWeight: 600,
                    fontSize: '13.5px',
                    cursor: 'pointer',
                  }}
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  disabled={submitting}
                  style={{
                    backgroundColor: '#059669',
                    color: '#ffffff',
                    border: '1px solid #047857',
                    borderRadius: '6px',
                    padding: '9px 22px',
                    fontWeight: 700,
                    fontSize: '13.5px',
                    cursor: 'pointer',
                    boxShadow: '0 2px 4px rgba(5, 150, 105, 0.3)',
                  }}
                >
                  {submitting ? 'Creating Chapter...' : 'Create Chapter & Provision Teams'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}
