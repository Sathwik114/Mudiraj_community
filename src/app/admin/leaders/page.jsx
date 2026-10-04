'use client';

import { useState, useEffect, useCallback, useMemo } from 'react';
import {
  Search,
  UserPlus,
  Trash2,
  Award,
  CheckCircle2,
  AlertCircle,
  MapPin,
  Users,
  ShieldCheck,
  RefreshCw,
  Building2,
} from 'lucide-react';
import {
  VALID_ORG_LEVELS,
  VALID_TEAM_TYPES,
  DEFAULT_TEAM_POSITIONS,
} from '@/lib/validation';

export default function AdminLeadersPage() {
  const [leaders, setLeaders] = useState([]);
  const [loading, setLoading] = useState(true);
  const [feedback, setFeedback] = useState({ type: '', text: '' });

  // Organization Hierarchy
  const [hierarchy, setHierarchy] = useState({
    states: [],
    districts: [],
    constitutions: [],
    mandals: [],
    gramams: [],
  });
  const [hierarchyLoading, setHierarchyLoading] = useState(true);

  // Filters for Leaders Directory Table
  const [search, setSearch] = useState('');
  const [orgLevel, setOrgLevel] = useState('');
  const [teamType, setTeamType] = useState('');
  const [positionCode, setPositionCode] = useState('');
  const [filterDistrictId, setFilterDistrictId] = useState('');
  const [filterConstitutionId, setFilterConstitutionId] = useState('');
  const [filterMandalId, setFilterMandalId] = useState('');

  // Leader Assignment Panel State
  const [memberSearchQuery, setMemberSearchQuery] = useState('');
  const [searchingMember, setSearchingMember] = useState(false);
  const [searchedMembers, setSearchedMembers] = useState([]);
  const [selectedMember, setSelectedMember] = useState(null);

  const [assignLevel, setAssignLevel] = useState('District');
  const [assignDistrictId, setAssignDistrictId] = useState('');
  const [assignConstitutionId, setAssignConstitutionId] = useState('');
  const [assignMandalId, setAssignMandalId] = useState('');
  const [assignNriUnitId, setAssignNriUnitId] = useState('unit-nri');
  const [assignTeamType, setAssignTeamType] = useState('Main Team');
  const [assignPositionCode, setAssignPositionCode] = useState('PRESIDENT');

  // Real-time Target Team Info
  const [targetTeamInfo, setTargetTeamInfo] = useState(null);
  const [loadingTeamInfo, setLoadingTeamInfo] = useState(false);
  const [submittingAssignment, setSubmittingAssignment] = useState(false);

  // Photo Lightbox
  const [previewPhoto, setPreviewPhoto] = useState(null);

  // Load Organization Hierarchy
  useEffect(() => {
    setHierarchyLoading(true);
    fetch('/api/admin/organizations')
      .then((r) => r.json())
      .then((data) => {
        if (data.hierarchy) {
          setHierarchy(data.hierarchy);
          if (data.hierarchy.districts?.[0]) {
            setAssignDistrictId(data.hierarchy.districts[0].id);
          }
        }
      })
      .catch((err) => console.error('Failed to load hierarchy:', err))
      .finally(() => setHierarchyLoading(false));
  }, []);

  // Sorted District Lists
  const districtsList = useMemo(() => {
    return (hierarchy.districts || [])
      .slice()
      .sort((a, b) => a.name.localeCompare(b.name));
  }, [hierarchy.districts]);

  // Filtered Constituencies for Assignment
  const assignConstitutionsList = useMemo(() => {
    if (!assignDistrictId) return [];
    return (hierarchy.constitutions || [])
      .filter((c) => c.parentId === assignDistrictId || c.districtId === assignDistrictId)
      .slice()
      .sort((a, b) => a.name.localeCompare(b.name));
  }, [hierarchy.constitutions, assignDistrictId]);

  // Filtered Mandals for Assignment
  const assignMandalsList = useMemo(() => {
    if (!assignConstitutionId) return [];
    return (hierarchy.mandals || [])
      .filter((m) => m.parentId === assignConstitutionId || m.constitutionId === assignConstitutionId)
      .slice()
      .sort((a, b) => a.name.localeCompare(b.name));
  }, [hierarchy.mandals, assignConstitutionId]);

  // Filtered Constituencies for Directory Table Filters
  const filterConstitutionsList = useMemo(() => {
    if (!filterDistrictId) return [];
    return (hierarchy.constitutions || [])
      .filter((c) => c.parentId === filterDistrictId || c.districtId === filterDistrictId)
      .slice()
      .sort((a, b) => a.name.localeCompare(b.name));
  }, [hierarchy.constitutions, filterDistrictId]);

  // Filtered Mandals for Directory Table Filters
  const filterMandalsList = useMemo(() => {
    if (!filterConstitutionId) return [];
    return (hierarchy.mandals || [])
      .filter((m) => m.parentId === filterConstitutionId || m.constitutionId === filterConstitutionId)
      .slice()
      .sort((a, b) => a.name.localeCompare(b.name));
  }, [hierarchy.mandals, filterConstitutionId]);

  // Resolved Target Unit for Assignment
  const resolvedTargetUnit = useMemo(() => {
    if (assignLevel === 'State') {
      return (
        hierarchy.states?.[0] || {
          id: 'unit-ap',
          name: 'Andhra Pradesh',
          locationLabel: 'Andhra Pradesh',
        }
      );
    }
    if (assignLevel === 'District') {
      return districtsList.find((d) => d.id === assignDistrictId) || null;
    }
    if (assignLevel === 'Constitution') {
      return assignConstitutionsList.find((c) => c.id === assignConstitutionId) || null;
    }
    if (assignLevel === 'Mandal') {
      return assignMandalsList.find((m) => m.id === assignMandalId) || null;
    }
    if (assignLevel === 'NRI') {
      return {
        id: assignNriUnitId || 'unit-nri',
        name: 'NRI Global Central Wing',
        locationLabel: 'International / Diaspora Wing',
      };
    }
    return null;
  }, [
    assignLevel,
    assignDistrictId,
    assignConstitutionId,
    assignMandalId,
    assignNriUnitId,
    hierarchy.states,
    districtsList,
    assignConstitutionsList,
    assignMandalsList,
  ]);

  // Computed team ID
  const computedTeamId = useMemo(() => {
    if (!resolvedTargetUnit) return '';
    const typeSlug = assignTeamType.toLowerCase().replace(/\s+/g, '-');
    if (assignLevel === 'NRI') {
      return `team-nri-${typeSlug}`;
    }
    return `team-${resolvedTargetUnit.id}-${typeSlug}`;
  }, [resolvedTargetUnit, assignTeamType, assignLevel]);

  // Fetch Live Target Team Details & Capacity
  useEffect(() => {
    if (!resolvedTargetUnit?.id) {
      setTargetTeamInfo(null);
      return;
    }

    let active = true;
    setLoadingTeamInfo(true);
    fetch(
      `/api/admin/teams?orgId=${encodeURIComponent(resolvedTargetUnit.id)}&exact=true&teamType=${encodeURIComponent(
        assignTeamType
      )}&includeNri=true`
    )
      .then((r) => r.json())
      .then((data) => {
        if (!active) return;
        const matchingTeam =
          (data.teams || []).find((t) => t.teamType === assignTeamType) ||
          data.teams?.[0] ||
          null;
        setTargetTeamInfo(matchingTeam);
      })
      .catch((err) => {
        console.error('Error fetching live team details:', err);
      })
      .finally(() => {
        if (active) setLoadingTeamInfo(false);
      });

    return () => {
      active = false;
    };
  }, [resolvedTargetUnit?.id, assignTeamType]);

  // Load Leaders Directory
  const loadLeaders = useCallback(async () => {
    setLoading(true);
    try {
      const params = new URLSearchParams();
      if (search) params.set('search', search);
      if (orgLevel) params.set('orgLevel', orgLevel);
      if (teamType) params.set('teamType', teamType);
      if (positionCode) params.set('positionCode', positionCode);
      if (filterDistrictId) params.set('districtId', filterDistrictId);
      if (filterConstitutionId) params.set('constitutionId', filterConstitutionId);
      if (filterMandalId) params.set('mandalId', filterMandalId);

      const res = await fetch(`/api/admin/leaders?${params.toString()}`);
      const data = await res.json();
      if (res.ok) {
        setLeaders(data.leaders || []);
      }
    } catch (err) {
      console.error('Failed to load leaders:', err);
    } finally {
      setLoading(false);
    }
  }, [
    search,
    orgLevel,
    teamType,
    positionCode,
    filterDistrictId,
    filterConstitutionId,
    filterMandalId,
  ]);

  useEffect(() => {
    loadLeaders();
  }, [loadLeaders]);

  // Step 1: Search Member
  async function handleSearchEligibleMembers(e) {
    e.preventDefault();
    if (!memberSearchQuery.trim()) return;
    setFeedback({ type: '', text: '' });
    setSearchingMember(true);
    try {
      const res = await fetch(
        `/api/admin/members?search=${encodeURIComponent(memberSearchQuery.trim())}&limit=12`
      );
      const data = await res.json();
      if (res.ok) {
        const eligible = (data.items || []).filter(
          (m) => m.membershipId && ['Active', 'Approved'].includes(m.status)
        );
        setSearchedMembers(eligible);
        if (eligible[0]) {
          handleSelectMember(eligible[0]);
        } else {
          setSelectedMember(null);
        }
      }
    } finally {
      setSearchingMember(false);
    }
  }

  // Handle Member Selection & Auto-fill Location
  function handleSelectMember(m) {
    setSelectedMember(m);
    if (!m) return;

    // Smart pre-fill: Check location path or district name
    if (m.locationPath) {
      const parts = m.locationPath.split('/').filter(Boolean);
      const matchedDist = districtsList.find((d) => parts.includes(d.id));
      if (matchedDist) {
        setAssignDistrictId(matchedDist.id);
        const consts = (hierarchy.constitutions || []).filter(
          (c) => c.parentId === matchedDist.id || c.districtId === matchedDist.id
        );
        const matchedConst = consts.find((c) => parts.includes(c.id));
        if (matchedConst) {
          setAssignConstitutionId(matchedConst.id);
          const mnds = (hierarchy.mandals || []).filter(
            (man) => man.parentId === matchedConst.id || man.constitutionId === matchedConst.id
          );
          const matchedMnd = mnds.find((man) => parts.includes(man.id));
          if (matchedMnd) {
            setAssignMandalId(matchedMnd.id);
          }
        }
      }
    } else if (m.districtName) {
      const matchedDist = districtsList.find(
        (d) => d.name.toLowerCase() === m.districtName.toLowerCase()
      );
      if (matchedDist) {
        setAssignDistrictId(matchedDist.id);
      }
    }
  }

  // Level Change Handler
  function handleLevelChange(lvl) {
    setAssignLevel(lvl);
    if (lvl === 'District') {
      if (!assignDistrictId && districtsList[0]) {
        setAssignDistrictId(districtsList[0].id);
      }
    } else if (lvl === 'Constitution') {
      if (!assignDistrictId && districtsList[0]) {
        setAssignDistrictId(districtsList[0].id);
      }
    } else if (lvl === 'Mandal') {
      if (!assignDistrictId && districtsList[0]) {
        setAssignDistrictId(districtsList[0].id);
      }
    }
  }

  // Check if position is currently available or occupied
  const currentOccupant = useMemo(() => {
    if (!targetTeamInfo?.leaders) return null;
    return targetTeamInfo.leaders.find(
      (l) => l.positionCode === assignPositionCode
    );
  }, [targetTeamInfo, assignPositionCode]);

  // Submit Leader Assignment
  async function handleAssignSubmit(e) {
    e.preventDefault();
    setFeedback({ type: '', text: '' });

    if (!selectedMember) {
      setFeedback({
        type: 'error',
        text: 'Please search and select a registered member in Step 1.',
      });
      return;
    }

    if (assignLevel === 'District' && !assignDistrictId) {
      setFeedback({ type: 'error', text: 'Please select a District from the dropdown.' });
      return;
    }

    if (assignLevel === 'Constitution') {
      if (!assignDistrictId) {
        setFeedback({ type: 'error', text: 'Please select a District first.' });
        return;
      }
      if (!assignConstitutionId) {
        setFeedback({
          type: 'error',
          text: 'Please select an Assembly Constituency from the dropdown.',
        });
        return;
      }
    }

    if (assignLevel === 'Mandal') {
      if (!assignDistrictId) {
        setFeedback({ type: 'error', text: 'Please select a District first.' });
        return;
      }
      if (!assignConstitutionId) {
        setFeedback({ type: 'error', text: 'Please select an Assembly Constituency first.' });
        return;
      }
      if (!assignMandalId) {
        setFeedback({ type: 'error', text: 'Please select a Mandal from the dropdown.' });
        return;
      }
    }

    if (!resolvedTargetUnit) {
      setFeedback({
        type: 'error',
        text: 'Please select all required location levels before assigning.',
      });
      return;
    }

    setSubmittingAssignment(true);
    try {
      const payload = {
        teamId: computedTeamId,
        orgUnitId: resolvedTargetUnit.id,
        teamType: assignTeamType,
        memberId: selectedMember.id,
        positionCode: assignPositionCode,
      };

      const res = await fetch('/api/admin/leaders', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      const data = await res.json();

      if (!res.ok) {
        setFeedback({ type: 'error', text: data.error || 'Failed to assign leader.' });
        return;
      }

      const posTitle =
        DEFAULT_TEAM_POSITIONS.find((p) => p.code === assignPositionCode)?.title ||
        assignPositionCode;
      setFeedback({
        type: 'success',
        text: `Successfully assigned ${selectedMember.fullName} (${selectedMember.membershipId}) as "${posTitle}" in ${resolvedTargetUnit.name} — ${assignTeamType}!`,
      });

      // Refresh target team info and directory list
      loadLeaders();
      if (resolvedTargetUnit?.id) {
        fetch(
          `/api/admin/teams?orgId=${encodeURIComponent(
            resolvedTargetUnit.id
          )}&exact=true&teamType=${encodeURIComponent(assignTeamType)}&includeNri=true`
        )
          .then((r) => r.json())
          .then((d) => {
            const matchingTeam =
              (d.teams || []).find((t) => t.teamType === assignTeamType) ||
              d.teams?.[0] ||
              null;
            setTargetTeamInfo(matchingTeam);
          });
      }
    } finally {
      setSubmittingAssignment(false);
    }
  }

  // Change Position
  async function handleChangeLeaderPosition(teamMemberId, newPositionCode) {
    setFeedback({ type: '', text: '' });
    const res = await fetch('/api/admin/leaders', {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ teamMemberId, newPositionCode }),
    });
    const data = await res.json();
    if (!res.ok) {
      setFeedback({ type: 'error', text: data.error || 'Failed to change position.' });
      return;
    }
    setFeedback({ type: 'success', text: data.message });
    loadLeaders();
  }

  // Remove Leader
  async function handleRemoveLeader(leader) {
    if (
      !window.confirm(
        `Remove ${leader.memberName} (${leader.membershipId}) from "${leader.positionTitle}" in ${leader.orgName} (${leader.teamType})?`
      )
    ) {
      return;
    }
    const res = await fetch(
      `/api/admin/leaders?teamMemberId=${encodeURIComponent(leader.id)}`,
      { method: 'DELETE' }
    );
    const data = await res.json();
    if (res.ok) {
      setFeedback({ type: 'success', text: data.message });
      loadLeaders();
    } else {
      setFeedback({ type: 'error', text: data.error || 'Failed to remove leader.' });
    }
  }

  return (
    <div>
      {/* Page Header */}
      <div style={{ marginBottom: '20px' }}>
        <h1 style={{ fontSize: '24px', fontWeight: 800, color: 'var(--primary-dark)' }}>
          Leader Assignment &amp; Office Bearers Directory
        </h1>
        <p style={{ fontSize: '13px', color: 'var(--text-secondary)' }}>
          Assign registered community members to leadership positions across Andhra Pradesh State,
          Districts (26), Constituencies (175), Mandals (677), and the NRI Wing.
        </p>
      </div>

      {/* Feedback Banner */}
      {feedback.text && (
        <div
          className={`alert ${feedback.type === 'error' ? 'alert-error' : 'alert-success'}`}
          style={{ marginBottom: '20px' }}
        >
          {feedback.type === 'error' ? (
            <AlertCircle size={18} style={{ flexShrink: 0 }} />
          ) : (
            <CheckCircle2 size={18} style={{ flexShrink: 0 }} />
          )}
          <span style={{ fontWeight: 600 }}>{feedback.text}</span>
        </div>
      )}

      {/* Leader Assignment Card */}
      <div
        className="card"
        style={{
          marginBottom: '28px',
          borderTop: '4px solid var(--primary)',
          boxShadow: '0 4px 6px -1px rgba(0, 0, 0, 0.1)',
        }}
      >
        <div className="card-header" style={{ paddingBottom: '12px' }}>
          <div>
            <h2 className="card-title" style={{ fontSize: '18px', display: 'flex', alignItems: 'center', gap: '8px' }}>
              <UserPlus size={20} style={{ color: 'var(--primary)' }} />
              Assign Registered Member to Leadership Position
            </h2>
            <p style={{ fontSize: '12px', color: 'var(--text-secondary)', marginTop: '2px' }}>
              Select a member, choose the organizational level (State, District, Constituency, Mandal, NRI), select the target area from cascading dropdowns, and assign the office bearer role.
            </p>
          </div>
        </div>

        <div className="grid-2" style={{ alignItems: 'start', gap: '24px' }}>
          {/* STEP 1: Search and Select Member */}
          <div
            style={{
              background: '#f8fafc',
              padding: '16px',
              borderRadius: '8px',
              border: '1px solid #e2e8f0',
            }}
          >
            <div style={{ fontWeight: 800, fontSize: '14px', color: 'var(--primary-dark)', marginBottom: '8px' }}>
              Step 1: Search &amp; Select Member
            </div>
            <p style={{ fontSize: '12px', color: 'var(--text-muted)', marginBottom: '12px' }}>
              Enter Membership ID (e.g. MUD-00000001), full name, or mobile number:
            </p>

            <form onSubmit={handleSearchEligibleMembers} style={{ marginBottom: '14px' }}>
              <div style={{ display: 'flex', gap: '8px' }}>
                <input
                  type="text"
                  className="form-control"
                  placeholder="Enter ID, Name, or Mobile..."
                  value={memberSearchQuery}
                  onChange={(e) => setMemberSearchQuery(e.target.value)}
                />
                <button
                  type="submit"
                  className="btn btn-primary"
                  disabled={searchingMember}
                  style={{ display: 'inline-flex', alignItems: 'center', gap: '6px' }}
                >
                  <Search size={15} /> {searchingMember ? 'Searching...' : 'Search'}
                </button>
              </div>
            </form>

            {/* Search Results List */}
            {searchedMembers.length > 0 && (
              <div
                style={{
                  border: '1px solid var(--border-color)',
                  borderRadius: '8px',
                  maxHeight: '230px',
                  overflowY: 'auto',
                  background: '#ffffff',
                  marginBottom: '14px',
                }}
              >
                {searchedMembers.map((m) => {
                  const isSel = selectedMember?.id === m.id;
                  return (
                    <div
                      key={m.id}
                      onClick={() => handleSelectMember(m)}
                      style={{
                        padding: '10px 12px',
                        borderBottom: '1px solid var(--border-color)',
                        cursor: 'pointer',
                        background: isSel ? 'var(--primary-light)' : '#ffffff',
                        display: 'flex',
                        justifyContent: 'space-between',
                        alignItems: 'center',
                        transition: 'background 0.15s ease',
                      }}
                    >
                      <div>
                        <div style={{ fontWeight: 700, fontSize: '13px', color: isSel ? 'var(--primary-dark)' : 'inherit' }}>
                          {m.fullName}
                        </div>
                        <div style={{ fontSize: '11px', color: 'var(--text-secondary)' }}>
                          📱 {m.mobile} {m.districtName ? `• 📍 ${m.districtName}` : ''}
                        </div>
                      </div>
                      <code
                        style={{
                          fontWeight: 700,
                          color: 'var(--primary)',
                          fontSize: '11px',
                          background: '#e0f2fe',
                          padding: '3px 7px',
                          borderRadius: '4px',
                        }}
                      >
                        {m.membershipId}
                      </code>
                    </div>
                  );
                })}
              </div>
            )}

            {/* Currently Selected Member Card */}
            <div
              style={{
                padding: '12px 14px',
                background: selectedMember ? '#eff6ff' : '#ffffff',
                border: `1px solid ${selectedMember ? '#bfdbfe' : '#e2e8f0'}`,
                borderRadius: '8px',
              }}
            >
              <div style={{ fontSize: '11px', textTransform: 'uppercase', fontWeight: 800, color: 'var(--text-muted)' }}>
                Target Member for Assignment
              </div>
              {selectedMember ? (
                <div style={{ marginTop: '6px' }}>
                  <div style={{ fontWeight: 800, fontSize: '15px', color: 'var(--primary-dark)' }}>
                    {selectedMember.fullName}
                  </div>
                  <div style={{ fontSize: '12px', color: 'var(--text-secondary)', marginTop: '2px' }}>
                    <strong>ID:</strong> <code style={{ color: 'var(--primary)', fontWeight: 700 }}>{selectedMember.membershipId}</code> • <strong>Mobile:</strong> {selectedMember.mobile}
                  </div>
                  {selectedMember.locationPath && (
                    <div style={{ fontSize: '11px', color: 'var(--text-muted)', marginTop: '4px' }}>
                      📍 {selectedMember.districtName || 'Registered Member'}
                    </div>
                  )}
                </div>
              ) : (
                <div style={{ fontSize: '13px', color: 'var(--text-muted)', marginTop: '4px', fontStyle: 'italic' }}>
                  No member selected. Search above and click a member to select.
                </div>
              )}
            </div>
          </div>

          {/* STEP 2: Cascading Dropdowns & Leadership Position */}
          <form
            onSubmit={handleAssignSubmit}
            style={{ display: 'flex', flexDirection: 'column', gap: '14px' }}
          >
            <div style={{ fontWeight: 800, fontSize: '14px', color: 'var(--primary-dark)' }}>
              Step 2: Select Hierarchy Level, Location &amp; Position
            </div>

            {/* Organizational Level & Team Wing */}
            <div className="form-grid-2">
              <div className="form-group">
                <label className="form-label" style={{ fontWeight: 700 }}>
                  1. Organizational Level <span style={{ color: 'var(--danger)' }}>*</span>
                </label>
                <select
                  className="form-control"
                  value={assignLevel}
                  onChange={(e) => handleLevelChange(e.target.value)}
                  style={{ fontWeight: 600 }}
                >
                  <option value="District">District Level (26 AP Districts)</option>
                  <option value="Constitution">Constituency Level (175 ACs)</option>
                  <option value="Mandal">Mandal Level (677 Mandals)</option>
                  <option value="State">State Level (Central Committee)</option>
                  <option value="NRI">NRI Wing (International Diaspora)</option>
                </select>
              </div>

              <div className="form-group">
                <label className="form-label" style={{ fontWeight: 700 }}>
                  2. Team Wing / Category <span style={{ color: 'var(--danger)' }}>*</span>
                </label>
                <select
                  className="form-control"
                  value={assignTeamType}
                  onChange={(e) => setAssignTeamType(e.target.value)}
                  style={{ fontWeight: 600 }}
                >
                  <option value="Main Team">Main Team (Core Leadership)</option>
                  <option value="Youth Team">Youth Team (Yuva Wing)</option>
                  <option value="Mahila Team">Mahila Team (Women's Wing)</option>
                </select>
              </div>
            </div>

            {/* CASCADING GEOGRAPHIC DROPDOWNS */}

            {/* Level: State */}
            {assignLevel === 'State' && (
              <div
                style={{
                  background: '#f0fdf4',
                  border: '1px solid #bbf7d0',
                  borderRadius: '8px',
                  padding: '12px 16px',
                }}
              >
                <div style={{ fontSize: '11px', textTransform: 'uppercase', fontWeight: 700, color: '#15803d' }}>
                  State Level Committee
                </div>
                <div style={{ fontWeight: 800, fontSize: '15px', color: '#166534', marginTop: '2px' }}>
                  🏛️ Andhra Pradesh State Central Committee
                </div>
                <div style={{ fontSize: '12px', color: '#15803d', marginTop: '2px' }}>
                  Apex State Body overseeing all 26 districts, 175 assembly constituencies, and 677 mandals.
                </div>
              </div>
            )}

            {/* Level: District (District Dropdown) */}
            {assignLevel === 'District' && (
              <div className="form-group">
                <label className="form-label" style={{ fontWeight: 700 }}>
                  Select District <span style={{ color: 'var(--danger)' }}>*</span>
                </label>
                <select
                  className="form-control"
                  value={assignDistrictId}
                  onChange={(e) => setAssignDistrictId(e.target.value)}
                  required
                >
                  <option value="">-- Select One of 26 Andhra Pradesh Districts --</option>
                  {districtsList.map((d) => (
                    <option key={d.id} value={d.id}>
                      {d.name} {d.code ? `[${d.code}]` : ''}
                    </option>
                  ))}
                </select>
              </div>
            )}

            {/* Level: Constitution (District -> Constituency) */}
            {assignLevel === 'Constitution' && (
              <div className="form-grid-2">
                <div className="form-group">
                  <label className="form-label" style={{ fontWeight: 700 }}>
                    Select District <span style={{ color: 'var(--danger)' }}>*</span>
                  </label>
                  <select
                    className="form-control"
                    value={assignDistrictId}
                    onChange={(e) => {
                      setAssignDistrictId(e.target.value);
                      setAssignConstitutionId('');
                    }}
                    required
                  >
                    <option value="">-- Step 1: Select District --</option>
                    {districtsList.map((d) => (
                      <option key={d.id} value={d.id}>
                        {d.name} {d.code ? `[${d.code}]` : ''}
                      </option>
                    ))}
                  </select>
                </div>

                <div className="form-group">
                  <label className="form-label" style={{ fontWeight: 700 }}>
                    Select Assembly Constituency <span style={{ color: 'var(--danger)' }}>*</span>
                  </label>
                  <select
                    className="form-control"
                    value={assignConstitutionId}
                    onChange={(e) => setAssignConstitutionId(e.target.value)}
                    disabled={!assignDistrictId}
                    required
                  >
                    <option value="">
                      {assignDistrictId
                        ? `-- Step 2: Select Constituency (${assignConstitutionsList.length} in this District) --`
                        : '-- First Select District on Left --'}
                    </option>
                    {assignConstitutionsList.map((c) => (
                      <option key={c.id} value={c.id}>
                        {c.name} {c.code ? `[${c.code}]` : ''}
                      </option>
                    ))}
                  </select>
                </div>
              </div>
            )}

            {/* Level: Mandal (District -> Constituency -> Mandal) */}
            {assignLevel === 'Mandal' && (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                <div className="form-grid-2">
                  <div className="form-group">
                    <label className="form-label" style={{ fontWeight: 700 }}>
                      1. Select District <span style={{ color: 'var(--danger)' }}>*</span>
                    </label>
                    <select
                      className="form-control"
                      value={assignDistrictId}
                      onChange={(e) => {
                        setAssignDistrictId(e.target.value);
                        setAssignConstitutionId('');
                        setAssignMandalId('');
                      }}
                      required
                    >
                      <option value="">-- Step 1: Select District --</option>
                      {districtsList.map((d) => (
                        <option key={d.id} value={d.id}>
                          {d.name} {d.code ? `[${d.code}]` : ''}
                        </option>
                      ))}
                    </select>
                  </div>

                  <div className="form-group">
                    <label className="form-label" style={{ fontWeight: 700 }}>
                      2. Select Constituency <span style={{ color: 'var(--danger)' }}>*</span>
                    </label>
                    <select
                      className="form-control"
                      value={assignConstitutionId}
                      onChange={(e) => {
                        setAssignConstitutionId(e.target.value);
                        setAssignMandalId('');
                      }}
                      disabled={!assignDistrictId}
                      required
                    >
                      <option value="">
                        {assignDistrictId
                          ? `-- Step 2: Select Constituency (${assignConstitutionsList.length}) --`
                          : '-- First Select District --'}
                      </option>
                      {assignConstitutionsList.map((c) => (
                        <option key={c.id} value={c.id}>
                          {c.name} {c.code ? `[${c.code}]` : ''}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>

                <div className="form-group">
                  <label className="form-label" style={{ fontWeight: 700 }}>
                    3. Select Mandal <span style={{ color: 'var(--danger)' }}>*</span>
                  </label>
                  <select
                    className="form-control"
                    value={assignMandalId}
                    onChange={(e) => setAssignMandalId(e.target.value)}
                    disabled={!assignConstitutionId}
                    required
                  >
                    <option value="">
                      {assignConstitutionId
                        ? `-- Step 3: Select Mandal (${assignMandalsList.length} in this Constituency) --`
                        : '-- First Select Constituency Above --'}
                    </option>
                    {assignMandalsList.map((m) => (
                      <option key={m.id} value={m.id}>
                        {m.name}
                      </option>
                    ))}
                  </select>
                </div>
              </div>
            )}

            {/* Level: NRI */}
            {assignLevel === 'NRI' && (
              <div className="form-group">
                <label className="form-label" style={{ fontWeight: 700 }}>
                  Select NRI Chapter / Wing <span style={{ color: 'var(--danger)' }}>*</span>
                </label>
                <select
                  className="form-control"
                  value={assignNriUnitId}
                  onChange={(e) => setAssignNriUnitId(e.target.value)}
                  required
                >
                  <option value="unit-nri">NRI Global Central Wing</option>
                </select>
              </div>
            )}

            {/* Leadership Position Selector */}
            <div className="form-group">
              <label className="form-label" style={{ fontWeight: 700 }}>
                3. Leadership Position / Office Bearer Role <span style={{ color: 'var(--danger)' }}>*</span>
              </label>
              <select
                className="form-control"
                value={assignPositionCode}
                onChange={(e) => setAssignPositionCode(e.target.value)}
                style={{ fontWeight: 600 }}
              >
                {DEFAULT_TEAM_POSITIONS.map((p) => (
                  <option key={p.code} value={p.code}>
                    {p.title} ({p.maxCount === 1 ? '1 slot' : `up to ${p.maxCount} slots`})
                  </option>
                ))}
              </select>
            </div>

            {/* Target Team & Slot Status Card */}
            {resolvedTargetUnit && (
              <div
                style={{
                  padding: '12px 16px',
                  borderRadius: '8px',
                  background: currentOccupant && assignPositionCode === 'PRESIDENT' ? '#fef2f2' : '#f0f9ff',
                  border: `1px solid ${
                    currentOccupant && assignPositionCode === 'PRESIDENT' ? '#fecaca' : '#bae6fd'
                  }`,
                  fontSize: '13px',
                }}
              >
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', flexWrap: 'wrap', gap: '8px' }}>
                  <div>
                    <div style={{ fontSize: '11px', textTransform: 'uppercase', fontWeight: 800, color: '#0369a1' }}>
                      Target Committee &amp; Position Details
                    </div>
                    <div style={{ fontWeight: 800, fontSize: '14px', color: 'var(--primary-dark)', marginTop: '2px' }}>
                      [{assignLevel}] {resolvedTargetUnit.name} — {assignTeamType}
                    </div>
                    <div style={{ fontSize: '12px', color: 'var(--text-secondary)', marginTop: '2px' }}>
                      📍 {resolvedTargetUnit.locationLabel || resolvedTargetUnit.name}
                    </div>
                  </div>

                  {targetTeamInfo && (
                    <div style={{ textAlign: 'right' }}>
                      <span className="badge badge-active" style={{ fontSize: '11px' }}>
                        {targetTeamInfo.filledPositionsCount || 0} / {targetTeamInfo.totalPositionsCapacity || 30} Filled
                      </span>
                    </div>
                  )}
                </div>

                {/* Conflict / Occupant Notice */}
                {currentOccupant && (
                  <div
                    style={{
                      marginTop: '8px',
                      padding: '8px 10px',
                      background: '#fff',
                      borderRadius: '6px',
                      fontSize: '12px',
                      color: '#b91c1c',
                      border: '1px solid #fecaca',
                    }}
                  >
                    ⚠️ <strong>Note:</strong> "{currentOccupant.positionTitle}" is currently held by{' '}
                    <strong>{currentOccupant.memberName}</strong> ({currentOccupant.membershipId}).
                    {assignPositionCode === 'PRESIDENT' || assignPositionCode === 'TREASURER' ? (
                      <span> Only 1 slot is allowed. Remove or change the existing leader's role first.</span>
                    ) : (
                      <span> Additional slots may be available up to capacity.</span>
                    )}
                  </div>
                )}
              </div>
            )}

            {/* Submit Button */}
            <button
              type="submit"
              className="btn btn-accent"
              disabled={submittingAssignment || !selectedMember || !resolvedTargetUnit}
              style={{
                padding: '12px',
                fontWeight: 800,
                fontSize: '14px',
                display: 'inline-flex',
                alignItems: 'center',
                justifyContent: 'center',
                gap: '8px',
              }}
            >
              <UserPlus size={18} />
              {submittingAssignment
                ? 'Assigning Leadership Position...'
                : 'Confirm & Assign Member to Position'}
            </button>
          </form>
        </div>
      </div>

      {/* ASSIGNED COMMUNITY LEADERS DIRECTORY */}
      <div className="card">
        <div className="card-header" style={{ flexWrap: 'wrap', gap: '8px' }}>
          <div>
            <h2 className="card-title" style={{ fontSize: '18px' }}>
              Assigned Community Leaders Directory ({leaders.length})
            </h2>
            <p style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>
              Filter by State, District, Constituency, Mandal, Team Wing, Position, or Member Search.
            </p>
          </div>
          <button
            type="button"
            className="btn btn-secondary btn-sm"
            onClick={loadLeaders}
            title="Refresh Directory"
            style={{ display: 'inline-flex', alignItems: 'center', gap: '6px' }}
          >
            <RefreshCw size={14} /> Refresh
          </button>
        </div>

        {/* DIRECTORY FILTER BAR */}
        <div
          style={{
            background: 'var(--bg-muted)',
            padding: '16px',
            borderRadius: '8px',
            marginBottom: '16px',
            display: 'flex',
            flexDirection: 'column',
            gap: '12px',
          }}
        >
          {/* Row 1: Search, Level, Wing, Position */}
          <div
            style={{
              display: 'grid',
              gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))',
              gap: '12px',
            }}
          >
            <div className="form-group">
              <label className="form-label" style={{ fontSize: '12px', fontWeight: 700 }}>
                Search Leader / ID / Mobile
              </label>
              <div style={{ display: 'flex', gap: '6px' }}>
                <input
                  type="text"
                  className="form-control"
                  placeholder="Name, MUD-00000001, Mobile..."
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') loadLeaders();
                  }}
                />
                <button
                  type="button"
                  className="btn btn-primary btn-sm"
                  onClick={loadLeaders}
                  title="Search"
                >
                  <Search size={14} />
                </button>
              </div>
            </div>

            <div className="form-group">
              <label className="form-label" style={{ fontSize: '12px', fontWeight: 700 }}>
                Organizational Level
              </label>
              <select
                className="form-control"
                value={orgLevel}
                onChange={(e) => {
                  setOrgLevel(e.target.value);
                  if (e.target.value === 'State' || e.target.value === 'NRI') {
                    setFilterDistrictId('');
                    setFilterConstitutionId('');
                    setFilterMandalId('');
                  }
                }}
              >
                <option value="">All Levels</option>
                {VALID_ORG_LEVELS.map((l) => (
                  <option key={l} value={l}>
                    {l} Level
                  </option>
                ))}
                <option value="NRI">NRI Wing</option>
              </select>
            </div>

            <div className="form-group">
              <label className="form-label" style={{ fontSize: '12px', fontWeight: 700 }}>
                Team Wing / Category
              </label>
              <select
                className="form-control"
                value={teamType}
                onChange={(e) => setTeamType(e.target.value)}
              >
                <option value="">All Wings</option>
                {VALID_TEAM_TYPES.map((t) => (
                  <option key={t} value={t}>
                    {t}
                  </option>
                ))}
              </select>
            </div>

            <div className="form-group">
              <label className="form-label" style={{ fontSize: '12px', fontWeight: 700 }}>
                Leadership Position
              </label>
              <select
                className="form-control"
                value={positionCode}
                onChange={(e) => setPositionCode(e.target.value)}
              >
                <option value="">All Positions</option>
                {DEFAULT_TEAM_POSITIONS.map((p) => (
                  <option key={p.code} value={p.code}>
                    {p.title}
                  </option>
                ))}
              </select>
            </div>
          </div>

          {/* Row 2: Location Cascading Filters (District, Constituency, Mandal) */}
          {(!orgLevel || ['District', 'Constitution', 'Mandal'].includes(orgLevel)) && (
            <div
              style={{
                display: 'grid',
                gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))',
                gap: '12px',
                paddingTop: '10px',
                borderTop: '1px dashed var(--border-color)',
              }}
            >
              <div className="form-group">
                <label className="form-label" style={{ fontSize: '12px', fontWeight: 700 }}>
                  Filter by District
                </label>
                <select
                  className="form-control"
                  value={filterDistrictId}
                  onChange={(e) => {
                    setFilterDistrictId(e.target.value);
                    setFilterConstitutionId('');
                    setFilterMandalId('');
                  }}
                >
                  <option value="">All 26 Districts</option>
                  {districtsList.map((d) => (
                    <option key={d.id} value={d.id}>
                      {d.name} {d.code ? `[${d.code}]` : ''}
                    </option>
                  ))}
                </select>
              </div>

              {(!orgLevel || ['Constitution', 'Mandal'].includes(orgLevel)) && (
                <div className="form-group">
                  <label className="form-label" style={{ fontSize: '12px', fontWeight: 700 }}>
                    Filter by Constituency
                  </label>
                  <select
                    className="form-control"
                    value={filterConstitutionId}
                    onChange={(e) => {
                      setFilterConstitutionId(e.target.value);
                      setFilterMandalId('');
                    }}
                    disabled={!filterDistrictId && (hierarchy.constitutions || []).length > 200}
                  >
                    <option value="">
                      {filterDistrictId
                        ? `All Constituencies in District (${filterConstitutionsList.length})`
                        : 'Select District to filter Constituencies'}
                    </option>
                    {filterConstitutionsList.map((c) => (
                      <option key={c.id} value={c.id}>
                        {c.name} {c.code ? `[${c.code}]` : ''}
                      </option>
                    ))}
                  </select>
                </div>
              )}

              {(!orgLevel || orgLevel === 'Mandal') && (
                <div className="form-group">
                  <label className="form-label" style={{ fontSize: '12px', fontWeight: 700 }}>
                    Filter by Mandal
                  </label>
                  <select
                    className="form-control"
                    value={filterMandalId}
                    onChange={(e) => setFilterMandalId(e.target.value)}
                    disabled={!filterConstitutionId}
                  >
                    <option value="">
                      {filterConstitutionId
                        ? `All Mandals in Constituency (${filterMandalsList.length})`
                        : 'Select Constituency to filter Mandals'}
                    </option>
                    {filterMandalsList.map((m) => (
                      <option key={m.id} value={m.id}>
                        {m.name}
                      </option>
                    ))}
                  </select>
                </div>
              )}

              {(filterDistrictId ||
                filterConstitutionId ||
                filterMandalId ||
                search ||
                orgLevel ||
                teamType ||
                positionCode) && (
                <div className="form-group" style={{ display: 'flex', alignItems: 'flex-end' }}>
                  <button
                    type="button"
                    className="btn btn-secondary btn-sm"
                    style={{ width: '100%', height: '36px' }}
                    onClick={() => {
                      setSearch('');
                      setOrgLevel('');
                      setTeamType('');
                      setPositionCode('');
                      setFilterDistrictId('');
                      setFilterConstitutionId('');
                      setFilterMandalId('');
                    }}
                  >
                    Reset All Filters
                  </button>
                </div>
              )}
            </div>
          )}
        </div>

        {/* LEADERS DIRECTORY TABLE */}
        <div className="table-container">
          <table className="data-table">
            <thead>
              <tr>
                <th>Member Name</th>
                <th>Membership ID</th>
                <th>Photo</th>
                <th>Position</th>
                <th>Wing / Team</th>
                <th>Organizational Level</th>
                <th>Location Details</th>
                <th style={{ textAlign: 'right' }}>Change / Remove</th>
              </tr>
            </thead>
            <tbody>
              {loading ? (
                <tr>
                  <td colSpan={8} style={{ textAlign: 'center', padding: '32px' }}>
                    Loading leaders directory...
                  </td>
                </tr>
              ) : leaders.length === 0 ? (
                <tr>
                  <td
                    colSpan={8}
                    style={{
                      textAlign: 'center',
                      padding: '36px',
                      color: 'var(--text-muted)',
                    }}
                  >
                    No leadership assignments match your filter criteria.
                  </td>
                </tr>
              ) : (
                leaders.map((l) => (
                  <tr key={l.id}>
                    <td>
                      <div style={{ fontWeight: 700 }}>{l.memberName}</div>
                      <div style={{ fontSize: '11px', color: 'var(--text-muted)' }}>
                        📱 {l.mobile}
                      </div>
                    </td>
                    <td>
                      <code
                        style={{
                          fontWeight: 800,
                          color: 'var(--primary)',
                          background: '#f1f5f9',
                          padding: '3px 8px',
                          borderRadius: '4px',
                        }}
                      >
                        {l.membershipId}
                      </code>
                    </td>
                    <td>
                      {l.photoUrl || l.memberPhotoUrl ? (
                        <img
                          src={l.photoUrl || l.memberPhotoUrl}
                          alt={l.memberName}
                          title="Click to view full image"
                          onClick={() =>
                            setPreviewPhoto({
                              src: l.photoUrl || l.memberPhotoUrl,
                              name: l.memberName,
                              membershipId: l.membershipId,
                              positionTitle: l.positionTitle,
                            })
                          }
                          style={{
                            width: '44px',
                            height: '44px',
                            borderRadius: '8px',
                            objectFit: 'cover',
                            border: '1px solid var(--border-strong)',
                            display: 'block',
                            cursor: 'pointer',
                          }}
                        />
                      ) : (
                        <span style={{ fontSize: '12px', color: 'var(--text-muted)' }}>
                          No Photo
                        </span>
                      )}
                    </td>
                    <td>
                      <span
                        className="badge badge-info"
                        style={{ display: 'inline-flex', alignItems: 'center', gap: '4px' }}
                      >
                        <Award size={12} /> {l.positionTitle}
                      </span>
                    </td>
                    <td style={{ fontWeight: 600 }}>{l.teamType}</td>
                    <td>
                      <span className="badge badge-active">{l.orgLevel}</span>
                    </td>
                    <td style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>
                      <div>
                        <strong>{l.orgName}</strong>
                      </div>
                      <div style={{ fontSize: '11px', color: 'var(--text-muted)' }}>
                        {l.location}
                      </div>
                    </td>
                    <td style={{ textAlign: 'right' }}>
                      <div
                        style={{
                          display: 'inline-flex',
                          gap: '6px',
                          alignItems: 'center',
                          justifyContent: 'flex-end',
                        }}
                      >
                        <select
                          className="form-control"
                          style={{ width: 'auto', padding: '4px 8px', fontSize: '12px' }}
                          value={l.positionCode}
                          onChange={(e) =>
                            handleChangeLeaderPosition(l.id, e.target.value)
                          }
                          title="Change Position"
                        >
                          {DEFAULT_TEAM_POSITIONS.map((p) => (
                            <option key={p.code} value={p.code}>
                              {p.title}
                            </option>
                          ))}
                        </select>
                        <button
                          type="button"
                          className="btn btn-danger btn-sm"
                          onClick={() => handleRemoveLeader(l)}
                          title="Remove from Position"
                        >
                          <Trash2 size={13} />
                        </button>
                      </div>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* PHOTO LIGHTBOX POPUP */}
      {previewPhoto && (
        <div
          className="modal-backdrop"
          onClick={() => setPreviewPhoto(null)}
          style={{
            zIndex: 200,
            cursor: 'pointer',
            background: 'rgba(15, 23, 42, 0.82)',
          }}
        >
          <div
            style={{
              background: '#ffffff',
              padding: '16px',
              borderRadius: '14px',
              boxShadow: '0 25px 50px -12px rgba(0, 0, 0, 0.5)',
              maxWidth: '90vw',
              maxHeight: '90vh',
              display: 'flex',
              flexDirection: 'column',
              alignItems: 'center',
              gap: '12px',
            }}
          >
            <img
              src={previewPhoto.src}
              alt={previewPhoto.name}
              style={{
                maxWidth: '420px',
                width: '100%',
                maxHeight: '72vh',
                objectFit: 'contain',
                borderRadius: '10px',
                display: 'block',
              }}
            />
            <div style={{ textAlign: 'center' }}>
              <div
                style={{
                  fontWeight: 800,
                  fontSize: '16px',
                  color: 'var(--primary-dark)',
                }}
              >
                {previewPhoto.name}
              </div>
              <div style={{ fontSize: '13px', color: 'var(--text-secondary)' }}>
                {previewPhoto.membershipId}
                {previewPhoto.positionTitle ? ` • ${previewPhoto.positionTitle}` : ''}
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
