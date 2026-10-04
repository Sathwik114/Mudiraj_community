import {
  getDatabase,
  saveDatabase,
  generateUuid,
  buildDefaultTeamsForOrg,
  recordAuditLog,
} from './database.js';
import { getHierarchyPath } from './organizations.js';
import { DEFAULT_TEAM_POSITIONS } from './validation.js';

/**
 * Reusable Team & Leadership Repository Layer (`lib/teams.js`)
 * ------------------------------------------------------------
 * Implements a single unified Team model across the dynamic org hierarchy.
 */

export function getPositionDefinitionsForTeam(team) {
  const execLimit = typeof team?.executiveMemberLimit === 'number' ? team.executiveMemberLimit : 14;
  return DEFAULT_TEAM_POSITIONS.map((pos) => {
    if (pos.code === 'EXECUTIVE_MEMBER') {
      return { ...pos, maxCount: execLimit };
    }
    return { ...pos };
  });
}

/**
 * Enriches a team record with full hierarchy names and assigned leadership roster.
 */
export function enrichTeamDetails(team, db = getDatabase()) {
  const path = team.orgUnitId ? getHierarchyPath(team.orgUnitId) : [];
  const unit = team.orgUnitId && db.orgUnits ? db.orgUnits.find((u) => u.id === team.orgUnitId) : null;
  const levelObj = unit && db.orgLevels ? db.orgLevels.find((l) => l.id === unit.orgLevelId) : null;
  const orgLevel = team.orgLevel || (levelObj ? levelObj.name : (team.isNri ? 'NRI' : ''));
  const orgLevelId = team.orgLevelId || (unit ? unit.orgLevelId : (team.isNri ? 'level-nri' : ''));

  const positionDefs = getPositionDefinitionsForTeam(team);
  const totalCapacity = positionDefs.reduce((acc, p) => acc + p.maxCount, 0);

  const assignments = (db.teamMembers || [])
    .filter((tm) => tm.teamId === team.id && tm.status === 'Active')
    .map((tm) => {
      const member = db.members ? db.members.find(
        (m) =>
          m.id === tm.memberId ||
          (tm.membershipId && m.membershipId === tm.membershipId)
      ) : null;
      const posDef = positionDefs.find((p) => p.code === tm.positionCode);
      const memberPath = member?.orgUnitId ? getHierarchyPath(member.orgUnitId) : [];
      const memberLocationLabel = memberPath.map((p) => p.name).join(' › ');
      return {
        ...tm,
        sortOrder: posDef ? posDef.sortOrder : 99,
        memberName: member ? member.fullName : (tm.memberName || 'Unknown Member'),
        memberMobile: member ? member.mobile : (tm.memberMobile || ''),
        memberStatus: member ? member.status : 'Active',
        memberPhotoUrl: member ? member.photoUrl : (tm.photoUrl || ''),
        photoUrl: member ? member.photoUrl : (tm.photoUrl || ''),
        memberLocationPath: member ? member.locationPath : '',
        memberMandalName: member?.mandalName || memberLocationLabel || team.orgName,
        country: member?.country || tm.country || (team.isNri ? 'NRI / International' : ''),
      };
    })
    .sort((a, b) => a.sortOrder - b.sortOrder || a.slotNumber - b.slotNumber);

  const locationLabel = path.length > 0 
    ? path.map((p) => p.name).join(' › ') 
    : (team.locationLabel || (team.isNri ? 'Global NRI Wing' : team.orgName));

  return {
    ...team,
    orgLevel: team.orgLevel || (team.isNri ? 'NRI' : ''),
    orgLevelId,
    locationLabel,
    hierarchyPath: path,
    positionDefinitions: positionDefs,
    totalPositionsCapacity: totalCapacity,
    filledPositionsCount: assignments.length,
    leaders: assignments,
  };
}

/**
 * Queries teams across any organizational level and category.
 */
export function queryTeams({
  orgUnitId = '',
  orgId = '',
  orgLevelId = '',
  orgLevel = '',
  teamType = '',
  search = '',
  includeNri = false,
} = {}) {
  const db = getDatabase();
  let filtered = db.teams;

  // By default, exclude NRI teams from standard AP hierarchical queries
  if (!includeNri && orgLevel !== 'NRI' && orgLevelId !== 'level-nri' && orgUnitId !== 'unit-nri') {
    filtered = filtered.filter((t) => !t.isNri && t.orgUnitId !== 'unit-nri');
  }

  const targetUnitId = orgUnitId || orgId;
  if (targetUnitId) {
    const unit = db.orgUnits.find((u) => u.id === targetUnitId);
    if (unit) {
      const matchingUnits = db.orgUnits.filter((u) =>
        getHierarchyPath(u.id).some((p) => p.id === targetUnitId)
      );
      const validUnitIds = new Set(matchingUnits.map((u) => u.id));
      filtered = filtered.filter((t) => validUnitIds.has(t.orgUnitId));
    }
  }

  if (orgLevelId) {
    const validUnitIds = new Set(
      db.orgUnits.filter((u) => u.orgLevelId === orgLevelId).map((u) => u.id)
    );
    filtered = filtered.filter((t) => validUnitIds.has(t.orgUnitId));
  } else if (orgLevel) {
    const normLevel = orgLevel.toLowerCase();
    const matchingLevels = db.orgLevels.filter(
      (l) =>
        l.name.toLowerCase() === normLevel ||
        (normLevel === 'constitution' && l.name.toLowerCase() === 'constituency') ||
        (normLevel === 'constituency' && l.name.toLowerCase() === 'constitution')
    );
    const levelIds = new Set(matchingLevels.map((l) => l.id));
    const validUnitIds = new Set(
      db.orgUnits.filter((u) => levelIds.has(u.orgLevelId)).map((u) => u.id)
    );
    filtered = filtered.filter((t) => validUnitIds.has(t.orgUnitId) || t.orgLevel === orgLevel);
  }

  if (teamType) {
    filtered = filtered.filter((t) => t.teamType === teamType);
  }

  if (search && search.trim()) {
    const q = search.trim().toLowerCase();
    filtered = filtered.filter(
      (t) =>
        t.orgName.toLowerCase().includes(q) ||
        t.teamType.toLowerCase().includes(q)
    );
  }

  return filtered.map((t) => enrichTeamDetails(t, db));
}

/**
 * Ensures all 3 teams (Main, Youth, Ladies) exist for a given organization unit.
 */
export function ensureTeamsForOrganization({ orgUnitId }) {
  const db = getDatabase();
  const unit = db.orgUnits.find(u => u.id === orgUnitId);
  if (!unit) throw new Error(`Organization unit not found (${orgUnitId})`);

  const existing = db.teams.filter((t) => t.orgUnitId === orgUnitId);
  if (existing.length === 3) {
    return existing.map((t) => enrichTeamDetails(t, db));
  }

  const defaults = buildDefaultTeamsForOrg({
    orgUnitId: unit.id,
    orgName: unit.name,
  });

  defaults.forEach((defTeam) => {
    const already = db.teams.some(
      (t) => t.orgUnitId === orgUnitId && t.teamType === defTeam.teamType
    );
    if (!already) {
      db.teams.push(defTeam);
    }
  });

  saveDatabase(db);
  return db.teams
    .filter((t) => t.orgUnitId === orgUnitId)
    .map((t) => enrichTeamDetails(t, db));
}

/**
 * Updates the Executive Member capacity for a specific team (since Executive Members are variable).
 */
export function updateTeamExecutiveCapacity({ teamId, executiveMemberLimit, actor = 'admin' }) {
  const db = getDatabase();
  const team = db.teams.find((t) => t.id === teamId);
  if (!team) throw new Error('Team not found.');

  const limitNum = Number(executiveMemberLimit);
  if (!Number.isInteger(limitNum) || limitNum < 1 || limitNum > 200) {
    throw new Error('Executive Member capacity must be an integer between 1 and 200.');
  }

  const currentExecCount = db.teamMembers.filter(
    (tm) => tm.teamId === teamId && tm.positionCode === 'EXECUTIVE_MEMBER' && tm.status === 'Active'
  ).length;

  if (limitNum < currentExecCount) {
    throw new Error(
      `Cannot reduce Executive Member slots below currently assigned count (${currentExecCount}).`
    );
  }

  team.executiveMemberLimit = limitNum;
  saveDatabase(db);

  recordAuditLog({
    action: 'TEAM_CAPACITY_UPDATED',
    entityType: 'Team',
    entityId: team.id,
    actor,
    details: `Updated Executive Member capacity for ${team.orgName} - ${team.teamType} to ${limitNum} positions.`,
  });

  return enrichTeamDetails(team, db);
}

/**
 * Assigns an approved/active member to a leadership position in a team.
 * Prevents duplicate/conflicting positions in the same team and enforces slot limits.
 */
export function assignLeaderToTeam({
  teamId,
  memberId,
  positionCode,
  actor = 'admin',
}) {
  const db = getDatabase();
  const team = db.teams.find((t) => t.id === teamId);
  if (!team) throw new Error('Selected Team does not exist.');

  const member = db.members.find((m) => m.id === memberId);
  if (!member) throw new Error('Selected Member does not exist.');

  if (!member.membershipId || !['Active', 'Approved'].includes(member.status)) {
    throw new Error(
      'Only Approved/Active members with a valid Membership ID can be assigned to leadership positions.'
    );
  }

  // Prevent same member from holding conflicting duplicate positions within the same team
  const existingInTeam = db.teamMembers.find(
    (tm) => tm.teamId === teamId && tm.memberId === memberId && tm.status === 'Active'
  );
  if (existingInTeam) {
    throw new Error(
      `${member.fullName} (${member.membershipId}) already holds the position of "${existingInTeam.positionTitle}" in ${team.orgName} - ${team.teamType}. Use "Change Position" to update their role.`
    );
  }

  const posDefs = getPositionDefinitionsForTeam(team);
  const posDef = posDefs.find((p) => p.code === positionCode);
  if (!posDef) {
    throw new Error(`Invalid leadership position code: ${positionCode}`);
  }

  const activeInPosition = db.teamMembers.filter(
    (tm) => tm.teamId === teamId && tm.positionCode === positionCode && tm.status === 'Active'
  );

  if (activeInPosition.length >= posDef.maxCount) {
    throw new Error(
      `All ${posDef.maxCount} slot(s) for "${posDef.title}" in ${team.orgName} - ${team.teamType} are currently filled.`
    );
  }

  const usedSlots = new Set(activeInPosition.map((tm) => tm.slotNumber));
  let slotNumber = 1;
  while (usedSlots.has(slotNumber)) {
    slotNumber += 1;
  }

  const assignment = {
    id: generateUuid('tm'),
    teamId: team.id,
    memberId: member.id,
    membershipId: member.membershipId,
    positionCode: posDef.code,
    positionTitle: posDef.title,
    slotNumber,
    assignedAt: new Date().toISOString(),
    assignedBy: actor,
    status: 'Active',
  };

  db.teamMembers.push(assignment);
  saveDatabase(db);

  recordAuditLog({
    action: 'LEADER_ASSIGNED',
    entityType: 'TeamMember',
    entityId: assignment.id,
    actor,
    details: `Assigned ${member.fullName} (${member.membershipId}) as ${posDef.title} in ${team.orgName} (${team.teamType}).`,
  });

  return enrichTeamDetails(team, db);
}

/**
 * Changes an existing leader's position within a team.
 */
export function changeLeaderPosition({ teamMemberId, newPositionCode, actor = 'admin' }) {
  const db = getDatabase();
  const tm = db.teamMembers.find((item) => item.id === teamMemberId && item.status === 'Active');
  if (!tm) throw new Error('Leadership assignment record not found.');

  const team = db.teams.find((t) => t.id === tm.teamId);
  if (!team) throw new Error('Associated Team not found.');

  if (tm.positionCode === newPositionCode) {
    return enrichTeamDetails(team, db);
  }

  const posDefs = getPositionDefinitionsForTeam(team);
  const newPosDef = posDefs.find((p) => p.code === newPositionCode);
  if (!newPosDef) throw new Error('Invalid target leadership position.');

  const activeInNewPosition = db.teamMembers.filter(
    (item) =>
      item.teamId === team.id &&
      item.positionCode === newPositionCode &&
      item.status === 'Active' &&
      item.id !== tm.id
  );

  if (activeInNewPosition.length >= newPosDef.maxCount) {
    throw new Error(
      `All ${newPosDef.maxCount} slot(s) for "${newPosDef.title}" in ${team.orgName} - ${team.teamType} are already occupied.`
    );
  }

  const usedSlots = new Set(activeInNewPosition.map((item) => item.slotNumber));
  let slotNumber = 1;
  while (usedSlots.has(slotNumber)) {
    slotNumber += 1;
  }

  const oldTitle = tm.positionTitle;
  tm.positionCode = newPosDef.code;
  tm.positionTitle = newPosDef.title;
  tm.slotNumber = slotNumber;
  tm.assignedAt = new Date().toISOString();

  saveDatabase(db);

  const member = db.members.find((m) => m.id === tm.memberId);
  recordAuditLog({
    action: 'LEADER_POSITION_CHANGED',
    entityType: 'TeamMember',
    entityId: tm.id,
    actor,
    details: `Changed position of ${member?.fullName || tm.membershipId} from "${oldTitle}" to "${newPosDef.title}" in ${team.orgName} (${team.teamType}).`,
  });

  return enrichTeamDetails(team, db);
}

/**
 * Removes a member from a leadership position.
 */
export function removeLeaderFromTeam({ teamMemberId, actor = 'admin' }) {
  const db = getDatabase();
  const index = db.teamMembers.findIndex((item) => item.id === teamMemberId);
  if (index === -1) throw new Error('Leadership assignment not found.');

  const removed = db.teamMembers[index];
  const team = db.teams.find((t) => t.id === removed.teamId);
  const member = db.members.find((m) => m.id === removed.memberId);

  db.teamMembers.splice(index, 1);
  saveDatabase(db);

  recordAuditLog({
    action: 'LEADER_REMOVED',
    entityType: 'TeamMember',
    entityId: teamMemberId,
    actor,
    details: `Removed ${member?.fullName || removed.membershipId} from ${removed.positionTitle} in ${team?.orgName || ''} (${team?.teamType || ''}).`,
  });

  return team ? enrichTeamDetails(team, db) : null;
}

/**
 * Lists all active leaders across all teams with filtering and search.
 */
export function queryAllLeaders({
  search = '',
  orgUnitId = '',
  orgLevel = '',
  teamType = '',
  positionCode = '',
} = {}) {
  const db = getDatabase();

  const results = db.teamMembers
    .filter((tm) => tm.status === 'Active')
    .map((tm) => {
      const team = db.teams.find((t) => t.id === tm.teamId);
      const member = db.members.find(
        (m) =>
          m.id === tm.memberId ||
          (tm.membershipId && m.membershipId === tm.membershipId)
      );
      if (!team || !member) return null;

      const unit = db.orgUnits.find((u) => u.id === team.orgUnitId);
      const levelObj = unit ? db.orgLevels.find((l) => l.id === unit.orgLevelId) : null;
      const teamOrgLevel = team.orgLevel || (levelObj ? levelObj.name : '');

      const path = getHierarchyPath(team.orgUnitId);
      const locationLabel = path.map((p) => p.name).join(' › ');

      return {
        id: tm.id,
        teamId: team.id,
        memberId: member.id,
        memberName: member.fullName,
        membershipId: member.membershipId,
        mobile: member.mobile,
        gender: member.gender,
        photoUrl: member.photoUrl,
        memberPhotoUrl: member.photoUrl,
        positionCode: tm.positionCode,
        positionTitle: tm.positionTitle,
        slotNumber: tm.slotNumber,
        teamType: team.teamType,
        orgLevel: teamOrgLevel,
        orgUnitId: team.orgUnitId,
        orgName: team.orgName,
        location: locationLabel,
        assignedAt: tm.assignedAt,
        hierarchyPath: path,
      };
    })
    .filter(Boolean);

  return results.filter((item) => {
    if (teamType && item.teamType !== teamType) return false;
    if (positionCode && item.positionCode !== positionCode) return false;
    if (orgLevel && item.orgLevel.toLowerCase() !== orgLevel.toLowerCase()) return false;
    if (orgUnitId) {
      if (!item.hierarchyPath.some((p) => p.id === orgUnitId)) return false;
    }

    if (search && search.trim()) {
      const q = search.trim().toLowerCase();
      const matchName = item.memberName.toLowerCase().includes(q);
      const matchId = (item.membershipId || '').toLowerCase().includes(q);
      const matchMobile = (item.mobile || '').includes(q);
      const matchPos = item.positionTitle.toLowerCase().includes(q);
      const matchOrg = item.orgName.toLowerCase().includes(q);
      if (!matchName && !matchId && !matchMobile && !matchPos && !matchOrg) {
        return false;
      }
    }
    return true;
  });
}

/**
 * =========================================================================
 * NRI (Non-Resident Indian) Specialized Team & Leadership Module
 * =========================================================================
 * NRI Team is kept strictly separate from the Andhra Pradesh state hierarchy.
 * Under this NRI wing, 3 distinct teams are available:
 * 1. Main Team (NRI Main Team)
 * 2. Youth Team (NRI Youth Team)
 * 3. Mahila Team (NRI Mahila Team)
 *
 * Each team features all 6 positions:
 * - President / Chairman (1 slot)
 * - Vice President (6 slots)
 * - General Secretary (2 slots)
 * - Secretary (6 slots)
 * - Treasurer (1 slot)
 * - Executive Member (variable slots, default 14)
 */

export function ensureNriTeamsExist(db = getDatabase()) {
  const now = new Date().toISOString();

  // 1. Ensure NRI Level exists in orgLevels
  if (!db.orgLevels.some((l) => l.id === 'level-nri' || l.name === 'NRI')) {
    db.orgLevels.push({
      id: 'level-nri',
      name: 'NRI',
      levelRank: 99,
      status: 'Active',
      isNri: true,
      createdAt: now,
    });
  }

  // 2. Ensure Global NRI Org Unit exists in orgUnits
  let nriUnit = db.orgUnits.find((u) => u.id === 'unit-nri');
  if (!nriUnit) {
    nriUnit = {
      id: 'unit-nri',
      orgLevelId: 'level-nri',
      parentId: null,
      name: 'NRI Global Central Wing',
      code: 'NRI-GLB',
      status: 'Active',
      isNri: true,
      createdAt: now,
    };
    db.orgUnits.push(nriUnit);
  }

  // 3. Ensure the 3 distinct NRI teams (Main, Youth, Mahila) exist
  const defaultNriTeams = [
    {
      id: 'team-nri-main-team',
      teamType: 'Main Team',
      orgUnitId: 'unit-nri',
      orgName: 'NRI Global Central Wing',
      isNri: true,
      orgLevel: 'NRI',
      locationLabel: 'International / Global Diaspora',
      executiveMemberLimit: 14,
      status: 'Active',
      createdAt: now,
    },
    {
      id: 'team-nri-youth-team',
      teamType: 'Youth Team',
      orgUnitId: 'unit-nri',
      orgName: 'NRI Global Central Wing',
      isNri: true,
      orgLevel: 'NRI',
      locationLabel: 'International / Global Diaspora',
      executiveMemberLimit: 14,
      status: 'Active',
      createdAt: now,
    },
    {
      id: 'team-nri-mahila-team',
      teamType: 'Mahila Team',
      orgUnitId: 'unit-nri',
      orgName: 'NRI Global Central Wing',
      isNri: true,
      orgLevel: 'NRI',
      locationLabel: 'International / Global Diaspora',
      executiveMemberLimit: 14,
      status: 'Active',
      createdAt: now,
    },
  ];

  let modified = false;
  defaultNriTeams.forEach((dt) => {
    const exists = db.teams.some((t) => t.id === dt.id);
    if (!exists) {
      db.teams.push(dt);
      modified = true;
    }
  });

  if (modified) {
    saveDatabase(db);
  }

  return db.teams.filter(
    (t) => t.isNri || t.orgUnitId === 'unit-nri' || t.orgUnitId?.startsWith('unit-nri-') || t.id.startsWith('team-nri-')
  );
}

/**
 * Retrieves NRI teams (enriched with positions & assigned leaders) and chapters.
 */
export function getNriTeams({ chapterId = '', teamType = '', search = '' } = {}) {
  const db = getDatabase();
  ensureNriTeamsExist(db);

  let nriTeams = db.teams.filter(
    (t) => t.isNri || t.orgUnitId === 'unit-nri' || t.orgUnitId?.startsWith('unit-nri-') || t.id.startsWith('team-nri-')
  );

  if (chapterId) {
    nriTeams = nriTeams.filter((t) => t.orgUnitId === chapterId);
  }

  if (teamType) {
    nriTeams = nriTeams.filter((t) => t.teamType === teamType);
  }

  if (search && search.trim()) {
    const q = search.trim().toLowerCase();
    nriTeams = nriTeams.filter(
      (t) =>
        t.orgName.toLowerCase().includes(q) ||
        t.teamType.toLowerCase().includes(q) ||
        (t.locationLabel && t.locationLabel.toLowerCase().includes(q))
    );
  }

  const enriched = nriTeams.map((t) => enrichTeamDetails(t, db));

  const chapters = db.orgUnits
    .filter((u) => u.isNri || u.id === 'unit-nri' || u.id.startsWith('unit-nri-'))
    .map((u) => ({
      id: u.id,
      name: u.name,
      code: u.code,
      parentId: u.parentId,
    }));

  return {
    teams: enriched,
    chapters,
  };
}

/**
 * Creates an optional regional/country NRI chapter with its 3 teams (Main, Youth, Mahila).
 */
export function createNriChapter({ name, countryCode = '', description = '', actor = 'admin' }) {
  const db = getDatabase();
  ensureNriTeamsExist(db);

  const cleanName = String(name || '').trim();
  if (!cleanName) throw new Error('Country or Regional Chapter name is required.');

  const slug = cleanName.toLowerCase().replace(/[^a-z0-9]+/g, '-');
  const unitId = `unit-nri-${slug}`;
  const code = (countryCode || cleanName.slice(0, 4)).trim().toUpperCase();

  if (db.orgUnits.some((u) => u.id === unitId || u.name.toLowerCase() === `nri — ${cleanName.toLowerCase()}`)) {
    throw new Error(`An NRI chapter named "${cleanName}" already exists.`);
  }

  const now = new Date().toISOString();
  const unit = {
    id: unitId,
    orgLevelId: 'level-nri',
    parentId: 'unit-nri',
    name: `NRI — ${cleanName}`,
    code: `NRI-${code}`,
    status: 'Active',
    isNri: true,
    description,
    createdAt: now,
  };
  db.orgUnits.push(unit);

  const newTeams = ['Main Team', 'Youth Team', 'Mahila Team'].map((type) => {
    const typeSlug = type.toLowerCase().replace(/\s+/g, '-');
    return {
      id: `team-${unitId}-${typeSlug}`,
      teamType: type,
      orgUnitId: unitId,
      orgName: `NRI — ${cleanName}`,
      isNri: true,
      orgLevel: 'NRI',
      locationLabel: `${cleanName} Diaspora Chapter`,
      executiveMemberLimit: 14,
      status: 'Active',
      createdAt: now,
    };
  });

  db.teams.push(...newTeams);
  saveDatabase(db);

  recordAuditLog({
    action: 'NRI_CHAPTER_CREATED',
    entityType: 'OrgUnit',
    entityId: unitId,
    actor,
    details: `Created NRI chapter "${cleanName}" with Main, Youth, and Mahila teams.`,
  });

  return {
    chapter: unit,
    teams: newTeams.map((t) => enrichTeamDetails(t, db)),
  };
}

/**
 * Directly registers and assigns an NRI leader with country & contact details.
 */
export function createDirectNriLeader({
  teamId,
  positionCode,
  fullName,
  country = 'NRI / Overseas',
  city = '',
  mobile = '',
  email = '',
  actor = 'admin',
}) {
  const db = getDatabase();
  ensureNriTeamsExist(db);

  const team = db.teams.find((t) => t.id === teamId);
  if (!team) throw new Error('Selected NRI team does not exist.');

  const cleanName = String(fullName || '').trim();
  if (!cleanName) throw new Error('Full Name is required for NRI leader assignment.');

  const now = new Date().toISOString();
  const seq = db.sequences?.membershipId || (db.members.length + 1);
  const membershipId = `NRI-${String(seq).padStart(6, '0')}`;
  if (db.sequences) {
    db.sequences.membershipId = seq + 1;
  }

  const memberId = generateUuid('mem');
  const newMember = {
    id: memberId,
    applicationNo: `NRI-APP-${seq}`,
    membershipId,
    fullName: cleanName,
    country: country.trim() || 'International',
    city: city.trim(),
    mobile: mobile.trim() || 'N/A',
    email: email.trim(),
    gender: team.teamType === 'Mahila Team' ? 'Female' : 'Male',
    status: 'Active',
    orgUnitId: team.orgUnitId || 'unit-nri',
    locationPath: '/unit-nri/',
    applicationDate: now,
    approvalDate: now,
    approvedBy: actor,
    remarks: `Direct NRI Community Leader for ${team.orgName} (${team.teamType})`,
    createdAt: now,
    updatedAt: now,
  };

  db.members.push(newMember);
  saveDatabase(db);

  return assignLeaderToTeam({
    teamId,
    memberId,
    positionCode,
    actor,
  });
}

