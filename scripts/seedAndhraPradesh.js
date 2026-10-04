import fs from 'fs';
import path from 'path';
import { ANDHRA_PRADESH_DATA } from '../src/lib/andhraPradeshData.js';
import { VALID_TEAM_TYPES } from '../src/lib/validation.js';

const DB_PATH = path.join(process.cwd(), 'data', 'mudiraj_db.json');

function slugify(text) {
  return String(text || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/(^-|-$)/g, '');
}

const customOverrides = {
  'unit-dist-ntr': 'unit-ntr',
  'unit-dist-kurnool': 'unit-knl',
  'unit-dist-nellore': 'unit-nlr',
  'unit-const-vijayawada-central': 'unit-vjacen',
  'unit-const-gudur': 'unit-gudur',
};

export function seedAndhraPradeshData() {
  if (!fs.existsSync(DB_PATH)) {
    console.error('Database file not found at:', DB_PATH);
    return;
  }

  const raw = fs.readFileSync(DB_PATH, 'utf8');
  const db = JSON.parse(raw);
  const now = new Date().toISOString();

  // 1. Ensure Levels
  db.orgLevels = [
    { id: 'level-1', name: 'State', levelRank: 1, status: 'Active', createdAt: now },
    { id: 'level-2', name: 'District', levelRank: 2, status: 'Active', createdAt: now },
    { id: 'level-3', name: 'Constituency', levelRank: 3, status: 'Active', createdAt: now },
    { id: 'level-4', name: 'Mandal', levelRank: 4, status: 'Active', createdAt: now },
    { id: 'level-5', name: 'Gramam', levelRank: 5, status: 'Active', createdAt: now },
    { id: 'level-nri', name: 'NRI', levelRank: 0, status: 'Active', createdAt: now },
  ];

  // 2. Build full Org Units list
  const orgUnits = [];

  // State Unit
  const stateId = 'unit-ap';
  orgUnits.push({
    id: stateId,
    orgLevelId: 'level-1',
    parentId: null,
    name: 'Andhra Pradesh',
    code: 'AP',
    status: 'Active',
    createdAt: now,
  });

  // 26 Districts, 175 Constituencies, 677 Mandals
  ANDHRA_PRADESH_DATA.districts.forEach((dist) => {
    const distId = customOverrides[dist.id] || dist.id;
    orgUnits.push({
      id: distId,
      orgLevelId: 'level-2',
      parentId: stateId,
      name: dist.name,
      code: dist.code,
      status: 'Active',
      createdAt: now,
    });

    dist.constituencies.forEach((ac) => {
      const acId = customOverrides[ac.id] || ac.id;
      orgUnits.push({
        id: acId,
        orgLevelId: 'level-3',
        parentId: distId,
        name: ac.name,
        code: ac.code,
        status: 'Active',
        createdAt: now,
      });

      ac.mandals.forEach((mandal) => {
        const mandalCleanName = mandal.endsWith('Mandal') ? mandal : `${mandal} Mandal`;
        const mKey = mandal.replace(/ Mandal$/i, '').trim();
        const mandalId =
          mKey === 'Vijayawada Urban'
            ? 'unit-vjaurban'
            : mKey === 'Naidupeta'
            ? 'unit-naidupeta'
            : `unit-mnd-${slugify(ac.name)}-${slugify(mKey)}`;

        orgUnits.push({
          id: mandalId,
          orgLevelId: 'level-4',
          parentId: acId,
          name: mandalCleanName,
          code: `MND-${slugify(mKey).slice(0, 8).toUpperCase()}`,
          status: 'Active',
          createdAt: now,
        });
      });
    });
  });

  // Preserve NRI Global Central Wing
  orgUnits.push({
    id: 'unit-nri',
    orgLevelId: 'level-nri',
    parentId: null,
    name: 'NRI Global Central Wing',
    code: 'NRI-GLB',
    status: 'Active',
    createdAt: now,
  });

  // Preserve any other custom NRI chapter units
  const existingCustomNriUnits = (db.orgUnits || []).filter(
    (u) => u.orgLevelId === 'level-nri' && u.id !== 'unit-nri'
  );
  orgUnits.push(...existingCustomNriUnits);

  db.orgUnits = orgUnits;

  // 3. Ensure Teams for all Units (Main Team, Youth Team, Mahila Team)
  const existingTeamsMap = new Map();
  (db.teams || []).forEach((t) => existingTeamsMap.set(t.id, t));

  const allTeams = [];

  // Preserve existing NRI teams
  const nriTeams = (db.teams || []).filter((t) => t.isNri || t.orgUnitId === 'unit-nri');
  nriTeams.forEach((t) => allTeams.push(t));

  // Build / retain 3 teams for each AP unit
  db.orgUnits.forEach((unit) => {
    if (unit.id === 'unit-nri' || unit.orgLevelId === 'level-nri') return;

    VALID_TEAM_TYPES.forEach((teamType) => {
      const slug = teamType.toLowerCase().replace(/\s+/g, '-');
      const teamId = `team-${unit.id}-${slug}`;

      if (existingTeamsMap.has(teamId)) {
        const existingTeam = existingTeamsMap.get(teamId);
        existingTeam.orgName = unit.name;
        existingTeam.status = 'Active';
        allTeams.push(existingTeam);
      } else {
        allTeams.push({
          id: teamId,
          teamType,
          orgUnitId: unit.id,
          orgName: unit.name,
          executiveMemberLimit: 14,
          status: 'Active',
          createdAt: now,
        });
      }
    });
  });

  db.teams = allTeams;

  // 4. Save Database
  fs.writeFileSync(DB_PATH, JSON.stringify(db, null, 2), 'utf8');

  console.log('Successfully seeded Andhra Pradesh administrative hierarchy:');
  console.log('- Total Org Units:', db.orgUnits.length);
  console.log('  • State:', db.orgUnits.filter((u) => u.orgLevelId === 'level-1').length);
  console.log('  • Districts:', db.orgUnits.filter((u) => u.orgLevelId === 'level-2').length);
  console.log('  • Constituencies:', db.orgUnits.filter((u) => u.orgLevelId === 'level-3').length);
  console.log('  • Mandals:', db.orgUnits.filter((u) => u.orgLevelId === 'level-4').length);
  console.log('  • NRI Wings:', db.orgUnits.filter((u) => u.orgLevelId === 'level-nri').length);
  console.log('- Total Teams:', db.teams.length);
  console.log('- Total Members:', db.members.length);
  console.log('- Total Leaders:', db.teamMembers.length);
}

seedAndhraPradeshData();
