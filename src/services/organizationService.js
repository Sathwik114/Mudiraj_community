import {
  getAllOrganizationHierarchy,
  getOrganizationById,
  createOrganizationUnit,
  updateOrganizationUnit,
  getHierarchyPath,
} from '@/lib/organizations';
import { getDatabase } from '@/lib/database';

/**
 * Organization Service (`services/organizationService.js`)
 * --------------------------------------------------------
 * Manages the dynamic organizational structure.
 */

export const organizationService = {
  getHierarchy({ includeInactive = true } = {}) {
    const raw = getAllOrganizationHierarchy({ includeInactive });
    const db = getDatabase();

    const unitMap = new Map();
    db.orgUnits.forEach((u) => unitMap.set(u.id, u));

    const childrenCountMap = new Map();
    db.orgUnits.forEach((u) => {
      if (u.parentId) {
        childrenCountMap.set(u.parentId, (childrenCountMap.get(u.parentId) || 0) + 1);
      }
    });

    const activeMembersCountMap = new Map();
    db.members.forEach((m) => {
      if (m.orgUnitId && ['Active', 'Approved'].includes(m.status)) {
        activeMembersCountMap.set(m.orgUnitId, (activeMembersCountMap.get(m.orgUnitId) || 0) + 1);
      }
    });

    // Attach counts and hierarchy path to each org unit
    const enrichedUnits = raw.orgUnits.map((u) => {
      const childrenCount = childrenCountMap.get(u.id) || 0;
      const membersCount = activeMembersCountMap.get(u.id) || 0;
      
      const path = [];
      let currentId = u.id;
      while (currentId) {
        const current = unitMap.get(currentId);
        if (current) {
          path.unshift(current);
          currentId = current.parentId;
        } else {
          break;
        }
      }

      const parentName = u.parentId ? unitMap.get(u.parentId)?.name : null;

      return {
        ...u,
        parentName,
        childrenCount,
        membersCount,
        hierarchyPath: path,
        locationLabel: path.map(p => p.name).join(' › '),
      };
    });

    const stateLevel = raw.orgLevels.find(l => l.levelRank === 1 || l.name === 'State');
    const districtLevel = raw.orgLevels.find(l => l.levelRank === 2 || l.name === 'District');
    const constLevel = raw.orgLevels.find(l => l.levelRank === 3 || l.name === 'Constituency');
    const mandalLevel = raw.orgLevels.find(l => l.levelRank === 4 || l.name === 'Mandal');
    const gramamLevel = raw.orgLevels.find(l => l.levelRank === 5 || l.name === 'Gramam');

    return {
      orgLevels: raw.orgLevels,
      orgUnits: enrichedUnits,
      states: enrichedUnits.filter(u => u.orgLevelId === stateLevel?.id),
      districts: enrichedUnits.filter(u => u.orgLevelId === districtLevel?.id),
      constitutions: enrichedUnits.filter(u => u.orgLevelId === constLevel?.id),
      mandals: enrichedUnits.filter(u => u.orgLevelId === mandalLevel?.id),
      gramams: enrichedUnits.filter(u => u.orgLevelId === gramamLevel?.id),
    };
  },

  getUnitById(id) {
    return getOrganizationById(id);
  },

  createUnit(payload) {
    return createOrganizationUnit(payload);
  },

  updateUnit(payload) {
    return updateOrganizationUnit(payload);
  },
};
