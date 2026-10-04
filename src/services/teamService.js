import {
  queryTeams,
  ensureTeamsForOrganization,
  updateTeamExecutiveCapacity,
  assignLeaderToTeam,
  changeLeaderPosition,
  removeLeaderFromTeam,
  queryAllLeaders,
  ensureNriTeamsExist,
  getNriTeams,
  createNriChapter,
  createDirectNriLeader,
} from '@/lib/teams';

/**
 * Team & Leadership Service (`services/teamService.js`)
 * -----------------------------------------------------
 * Reusable team management across Main Team, Youth Team, and Mahila Team
 * for State, District, Constitution, Mandal, and Gramam levels,
 * plus dedicated separate NRI Teams (Main, Youth, Mahila).
 */

export const teamService = {
  getTeams(filters = {}) {
    return queryTeams(filters);
  },

  ensureTeamsExist({ orgLevel, orgId }) {
    return ensureTeamsForOrganization({ orgLevel, orgId });
  },

  setExecutiveCapacity({ teamId, executiveMemberLimit, actor }) {
    return updateTeamExecutiveCapacity({ teamId, executiveMemberLimit, actor });
  },

  assignLeader({ teamId, memberId, positionCode, actor }) {
    return assignLeaderToTeam({ teamId, memberId, positionCode, actor });
  },

  updateLeaderRole({ teamMemberId, newPositionCode, actor }) {
    return changeLeaderPosition({ teamMemberId, newPositionCode, actor });
  },

  removeLeader({ teamMemberId, actor }) {
    return removeLeaderFromTeam({ teamMemberId, actor });
  },

  getLeadersDirectory(filters = {}) {
    return queryAllLeaders(filters);
  },

  // NRI Dedicated Methods
  ensureNriTeamsExist() {
    return ensureNriTeamsExist();
  },

  getNriTeams(filters = {}) {
    return getNriTeams(filters);
  },

  createNriChapter(params) {
    return createNriChapter(params);
  },

  assignDirectNriLeader(params) {
    return createDirectNriLeader(params);
  },
};

