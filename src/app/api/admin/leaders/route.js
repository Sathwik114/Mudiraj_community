import { NextResponse } from 'next/server';
import { requireAdminAuth } from '@/lib/auth';
import { teamService } from '@/services/teamService';

export async function GET(request) {
  const { authorized } = await requireAdminAuth();
  if (!authorized) {
    return NextResponse.json({ error: 'Unauthorized administrator access.' }, { status: 401 });
  }

  const { searchParams } = new URL(request.url);
  const search = searchParams.get('search') || '';
  const orgLevel = searchParams.get('orgLevel') || '';
  const teamType = searchParams.get('teamType') || '';
  const positionCode = searchParams.get('positionCode') || '';
  const districtId = searchParams.get('districtId') || '';
  const constitutionId = searchParams.get('constitutionId') || '';
  const mandalId = searchParams.get('mandalId') || '';

  const leaders = teamService.getLeadersDirectory({
    search,
    orgLevel,
    teamType,
    positionCode,
    districtId,
    constitutionId,
    mandalId,
  });

  return NextResponse.json({ leaders });
}

export async function POST(request) {
  const { authorized, admin } = await requireAdminAuth();
  if (!authorized) {
    return NextResponse.json({ error: 'Unauthorized administrator access.' }, { status: 401 });
  }

  try {
    const { teamId, orgUnitId, teamType, memberId, positionCode } = await request.json();
    let targetTeamId = teamId;

    if (!targetTeamId && orgUnitId && teamType) {
      const teams = teamService.getTeams({ orgId: orgUnitId, exact: true, includeNri: true });
      const found = teams.find((t) => t.teamType === teamType);
      if (found) {
        targetTeamId = found.id;
      } else {
        const ensured = teamService.ensureTeamsExist({ orgId: orgUnitId });
        const created = (ensured || []).find((t) => t.teamType === teamType);
        targetTeamId = created?.id;
      }
    }

    if (!targetTeamId) {
      return NextResponse.json(
        { error: 'Please select a valid organizational level, target area, and team category.' },
        { status: 400 }
      );
    }

    const updatedTeam = teamService.assignLeader({
      teamId: targetTeamId,
      memberId,
      positionCode,
      actor: admin.username,
    });

    return NextResponse.json(
      {
        success: true,
        message: 'Member successfully assigned to leadership position.',
        team: updatedTeam,
      },
      { status: 201 }
    );
  } catch (err) {
    return NextResponse.json(
      { error: err.message || 'Failed to assign leader.' },
      { status: 400 }
    );
  }
}

export async function PATCH(request) {
  const { authorized, admin } = await requireAdminAuth();
  if (!authorized) {
    return NextResponse.json({ error: 'Unauthorized administrator access.' }, { status: 401 });
  }

  try {
    const { teamMemberId, newPositionCode } = await request.json();
    const updatedTeam = teamService.updateLeaderRole({
      teamMemberId,
      newPositionCode,
      actor: admin.username,
    });

    return NextResponse.json({
      success: true,
      message: 'Leadership position updated successfully.',
      team: updatedTeam,
    });
  } catch (err) {
    return NextResponse.json(
      { error: err.message || 'Failed to change leader position.' },
      { status: 400 }
    );
  }
}

export async function DELETE(request) {
  const { authorized, admin } = await requireAdminAuth();
  if (!authorized) {
    return NextResponse.json({ error: 'Unauthorized administrator access.' }, { status: 401 });
  }

  try {
    const { searchParams } = new URL(request.url);
    const teamMemberId = searchParams.get('teamMemberId');
    if (!teamMemberId) {
      return NextResponse.json({ error: 'teamMemberId is required.' }, { status: 400 });
    }

    const updatedTeam = teamService.removeLeader({
      teamMemberId,
      actor: admin.username,
    });

    return NextResponse.json({
      success: true,
      message: 'Member removed from leadership position.',
      team: updatedTeam,
    });
  } catch (err) {
    return NextResponse.json(
      { error: err.message || 'Failed to remove leader.' },
      { status: 400 }
    );
  }
}
