import { NextResponse } from 'next/server';
import { requireAdminAuth } from '@/lib/auth';
import { teamService } from '@/services/teamService';

export async function GET(request) {
  const { authorized } = await requireAdminAuth();
  if (!authorized) {
    return NextResponse.json({ error: 'Unauthorized administrator access.' }, { status: 401 });
  }

  const { searchParams } = new URL(request.url);
  const chapterId = searchParams.get('chapterId') || '';
  const teamType = searchParams.get('teamType') || '';
  const search = searchParams.get('search') || '';

  const data = teamService.getNriTeams({ chapterId, teamType, search });
  return NextResponse.json(data);
}

export async function POST(request) {
  const { authorized, admin } = await requireAdminAuth();
  if (!authorized) {
    return NextResponse.json({ error: 'Unauthorized administrator access.' }, { status: 401 });
  }

  try {
    const body = await request.json();
    const { action } = body;

    if (action === 'CREATE_CHAPTER') {
      const { name, countryCode, description } = body;
      const result = teamService.createNriChapter({
        name,
        countryCode,
        description,
        actor: admin.username,
      });
      return NextResponse.json({
        success: true,
        message: `NRI chapter "${name}" created successfully with Main, Youth, and Mahila teams.`,
        ...result,
      }, { status: 201 });
    }

    if (action === 'ASSIGN_DIRECT_LEADER') {
      const { teamId, positionCode, fullName, country, city, mobile, email } = body;
      const updatedTeam = teamService.assignDirectNriLeader({
        teamId,
        positionCode,
        fullName,
        country,
        city,
        mobile,
        email,
        actor: admin.username,
      });
      return NextResponse.json({
        success: true,
        message: `NRI Leader "${fullName}" successfully assigned to ${positionCode}.`,
        team: updatedTeam,
      }, { status: 201 });
    }

    if (action === 'ASSIGN_LEADER') {
      const { teamId, memberId, positionCode } = body;
      const updatedTeam = teamService.assignLeader({
        teamId,
        memberId,
        positionCode,
        actor: admin.username,
      });
      return NextResponse.json({
        success: true,
        message: 'Member successfully assigned to NRI leadership position.',
        team: updatedTeam,
      }, { status: 201 });
    }

    return NextResponse.json({ error: 'Invalid action specified.' }, { status: 400 });
  } catch (err) {
    return NextResponse.json(
      { error: err.message || 'Operation failed.' },
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
    const { teamId, executiveMemberLimit } = await request.json();
    const updatedTeam = teamService.setExecutiveCapacity({
      teamId,
      executiveMemberLimit,
      actor: admin.username,
    });
    return NextResponse.json({
      success: true,
      message: `Executive Member capacity updated to ${updatedTeam.executiveMemberLimit}.`,
      team: updatedTeam,
    });
  } catch (err) {
    return NextResponse.json(
      { error: err.message || 'Failed to update NRI team settings.' },
      { status: 400 }
    );
  }
}
