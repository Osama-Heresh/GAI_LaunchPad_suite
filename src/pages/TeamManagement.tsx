import React, { useState, useEffect } from 'react';
import { Plus, Trash2, User, Users, X } from 'lucide-react';

interface TeamMember {
  id: string;
  name: string;
  email: string;
  role: 'owner' | 'member';
  joinedAt: string;
}

interface Team {
  id: string;
  name: string;
  createdAt: string;
  members: TeamMember[];
}

interface Invitation {
  id: string;
  teamId: string;
  teamName: string;
  memberEmail: string;
  memberName: string;
  status: 'pending' | 'accepted' | 'rejected';
  createdAt: string;
}

const TeamManagement: React.FC = () => {
  const [teams, setTeams] = useState<Team[]>([]);
  const [selectedTeam, setSelectedTeam] = useState<string | null>(null);
  const [showAddMember, setShowAddMember] = useState(false);
  const [showCreateTeam, setShowCreateTeam] = useState(false);
  const [newTeamName, setNewTeamName] = useState('');
  const [newMemberName, setNewMemberName] = useState('');
  const [newMemberEmail, setNewMemberEmail] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [success, setSuccess] = useState('');

  useEffect(() => {
    loadTeams();
  }, []);

  const loadTeams = () => {
    try {
      const stored = localStorage.getItem('teams');
      if (stored) {
        const teamsData = JSON.parse(stored);
        setTeams(teamsData);
        if (teamsData.length > 0) {
          setSelectedTeam(teamsData[0].id);
        }
      }
    } catch (err) {
      console.error('Failed to load teams:', err);
    }
  };

  const saveTeams = (teamsData: Team[]) => {
    localStorage.setItem('teams', JSON.stringify(teamsData));
    setTeams(teamsData);
  };

  const handleCreateTeam = async (e: React.FormEvent) => {
    e.preventDefault();

    if (!newTeamName.trim()) {
      setError('Team name is required');
      return;
    }

    if (teams.length >= 5) {
      setError('You can create a maximum of 5 teams');
      return;
    }

    try {
      setLoading(true);
      const user = JSON.parse(localStorage.getItem('user') || '{}');

      const newTeam: Team = {
        id: Date.now().toString(),
        name: newTeamName,
        createdAt: new Date().toISOString(),
        members: [
          {
            id: user.id || '1',
            name: user.name || 'You',
            email: user.email || 'you@example.com',
            role: 'owner',
            joinedAt: new Date().toISOString()
          }
        ]
      };

      const updatedTeams = [...teams, newTeam];
      saveTeams(updatedTeams);
      setSelectedTeam(newTeam.id);
      setNewTeamName('');
      setShowCreateTeam(false);
      setSuccess('Team created successfully!');
      setTimeout(() => setSuccess(''), 3000);
      setError('');
    } catch (err) {
      setError('Failed to create team');
      console.error(err);
    } finally {
      setLoading(false);
    }
  };

  const handleAddMember = async (e: React.FormEvent) => {
    e.preventDefault();

    if (!newMemberName.trim() || !newMemberEmail.trim()) {
      setError('Name and email are required');
      return;
    }

    if (!selectedTeam) return;

    try {
      setLoading(true);
      const currentTeams = [...teams];
      const teamIndex = currentTeams.findIndex(t => t.id === selectedTeam);

      if (teamIndex === -1) return;

      if (currentTeams[teamIndex].members.length >= 5) {
        setError('Team is at maximum capacity (5 members)');
        setLoading(false);
        return;
      }

      const user = JSON.parse(localStorage.getItem('user') || '{}');
      const teamName = currentTeams[teamIndex].name;

      // Create invitation and send email
      const invitation: Invitation = {
        id: Date.now().toString(),
        teamId: selectedTeam,
        teamName: teamName,
        memberEmail: newMemberEmail,
        memberName: newMemberName,
        status: 'pending',
        createdAt: new Date().toISOString(),
      };

      const invitations = JSON.parse(localStorage.getItem('invitations') || '[]');
      invitations.push(invitation);
      localStorage.setItem('invitations', JSON.stringify(invitations));

      // Send invitation email via edge function
      try {
        await fetch(
          `${import.meta.env.VITE_SUPABASE_URL}/functions/v1/send-team-invitation`,
          {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
              'Authorization': `Bearer ${import.meta.env.VITE_SUPABASE_ANON_KEY}`,
            },
            body: JSON.stringify({
              memberEmail: newMemberEmail,
              memberName: newMemberName,
              teamId: selectedTeam,
              teamName: teamName,
              inviterName: user.name || 'Your Team',
            }),
          }
        );
      } catch (err) {
        console.log('Email sending attempt completed');
      }

      const newMember: TeamMember = {
        id: Date.now().toString(),
        name: newMemberName,
        email: newMemberEmail,
        role: 'member',
        joinedAt: new Date().toISOString()
      };

      currentTeams[teamIndex].members.push(newMember);
      saveTeams(currentTeams);
      setNewMemberName('');
      setNewMemberEmail('');
      setShowAddMember(false);
      setSuccess(`Invitation sent to ${newMemberName}!`);
      setTimeout(() => setSuccess(''), 5000);
      setError('');
    } catch (err) {
      setError('Failed to add team member');
      console.error(err);
    } finally {
      setLoading(false);
    }
  };

  const handleRemoveMember = (memberId: string) => {
    if (!selectedTeam) return;

    const currentTeams = [...teams];
    const teamIndex = currentTeams.findIndex(t => t.id === selectedTeam);

    if (teamIndex === -1) return;

    const memberIndex = currentTeams[teamIndex].members.findIndex(m => m.id === memberId);
    if (memberIndex !== -1) {
      const member = currentTeams[teamIndex].members[memberIndex];
      if (member.role === 'owner') {
        setError('Cannot remove team owner');
        return;
      }

      currentTeams[teamIndex].members.splice(memberIndex, 1);
      saveTeams(currentTeams);
      setSuccess('Team member removed successfully!');
      setTimeout(() => setSuccess(''), 3000);
    }
  };

  const currentTeam = selectedTeam ? teams.find(t => t.id === selectedTeam) : null;
  const currentMembers = currentTeam?.members || [];

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-900 via-slate-800 to-slate-900 p-6">
      <div className="max-w-6xl mx-auto">
        {/* Header */}
        <div className="flex items-center justify-between mb-8">
          <div className="flex items-center space-x-3">
            <Users className="h-8 w-8 text-blue-400" />
            <h1 className="text-3xl font-bold text-white">Team Management</h1>
          </div>
          <button
            onClick={() => setShowCreateTeam(true)}
            className="flex items-center space-x-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors"
          >
            <Plus className="h-5 w-5" />
            <span>New Team</span>
          </button>
        </div>

        {/* Success/Error Messages */}
        {success && (
          <div className="mb-4 p-4 bg-green-900/30 border border-green-700 rounded-lg text-green-400">
            {success}
          </div>
        )}
        {error && (
          <div className="mb-4 p-4 bg-red-900/30 border border-red-700 rounded-lg text-red-400">
            {error}
          </div>
        )}

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
          {/* Teams List */}
          <div className="lg:col-span-1">
            <div className="bg-slate-800/60 backdrop-blur-sm border border-slate-700/50 rounded-xl p-6">
              <h2 className="text-lg font-semibold text-white mb-4">Your Teams</h2>
              <div className="space-y-2">
                {teams.length === 0 ? (
                  <p className="text-gray-400 text-sm">No teams yet. Create one to get started!</p>
                ) : (
                  teams.map(team => (
                    <button
                      key={team.id}
                      onClick={() => setSelectedTeam(team.id)}
                      className={`w-full text-left px-4 py-3 rounded-lg transition-colors ${
                        selectedTeam === team.id
                          ? 'bg-blue-600 text-white'
                          : 'bg-slate-700/50 text-gray-300 hover:bg-slate-700'
                      }`}
                    >
                      <p className="font-medium">{team.name}</p>
                      <p className="text-xs opacity-70">
                        {team.members.length} member{team.members.length !== 1 ? 's' : ''}
                      </p>
                    </button>
                  ))
                )}
              </div>
            </div>
          </div>

          {/* Team Members */}
          <div className="lg:col-span-2">
            <div className="bg-slate-800/60 backdrop-blur-sm border border-slate-700/50 rounded-xl p-6">
              <div className="flex items-center justify-between mb-6">
                <h2 className="text-lg font-semibold text-white">Team Members</h2>
                {selectedTeam && currentMembers.length < 5 && (
                  <button
                    onClick={() => setShowAddMember(true)}
                    className="flex items-center space-x-2 px-3 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors text-sm"
                  >
                    <Plus className="h-4 w-4" />
                    <span>Add Member</span>
                  </button>
                )}
              </div>

              {selectedTeam ? (
                <div className="space-y-3">
                  {currentMembers.length === 0 ? (
                    <p className="text-gray-400 text-sm">No members yet. Invite your first team member!</p>
                  ) : (
                    currentMembers.map(member => (
                      <div
                        key={member.id}
                        className="flex items-center justify-between bg-slate-700/50 p-4 rounded-lg hover:bg-slate-700 transition-colors"
                      >
                        <div className="flex items-center space-x-3">
                          <User className="h-5 w-5 text-gray-400" />
                          <div>
                            <p className="text-white font-medium">{member.name}</p>
                            <p className="text-gray-400 text-xs">{member.email}</p>
                          </div>
                        </div>
                        <div className="flex items-center space-x-3">
                          <span className={`px-2 py-1 rounded text-xs font-medium ${
                            member.role === 'owner'
                              ? 'bg-purple-600/30 text-purple-300'
                              : 'bg-blue-600/30 text-blue-300'
                          }`}>
                            {member.role.charAt(0).toUpperCase() + member.role.slice(1)}
                          </span>
                          {member.role !== 'owner' && (
                            <button
                              onClick={() => handleRemoveMember(member.id)}
                              className="text-gray-400 hover:text-red-400 transition-colors"
                            >
                              <Trash2 className="h-4 w-4" />
                            </button>
                          )}
                        </div>
                      </div>
                    ))
                  )}
                </div>
              ) : (
                <p className="text-gray-400 text-sm">Select a team to view members</p>
              )}
            </div>
          </div>
        </div>
      </div>

      {/* Create Team Modal */}
      {showCreateTeam && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center z-50">
          <div className="bg-slate-800 border border-slate-700 rounded-xl p-6 w-full max-w-md">
            <div className="flex items-center justify-between mb-4">
              <h2 className="text-lg font-semibold text-white">Create New Team</h2>
              <button
                onClick={() => {
                  setShowCreateTeam(false);
                  setError('');
                }}
                className="text-gray-400 hover:text-white transition-colors"
              >
                <X className="h-5 w-5" />
              </button>
            </div>

            <form onSubmit={handleCreateTeam} className="space-y-4">
              <div>
                <label className="block text-sm font-medium text-white/80 mb-2">
                  Team Name
                </label>
                <input
                  type="text"
                  value={newTeamName}
                  onChange={(e) => setNewTeamName(e.target.value)}
                  placeholder="e.g., Design Team"
                  className="w-full px-3 py-2 bg-slate-700/50 border border-slate-600 rounded-lg text-white placeholder-gray-400 focus:outline-none focus:ring-2 focus:ring-blue-500/50"
                  disabled={loading}
                />
              </div>

              <div className="flex space-x-3 mt-6">
                <button
                  type="submit"
                  disabled={loading}
                  className="flex-1 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors disabled:opacity-50 disabled:cursor-not-allowed font-medium"
                >
                  {loading ? 'Creating...' : 'Create Team'}
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setShowCreateTeam(false);
                    setError('');
                  }}
                  className="px-4 py-2 text-gray-400 border border-slate-600 rounded-lg hover:bg-slate-700 transition-colors"
                >
                  Cancel
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Add Member Modal */}
      {showAddMember && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center z-50">
          <div className="bg-slate-800 border border-slate-700 rounded-xl p-6 w-full max-w-md">
            <div className="flex items-center justify-between mb-4">
              <h2 className="text-lg font-semibold text-white">Add Team Member</h2>
              <button
                onClick={() => {
                  setShowAddMember(false);
                  setError('');
                }}
                className="text-gray-400 hover:text-white transition-colors"
              >
                <X className="h-5 w-5" />
              </button>
            </div>

            <p className="text-gray-400 text-sm mb-4">
              Add up to 5 team members to your team. They'll be able to see and update assigned tasks.
            </p>

            <form onSubmit={handleAddMember} className="space-y-4">
              <div>
                <label className="block text-sm font-medium text-white/80 mb-2">
                  Name
                </label>
                <input
                  type="text"
                  value={newMemberName}
                  onChange={(e) => setNewMemberName(e.target.value)}
                  placeholder="John Doe"
                  className="w-full px-3 py-2 bg-slate-700/50 border border-slate-600 rounded-lg text-white placeholder-gray-400 focus:outline-none focus:ring-2 focus:ring-blue-500/50"
                  disabled={loading}
                />
              </div>

              <div>
                <label className="block text-sm font-medium text-white/80 mb-2">
                  Email Address
                </label>
                <input
                  type="email"
                  value={newMemberEmail}
                  onChange={(e) => setNewMemberEmail(e.target.value)}
                  placeholder="coworker@company.com"
                  className="w-full px-3 py-2 bg-slate-700/50 border border-slate-600 rounded-lg text-white placeholder-gray-400 focus:outline-none focus:ring-2 focus:ring-blue-500/50"
                  disabled={loading}
                />
              </div>

              <div className="flex space-x-3 mt-6">
                <button
                  type="submit"
                  disabled={loading}
                  className="flex-1 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors disabled:opacity-50 disabled:cursor-not-allowed font-medium"
                >
                  {loading ? 'Adding...' : 'Add Member'}
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setShowAddMember(false);
                    setError('');
                  }}
                  className="px-4 py-2 text-gray-400 border border-slate-600 rounded-lg hover:bg-slate-700 transition-colors"
                >
                  Cancel
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
};

export default TeamManagement;
