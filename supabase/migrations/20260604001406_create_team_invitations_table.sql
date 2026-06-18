/*
  # Create Team Invitations Table

  1. New Tables
    - `team_invitations`
      - `id` (uuid, primary key)
      - `team_id` (uuid, foreign key to teams)
      - `invited_email` (text, email being invited)
      - `invited_name` (text, name of person being invited)
      - `invited_by` (uuid, foreign key to auth.users - who sent the invitation)
      - `status` (text, 'pending', 'accepted', 'rejected')
      - `created_at` (timestamptz, when invitation was created)
      - `accepted_at` (timestamptz, when invitation was accepted)

  2. Security
    - Enable RLS on `team_invitations` table
    - Only team members and the invitee can view invitations
    - Only team owners can create invitations
*/

CREATE TABLE IF NOT EXISTS team_invitations (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  team_id uuid NOT NULL REFERENCES public.teams(id) ON DELETE CASCADE,
  invited_email text NOT NULL,
  invited_name text NOT NULL,
  invited_by uuid NOT NULL REFERENCES auth.users(id) ON DELETE CASCADE,
  status text NOT NULL DEFAULT 'pending' CHECK (status IN ('pending', 'accepted', 'rejected')),
  created_at timestamptz DEFAULT now(),
  accepted_at timestamptz,
  UNIQUE(team_id, invited_email)
);

ALTER TABLE team_invitations ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Users can view invitations sent to them"
  ON team_invitations FOR SELECT
  TO authenticated
  USING (
    invited_email = (SELECT email FROM auth.users WHERE id = auth.uid())
    OR invited_by = auth.uid()
    OR EXISTS (
      SELECT 1 FROM team_members
      WHERE team_members.team_id = team_invitations.team_id
      AND team_members.user_id = auth.uid()
    )
  );

CREATE POLICY "Team owners can create invitations"
  ON team_invitations FOR INSERT
  TO authenticated
  WITH CHECK (
    EXISTS (
      SELECT 1 FROM teams
      WHERE teams.id = team_id
      AND teams.user_id = auth.uid()
    )
  );

CREATE POLICY "Can update own invitation status"
  ON team_invitations FOR UPDATE
  TO authenticated
  USING (invited_email = (SELECT email FROM auth.users WHERE id = auth.uid()))
  WITH CHECK (invited_email = (SELECT email FROM auth.users WHERE id = auth.uid()));
