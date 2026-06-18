/*
  # Update tasks table for team collaboration

  1. Modified Tables
    - `tasks` - Add fields for team support
      - `team_id` (uuid, foreign key to teams)
      - `assigned_to` (uuid, foreign key to auth.users)
      - `created_by` (uuid, foreign key to auth.users)

  2. Changes
    - Tasks now belong to teams
    - Tasks are assigned to specific team members
    - Track who created the task
    - Add RLS policies for team access
*/

CREATE TABLE IF NOT EXISTS tasks (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  team_id uuid NOT NULL REFERENCES teams(id) ON DELETE CASCADE,
  created_by uuid NOT NULL REFERENCES auth.users(id) ON DELETE CASCADE,
  assigned_to uuid REFERENCES auth.users(id) ON DELETE SET NULL,
  title text NOT NULL,
  description text,
  status text NOT NULL DEFAULT 'backlog' CHECK (status IN ('backlog', 'todo', 'inProgress', 'review', 'done')),
  priority text NOT NULL DEFAULT 'medium' CHECK (priority IN ('low', 'medium', 'high')),
  due_date date,
  tags text[] DEFAULT '{}',
  comments integer DEFAULT 0,
  attachments integer DEFAULT 0,
  created_at timestamptz DEFAULT now(),
  updated_at timestamptz DEFAULT now()
);

-- Enable RLS
ALTER TABLE tasks ENABLE ROW LEVEL SECURITY;

-- Tasks policies
CREATE POLICY "Team members can view their team tasks"
  ON tasks FOR SELECT
  TO authenticated
  USING (
    EXISTS (
      SELECT 1 FROM team_members
      WHERE team_members.team_id = tasks.team_id
      AND team_members.user_id = auth.uid()
    ) OR
    EXISTS (
      SELECT 1 FROM teams
      WHERE teams.id = tasks.team_id
      AND teams.user_id = auth.uid()
    )
  );

CREATE POLICY "Team members can create tasks"
  ON tasks FOR INSERT
  TO authenticated
  WITH CHECK (
    auth.uid() = created_by AND
    (
      EXISTS (
        SELECT 1 FROM team_members
        WHERE team_members.team_id = tasks.team_id
        AND team_members.user_id = auth.uid()
      ) OR
      EXISTS (
        SELECT 1 FROM teams
        WHERE teams.id = tasks.team_id
        AND teams.user_id = auth.uid()
      )
    )
  );

CREATE POLICY "Team members can update tasks"
  ON tasks FOR UPDATE
  TO authenticated
  USING (
    EXISTS (
      SELECT 1 FROM team_members
      WHERE team_members.team_id = tasks.team_id
      AND team_members.user_id = auth.uid()
    ) OR
    EXISTS (
      SELECT 1 FROM teams
      WHERE teams.id = tasks.team_id
      AND teams.user_id = auth.uid()
    )
  )
  WITH CHECK (
    EXISTS (
      SELECT 1 FROM team_members
      WHERE team_members.team_id = tasks.team_id
      AND team_members.user_id = auth.uid()
    ) OR
    EXISTS (
      SELECT 1 FROM teams
      WHERE teams.id = tasks.team_id
      AND teams.user_id = auth.uid()
    )
  );

CREATE POLICY "Task creator or team owner can delete tasks"
  ON tasks FOR DELETE
  TO authenticated
  USING (
    auth.uid() = created_by OR
    EXISTS (
      SELECT 1 FROM teams
      WHERE teams.id = tasks.team_id
      AND teams.user_id = auth.uid()
    )
  );

-- Indexes
CREATE INDEX IF NOT EXISTS idx_tasks_team_id ON tasks(team_id);
CREATE INDEX IF NOT EXISTS idx_tasks_assigned_to ON tasks(assigned_to);
CREATE INDEX IF NOT EXISTS idx_tasks_created_by ON tasks(created_by);
CREATE INDEX IF NOT EXISTS idx_tasks_status ON tasks(status);
