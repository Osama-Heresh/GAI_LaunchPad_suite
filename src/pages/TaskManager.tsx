import React, { useState, useEffect } from 'react';
import { DragDropContext, Droppable, Draggable, DropResult } from 'react-beautiful-dnd';
import { Plus, Search, Filter, Calendar, User, MessageSquare, Paperclip, MoreHorizontal, Star, CreditCard as Edit, Check, X } from 'lucide-react';
import { useLanguage } from '../contexts/LanguageContext';

interface Task {
  id: string;
  title: string;
  description: string;
  assignee: string;
  assigneeId: string;
  avatar?: string;
  dueDate: string;
  priority: 'low' | 'medium' | 'high';
  status: 'backlog' | 'todo' | 'inProgress' | 'review' | 'done';
  comments: number;
  attachments: number;
  tags: string[];
  color: string;
}

interface TeamMember {
  id: string;
  name: string;
  email: string;
}

const TaskManager: React.FC = () => {
  const { t } = useLanguage();
  const [showAddTask, setShowAddTask] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [filterAssignee, setFilterAssignee] = useState('all');
  const [filterPriority, setFilterPriority] = useState('all');
  const [filterStage, setFilterStage] = useState('all');
  const [teamMembers, setTeamMembers] = useState<TeamMember[]>([]);
  const [editingTaskId, setEditingTaskId] = useState<string | null>(null);
  const [editingStatus, setEditingStatus] = useState('');
  const [newTask, setNewTask] = useState({
    title: '',
    description: '',
    assigneeId: '',
    dueDate: '',
    priority: 'medium' as const
  });

  const [tasks, setTasks] = useState<Task[]>([
    {
      id: '1',
      title: 'Create Blog Content',
      description: 'Write a blog about our design process and publish it on our website',
      assignee: 'John Doe',
      assigneeId: '1',
      avatar: 'JD',
      dueDate: '2024-01-20',
      priority: 'high',
      status: 'backlog',
      comments: 3,
      attachments: 2,
      tags: ['Content', 'Marketing'],
      color: 'border-l-green-500'
    },
    {
      id: '2',
      title: 'Send Email to John',
      description: 'Make sure to talk about what the design process and get his approval',
      assignee: 'Sarah Smith',
      assigneeId: '2',
      avatar: 'SS',
      dueDate: '2024-01-25',
      priority: 'medium',
      status: 'backlog',
      comments: 2,
      attachments: 1,
      tags: ['Email', 'Communication'],
      color: 'border-l-orange-500'
    },
    {
      id: '4',
      title: 'Send Out Proposals',
      description: 'Find a new client and send them a detailed proposal for our services',
      assignee: 'Emily Brown',
      assigneeId: '4',
      avatar: 'EB',
      dueDate: '2024-01-18',
      priority: 'high',
      status: 'todo',
      comments: 1,
      attachments: 3,
      tags: ['Sales', 'Proposal'],
      color: 'border-l-blue-500'
    },
    {
      id: '5',
      title: 'Design the Dashboard',
      description: 'Be sure to look at what\'s been done and implement the design',
      assignee: 'Alex Wilson',
      assigneeId: '5',
      avatar: 'AW',
      dueDate: '2024-01-22',
      priority: 'medium',
      status: 'todo',
      comments: 2,
      attachments: 1,
      tags: ['Design', 'UI'],
      color: 'border-l-purple-500'
    }
  ]);

  const columns = [
    { id: 'backlog', title: 'BACKLOG', color: 'text-gray-400' },
    { id: 'todo', title: 'TO DO', color: 'text-blue-400' },
    { id: 'inProgress', title: 'IN PROGRESS', color: 'text-yellow-400' },
    { id: 'review', title: 'REVIEW', color: 'text-purple-400' },
    { id: 'done', title: 'DONE', color: 'text-green-400' }
  ];

  useEffect(() => {
    loadTeamMembers();
    loadTasks();
  }, []);

  const loadTeamMembers = () => {
    setTeamMembers([
      { id: '1', name: 'John Doe', email: 'john@example.com' },
      { id: '2', name: 'Sarah Smith', email: 'sarah@example.com' },
      { id: '4', name: 'Emily Brown', email: 'emily@example.com' },
      { id: '5', name: 'Alex Wilson', email: 'alex@example.com' }
    ]);
  };

  const loadTasks = () => {
    try {
      const stored = localStorage.getItem('tasks');
      if (stored) {
        setTasks(JSON.parse(stored));
      }
    } catch (err) {
      console.error('Failed to load tasks:', err);
    }
  };

  const saveTasks = (newTasks: Task[]) => {
    setTasks(newTasks);
    localStorage.setItem('tasks', JSON.stringify(newTasks));
  };

  const filteredTasks = tasks.filter(task => {
    const matchesSearch = task.title.toLowerCase().includes(searchTerm.toLowerCase()) ||
                         task.description.toLowerCase().includes(searchTerm.toLowerCase());
    const matchesAssignee = filterAssignee === 'all' || task.assigneeId === filterAssignee;
    const matchesPriority = filterPriority === 'all' || task.priority === filterPriority;
    const matchesStage = filterStage === 'all' || task.status === filterStage;

    return matchesSearch && matchesAssignee && matchesPriority && matchesStage;
  });

  const getTasksByStatus = (status: string) => {
    return filteredTasks.filter(task => task.status === status);
  };

  const handleDragEnd = (result: DropResult) => {
    const { destination, source, draggableId } = result;

    if (!destination) return;

    if (destination.droppableId === source.droppableId && destination.index === source.index) {
      return;
    }

    const newTasks = tasks.map(task => {
      if (task.id === draggableId) {
        return { ...task, status: destination.droppableId as Task['status'] };
      }
      return task;
    });

    saveTasks(newTasks);
  };

  const handleAddTask = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newTask.title.trim()) return;

    const assignedMember = teamMembers.find(m => m.id === newTask.assigneeId);
    const task: Task = {
      id: Date.now().toString(),
      title: newTask.title,
      description: newTask.description,
      assignee: assignedMember?.name || 'Unassigned',
      assigneeId: newTask.assigneeId || 'unassigned',
      status: 'backlog',
      comments: 0,
      attachments: 0,
      tags: ['New'],
      color: 'border-l-blue-500',
      avatar: assignedMember?.name.split(' ').map(n => n[0]).join('') || 'UN',
      dueDate: newTask.dueDate,
      priority: newTask.priority
    };

    saveTasks([...tasks, task]);
    setNewTask({ title: '', description: '', assigneeId: '', dueDate: '', priority: 'medium' });
    setShowAddTask(false);
  };

  const handleUpdateTaskStatus = (taskId: string, newStatus: string) => {
    const newTasks = tasks.map(task => {
      if (task.id === taskId) {
        return { ...task, status: newStatus as Task['status'] };
      }
      return task;
    });
    saveTasks(newTasks);
    setEditingTaskId(null);
  };

  const getPriorityColor = (priority: string) => {
    switch (priority) {
      case 'high': return 'bg-red-500';
      case 'medium': return 'bg-yellow-500';
      case 'low': return 'bg-green-500';
      default: return 'bg-gray-500';
    }
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-900 via-slate-800 to-slate-900 p-6">
      {/* Header */}
      <div className="flex items-center justify-between mb-6">
        <div className="flex items-center space-x-3">
          <Star className="h-6 w-6 text-blue-400" />
          <h1 className="text-2xl font-bold text-white">Task Manager</h1>
        </div>

        <div className="flex items-center space-x-4">
          <div className="relative">
            <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 h-4 w-4 text-gray-400" />
            <input
              type="text"
              placeholder="Search Tasks"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="pl-10 pr-4 py-2 w-64 bg-slate-800/50 border border-slate-700 rounded-lg text-white placeholder-gray-400 focus:outline-none focus:ring-2 focus:ring-blue-500/50"
            />
          </div>
        </div>
      </div>

      {/* Filters */}
      <div className="bg-slate-800/30 backdrop-blur-sm border border-slate-700/50 rounded-xl p-4 mb-6">
        <div className="flex items-center space-x-2 mb-4">
          <Filter className="h-5 w-5 text-white/70" />
          <span className="text-white/70 font-medium">Filter Tasks</span>
        </div>

        <div className="grid grid-cols-2 md:grid-cols-3 gap-4">
          <div>
            <label className="block text-white/60 text-sm font-medium mb-2">Assigned To</label>
            <select
              value={filterAssignee}
              onChange={(e) => setFilterAssignee(e.target.value)}
              className="w-full px-3 py-2 bg-slate-700/50 border border-slate-600 rounded-lg text-white text-sm focus:outline-none focus:ring-2 focus:ring-blue-500/50"
            >
              <option value="all">All Users</option>
              {teamMembers.map(member => (
                <option key={member.id} value={member.id}>{member.name}</option>
              ))}
            </select>
          </div>

          <div>
            <label className="block text-white/60 text-sm font-medium mb-2">Priority</label>
            <select
              value={filterPriority}
              onChange={(e) => setFilterPriority(e.target.value)}
              className="w-full px-3 py-2 bg-slate-700/50 border border-slate-600 rounded-lg text-white text-sm focus:outline-none focus:ring-2 focus:ring-blue-500/50"
            >
              <option value="all">All Priorities</option>
              <option value="high">High</option>
              <option value="medium">Medium</option>
              <option value="low">Low</option>
            </select>
          </div>

          <div>
            <label className="block text-white/60 text-sm font-medium mb-2">Stage</label>
            <select
              value={filterStage}
              onChange={(e) => setFilterStage(e.target.value)}
              className="w-full px-3 py-2 bg-slate-700/50 border border-slate-600 rounded-lg text-white text-sm focus:outline-none focus:ring-2 focus:ring-blue-500/50"
            >
              <option value="all">All Stages</option>
              <option value="backlog">Backlog</option>
              <option value="todo">To Do</option>
              <option value="inProgress">In Progress</option>
              <option value="review">Review</option>
              <option value="done">Done</option>
            </select>
          </div>
        </div>
      </div>

      {/* Kanban Board */}
      <DragDropContext onDragEnd={handleDragEnd}>
        <div className="flex gap-4 overflow-x-auto pb-4">
          {columns.map((column) => (
            <div
              key={column.id}
              className="flex-shrink-0 w-80"
            >
              <div className="flex items-center justify-between mb-4">
                <div className="flex items-center space-x-2">
                  <h2 className={`text-sm font-bold tracking-wider ${column.color}`}>
                    {column.title}
                  </h2>
                  <span className="text-xs text-white/40">
                    ({getTasksByStatus(column.id).length})
                  </span>
                </div>
              </div>

              <Droppable droppableId={column.id}>
                {(provided, snapshot) => (
                  <div
                    ref={provided.innerRef}
                    {...provided.droppableProps}
                    className={`rounded-lg p-3 min-h-96 transition-all ${
                      snapshot.isDraggingOver
                        ? 'bg-blue-900/30 ring-2 ring-blue-500/50'
                        : 'bg-slate-800/40'
                    }`}
                  >
                    {getTasksByStatus(column.id).map((task, index) => (
                      <Draggable
                        key={task.id}
                        draggableId={task.id}
                        index={index}
                      >
                        {(provided, snapshot) => (
                          <div
                            ref={provided.innerRef}
                            {...provided.draggableProps}
                            {...provided.dragHandleProps}
                            className={`mb-3 bg-slate-800/80 backdrop-blur border border-slate-700/60 rounded-lg p-4 cursor-move transition-all ${
                              snapshot.isDragging
                                ? 'shadow-2xl ring-2 ring-blue-400 opacity-100'
                                : 'hover:shadow-lg hover:bg-slate-800'
                            } ${task.color} border-l-4`}
                          >
                            <div className="flex items-start justify-between mb-2">
                              <h3 className="text-white font-semibold text-sm leading-tight flex-1 pr-2">
                                {task.title}
                              </h3>
                              <button className="text-gray-400 hover:text-white transition-colors flex-shrink-0">
                                <MoreHorizontal className="h-4 w-4" />
                              </button>
                            </div>

                            <p className="text-gray-400 text-xs mb-3 leading-relaxed">
                              {task.description}
                            </p>

                            <div className="flex flex-wrap gap-1 mb-3">
                              {task.tags.map((tag, idx) => (
                                <span
                                  key={idx}
                                  className="px-2 py-1 bg-slate-700/60 text-gray-300 text-xs rounded"
                                >
                                  {tag}
                                </span>
                              ))}
                            </div>

                            {editingTaskId === task.id ? (
                              <div className="flex gap-2 mb-3">
                                <select
                                  value={editingStatus}
                                  onChange={(e) => setEditingStatus(e.target.value)}
                                  className="flex-1 px-2 py-1 bg-slate-700 border border-slate-600 rounded text-white text-xs focus:outline-none focus:ring-2 focus:ring-blue-500/50"
                                >
                                  <option value="">Change status...</option>
                                  {columns.map(col => (
                                    <option key={col.id} value={col.id}>{col.title}</option>
                                  ))}
                                </select>
                                <button
                                  onClick={() => {
                                    if (editingStatus) {
                                      handleUpdateTaskStatus(task.id, editingStatus);
                                    }
                                  }}
                                  className="px-2 py-1 bg-green-600 hover:bg-green-700 text-white rounded text-xs transition-colors"
                                >
                                  <Check className="h-4 w-4" />
                                </button>
                                <button
                                  onClick={() => setEditingTaskId(null)}
                                  className="px-2 py-1 bg-red-600 hover:bg-red-700 text-white rounded text-xs transition-colors"
                                >
                                  <X className="h-4 w-4" />
                                </button>
                              </div>
                            ) : null}

                            <div className="flex items-center justify-between pt-3 border-t border-slate-700/50">
                              <div className="flex items-center space-x-2 text-xs">
                                <MessageSquare className="h-3 w-3 text-gray-400" />
                                <span className="text-gray-400">{task.comments}</span>
                              </div>

                              <div className="flex items-center space-x-2">
                                {!editingTaskId && (
                                  <button
                                    onClick={() => {
                                      setEditingTaskId(task.id);
                                      setEditingStatus(task.status);
                                    }}
                                    className="text-gray-400 hover:text-blue-400 transition-colors"
                                    title="Update task status"
                                  >
                                    <Edit className="h-4 w-4" />
                                  </button>
                                )}

                                <div className={`w-2 h-2 rounded-full ${getPriorityColor(task.priority)}`}></div>
                                <div className="w-6 h-6 bg-gradient-to-r from-blue-500 to-purple-600 rounded-full flex items-center justify-center text-white text-xs font-bold">
                                  {task.avatar}
                                </div>
                              </div>
                            </div>
                          </div>
                        )}
                      </Draggable>
                    ))}
                    {provided.placeholder}
                  </div>
                )}
              </Droppable>
            </div>
          ))}
        </div>
      </DragDropContext>

      {/* Add Task Modal */}
      {showAddTask && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center z-50">
          <div className="bg-slate-800 border border-slate-700 rounded-xl p-6 w-full max-w-md">
            <h2 className="text-lg font-semibold text-white mb-4">Create New Task</h2>

            <form onSubmit={handleAddTask} className="space-y-4">
              <div>
                <label className="block text-sm font-medium text-white/80 mb-1">Title</label>
                <input
                  type="text"
                  value={newTask.title}
                  onChange={(e) => setNewTask(prev => ({ ...prev, title: e.target.value }))}
                  placeholder="Task title"
                  className="w-full px-3 py-2 bg-slate-700/50 border border-slate-600 rounded-lg text-white placeholder-gray-400 focus:outline-none focus:ring-2 focus:ring-blue-500/50"
                />
              </div>

              <div>
                <label className="block text-sm font-medium text-white/80 mb-1">Description</label>
                <textarea
                  value={newTask.description}
                  onChange={(e) => setNewTask(prev => ({ ...prev, description: e.target.value }))}
                  placeholder="Task description"
                  rows={3}
                  className="w-full px-3 py-2 bg-slate-700/50 border border-slate-600 rounded-lg text-white placeholder-gray-400 focus:outline-none focus:ring-2 focus:ring-blue-500/50 resize-none"
                />
              </div>

              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block text-sm font-medium text-white/80 mb-1">Assign To</label>
                  <select
                    value={newTask.assigneeId}
                    onChange={(e) => setNewTask(prev => ({ ...prev, assigneeId: e.target.value }))}
                    className="w-full px-3 py-2 bg-slate-700/50 border border-slate-600 rounded-lg text-white focus:outline-none focus:ring-2 focus:ring-blue-500/50"
                  >
                    <option value="">Select Member</option>
                    {teamMembers.map(member => (
                      <option key={member.id} value={member.id}>{member.name}</option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-sm font-medium text-white/80 mb-1">Due Date</label>
                  <input
                    type="date"
                    value={newTask.dueDate}
                    onChange={(e) => setNewTask(prev => ({ ...prev, dueDate: e.target.value }))}
                    className="w-full px-3 py-2 bg-slate-700/50 border border-slate-600 rounded-lg text-white focus:outline-none focus:ring-2 focus:ring-blue-500/50"
                  />
                </div>
              </div>

              <div>
                <label className="block text-sm font-medium text-white/80 mb-1">Priority</label>
                <select
                  value={newTask.priority}
                  onChange={(e) => setNewTask(prev => ({ ...prev, priority: e.target.value as Task['priority'] }))}
                  className="w-full px-3 py-2 bg-slate-700/50 border border-slate-600 rounded-lg text-white focus:outline-none focus:ring-2 focus:ring-blue-500/50"
                >
                  <option value="low">Low</option>
                  <option value="medium">Medium</option>
                  <option value="high">High</option>
                </select>
              </div>

              <div className="flex gap-3 mt-6">
                <button
                  type="submit"
                  className="flex-1 px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-lg transition-colors font-medium"
                >
                  Create Task
                </button>
                <button
                  type="button"
                  onClick={() => setShowAddTask(false)}
                  className="px-4 py-2 bg-slate-700 hover:bg-slate-600 text-white rounded-lg transition-colors"
                >
                  Cancel
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Floating Add Button */}
      <button
        onClick={() => setShowAddTask(true)}
        className="fixed bottom-8 right-8 w-14 h-14 bg-gradient-to-r from-blue-600 to-purple-600 rounded-full flex items-center justify-center shadow-lg hover:shadow-xl transition-all hover:scale-110"
      >
        <Plus className="h-6 w-6 text-white" />
      </button>
    </div>
  );
};

export default TaskManager;
