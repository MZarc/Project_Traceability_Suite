import React, { useCallback, useEffect, useMemo, useRef, useState, useDeferredValue } from 'react';
import axios from 'axios';
import { Check, ClipboardList, ListChecks, Minus, Pencil, Plus, Search, Trash2, Upload, X, Calendar, Tag, AlertCircle, Clock, Bell, User, ChevronRight, ChevronLeft, ChevronDown, ChevronsLeft, ChevronsRight, ShieldAlert, BarChart2, HelpCircle, Settings, MessageSquare, Edit3, Building2 } from 'lucide-react';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import * as XLSX from 'xlsx';
import { motion, AnimatePresence } from 'framer-motion';
import ChangeLog from './components/ChangeLog';
import DropdownMenu from './components/DropdownMenu';
import ProjectAnalyticsModal from './components/ProjectAnalyticsModal';
import CustomerModal from './components/CustomerModal';
import appLogo from './assets/icon.png';

const API_URL = 'http://localhost:5000/api';

// SPIS master palette — matches ProjectsHub.tsx phase distribution colors
const SPIS_PALETTE = [
  '#e5383b','#0466c8','#f77f00','#72ddf7','#fcf300',
  '#f7aef8','#89fc00','#38b000','#9d4edd','#07beb8',
  '#ffddd2','#99582a','#003f88','#adb5bd',
];

const PROJECT_REMARKS_MAX_LENGTH = 500;
const STAGE_NOTE_MAX_LENGTH = 180;

const MOTION_EASE: [number, number, number, number] = [0.22, 1, 0.36, 1];
const MOTION_FADE = { duration: 0.16, ease: MOTION_EASE };
const MOTION_PANEL = { type: 'tween' as const, duration: 0.2, ease: MOTION_EASE };
const MOTION_SIDEBAR = { type: 'tween' as const, duration: 0.22, ease: MOTION_EASE };

const STAGE_STRUCTURE = [
  { name: 'Procurement / ordersheet', sub: ['status', 'Start date', 'Due date', 'Done date'] },
  { name: 'Kitting', sub: ['status', 'Start date', 'Due date', 'Done date'] },
  { name: 'SMT', sub: ['status', 'Start date', 'Due date', 'Done date'] },
  { name: 'AOI', sub: ['status', 'Start date', 'Due date', 'Done date'] },
  { name: 'TH', sub: ['status', 'Start date', 'Due date', 'Done date'] },
  { name: 'QC', sub: ['status', 'Start date', 'Due date', 'Done date'] },
  { name: 'Programming', sub: ['status', 'Start date', 'Due date', 'Done date'] },
  { name: 'TESTING', sub: ['status', 'Start date', 'Due date', 'Done date'] },
  { name: 'COATING', sub: ['status', 'Start date', 'Due date', 'Done date'] },
  { name: 'FG', sub: ['status', 'Start date', 'Due date', 'Done date'] },
  { name: 'Dispatch', sub: ['status', 'Start date', 'Due date', 'Done date', 'Shipped Qty', 'Remaining Qty'] },
] as const;

const CORE_FIELDS = [
  { key: 'sr_no', label: 'SR No' },
  { key: 'customer', label: 'Customer' },
  { key: 'project_name', label: 'Project Name' },
  { key: 'qty', label: 'Quantity' },
  { key: 'priority', label: 'Priority' },
  { key: 'remarks', label: 'Remarks' },
] as const;

const HEADER_MAP: Record<string, string> = {
  'sr.no': 'sr_no',
  'sr no': 'sr_no',
  'sr_no': 'sr_no',
  customer: 'customer',
  'project name': 'project_name',
  'project_name': 'project_name',
  'job no. / project name': 'project_name',
  'job no / project name': 'project_name',
  'job no.': 'project_name',
  'job no': 'project_name',
  'job_no': 'project_name',
  qty: 'qty',
  quantity: 'qty',
  priority: 'priority',
  remarks: 'remarks',
  'data received': 'due_date',
  'data_received': 'due_date',
  'due date': 'due_date',
  'due_date': 'due_date',
  'data shared': 'completion_date',
  'data_shared': 'completion_date',
  'completion date': 'completion_date',
  'completion_date': 'completion_date',
};

const FORM_LABELS: Record<string, string> = {
  sr_no: 'Sr.no',
  customer: 'Customer',
  project_name: 'Project name',
  qty: 'Qty',
  priority: 'Priority',
  due_date: 'Data received',
  completion_date: 'Data shared',
  remarks: 'Remarks',
};

type Stages = Record<string, string | number>;

interface ProjectApi {
  id?: number;
  sr_no: string;
  customer: string;
  project_name: string;
  qty: number;
  priority: string;
  due_date: string;
  completion_date: string;
  remarks: string;
  stages: string | Stages;
}

interface Todo {
  id: number;
  task: string;
  priority: 'High' | 'Medium' | 'Low';
  category: 'Production' | 'QC' | 'Testing' | 'Dispatch' | 'Procurement' | 'Meeting' | 'Misc';
  due_date: string;
  due_time: string;
  completed: boolean;
  created_at: string;
}

type ProjectStatus = 'Active' | 'Completed' | 'Overdue';

interface Project extends Omit<ProjectApi, 'stages'> {
  stages: Stages;
  status: ProjectStatus;
}

type FormState = {
  sr_no: string;
  customer: string;
  project_name: string;
  qty: string;
  priority: string;
  due_date: string;
  completion_date: string;
  remarks: string;
  stages: Record<string, string>;
};

type AppDialogState = {
  open: boolean;
  title: string;
  message: string;
  confirmLabel: string;
  cancelLabel?: string;
  tone?: 'default' | 'danger';
  onConfirm?: (() => void | Promise<void>) | null;
};

type DueAlertSeverity = 'TODAY' | 'SOON' | 'OVERDUE';

type DueAlert = {
  key: string;
  projectId: number;
  qty: number;
  srNo: string;
  projectName: string;
  customer: string;
  stageName: string;
  deadlineLabel: string;
  deadlineDate: string;
  completionDate: string;
  daysLeft: number;
  severity: DueAlertSeverity;
};

type EditingCellState = {
  projectId: number;
  key: string;
  scope: 'project' | 'stage';
  location: 'table' | 'sidebar';
};

declare global {
  interface Window {
    electronAPI?: {
      notify?: (title: string, body: string) => void;
      updateStats?: (stats: { pending: number; completed: number; overdue: number }) => void;
      windowControl?: (action: 'minimize' | 'close' | 'maximize') => void;
      bringToFront?: () => void;
      onMaximizedStateChange?: (callback: (isMaximized: boolean) => void) => void;
    };
  }
}

const parseStages = (value: string | Stages): Stages => {
  if (typeof value !== 'string') return value || {};
  try {
    return value ? JSON.parse(value) : {};
  } catch {
    return {};
  }
};

const normalizeText = (value: unknown) => String(value ?? '').trim();
const normalizeSrNo = (value: unknown) => normalizeText(value).replace(/\.0+$/, '');
const isDateHeader = (header: string) => header.toLowerCase().includes('date');
const getEndOfLocalDay = (value = new Date()) => new Date(value.getFullYear(), value.getMonth(), value.getDate(), 23, 59, 59, 999).getTime();

const excelSerialToIso = (value: number) => {
  const parsed = XLSX.SSF.parse_date_code(value);
  if (!parsed) return String(value);
  return `${String(parsed.y).padStart(4, '0')}-${String(parsed.m).padStart(2, '0')}-${String(parsed.d).padStart(2, '0')}`;
};

const normalizeCellValue = (header: string, value: unknown) => {
  if (value === null || value === undefined || value === '') return '';
  if (header === 'Sr.no') return normalizeSrNo(value);
  if (header === 'Qty') return typeof value === 'number' ? value : Number(value) || 0;
  if (typeof value === 'number' && isDateHeader(header)) return excelSerialToIso(value);
  return normalizeText(value);
};

const formatDate = (value: unknown) => {
  const text = normalizeText(value);
  if (!text) return '---';
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    const [y, m, d] = text.split('-');
    return `${d}-${m}-${y}`;
  }
  return text;
};


const MS_PER_DAY = 24 * 60 * 60 * 1000;

const parseFlexibleDate = (value: unknown) => {
  const text = normalizeText(value);
  if (!text) return null;

  const iso = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));

  const dmy = text.match(/^(\d{2})-(\d{2})-(\d{4})$/);
  if (dmy) return new Date(Number(dmy[3]), Number(dmy[2]) - 1, Number(dmy[1]));

  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
};

const dayDiff = (left: Date, right: Date) => {
  const a = new Date(left.getFullYear(), left.getMonth(), left.getDate());
  const b = new Date(right.getFullYear(), right.getMonth(), right.getDate());
  return Math.round((b.getTime() - a.getTime()) / MS_PER_DAY);
};

const formatDueLabel = (daysLeft: number) => {
  if (daysLeft < 0) return `OVERDUE BY ${Math.abs(daysLeft)} DAY${Math.abs(daysLeft) === 1 ? '' : 'S'}`;
  if (daysLeft === 0) return 'DUE TODAY';
  return `DUE IN ${daysLeft} DAY${daysLeft === 1 ? '' : 'S'}`;
};

const dueAlertTone = (severity: DueAlertSeverity) => {
  if (severity === 'OVERDUE') return 'border-rose-200 bg-rose-50 text-rose-700';
  if (severity === 'TODAY') return 'border-amber-200 bg-amber-50 text-amber-700';
  return 'border-sky-200 bg-sky-50 text-sky-700';
};

const datePillClass = (value: unknown, completion: unknown, kind: 'due' | 'completion' | 'neutral' | 'info') => {
  if (kind === 'neutral') return 'date-pill';
  if (kind === 'info') return 'date-pill date-pill-info';
  const dateValue = parseFlexibleDate(value);
  if (!dateValue) return 'date-pill';
  const completionValue = parseFlexibleDate(completion);
  if (kind === 'completion') {
    if (!completionValue) return 'date-pill';
    return dayDiff(dateValue, completionValue) <= 0 ? 'date-pill date-pill-comp-ok' : 'date-pill date-pill-comp-late';
  }
  if (completionValue) {
    return dayDiff(dateValue, completionValue) <= 0 ? 'date-pill date-pill-due-ok' : 'date-pill date-pill-due-late';
  }
  const remaining = dayDiff(new Date(), dateValue);
  if (remaining < 0) return 'date-pill date-pill-due-late';
  if (remaining <= 3) return 'date-pill date-pill-due-warn';
  return 'date-pill';
};




const renderSheetDate = (value: unknown, completion: unknown, kind: 'due' | 'completion' | 'neutral' | 'info', customClass?: string) => {
  const raw = normalizeText(value);
  const text = formatDate(value);
  if (!raw || raw === '---') {
    return <span className="date-pill">---</span>;
  }
  const baseClass = datePillClass(value, completion, kind);
  const finalClass = customClass ? `${baseClass} ${customClass}` : baseClass;
  return <span title={text} className={finalClass}>{text}</span>;
};

const PROJECT_FIELD_KEYS = new Set(['sr_no', 'customer', 'project_name', 'qty', 'priority', 'due_date', 'completion_date', 'remarks']);
const isProjectFieldKey = (key: string) => PROJECT_FIELD_KEYS.has(key);

const isDateFieldKey = (key: string) =>
  key === 'due_date' ||
  key === 'completion_date' ||
  key.toLowerCase().endsWith('date') ||
  key.endsWith('_Date completion');

const isPriorityFieldKey = (key: string) => key === 'priority';
const isStatusFieldKey = (key: string) => key.toLowerCase().endsWith('_status');
const isStageNoteKey = (key: string) => key.toLowerCase().endsWith('_remarks');

const toDateInputValue = (value: unknown) => {
  const parsed = parseFlexibleDate(value);
  if (!parsed) return '';
  const y = parsed.getFullYear();
  const m = String(parsed.getMonth() + 1).padStart(2, '0');
  const d = String(parsed.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
};

const buildCompositeHeaders = (rows: unknown[][]) => {
  const top = rows[0] || [];
  const bottom = rows[1] || [];
  const headers: string[] = [];
  let parent = '';
  for (let i = 0; i < Math.max(top.length, bottom.length); i += 1) {
    const first = normalizeText(top[i]);
    const second = normalizeText(bottom[i]);
    if (first) parent = first;
    if (first && !second) headers.push(first);
    else if (second && parent) headers.push(`${parent}_${second}`);
    else headers.push(first || second || `Column_${i + 1}`);
  }
  return headers;
};

const suggestMappings = (headers: string[]) =>
  Object.fromEntries(headers.map((header) => [HEADER_MAP[header.toLowerCase()], header] as const).filter(([key]) => Boolean(key)));

const makeEmptyForm = (): FormState => ({
  sr_no: '',
  customer: '',
  project_name: '',
  qty: '',
  priority: 'Medium',
  due_date: '',
  completion_date: '',
  remarks: '',
  stages: {},
});

const nextSrNo = (projects: Project[]) => {
  const max = projects.reduce((highest, project) => {
    const value = Number(normalizeSrNo(project.sr_no));
    return Number.isFinite(value) ? Math.max(highest, value) : highest;
  }, 0);
  return String(max + 1);
};

const getStatus = (project: ProjectApi, stages: Stages, allDueAlerts?: DueAlert[]): Project['status'] => {
  const dispatch = normalizeText(stages['Dispatch_status'] || stages['Dispatch/FG_status'] || '').toLowerCase();
  const dispatchComplete = normalizeText(stages['Dispatch_Done date'] || stages['Dispatch/FG_Date completion'] || '');
  if (project.completion_date || dispatchComplete || dispatch.includes('done') || dispatch.includes('dispatch')) return 'Completed';
  
  // If we have pre-calculated alerts for this project, it's overdue
  if (allDueAlerts && allDueAlerts.some(a => a.projectId === project.id && a.severity === 'OVERDUE')) {
    return 'Overdue';
  }

  // Fallback for initial load if alerts aren't ready
  const dispatchDue = normalizeText(stages['Dispatch_Due date']) || normalizeText(stages['Dispatch/FG_Due date']);
  const parsedDue = parseFlexibleDate(dispatchDue);
  if (parsedDue && parsedDue.getTime() < Date.now()) return 'Overdue';

  return 'Active';
};

const priorityBadge: Record<string, string> = {
  high: 'badge badge-priority-high',
  medium: 'badge badge-priority-medium',
  low: 'badge badge-priority-low',
};

const getPriorityBadgeClass = (priority: string) => {
  const p = String(priority || '').trim().toLowerCase();
  return priorityBadge[p] || 'badge badge-priority-medium';
};

const getStageColorClass = (value: unknown) => {
  const v = String(value || '').toLowerCase().trim();
  if (!v || v === 'pending') return 'badge badge-pending';
  if (v === 'done' || v === 'completed' || v.includes('dispatch')) return 'badge badge-done';
  if (v === 'on hold' || v === 'hold' || v === 'halt') return 'badge badge-hold';
  if (v === 'n/a' || v === 'na') return 'badge badge-na';
  return 'badge badge-active';
};

const createAlertDialog = (title: string, message: string): AppDialogState => ({
  open: true,
  title,
  message,
  confirmLabel: 'Understood',
  tone: 'default',
});

const AppDialog = React.memo(({ state, onConfirm, onClose }: { state: AppDialogState, onConfirm: () => void, onClose: () => void }) => {
  if (!state.open) return null;
  return (
    <div className="fixed inset-0 z-[100] flex items-center justify-center bg-slate-950/60 p-4 backdrop-blur-sm">
      <div className="w-full max-w-md rounded-[28px] bg-white p-8 shadow-[0_35px_120px_rgba(15,23,42,0.28)]">
        <p className={`text-xs font-semibold uppercase tracking-[0.24em] ${state.tone === 'danger' ? 'text-rose-600' : 'text-teal-700'}`}>System Dialog</p>
        <h3 className="mt-2 text-2xl font-semibold tracking-tight text-slate-900">{state.title}</h3>
        <div className="mt-6 space-y-1">
          {state.message.split('\n').map((line, i) => {
            if (!line.trim()) return <div key={i} className="h-2" />;
            if (line.includes(': ')) {
              const [key, ...rest] = line.split(': ');
              const val = rest.join(': ');
              return (
                <div key={i} className="flex items-center justify-between rounded-xl bg-slate-50 px-4 py-2 border border-slate-100">
                  <span className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">{key}</span>
                  <span className="text-[12px] font-bold text-slate-800">{val}</span>
                </div>
              );
            }
            const isBriefing = line.includes('MISSION BRIEFING');
            return (
              <div key={i} className={`text-sm leading-relaxed ${isBriefing ? 'text-[11px] font-black tracking-[0.3em] text-teal-700 border-b border-teal-100 pb-1 mb-4' : 'text-slate-500 font-medium'}`}>
                {line}
              </div>
            );
          })}
        </div>
        <div className="mt-8 flex justify-end gap-3">
          {state.cancelLabel && (
            <button onClick={onClose} className="rounded-2xl border border-slate-200 px-5 py-3 text-sm font-semibold text-slate-600">
              {state.cancelLabel}
            </button>
          )}
          <button onClick={onConfirm} className={`rounded-2xl px-5 py-3 text-sm font-semibold ${state.tone === 'danger' ? 'bg-rose-600 text-white' : 'bg-slate-950 text-white'}`}>
            {state.confirmLabel}
          </button>
        </div>
      </div>
    </div>
  );
});

const ProjectRow = React.memo(({
  project,
  rowIndex = 1,
  selected,
  onSelect,
  renderCell,
  renderDate,
  openEditModal,
  handleDelete,
  onStageNoteEdit,
  editingCell,
  renderInlineEditor,
  getEditableCellValue,
  onStageCellEdit,
  onOpenAnalytics,
  onOpenCustomerDetails
}: {
  project: Project,
  rowIndex?: number,
  selected: boolean,
  onSelect: (p: Project, e: React.MouseEvent<HTMLTableRowElement>) => void,
  renderCell: any,
  renderDate: any,
  openEditModal: any,
  handleDelete: any,
  onStageNoteEdit: (project: Project, noteKey: string) => void,
  editingCell: EditingCellState | null,
  renderInlineEditor: (project: Project, key: string, value: string, isSidebar: boolean, label?: string) => React.ReactNode,
  getEditableCellValue: (project: Project, key: string) => string,
  onStageCellEdit: (project: Project, key: string, event: React.MouseEvent<HTMLDivElement>) => void,
  onOpenAnalytics: (project: Project) => void,
  onOpenCustomerDetails?: (customerName?: string) => void
}) => {
  const isRowEditing = editingCell?.location === 'table' && String(editingCell?.projectId) === String(project.id);
  const isFirstRow = rowIndex === 0;

  return (
    <tr
      onClick={(e) => onSelect(project, e)}
      className={`data-row cursor-pointer ${selected ? 'row-selected' : ''} ${isRowEditing ? 'z-[100] relative' : ''}`}
    >
      {renderCell(project, 'sr_no', 'project', <span className="text-[13px] font-bold text-slate-600">{project.sr_no}</span>, `sticky left-0 z-10 ${selected ? 'bg-[#ecfdf5]' : 'bg-white'} border-b border-r border-slate-300 px-2 py-3.5 text-center`)}
      {renderCell(
        project, 
        'project_name', 
        'project', 
        <button
          type="button"
          onClick={(e) => {
            e.stopPropagation();
            onOpenAnalytics(project);
          }}
          title="Click to view Project Analytics & Gantt Chart"
          className="group inline-flex items-center gap-1.5 rounded-lg px-2 py-0.5 font-black text-slate-900 text-[12px] leading-snug hover:bg-emerald-500/10 hover:text-emerald-700 transition-all cursor-pointer text-left"
        >
          <span className="underline decoration-emerald-500/30 underline-offset-4 group-hover:decoration-emerald-500">{project.project_name || '---'}</span>
          <BarChart2 className="h-3.5 w-3.5 text-emerald-500 opacity-60 group-hover:opacity-100 group-hover:scale-110 transition-all shrink-0" />
        </button>, 
        `sticky left-[56px] z-10 ${selected ? 'bg-[#ecfdf5]' : 'bg-white'} border-b border-r border-slate-300 px-2 py-3.5 text-center`
      )}
      {renderCell(
        project, 
        'customer', 
        'project', 
        <button
          type="button"
          onClick={(e) => {
            e.stopPropagation();
            if (onOpenCustomerDetails) onOpenCustomerDetails(project.customer);
          }}
          title={`Click to view customer details for: ${project.customer || 'N/A'}`}
          className="group inline-flex items-center gap-1 rounded-md px-1.5 py-0.5 font-bold text-slate-800 text-[11px] hover:bg-sky-500/10 hover:text-sky-700 transition-all cursor-pointer truncate max-w-full text-center"
        >
          <span className="truncate group-hover:underline underline-offset-2 decoration-sky-400">{project.customer || '---'}</span>
        </button>, 
        `sticky left-[266px] z-10 ${selected ? 'bg-[#ecfdf5]' : 'bg-white'} border-b border-r border-slate-300 px-3 py-3.5 text-center`
      )}
      {renderCell(project, 'qty', 'project', <span className="text-[13px] font-bold text-slate-700">{project.qty || 0}</span>, `sticky left-[436px] z-10 ${selected ? 'bg-[#ecfdf5]' : 'bg-white'} border-b border-r border-slate-300 px-4 py-3.5 text-center`)}

      {renderCell(project, 'priority', 'project', (
        <span className={getPriorityBadgeClass(project.priority)}>
          {project.priority}
        </span>
      ), 'border-b border-r border-slate-300 px-4 py-3.5 text-center')}

      {renderCell(project, 'remarks', 'project', (
        <span className={`text-[11px] font-black ${project.remarks ? 'text-rose-600' : 'text-slate-400'}`}>
          {project.remarks || '---'}
        </span>
      ), 'border-b border-r border-slate-300 px-4 py-3.5 text-center')}

      {renderCell(project, 'due_date', 'project', renderDate(project.due_date, '', 'neutral', 'date-pill-violet'), 'border-b border-r border-slate-300 px-4 py-3.5 text-center')}
      {renderCell(project, 'completion_date', 'project', renderDate(project.completion_date, '', 'neutral', 'date-pill-pink'), 'border-b border-r border-slate-300 px-4 py-3.5 text-center stage-separator-cell')}
      {STAGE_STRUCTURE.map((stage) => {
        const noteKey = `${stage.name}_remarks`;
        const noteValue = String(project.stages[noteKey] || '');
        const stageGridColumns = `repeat(${stage.sub.length}, minmax(0, 1fr))`;
        const isStageEditing = isRowEditing && editingCell?.key.startsWith(stage.name);

        return (
          <td
            key={stage.name}
            colSpan={stage.sub.length}
            className={`stage-separator-cell border-b border-r border-slate-200 p-0 align-middle ${isStageEditing ? 'editing-stage-cell !z-[9999] relative overflow-visible' : ''}`}
          >
            <div className="group/stage relative grid h-full" style={{ gridTemplateColumns: stageGridColumns }}>
              <div className={`stage-note-tooltip pointer-events-none absolute left-1/2 z-40 -translate-x-1/2 opacity-0 transition-all duration-150 group-hover/stage:opacity-100 group-hover/stage:pointer-events-auto ${
                isFirstRow 
                  ? 'top-0.5 stage-note-tooltip-first-row' 
                  : 'top-[-11px] group-hover/stage:-translate-y-1'
              }`}>
                {noteValue ? (
                  <div className="stage-note-chip stage-note-chip-note pointer-events-auto">
                    <span
                      title={noteValue}
                      className="max-w-[185px] truncate text-[9px] font-black tracking-[0.12em] text-white/95"
                    >
                      {noteValue}
                    </span>
                    <button
                      type="button"
                      title="Edit stage note"
                      onClick={(e) => {
                        e.preventDefault();
                        e.stopPropagation();
                        onStageNoteEdit(project, noteKey);
                      }}
                      className="stage-note-edit-btn inline-flex h-4 w-4 items-center justify-center rounded-full border border-white/10 bg-white/10 text-white transition-colors hover:bg-white/20"
                    >
                      <Edit3 className="h-2.5 w-2.5" />
                    </button>
                  </div>
                ) : (
                  <button
                    type="button"
                    title="Add stage note"
                    onClick={(e) => {
                      e.preventDefault();
                      e.stopPropagation();
                      onStageNoteEdit(project, noteKey);
                    }}
                    className="stage-note-chip stage-note-chip-add pointer-events-auto inline-flex items-center gap-1"
                  >
                    <MessageSquare className="h-2 w-2" />
                    <span>ADD NOTE</span>
                  </button>
                )}
              </div>
              {stage.sub.map((sub, idx) => {
                const isLast = idx === stage.sub.length - 1;
                const key = `${stage.name}_${sub}`;
                let value = String(project.stages[key] || '');
                if (key === 'Dispatch_Remaining Qty') {
                  const totalQty = project.qty || 0;
                  const shippedQty = parseFloat(String(project.stages['Dispatch_Shipped Qty'] || '0')) || 0;
                  value = String(Math.max(0, totalQty - shippedQty));
                }
                const completionValue = String(project.stages[`${stage.name}_Done date`] || project.stages[`${stage.name}_Date completion`] || '');
                const dueKind = (sub as string) === 'status' ? 'neutral' : ((sub as string) === 'Date completion' || (sub as string) === 'Done date' ? 'completion' : 'due');

                let children: React.ReactNode;
                if (sub === 'status') {
                  const formatStageValue = (k: string, v: string) => {
                    if (!v) return 'PENDING';
                    if (k.startsWith('Procurement') && v === 'NA') return 'NA';
                    return v.toUpperCase();
                  };
                  children = (
                    <span className={getStageColorClass(value) + ' min-w-[84px]'}>
                      {formatStageValue(key, value)}
                    </span>
                  );
                } else {
                  if (key === 'Dispatch_Remaining Qty') {
                    const valNum = parseFloat(value) || 0;
                    const totalQty = project.qty || 0;
                    let qtyClass = "qty-pill-rose";
                    if (valNum === 0) {
                      qtyClass = "qty-pill-emerald";
                    } else if (valNum < totalQty) {
                      qtyClass = "qty-pill-amber";
                    }
                    children = <span className={qtyClass}>{value}</span>;
                  } else if (key === 'Dispatch_Shipped Qty') {
                    children = <span className="qty-pill-amber">{value || '0'}</span>;
                  } else if (sub === 'Start date') {
                    children = renderDate(value, completionValue, 'neutral', 'date-pill-blue');
                  } else {
                    children = renderDate(value, completionValue, dueKind === 'completion' ? 'completion' : 'due');
                  }
                }

                const isEditing = editingCell?.location === 'table'
                  && String(editingCell?.projectId) === String(project.id)
                  && editingCell?.key === key;

                return (
                  <div
                    key={key}
                    className={`relative select-none active:bg-sky-100/30 ${isLast ? '' : 'border-r border-slate-200'} ${
                      isEditing
                        ? 'z-[9999] relative shadow-2xl bg-white rounded-xl p-1 border-2 border-emerald-500 flex items-center justify-center w-full min-w-[135px]'
                        : 'min-w-0 px-3 py-3.5 text-center'
                    }`}
                    title={key === 'Dispatch_Remaining Qty' ? 'AUTO CALCULATED' : 'DOUBLE CLICK TO EDIT CELL'}
                    onMouseDown={(e) => {
                      if (e.detail > 1) e.preventDefault();
                    }}
                    onDoubleClick={(event) => {
                      if (key === 'Dispatch_Remaining Qty') return;
                      onStageCellEdit(project, key, event);
                    }}
                  >
                    {isEditing ? renderInlineEditor(project, key, getEditableCellValue(project, key), false) : children}
                  </div>
                );
              })}
            </div>
          </td>
        );
      })}

      <td className="border-b border-slate-300 px-4 py-3.5">
        <div className="flex justify-center gap-1.5">
          <button title="EDIT FULL ROW" onClick={(e) => openEditModal(project, e)} className="inline-flex h-7 w-7 items-center justify-center rounded-md border border-slate-200 text-slate-400 hover:bg-slate-800 hover:text-white hover:border-slate-800 transition-all">
            <Pencil className="h-3.5 w-3.5" />
          </button>
          <button title="DELETE ROW" onClick={(e) => handleDelete(project.id!, e)} className="inline-flex h-7 w-7 items-center justify-center rounded-md border border-rose-200 text-rose-400 hover:bg-rose-600 hover:text-white hover:border-rose-600 transition-all">
            <Trash2 className="h-3.5 w-3.5" />
          </button>
        </div>
      </td>
    </tr>
  );
});

const InfoTip = ({ text }: { text: string }) => (
  <div className="group relative inline-block ml-2 align-middle cursor-help">
    <HelpCircle size={13} className="text-slate-400 hover:text-slate-600 transition-colors" />
    <div className="absolute bottom-full left-0 mb-3 w-64 p-3 bg-[#0f172a] text-slate-100 text-[10px] font-bold leading-relaxed rounded-2xl opacity-0 invisible group-hover:opacity-100 group-hover:visible transition-all z-[100] shadow-2xl pointer-events-none border border-slate-700/50 backdrop-blur-xl transition-opacity duration-200">
      {text}
      <div className="absolute top-full left-3 -mt-1 border-[5px] border-transparent border-t-[#0f172a]" />
    </div>
  </div>
);

export default function App() {
  const [projects, setProjects] = useState<Project[]>([]);
  const [query, setQuery] = useState('');
  const deferredQuery = useDeferredValue(query);
  const [priorityFilter, setPriorityFilter] = useState('All');
  const [statusFilter, setStatusFilter] = useState('All');
  const [phaseFilter, setPhaseFilter] = useState('All');
  const [analyticsProject, setAnalyticsProject] = useState<Project | null>(null);
  const [groupBy, setGroupBy] = useState<'None' | 'customer' | 'priority' | 'status' | 'qty' | 'Phase'>('None');
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState(100);
  const [isPageSizeDropdownOpen, setIsPageSizeDropdownOpen] = useState(false);
  const [modalVisibleCount, setModalVisibleCount] = useState(30);
  const [modalQuery, setModalQuery] = useState('');
  const [visibleLaunchAlertCount, setVisibleLaunchAlertCount] = useState(15);
  const [statusModalType, setStatusModalType] = useState<'Total' | 'Active' | 'Done' | 'Alerts' | null>(null);
  const [isStatsOpen, setIsStatsOpen] = useState(false);
  const [isCustomerModalOpen, setIsCustomerModalOpen] = useState(false);
  const [isDeveloperModalOpen, setIsDeveloperModalOpen] = useState(false);
  const [alertSnoozeHours, setAlertSnoozeHours] = useState<number>(() => {
    const raw = localStorage.getItem('traceability-alert-snooze-hours');
    const s = raw === null ? NaN : Number(raw);
    return Number.isFinite(s) ? s : 2;
  });
  const [alertSnoozeMinutes, setAlertSnoozeMinutes] = useState<number>(() => {
    const raw = localStorage.getItem('traceability-alert-snooze-minutes');
    const s = raw === null ? NaN : Number(raw);
    return Number.isFinite(s) ? s : 0;
  });
  const [selectedProject, setSelectedProject] = useState<Project | null>(null);
  const [isProjectModalOpen, setIsProjectModalOpen] = useState(false);
  const [editingId, setEditingId] = useState<number | null>(null);
  const [form, setForm] = useState<FormState>(makeEmptyForm());
  const [isImportOpen, setIsImportOpen] = useState(false);
  const [isToolsOpen, setIsToolsOpen] = useState(false);
  const [settingsTab, setSettingsTab] = useState<'General' | 'Logs'>('General');
  const [isTodoOpen, setIsTodoOpen] = useState(false);
  const [todos, setTodos] = useState<Todo[]>([]);
  const [visibleTodoCount, setVisibleTodoCount] = useState(10);

  // Reset scroll pagination limit on To-Do modal close
  useEffect(() => {
    if (!isTodoOpen) {
      setVisibleTodoCount(10);
    }
  }, [isTodoOpen]);

  const handleTodoScroll = (e: React.UIEvent<HTMLDivElement>) => {
    const { scrollTop, scrollHeight, clientHeight } = e.currentTarget;
    if (scrollHeight - scrollTop - clientHeight < 60) {
      if (visibleTodoCount < todos.length) {
        setVisibleTodoCount((prev) => Math.min(todos.length, prev + 10));
      }
    }
  };
  const [newTodoTask, setNewTodoTask] = useState('');
  const [newTodoPriority, setNewTodoPriority] = useState<Todo['priority']>('Medium');
  const [newTodoCategory, setNewTodoCategory] = useState<Todo['category']>('Production');
  const [newTodoDueDate, setNewTodoDueDate] = useState('');
  const [newTodoDueTime, setNewTodoDueTime] = useState('');
  const [editingTodoId, setEditingTodoId] = useState<number | null>(null);
  const notifiedTodos = useRef<Set<string>>(new Set());
  const [importRows, setImportRows] = useState<Record<string, string | number>[]>([]);
  const [importHeaders, setImportHeaders] = useState<string[]>([]);
  const [mappings, setMappings] = useState<Record<string, string>>({});
  const [editingCell, setEditingCell] = useState<EditingCellState | null>(null);
  const [editingCellValue, setEditingCellValue] = useState('');
  const [isCommitting, setIsCommitting] = useState(false);
  const editorRef = useRef<HTMLDivElement>(null);
  const noteEditorRef = useRef<HTMLTextAreaElement | null>(null);
  const lastNoteFocusRef = useRef('');
  const [acknowledgedAlertExpiries, setAcknowledgedAlertExpiries] = useState<Record<string, number>>(() => {
    try {
      const stored = localStorage.getItem('traceability-acknowledged-alerts');
      if (!stored) return {};
      const parsed = JSON.parse(stored) as string[] | Record<string, number>;
      if (Array.isArray(parsed)) {
        const expiry = getEndOfLocalDay();
        return Object.fromEntries(parsed.map((key) => [key, expiry]));
      }
      const now = Date.now();
      return Object.fromEntries(Object.entries(parsed).filter(([, until]) => Number(until) > now));
    } catch {
      return {};
    }
  });
  const [isDueAlertOpen, setIsDueAlertOpen] = useState(false);
  const [dialogState, setDialogState] = useState<AppDialogState>({
    open: false,
    title: '',
    message: '',
    confirmLabel: 'OK',
    cancelLabel: undefined,
    tone: 'default',
    onConfirm: null,
  });
  const alertSnoozeTimerRef = useRef<number | null>(null);
  const rowSelectionTimerRef = useRef<number | null>(null);
  const lastAlertSignatureRef = useRef('');
  const [alertSnoozes, setAlertSnoozes] = useState<Record<string, number>>(() => {
    try {
      return JSON.parse(sessionStorage.getItem('traceability-due-alert-snoozes') || '{}') as Record<string, number>;
    } catch {
      return {};
    }
  });
  const [snoozedAlertKeys, setSnoozedAlertKeys] = useState<string[]>(() => {
    try {
      return JSON.parse(sessionStorage.getItem('traceability-due-alert-session-hides') || '[]') as string[];
    } catch {
      return [];
    }
  });
  const [currentUser, setCurrentUser] = useState<string>(() => localStorage.getItem('pms-current-user') || 'ADMIN');
  const [isMaximized, setIsMaximized] = useState(false);

  // Requirement 2: Permanent Default Zoom vs Temporary Session Zoom
  const [permanentDefaultZoom, setPermanentDefaultZoom] = useState<number>(() => {
    const saved = Number(localStorage.getItem('pms-permanent-default-zoom'));
    return Number.isFinite(saved) && saved > 0 ? saved : 1.0;
  });
  const [zoomLevel, setZoomLevel] = useState<number>(permanentDefaultZoom);

  const handleZoomIn = () => setZoomLevel(prev => Math.min(1.5, Math.round((prev + 0.1) * 10) / 10));
  const handleZoomOut = () => setZoomLevel(prev => Math.max(0.7, Math.round((prev - 0.1) * 10) / 10));
  const handleZoomReset = () => setZoomLevel(permanentDefaultZoom);

  // Requirement 3: Redirect to Customer Details from Table
  const [customerModalSearch, setCustomerModalSearch] = useState('');
  const openCustomerModalFor = (customerName?: string) => {
    setCustomerModalSearch(customerName || '');
    setIsCustomerModalOpen(true);
  };

  // Requirement 3: Customer auto-complete suggestions (ONLY when user types a letter!)
  const [showCustomerSuggestions, setShowCustomerSuggestions] = useState(false);
  const customerSuggestions = useMemo(() => {
    const q = (form.customer || '').trim().toLowerCase();
    if (!q || q.length === 0) return []; // ONLY when user types a letter!

    const set = new Set<string>();
    projects.forEach(p => {
      if (p.customer && typeof p.customer === 'string') {
        const name = p.customer.trim();
        if (name && name.toLowerCase() !== 'n/a' && name.toLowerCase().includes(q)) {
          set.add(name);
        }
      }
    });
    return Array.from(set).slice(0, 8);
  }, [form.customer, projects]);

  const stageColumns = useMemo(
    () => STAGE_STRUCTURE.flatMap((stage) => stage.sub.map((sub) => `${stage.name}_${sub}`)),
    []
  );

  const loadData = async () => {
    try {
      const response = await axios.get<ProjectApi[]>(`${API_URL}/projects`);
      setProjects(
        response.data.map((project) => {
          const stages = parseStages(project.stages);
          return {
            ...project,
            sr_no: normalizeSrNo(project.sr_no),
            priority: project.priority || 'Medium',
            due_date: normalizeText(project.due_date) || normalizeText(stages['Dispatch_Due date']) || normalizeText(stages['Dispatch/FG_Due date']),
            completion_date: normalizeText(project.completion_date) || normalizeText(stages['Dispatch_Done date']) || normalizeText(stages['Dispatch/FG_Date completion']),
            stages,
            status: 'Active', // Temp status, will be updated by useEffect
          };
        })
      );
    } catch (err) {
      console.error(err);
      setDialogState(createAlertDialog('Load Failed', 'Unable to load dashboard data.'));
    }
  };

  const loadTodos = async () => {
    try {
      const response = await axios.get<any[]>(`${API_URL}/todos`);
      setTodos(response.data.map(t => ({ ...t, completed: Boolean(t.completed) })));
    } catch (err) {
      console.error('Failed to load todos:', err);
    }
  };

  useEffect(() => {
    loadData();
    loadTodos();

    // Request notification permission
    if (typeof Notification !== 'undefined' && Notification.permission === 'default') {
      Notification.requestPermission();
    }

    // Subscribe to window maximize events from Electron main process
    if (window.electronAPI?.onMaximizedStateChange) {
      window.electronAPI.onMaximizedStateChange((maximized) => {
        setIsMaximized(maximized);
      });
    }
  }, []);

  // Notification Watchdog
  useEffect(() => {
    const checkAlerts = () => {
      const now = new Date();
      // Use the user's local calendar date (toISOString() would shift by TZ offset
      // and desync from the local date stored in todo.due_date).
      const nowDateStr = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
      const nowTimeStr = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`; // HH:mm

      todos.forEach(todo => {
        if (!todo.completed && todo.due_date === nowDateStr && todo.due_time === nowTimeStr) {
          const alertKey = `${todo.id}_${todo.due_date}_${todo.due_time}`;
          if (!notifiedTodos.current.has(alertKey)) {
            // Native Windows Notification
            if (window.electronAPI?.notify) {
              window.electronAPI.notify('PMS TASK ALERT', `TASK: ${todo.task.toUpperCase()}\nCATEGORY: ${todo.category.toUpperCase()}`);
            }
            // High Fidelity Mission Alert
            setDialogState({
              open: true,
              title: 'CRITICAL MISSION ALERT',
              message: `MISSION BRIEFING\n\nTARGET: ${todo.task.toUpperCase()}\nCATEGORY: ${todo.category.toUpperCase()}\nPRIORITY: ${todo.priority.toUpperCase()}\n\nSTRIKE TIME: REACHED\n\nCommand suggests immediate fulfillment of this objective. Mark as completed or acknowledge to silence.`,
              confirmLabel: 'MARK AS COMPLETED',
              cancelLabel: 'ACKNOWLEDGE',
              tone: 'danger',
              onConfirm: async () => {
                await axios.put(`${API_URL}/todos/${todo.id}`, { ...todo, completed: true });
                await loadTodos();
              }
            });
            notifiedTodos.current.add(alertKey);
          }
        }
      });
    };

    const interval = setInterval(checkAlerts, 20000); // Check every 20s
    return () => clearInterval(interval);
  }, [todos]);

  useEffect(() => () => {
    if (alertSnoozeTimerRef.current) window.clearTimeout(alertSnoozeTimerRef.current);
    if (rowSelectionTimerRef.current) window.clearTimeout(rowSelectionTimerRef.current);
  }, []);

  useEffect(() => {
    if (selectedProject?.id == null) return;
    const refreshed = projects.find((project) => project.id === selectedProject.id);
    if (refreshed && refreshed !== selectedProject) setSelectedProject(refreshed);
  }, [projects, selectedProject]);

  useEffect(() => {
    const now = Date.now();
    const nextAcknowledged = Object.fromEntries(Object.entries(acknowledgedAlertExpiries).filter(([, until]) => until > now));
    if (Object.keys(nextAcknowledged).length !== Object.keys(acknowledgedAlertExpiries).length) {
      setAcknowledgedAlertExpiries(nextAcknowledged);
      localStorage.setItem('traceability-acknowledged-alerts', JSON.stringify(nextAcknowledged));
    }

    const nextSnoozes = Object.fromEntries(Object.entries(alertSnoozes).filter(([, until]) => (until as number) > now));
    if (Object.keys(nextSnoozes).length !== Object.keys(alertSnoozes).length) {
      setAlertSnoozes(nextSnoozes);
      sessionStorage.setItem('traceability-due-alert-snoozes', JSON.stringify(nextSnoozes));
    }
  }, [acknowledgedAlertExpiries, alertSnoozes]);

  // Handle immediate snooze re-scheduling when settings change
  useEffect(() => {
    const activeKeys = Object.keys(alertSnoozes);
    if (activeKeys.length === 0) return;

    const durationMs = Math.max(1, (alertSnoozeHours * 60) + alertSnoozeMinutes) * 60 * 1000;
    const nextUntil = Date.now() + durationMs;

    setAlertSnoozes(prev => {
      const next = { ...prev };
      activeKeys.forEach(k => { next[k] = nextUntil; });
      return next;
    });
    // We only want to trigger this when the settings actually change from the UI
    // Excluding alertSnoozes from deps to avoid loop
  }, [alertSnoozeHours, alertSnoozeMinutes]);

  useEffect(() => {
    const now = Date.now();
    const activeKeys = Object.keys(alertSnoozes).filter(key => (alertSnoozes[key] as number) > now);
    
    if (alertSnoozeTimerRef.current) {
      window.clearTimeout(alertSnoozeTimerRef.current);
      alertSnoozeTimerRef.current = null;
    }

    const hiddenKeys = snoozedAlertKeys.filter((key) => activeKeys.includes(key));
    if (hiddenKeys.length !== snoozedAlertKeys.length) {
      setSnoozedAlertKeys(hiddenKeys);
      sessionStorage.setItem('traceability-due-alert-session-hides', JSON.stringify(hiddenKeys));
    }

    if (!activeKeys.length) return;

    const nextExpiry = Math.min(...activeKeys.map(k => alertSnoozes[k] as number));
    const delay = Math.max(0, nextExpiry - now);
    alertSnoozeTimerRef.current = window.setTimeout(() => {
      const refreshed = Object.fromEntries(Object.entries(alertSnoozes).filter(([, until]) => (until as number) > Date.now()));
      setAlertSnoozes(refreshed);
      sessionStorage.setItem('traceability-due-alert-snoozes', JSON.stringify(refreshed));
      const remaining = snoozedAlertKeys.filter((key) => refreshed[key] && (refreshed[key] as number) > Date.now());
      setSnoozedAlertKeys(remaining);
      sessionStorage.setItem('traceability-due-alert-session-hides', JSON.stringify(remaining));
      setIsDueAlertOpen(true);
    }, delay);
  }, [alertSnoozes, snoozedAlertKeys]);
  
  useEffect(() => {
    if (!editingCell) return;
    const handleOutsideClick = (e: MouseEvent) => {
      if (!editingCell || isCommitting) return;
      // Native date pickers trigger document mousedown events that can clear state prematurely.
      // We let the input's own onBlur handle closing for date fields.
      if (isDateFieldKey(editingCell.key)) return;

      if (editorRef.current && !editorRef.current.contains(e.target as Node)) {
        void commitCellEdit();
      }
    };
    document.addEventListener('mousedown', handleOutsideClick);
    return () => document.removeEventListener('mousedown', handleOutsideClick);
  }, [editingCell]);

  // Click outside table & sidebar to close project details
  useEffect(() => {
    const handleOutsideClick = (e: MouseEvent) => {
      if (!selectedProject) return;

      const target = e.target as HTMLElement;
      
      const isTableClick = target.closest('.sheet-scroll') || target.closest('.table-header-row');
      const isSidebarClick = target.closest('.sidebar-stage-card') || target.closest('.w-\\[460px\\]') || target.closest('.sidebar-field-card');
      const isModalClick = target.closest('.fixed') || target.closest('.modal-container');
      const isHeaderClick = target.closest('header') || target.closest('.ctrl-btn');
      
      if (!isTableClick && !isSidebarClick && !isModalClick && !isHeaderClick) {
        setSelectedProject(null);
      }
    };
    document.addEventListener('mousedown', handleOutsideClick);
    return () => document.removeEventListener('mousedown', handleOutsideClick);
  }, [selectedProject]);

  // Click outside to close page size dropdown
  useEffect(() => {
    if (!isPageSizeDropdownOpen) return;
    const handleDropdownOutside = (e: MouseEvent) => {
      const target = e.target as HTMLElement;
      if (!target.closest('#pagesize-dropdown-container')) {
        setIsPageSizeDropdownOpen(false);
      }
    };
    document.addEventListener('click', handleDropdownOutside);
    return () => document.removeEventListener('click', handleDropdownOutside);
  }, [isPageSizeDropdownOpen]);

  const filteredProjects = useMemo(() => {
    const q = deferredQuery.toLowerCase().trim();
    return projects.filter((project) => {
      const matchesPriority = priorityFilter === 'All' || project.priority === priorityFilter;
      if (!matchesPriority) return false;

      const matchesStatus = 
        statusFilter === 'All' || 
        project.status === statusFilter ||
        (statusFilter === 'Active' && project.status === 'Overdue');
      if (!matchesStatus) return false;

      if (phaseFilter !== 'All') {
        const currentStage = STAGE_STRUCTURE.find(s => {
          const status = String(project.stages[`${s.name}_status`] || '').toLowerCase();
          return status !== 'done' && status !== 'completed' && status !== 'n/a' && status !== 'na';
        });
        const projectPhase = currentStage ? currentStage.name : 'Completed';
        if (projectPhase !== phaseFilter) return false;
      }

      if (!q) return true;

      // Short-circuit search checks for maximum performance
      if (
        (project.sr_no && String(project.sr_no).toLowerCase().includes(q)) ||
        (project.project_name && project.project_name.toLowerCase().includes(q)) ||
        (project.customer && project.customer.toLowerCase().includes(q)) ||
        (project.remarks && project.remarks.toLowerCase().includes(q))
      ) {
        return true;
      }

      for (const val of Object.values(project.stages)) {
        if (val && String(val).toLowerCase().includes(q)) return true;
      }

      return false;
    });
  }, [deferredQuery, phaseFilter, priorityFilter, projects, statusFilter]);

  // Reset page when filters change
  useEffect(() => {
    setCurrentPage(1);
  }, [query, priorityFilter, statusFilter, phaseFilter, pageSize]);

  // Reset modal search and visible count when modal type changes
  useEffect(() => {
    setModalVisibleCount(30);
    setModalQuery('');
  }, [statusModalType]);

  // Reset launch alert visible count when closed
  useEffect(() => {
    if (!isDueAlertOpen) {
      setVisibleLaunchAlertCount(15);
    }
  }, [isDueAlertOpen]);

  const paginatedProjects = useMemo(() => {
    if (pageSize === -1) return filteredProjects;
    const startIndex = (currentPage - 1) * pageSize;
    return filteredProjects.slice(startIndex, startIndex + pageSize);
  }, [filteredProjects, currentPage, pageSize]);

  const groupedProjects = useMemo(() => {
    if (groupBy === 'None') return { Records: paginatedProjects };
    
    const groups = paginatedProjects.reduce((acc, project) => {
      let key = 'Unassigned';
      
      if (groupBy === 'Phase') {
        // Find the first stage that isn't 'Done'
        const currentStage = STAGE_STRUCTURE.find(s => {
          const status = String(project.stages[`${s.name}_status`] || '').toLowerCase();
          return status !== 'done' && status !== 'completed' && status !== 'n/a' && status !== 'na';
        });
        key = currentStage ? currentStage.name : 'Completed';
      } else {
        key = String(project[groupBy as keyof Project] || 'Unassigned');
      }

      if (!acc[key]) acc[key] = [];
      acc[key].push(project);
      return acc;
    }, {} as Record<string, Project[]>);

    // Sort groups: Priority High->Low, Status Overdue->Active->Done, others alphabetically
    const sortedKeys = Object.keys(groups).sort((a, b) => {
      if (groupBy === 'priority') {
        const order = { High: 0, Medium: 1, Low: 2, Unassigned: 3 };
        return (order[a as keyof typeof order] ?? 4) - (order[b as keyof typeof order] ?? 4);
      }
      if (groupBy === 'status') {
        const order = { Overdue: 0, Active: 1, Completed: 2, Unassigned: 3 };
        return (order[a as keyof typeof order] ?? 4) - (order[b as keyof typeof order] ?? 4);
      }
      return a.localeCompare(b);
    });

    return Object.fromEntries(sortedKeys.map(k => [k, groups[k]]));
  }, [paginatedProjects, groupBy]);


  const allDueAlerts = useMemo(() => {
    const alerts: DueAlert[] = [];
    projects.forEach((project) => {
      STAGE_STRUCTURE.forEach((stage) => {
        const statusKey = `${stage.name}_status`;
        const status = String(project.stages[statusKey] || '').toLowerCase();
        if (status === 'done' || status === 'completed' || status === 'n/a' || status === 'na') return;

        const dueVal = project.stages[`${stage.name}_Due date`] || (stage.name === 'SMT' || stage.name === 'TH' ? project.stages[`${stage.name}_Plan given date`] : '') || '';
        const parsedDue = parseFlexibleDate(dueVal);
        if (!parsedDue) return;
        const dueDate = parsedDue.getTime();
        if (isNaN(dueDate)) return;

        const now = new Date();
        const todayEnd = getEndOfLocalDay(now);
        const diff = dueDate - todayEnd;
        const daysLeft = Math.ceil(diff / (1000 * 60 * 60 * 24));

        if (daysLeft <= 3) {
          alerts.push({
            key: `${project.id}-${stage.name}`,
            projectId: project.id!,
            qty: project.qty,
            srNo: project.sr_no,
            projectName: project.project_name,
            customer: project.customer,
            stageName: stage.name,
            deadlineLabel: 'DUE',
            deadlineDate: String(dueVal),
            completionDate: String(project.stages[`${stage.name}_Done date`] || project.stages[`${stage.name}_Date completion`] || ''),
            daysLeft,
            severity: daysLeft < 0 ? 'OVERDUE' : daysLeft === 0 ? 'TODAY' : 'SOON',
          });
        }
      });
    });
    return alerts.sort((a, b) => a.daysLeft - b.daysLeft);
  }, [projects]);

  const visibleDueAlerts = useMemo(() => {
    const now = Date.now();
    return allDueAlerts.filter((alert) => {
      const expiry = acknowledgedAlertExpiries[alert.key];
      if (expiry && expiry > now) return false;
      const until = alertSnoozes[alert.key] as number;
      if (until && until > now) return false;
      if (snoozedAlertKeys.includes(alert.key)) return false;
      return true;
    });
  }, [allDueAlerts, acknowledgedAlertExpiries, alertSnoozes, snoozedAlertKeys]);

  const visibleDueAlertGroups = useMemo(() => {
    const groups = new Map<string, DueAlert[]>();
    visibleDueAlerts.forEach((alert) => {
      const groupKey = `${alert.projectId}:${alert.projectName}:${alert.srNo}`;
      const existing = groups.get(groupKey) || [];
      existing.push(alert);
      groups.set(groupKey, existing);
    });

    return Array.from(groups.entries()).map(([groupKey, alerts]) => ({
      groupKey,
      alerts: alerts.sort((a, b) => a.daysLeft - b.daysLeft),
    }));
  }, [visibleDueAlerts]);

  const slicedLaunchAlertGroups = useMemo(() => {
    return visibleDueAlertGroups.slice(0, visibleLaunchAlertCount);
  }, [visibleDueAlertGroups, visibleLaunchAlertCount]);

  const dueAlertCount = allDueAlerts.length;

  const stats = useMemo(() => {
    return {
      total: projects.length,
      active: projects.filter(p => p.status !== 'Completed').length,
      completed: projects.filter(p => p.status === 'Completed').length,
      overdue: allDueAlerts.length,
    };
  }, [projects, allDueAlerts.length]);

  useEffect(() => {
    window.electronAPI?.updateStats?.({
      pending: stats.active,
      completed: stats.completed,
      overdue: stats.overdue,
    });
  }, [stats.active, stats.completed, stats.overdue]);

  // Re-derive project statuses whenever the alert set changes. Bail out (return
  // the previous array reference) when nothing actually changed so this effect
  // cannot create an update loop with the `allDueAlerts` memo below it.
  useEffect(() => {
    setProjects(prev => {
      let changed = false;
      const next = prev.map(p => {
        const newStatus = getStatus(p, p.stages, allDueAlerts);
        if (newStatus !== p.status) {
          changed = true;
          return { ...p, status: newStatus };
        }
        return p;
      });
      return changed ? next : prev;
    });
  }, [allDueAlerts]);

  const getProjectById = (id: number | string | undefined | null) => {
    if (id === undefined || id === null) return null;
    return projects.find((p) => String(p.id) === String(id)) || null;
  };

  const getEditableCellValue = (project: Project, key: string) => {
    if (isProjectFieldKey(key)) return String(project[key as keyof Project] ?? '');
    return String(project.stages[key] || '');
  };

  const cancelPendingRowSelection = () => {
    if (rowSelectionTimerRef.current) {
      window.clearTimeout(rowSelectionTimerRef.current);
      rowSelectionTimerRef.current = null;
    }
  };

  const handleRowSelect = (project: Project, event: React.MouseEvent<HTMLTableRowElement>) => {
    if (editingCell) return;
    if (event.detail > 1) return; // Prevent row selection on double click

    // Don't select row if clicking interactive elements
    const target = event.target as HTMLElement | null;
    if (target?.closest('button, input, textarea, select, [data-no-row-select="true"]')) return;

    cancelPendingRowSelection();
    rowSelectionTimerRef.current = window.setTimeout(() => {
      setSelectedProject(project);
      rowSelectionTimerRef.current = null;
    }, 400);
  };

  const startCellEdit = (project: Project, key: string, scope: 'project' | 'stage', location: 'table' | 'sidebar' = 'table') => {
    cancelPendingRowSelection();
    setEditingCell({ projectId: project.id || 0, key, scope, location });
    const currentValue = getEditableCellValue(project, key);
    setEditingCellValue(isDateFieldKey(key) ? toDateInputValue(currentValue) : currentValue);
  };

  const cancelCellEdit = () => {
    setEditingCell(null);
    setEditingCellValue('');
  };

  useEffect(() => {
    if (!editingCell || !editingCell.key.toLowerCase().includes('remarks')) {
      lastNoteFocusRef.current = '';
      return;
    }

    const focusKey = `${editingCell.location}:${editingCell.projectId}:${editingCell.key}`;
    if (lastNoteFocusRef.current === focusKey) return;
    lastNoteFocusRef.current = focusKey;

    const timer = window.setTimeout(() => {
      const el = noteEditorRef.current;
      if (!el) return;
      el.focus();
      const len = el.value.length;
      if (typeof el.setSelectionRange === 'function') {
        el.setSelectionRange(len, len);
      }
    }, 0);

    return () => window.clearTimeout(timer);
  }, [editingCell?.location, editingCell?.projectId, editingCell?.key]);

  const commitCellEdit = async (overrideValue?: string) => {
    if (!editingCell || isCommitting) return;
    const project = getProjectById(editingCell.projectId);
    if (!project) return;

    const editAtStart = { ...editingCell };
    const rawValue = (overrideValue ?? editingCellValue).trim();
    const maxLength = editingCell.key === 'remarks'
      ? PROJECT_REMARKS_MAX_LENGTH
      : isStageNoteKey(editingCell.key)
        ? STAGE_NOTE_MAX_LENGTH
        : undefined;
    const nextValue = maxLength ? rawValue.slice(0, maxLength) : rawValue;
    const payload = {
      ...project,
      qty: project.qty || 0,
      stages: { ...(project.stages as Stages) },
    } as ProjectApi & { stages: Stages } & Record<string, unknown>;

    if (editingCell.scope === 'project') {
      if (editingCell.key === 'qty') {
        payload.qty = Number(nextValue) || 0;
        const totalQty = payload.qty;
        const shippedQty = Number(payload.stages['Dispatch_Shipped Qty']) || 0;
        payload.stages['Dispatch_Remaining Qty'] = String(Math.max(0, totalQty - shippedQty));
      } else if (editingCell.key === 'Dispatch_Shipped Qty') {
        payload.stages[editingCell.key] = nextValue === '' ? '' : String(Number(nextValue) || 0);
        const totalQty = Number(payload.qty) || 0;
        const shippedQty = Number(payload.stages['Dispatch_Shipped Qty']) || 0;
        payload.stages['Dispatch_Remaining Qty'] = String(Math.max(0, totalQty - shippedQty));
      } else if (editingCell.key === 'priority') {
        payload.priority = nextValue || 'Medium';
      } else {
        payload[editingCell.key] = nextValue;
      }
    } else {
      payload.stages[editingCell.key] = nextValue;
      if (editingCell.key === 'Dispatch_Shipped Qty') {
        const totalQty = Number(payload.qty) || 0;
        const shippedQty = nextValue === '' ? 0 : Number(nextValue) || 0;
        payload.stages['Dispatch_Remaining Qty'] = String(Math.max(0, totalQty - shippedQty));
      }
    }

    try {
      setIsCommitting(true);
      await axios.put(`${API_URL}/projects/${project.id}`, payload);
      
      setEditingCell((prev) => {
        if (prev && prev.projectId === editAtStart.projectId && prev.key === editAtStart.key) {
          return null;
        }
        return prev;
      });
      
      await loadData();
    } catch (err) {
      console.error(err);
      setDialogState(createAlertDialog('Update Failed', 'Could not save the edited cell.'));
    } finally {
      setIsCommitting(false);
    }
  };

  const renderInlineEditor = (project: Project, key: string, value: string, isSidebar: boolean, label?: string) => {
    const isStatus = isStatusFieldKey(key);
    const currentEdit = editingCell;
    if (!currentEdit || String(currentEdit.projectId) !== String(project.id) || currentEdit.key !== key) {
      return null;
    }

    const isRemarks = isStageNoteKey(key) || key === 'remarks';
    const isQtyField = key === 'qty' || key === 'Dispatch_Shipped Qty';
    const inputType = isQtyField ? 'number' : isDateFieldKey(key) ? 'date' : 'text';
    const maxLength = key === 'remarks' ? PROJECT_REMARKS_MAX_LENGTH : isStageNoteKey(key) ? STAGE_NOTE_MAX_LENGTH : undefined;

    const options = isPriorityFieldKey(key)
      ? [
          { value: 'High', label: 'High' },
          { value: 'Medium', label: 'Medium' },
          { value: 'Low', label: 'Low' },
        ]
      : [
          { value: 'Pending', label: 'Pending' },
          { value: 'Running', label: 'Running' },
          { value: 'On Hold', label: 'On Hold' },
          { value: 'Done', label: 'Done' },
          { value: 'N/A', label: 'N/A' },
        ];

    return (
      <div ref={editorRef} className="w-full">
        {isPriorityFieldKey(key) || isStatus ? (
          <div 
            className={`flex w-full flex-col gap-2 ${isSidebar ? 'animate-in fade-in zoom-in-95 duration-200' : ''}`}
            onMouseDown={(e) => e.stopPropagation()}
            onClick={(e) => e.stopPropagation()}
          >
            {isSidebar && label && <p className="text-[10px] font-black uppercase tracking-[0.2em] text-teal-600 mb-1">{label}</p>}
            <DropdownMenu
              variant="light"
              initialOpen={true}
              value={editingCellValue || value || (isStatus ? 'Pending' : 'Medium')}
              onChange={(next) => {
                setEditingCellValue(next);
                if (!isSidebar) void commitCellEdit(next);
              }}
              options={options}
              className="w-full"
              buttonClassName={isSidebar 
                ? "h-9 w-full rounded-xl border border-slate-200 bg-slate-50 px-3 text-[12px] font-bold text-slate-800" 
                : "h-[28px] w-full rounded-md border border-slate-300 bg-white px-2 text-[11px] font-bold text-slate-800"}
              menuClassName={isSidebar ? "w-full" : "w-[130px] left-1/2 -translate-x-1/2 shadow-xl z-[999]"}
            />
            {isSidebar && (
              <div className="flex justify-end gap-2 pt-2 mt-1 border-t border-slate-100">
                <button onClick={() => cancelCellEdit()} className="flex h-8 items-center px-3 rounded-lg border border-slate-200 bg-white text-[10px] font-bold text-slate-400 hover:bg-slate-50">CANCEL</button>
                <button
                  onMouseDown={(e) => e.preventDefault()}
                  onClick={() => void commitCellEdit()}
                  className="flex h-8 items-center px-4 rounded-lg bg-teal-600 text-[10px] font-bold text-white hover:bg-teal-700"
                >
                  SAVE
                </button>
              </div>
            )}
          </div>
        ) : isRemarks ? (
          <div className="flex w-full flex-col gap-2" onMouseDown={(e) => e.stopPropagation()} onClick={(e) => e.stopPropagation()}>
            {isSidebar && label && <p className="text-[10px] font-black uppercase tracking-[0.2em] text-teal-600 mb-1">{label}</p>}
            <textarea
              ref={noteEditorRef}
              autoFocus
              value={editingCellValue}
              onChange={(e) => setEditingCellValue(maxLength ? e.target.value.slice(0, maxLength) : e.target.value)}
              onBlur={(e) => {
                const related = e.relatedTarget as HTMLElement;
                if (!isSidebar || !related?.closest('button')) void commitCellEdit();
              }}
              onKeyDown={(e) => {
                if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); void commitCellEdit(); }
                if (e.key === 'Escape') cancelCellEdit();
              }}
              rows={isSidebar ? 4 : 2}
              maxLength={maxLength}
              className={isSidebar ? "sidebar-textarea" : "ctrl-input !h-auto py-1"}
              placeholder="Enter remarks..."
            />
            {maxLength && (
              <div className="flex justify-end px-1">
                <span className="text-[9px] font-bold uppercase tracking-[0.18em] text-slate-400">
                  {editingCellValue.length}/{maxLength}
                </span>
              </div>
            )}
            {isSidebar && (
              <div className="flex justify-end gap-2 pt-2 mt-1 border-t border-slate-100">
                <button onClick={() => cancelCellEdit()} className="flex h-8 items-center px-3 rounded-lg border border-slate-200 bg-white text-[10px] font-bold text-slate-400 hover:bg-slate-50">CANCEL</button>
                <button onMouseDown={(e) => e.preventDefault()} onClick={() => void commitCellEdit()} className="flex h-8 items-center px-4 rounded-lg bg-teal-600 text-[10px] font-bold text-white">SAVE</button>
              </div>
            )}
          </div>
        ) : (
          <div className="flex w-full flex-col gap-2" onMouseDown={(e) => e.stopPropagation()} onClick={(e) => e.stopPropagation()}>
            {isSidebar && label && <p className="text-[10px] font-black uppercase tracking-[0.2em] text-teal-600 mb-1">{label}</p>}
            <div className="relative z-[220] w-full">
              <input
                autoFocus
                value={editingCellValue}
                onChange={(e) => {
                  setEditingCellValue(isQtyField ? e.target.value.replace(/[^\d.]/g, '') : e.target.value);
                  if (!isSidebar && inputType === 'date' && e.target.value) {
                    void commitCellEdit(e.target.value);
                  }
                }}
                onBlur={(e) => {
                  const related = e.relatedTarget as HTMLElement;
                  if (!isSidebar || !related?.closest('button')) void commitCellEdit();
                }}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') void commitCellEdit();
                  if (e.key === 'Escape') cancelCellEdit();
                }}
                type={inputType}
                onClick={(e) => (e.target as any).showPicker?.()}
                className={isSidebar ? "sidebar-input" : "ctrl-input !h-[30px] font-bold text-slate-900 border-2 border-emerald-500 shadow-md bg-white text-center w-full"}
              />
            </div>
            {isQtyField && (
              <div className="flex justify-end px-1">
                <span className="text-[9px] font-bold uppercase tracking-[0.18em] text-slate-400">
                  Numeric only
                </span>
              </div>
            )}
            {isSidebar && (
              <div className="flex justify-end gap-2 pt-2 mt-1 border-t border-slate-100">
                <button onClick={() => cancelCellEdit()} className="flex h-8 items-center px-3 rounded-lg border border-slate-200 bg-white text-[10px] font-bold text-slate-400 hover:bg-slate-50">CANCEL</button>
                <button onMouseDown={(e) => e.preventDefault()} onClick={() => void commitCellEdit()} className="flex h-8 items-center px-4 rounded-lg bg-teal-600 text-[10px] font-bold text-white">SAVE</button>
              </div>
            )}
          </div>
        )}
      </div>
    );

  };



  const renderEditableTableCell = useCallback((
    project: Project,
    key: string,
    scope: 'project' | 'stage',
    children: React.ReactNode,
    className: string,
    titleText = 'DOUBLE CLICK TO EDIT CELL'
  ) => {
    const isEditing = editingCell?.location === 'table' && String(editingCell?.projectId) === String(project.id) && editingCell?.key === key;
    return (
      <td
        className={`${className} select-none active:bg-sky-100/30 ${isEditing ? 'editing-stage-cell !z-[9999] relative overflow-visible bg-white shadow-xl' : ''}`}
        title={titleText}
        onMouseDown={(e) => {
          if (e.detail > 1) e.preventDefault(); // Prevent text selection on double click
        }}
        onDoubleClick={(event) => {
          if (key === 'Dispatch_Remaining Qty') return;
          event.preventDefault();
          event.stopPropagation();
          cancelPendingRowSelection();
          startCellEdit(project, key, scope, 'table');
        }}
      >
        {isEditing ? renderInlineEditor(project, key, getEditableCellValue(project, key), false) : children}
      </td>
    );
  }, [editingCell, startCellEdit, editingCellValue, renderInlineEditor]);

  const handleStageTableCellEdit = useCallback((
    project: Project,
    key: string,
    event: React.MouseEvent<HTMLDivElement>
  ) => {
    event.preventDefault();
    event.stopPropagation();
    cancelPendingRowSelection();
    startCellEdit(project, key, 'stage', 'table');
  }, [startCellEdit]);

  const renderSidebarEditableField = (
    project: Project,
    key: string,
    scope: 'project' | 'stage',
    value: string,
    label: string,
    display: React.ReactNode,
    className = 'rounded-2xl border border-slate-200 bg-white p-4'
  ) => {
    const isEditing = editingCell?.location === 'sidebar' && String(editingCell?.projectId) === String(project.id) && editingCell?.key === key;
    return (
      <div
        data-no-row-select="true"
        className={`${className} transition-all duration-300 ${isEditing ? 'col-span-2 border-teal-400 bg-white shadow-2xl ring-4 ring-teal-500/10 z-[60] overflow-visible' : 'hover:border-teal-400 hover:shadow-md hover:ring-2 hover:ring-teal-100 group/sidebar-field'} relative`}
        onDoubleClick={(event) => {
          if (key === 'Dispatch_Remaining Qty') return;
          event.stopPropagation();
          cancelPendingRowSelection();
          startCellEdit(project, key, scope, 'sidebar');
        }}
      >
        {!isEditing && key !== 'Dispatch_Remaining Qty' && (
          <div className="absolute top-3 right-3 opacity-0 group-hover/sidebar-field:opacity-100 transition-opacity">
            <div className="rounded-full bg-teal-50 p-1.5 text-teal-600 shadow-sm">
              <Pencil className="h-3 w-3" />
            </div>
          </div>
        )}
        {isEditing ? (
          <div className="p-1">{renderInlineEditor(project, key, value, true, label)}</div>
        ) : (
          display
        )}
      </div>
    );
  };

  useEffect(() => {
    if (selectedProject) {
      const latest = getProjectById(selectedProject.id);
      if (latest && JSON.stringify(latest) !== JSON.stringify(selectedProject)) {
        setSelectedProject(latest);
      }
    }
  }, [projects, selectedProject]);

  useEffect(() => {
    const now = Date.now();
    const activeSnoozes = Object.fromEntries(Object.entries(alertSnoozes).filter(([, until]) => until > now));
    if (Object.keys(activeSnoozes).length !== Object.keys(alertSnoozes).length) {
      setAlertSnoozes(activeSnoozes);
      sessionStorage.setItem('traceability-due-alert-snoozes', JSON.stringify(activeSnoozes));
    }

    if (!visibleDueAlerts.length) {
      setIsDueAlertOpen(false);
      return;
    }

    const signature = visibleDueAlerts.map((alert) => alert.key).join('|');
    if (signature !== lastAlertSignatureRef.current) {
      lastAlertSignatureRef.current = signature;
      setIsDueAlertOpen(true);
      const leadAlert = visibleDueAlerts[0];
      if (leadAlert) {
        window.electronAPI?.notify?.(
          `⚠️ ${leadAlert.severity} - DEADLINE ALERT`,
          `SR ${leadAlert.srNo} | ${leadAlert.projectName} | ${visibleDueAlerts.length} Action${visibleDueAlerts.length === 1 ? '' : 's'} Required (View in Dashboard)`
        );
        window.electronAPI?.bringToFront?.();
      }
    }

    setIsDueAlertOpen(true);
  }, [visibleDueAlerts, alertSnoozes]);

  // Window Focus Logic for Alert Persistence
  useEffect(() => {
    const handleFocus = () => {
      // If we have visible alerts but the modal is closed, show it again
      // (Unless it was snoozed for the session)
      if (visibleDueAlerts.length > 0 && !isDueAlertOpen) {
        setIsDueAlertOpen(true);
      }
    };

    window.addEventListener('focus', handleFocus);
    return () => window.removeEventListener('focus', handleFocus);
  }, [visibleDueAlerts, isDueAlertOpen]);

  const handleImportFile = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (loadEvent) => {
      const binary = loadEvent.target?.result;
      if (typeof binary !== 'string') return;
      const workbook = XLSX.read(binary, { type: 'binary' });
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as unknown[][];
      if (rows.length < 2) return;
      const headers = buildCompositeHeaders(rows);
      const data = rows.slice(2).map((row) => {
        const obj: Record<string, string | number> = {};
        row.forEach((value, index) => {
          const key = headers[index] || `Column_${index + 1}`;
          obj[key] = normalizeCellValue(key, value);
        });
        return obj;
      });
      setImportHeaders(headers);
      setImportRows(data);
      setMappings(suggestMappings(headers));
    };
    reader.readAsBinaryString(file);
  };

  const confirmImport = async () => {
    try {
      await axios.post(`${API_URL}/import-excel`, { projects: importRows, mappings });
      setIsImportOpen(false);
      setIsToolsOpen(false);
      setImportRows([]);
      setImportHeaders([]);
      setMappings({});
      await loadData();
    } catch (err) {
      console.error(err);
      setDialogState(createAlertDialog('Import Failed', 'Excel import failed.'));
    }
  };

  const exportReport = async () => {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Meet Mistry';
    workbook.lastModifiedBy = 'Meet Mistry';
    workbook.created = new Date();
    workbook.modified = new Date();

    const worksheet = workbook.addWorksheet('Traceability Report');

    // 1. Setup Column Headers (PMS Standard)
    const columns = [
      { header: 'SR.NO', key: 'sr_no', width: 8 },
      { header: 'PROJECT NAME', key: 'project_name', width: 30 },
      { header: 'CUSTOMER', key: 'customer', width: 25 },
      { header: 'QTY', key: 'qty', width: 10 },
      { header: 'PRIORITY', key: 'priority', width: 15 },
      { header: 'REMARKS', key: 'remarks', width: 40 },
      { header: 'DATA RECEIVED', key: 'due_date', width: 18 },
      { header: 'DATA SHARED', key: 'completion_date', width: 18 },
    ];

    // Add Stage Columns dynamically based on sub-keys count
    STAGE_STRUCTURE.forEach(stage => {
      stage.sub.forEach(sub => {
        columns.push({ header: stage.name.toUpperCase(), key: `${stage.name}_${sub}`, width: 15 });
      });
    });

    // 2. Build Header Rows
    const headerRow1 = worksheet.addRow(columns.map(c => c.header));
    const subHeaderRow: (string | null)[] = [];
    // Main field column offsets are null in subheader
    for (let i = 0; i < 8; i++) subHeaderRow.push(null);
    STAGE_STRUCTURE.forEach(stage => {
      stage.sub.forEach(sub => {
        let subLabel = sub.toUpperCase();
        if ((sub as string) === 'Date completion' || (sub as string) === 'Done date') subLabel = 'DONE DATE';
        subHeaderRow.push(subLabel);
      });
    });
    const headerRow2 = worksheet.addRow(subHeaderRow);

    // 3. Apply Merges
    // Merge main headers (SR No to Data Shared)
    for (let i = 1; i <= 8; i++) {
      worksheet.mergeCells(1, i, 2, i);
    }
    // Merge Stage headers (span stage.sub.length cols)
    let stageColStart = 9;
    STAGE_STRUCTURE.forEach(stage => {
      worksheet.mergeCells(1, stageColStart, 1, stageColStart + stage.sub.length - 1);
      stageColStart += stage.sub.length;
    });

    // 4. Style Headers
    [headerRow1, headerRow2].forEach(row => {
      row.eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };
        cell.font = { color: { argb: 'FFFFFFFF' }, bold: true, size: 10 };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.border = {
          top: { style: 'thin', color: { argb: 'FF334155' } },
          left: { style: 'thin', color: { argb: 'FF334155' } },
          bottom: { style: 'thin', color: { argb: 'FF334155' } },
          right: { style: 'thin', color: { argb: 'FF334155' } }
        };
      });
    });

    // 5. Add Data Rows
    filteredProjects.forEach(project => {
      const rowData = [
        normalizeSrNo(project.sr_no),
        String(project.project_name || '').toUpperCase(),
        String(project.customer || '').toUpperCase(),
        project.qty,
        project.priority.toUpperCase(),
        String(project.remarks || '').toUpperCase(),
        formatDate(project.due_date).toUpperCase(),
        formatDate(project.completion_date).toUpperCase(),
      ];

      STAGE_STRUCTURE.forEach(stage => {
        stage.sub.forEach(sub => {
          const fieldKey = `${stage.name}_${sub}`;
          // Fallbacks for backward compatibility
          let val = project.stages[fieldKey];
          if (val === undefined || val === null) {
            if (sub === 'Done date') val = project.stages[`${stage.name}_Date completion`];
            else if (sub === 'Due date' && (stage.name === 'SMT' || stage.name === 'TH')) val = project.stages[`${stage.name}_Plan given date`];
          }
          if (sub === 'status' && !val) val = 'PENDING';
          rowData.push(String(val || '').toUpperCase());
        });
      });

      const row = worksheet.addRow(rowData);

      // Data Cell Styling
      row.eachCell((cell, colNumber) => {
        cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: colNumber === row.cellCount };
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
        };
        cell.font = { size: 10 };

        if (colNumber === 6) {
          cell.font = { color: { argb: 'FFE11D48' }, bold: true, size: 10 };
        }

        // Conditional Color Coding for Status Columns
        if (colNumber >= 9 && subHeaderRow[colNumber - 1] === 'STATUS') {
          const status = String(cell.value).toLowerCase();
          let color = 'FFF1F5F9'; // Default slate 100
          if (status.includes('done') || status.includes('completed') || status.includes('dispatch')) color = 'FFD1FAE5'; // emerald 100
          else if (status.includes('on hold') || status.includes('hold')) color = 'FFFEE2E2'; // rose 100
          else if (status.includes('running') || status.includes('process')) color = 'FFE0F2FE'; // sky 100
          else if (status === 'pending' || !status) color = 'FFFEF3C7'; // amber 100
          else if (status === 'n/a') color = 'FFF1F5F9'; // slate 100

          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };
          cell.font = { bold: true, size: 9, color: { argb: color === 'FFD1FAE5' ? 'FF065F46' : 'FF475569' } };
        }
      });
    });

    // 6. Freeze Panes
    worksheet.views = [{ state: 'frozen', xSplit: 3, ySplit: 2 }];

    // 7. Auto-Width Adjustments
    worksheet.columns.forEach(col => {
      let maxLen = 0;
      col.eachCell?.({ includeEmpty: true }, (cell) => {
        const len = cell.value ? String(cell.value).length : 10;
        if (len > maxLen) maxLen = len;
      });
      col.width = Math.min(Math.max(maxLen + 4, 12), 50);
    });

    // 8. Generate and Save
    const buffer = await workbook.xlsx.writeBuffer();
    const now = new Date();
    const yy = String(now.getFullYear()).slice(-2);
    const mm = String(now.getMonth() + 1).padStart(2, '0');
    const dd = String(now.getDate()).padStart(2, '0');
    const hh = String(now.getHours()).padStart(2, '0');
    const min = String(now.getMinutes()).padStart(2, '0');
    const ss = String(now.getSeconds()).padStart(2, '0');

    const fileName = `PTS_Export_${yy}${mm}${dd}_${hh}${min}${ss}.xlsx`;
    saveAs(new Blob([buffer]), fileName);
  };

  const addTodo = async () => {
    if (!newTodoTask.trim()) return;

    try {
      if (editingTodoId) {
        await axios.put(`${API_URL}/todos/${editingTodoId}`, {
          task: newTodoTask,
          priority: newTodoPriority,
          category: newTodoCategory,
          due_date: newTodoDueDate,
          due_time: newTodoDueTime,
          completed: false
        });
        setEditingTodoId(null);
      } else {
        await axios.post(`${API_URL}/todos`, {
          task: newTodoTask,
          priority: newTodoPriority,
          category: newTodoCategory,
          due_date: newTodoDueDate,
          due_time: newTodoDueTime
        });
      }

      setNewTodoTask('');
      setNewTodoDueDate('');
      setNewTodoDueTime('');
      setNewTodoPriority('Medium');
      setNewTodoCategory('Production');
      await loadTodos();

      setDialogState({
        open: true,
        title: editingTodoId ? 'MISSION UPDATED' : 'MISSION CAPTURED',
        message: 'The task parameters have been successfully synchronized with the command system.',
        confirmLabel: 'Understood',
        tone: 'default'
      });
    } catch (err) {
      console.error('[TODO] Operation failed:', err);
      setDialogState({
        open: true,
        title: 'COMMAND FAILURE',
        message: 'The system was unable to synchronize the task. Please verify backend connectivity.',
        confirmLabel: 'Retry',
        tone: 'danger'
      });
    }
  };

  const startEditTodo = (todo: Todo) => {
    setEditingTodoId(todo.id);
    setNewTodoTask(todo.task);
    setNewTodoPriority(todo.priority);
    setNewTodoCategory(todo.category);
    setNewTodoDueDate(todo.due_date);
    setNewTodoDueTime(todo.due_time);
  };

  const toggleTodo = async (todo: Todo) => {
    try {
      await axios.put(`${API_URL}/todos/${todo.id}`, {
        ...todo,
        completed: !todo.completed
      });
      loadTodos();
    } catch (err) {
      console.error('Toggle todo failed:', err);
    }
  };

  const deleteTodo = async (id: number) => {
    try {
      await axios.delete(`${API_URL}/todos/${id}`);
      loadTodos();
    } catch (err) {
      console.error('Delete todo failed:', err);
    }
  };

  const handleDelete = async (id: number, event: React.MouseEvent) => {
    event.stopPropagation();
    setDialogState({
      open: true,
      title: 'Delete Project',
      message: 'Delete this project record? This action cannot be undone.',
      confirmLabel: 'Delete',
      cancelLabel: 'Cancel',
      tone: 'danger',
      onConfirm: async () => {
        await axios.delete(`${API_URL}/projects/${id}`);
        await loadData();
      },
    });
  };

  const handleClearAll = async () => {
    setDialogState({
      open: true,
      title: 'Wipe All Projects',
      message: 'This will permanently delete all imported and manually created records.',
      confirmLabel: 'Wipe All',
      cancelLabel: 'Cancel',
      tone: 'danger',
      onConfirm: async () => {
        await axios.delete(`${API_URL}/projects-clear`);
        setSelectedProject(null);
        await loadData();
      },
    });
  };

  const openEditModal = (project: Project, event: React.MouseEvent) => {
    event.stopPropagation();
    setEditingId(project.id || null);
    setForm({
      sr_no: project.sr_no,
      customer: project.customer,
      project_name: project.project_name,
      qty: String(project.qty || ''),
      priority: project.priority,
      due_date: project.due_date,
      completion_date: project.completion_date,
      remarks: project.remarks,
      stages: Object.fromEntries(stageColumns.map((key) => [key, String(project.stages[key] || '')])),
    });
    setIsProjectModalOpen(true);
  };

  const openCreateModal = () => {
    setEditingId(null);
    setForm({
      ...makeEmptyForm(),
      sr_no: nextSrNo(projects),
    });
    setIsProjectModalOpen(true);
  };

  const saveProject = async (event: React.FormEvent) => {
    event.preventDefault();
    const totalQty = Number(form.qty) || 0;
    const shippedQty = Number(form.stages['Dispatch_Shipped Qty']) || 0;
    const remainingQty = String(Math.max(0, totalQty - shippedQty));

    const payload = {
      ...form,
      qty: totalQty,
      stages: {
        ...form.stages,
        'Dispatch_Remaining Qty': remainingQty
      }
    };
    try {
      if (editingId) await axios.put(`${API_URL}/projects/${editingId}`, payload);
      else await axios.post(`${API_URL}/projects`, payload);
      setIsProjectModalOpen(false);
      setEditingId(null);
      await loadData();
    } catch (err) {
      console.error(err);
      setDialogState(createAlertDialog('Save Failed', 'Could not save this project record.'));
    }
  };

  const handleDialogConfirm = async () => {
    const action = dialogState.onConfirm;
    setDialogState((current) => ({ ...current, open: false }));
    if (action) {
      try {
        await action();
      } catch (err) {
        console.error(err);
        setDialogState(createAlertDialog('Action Failed', 'The requested action could not be completed.'));
      }
    }
  };

  const handleAcknowledgeDueAlerts = (keys: string[] = visibleDueAlerts.map((alert) => alert.key)) => {
    if (!keys.length) {
      setIsDueAlertOpen(false);
      return;
    }

    const acknowledgedUntil = getEndOfLocalDay();
    setAcknowledgedAlertExpiries((current) => {
      const next = { ...current };
      keys.forEach((key) => {
        next[key] = acknowledgedUntil;
      });
      localStorage.setItem('traceability-acknowledged-alerts', JSON.stringify(next));
      return next;
    });
    setIsDueAlertOpen(false);
    const nextSnoozes = { ...alertSnoozes };
    keys.forEach((key) => {
      delete nextSnoozes[key];
    });
    setAlertSnoozes(nextSnoozes);
    sessionStorage.setItem('traceability-due-alert-snoozes', JSON.stringify(nextSnoozes));
    if (alertSnoozeTimerRef.current) window.clearTimeout(alertSnoozeTimerRef.current);
  };

  const handleRemindDueAlerts = (keys: string[] = visibleDueAlerts.map((alert) => alert.key)) => {
    if (!keys.length) {
      setIsDueAlertOpen(false);
      return;
    }

    const durationMs = Math.max(1, (alertSnoozeHours * 60) + alertSnoozeMinutes) * 60 * 1000;
    const snoozeUntil = Date.now() + durationMs;
    const nextSnoozes = { ...alertSnoozes };
    keys.forEach((key) => {
      nextSnoozes[key] = snoozeUntil;
    });
    setAlertSnoozes(nextSnoozes);
    sessionStorage.setItem('traceability-due-alert-snoozes', JSON.stringify(nextSnoozes));
    const nextHidden = Array.from(new Set([...snoozedAlertKeys, ...keys]));
    setSnoozedAlertKeys(nextHidden);
    sessionStorage.setItem('traceability-due-alert-session-hides', JSON.stringify(nextHidden));
    setIsDueAlertOpen(false);
  };

  return (
    <div className="flex h-screen flex-col overflow-hidden bg-slate-100 text-slate-900">
      {/* ────── PREMIUM HEADER ────── */}
      <header 
        className="custom-titlebar header-chrome px-5 py-0" 
        style={{ overflow: 'visible' }}
      >
        {/* ── Top Row: Brand + Stats + Window Controls ── */}
        <div className="flex items-center gap-3 py-2 border-b border-white/5">
          <div className="flex items-center gap-3 min-w-0 flex-1">
            {/* APP ICON - Clickable for Developer Info (Requirement 7: Borderless) */}
            <button
              type="button"
              title="Developed by Meet Mistry"
              onClick={() => setIsDeveloperModalOpen(true)}
              className="no-drag group flex h-8 w-8 shrink-0 items-center justify-center p-0.5 transition-all duration-200 hover:scale-105 active:scale-95 cursor-pointer"
            >
              <img src={appLogo} alt="App Icon" className="h-full w-full object-contain transition-transform duration-200 group-hover:scale-105" />
            </button>

            <div className="flex flex-col leading-none">
              <span className="text-[7px] font-black tracking-[0.35em] text-emerald-400/70 uppercase">Enterprise Manufacturing Control</span>
              <h1 className="text-[14px] font-black tracking-tight text-white leading-none mt-0.5">Project Traceability Suite</h1>
            </div>
            <span className="brand-divider" />
            <button
              type="button"
              onClick={() => setIsDeveloperModalOpen(true)}
              className="no-drag flex flex-col justify-end items-start leading-none gap-0.5 text-left group cursor-pointer hover:opacity-80 transition-opacity"
              title="Click to view Developer Credits"
            >
              <span className="text-[7px] font-black tracking-[0.1em] text-sky-300/80 italic uppercase mb-[1px]">v1.4</span>
              <span className="text-[7px] font-black text-sky-400/80 tracking-widest uppercase italic">Developed by</span>
              <span className="text-[9px] font-black text-emerald-400 tracking-wider uppercase group-hover:text-emerald-300 group-hover:underline">MEET MISTRY</span>
            </button>
          </div>

          {/* Action Center - Header Buttons with Compact Labels */}
          <div className="flex items-center gap-1.5 shrink-0">
            {/* TO-DO Button */}
            <button
              id="btn-todo-list"
              type="button"
              title="TO-DO LIST"
              onClick={() => setIsTodoOpen(true)}
              className="no-drag flex h-7 items-center gap-1.5 px-2.5 rounded-lg border border-indigo-400/30 bg-indigo-500/10 text-indigo-300 hover:bg-indigo-500/20 hover:text-white transition-all relative cursor-pointer"
            >
              <ListChecks className="h-3.5 w-3.5" />
              <span className="text-[10px] font-black uppercase tracking-wider">To-Do</span>
              {todos.filter(t => !t.completed).length > 0 && (
                <span className="flex h-4 min-w-[16px] px-1 items-center justify-center rounded-full bg-rose-600 text-[8px] font-black text-white ring-1 ring-[#0c1322] shadow-sm ml-0.5">
                  {todos.filter(t => !t.completed).length}
                </span>
              )}
            </button>

            {/* CUSTOMER DATABASE Button */}
            <button
              type="button"
              onClick={() => {
                setCustomerModalSearch('');
                setIsCustomerModalOpen(true);
              }}
              title="CUSTOMER DATABASE"
              className="no-drag flex h-7 items-center gap-1.5 px-2.5 rounded-lg border border-sky-400/30 bg-sky-500/10 text-sky-300 hover:bg-sky-500/20 hover:text-white transition-all cursor-pointer"
            >
              <Building2 className="h-3.5 w-3.5 text-sky-400" />
              <span className="text-[10px] font-black uppercase tracking-wider">Customers</span>
            </button>

            {/* STATS Button */}
            <button
              type="button"
              onClick={() => setIsStatsOpen(true)}
              title="ANALYTICS DASHBOARD"
              className="no-drag flex h-7 items-center gap-1.5 px-2.5 rounded-lg border border-amber-400/30 bg-amber-500/10 text-amber-300 hover:bg-amber-500/20 hover:text-white transition-all cursor-pointer"
            >
              <BarChart2 className="h-3.5 w-3.5 text-amber-400" />
              <span className="text-[10px] font-black uppercase tracking-wider">Stats</span>
            </button>
            {currentUser && (
              <div className="user-chip no-drag">
                <User className="h-3 w-3" />
                <span>{currentUser}</span>
              </div>
            )}
            <div className="flex items-center gap-1.5">
              {([
                ['Total', stats.total, 'text-slate-200', () => setStatusModalType('Total')],
                ['Active', stats.active, 'text-emerald-300', () => setStatusModalType('Active')],
                ['Done', stats.completed, 'text-sky-300', () => setStatusModalType('Done')],
                ['Alerts', dueAlertCount, 'text-rose-300', () => {
                  setSnoozedAlertKeys([]); 
                  setStatusModalType('Alerts');
                }],
              ] as const).map(([label, value, tone, onClick]) => (
                <button 
                  key={label} 
                  onClick={onClick}
                  title={`View ${label} Details`}
                  className="stat-chip hover:scale-105 active:scale-95 transition-all group"
                >
                  <span className="stat-chip-label group-hover:text-white transition-colors">{label}</span>
                  <span className={`stat-chip-value ${tone}`}>{value}</span>
                </button>
              ))}
            </div>
          </div>

          {/* Window Controls */}
          <div className="no-drag flex items-center gap-1 pl-1">
            <button type="button" title="MINIMIZE" onClick={() => window.electronAPI?.windowControl?.('minimize')} className="win-btn"><Minus className="h-3.5 w-3.5" /></button>
            <button 
              type="button" 
              title={isMaximized ? "RESTORE" : "MAXIMIZE"} 
              onClick={() => window.electronAPI?.windowControl?.('maximize')} 
              className="win-btn"
            >
              {isMaximized ? (
                <svg className="h-2.5 w-2.5" viewBox="0 0 10 10">
                  <path fill="none" stroke="currentColor" strokeWidth="1.2" d="M2,4 L2,8 L6,8 L6,4 Z M4,2 L8,2 L8,6 L6,6" />
                </svg>
              ) : (
                <svg className="h-2.5 w-2.5" viewBox="0 0 10 10">
                  <path fill="none" stroke="currentColor" strokeWidth="1.2" d="M1,1 L9,1 L9,9 L1,9 Z" />
                </svg>
              )}
            </button>
            <button type="button" title="CLOSE TO TRAY" onClick={() => window.electronAPI?.windowControl?.('close')} className="win-btn win-btn-close"><X className="h-3.5 w-3.5" /></button>
          </div>
        </div>

        {/* ── Bottom Row: Toolbar ── */}
        <div className="no-drag flex items-center gap-1.5 py-2 relative z-50" style={{ overflow: 'visible' }}>

          {/* Search — fixed width */}
          <div className="relative shrink-0">
            <Search className="absolute left-2.5 top-1/2 h-3.5 w-3.5 -translate-y-1/2 text-slate-500 pointer-events-none" />
            <input
              id="input-search"
              title="SEARCH PROJECTS"
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              placeholder="Search..."
              className="search-input"
            />
          </div>

          {/* ── Filters group — flex, auto-size, equal share ── */}
          <div className="flex flex-1 items-center gap-1.5 min-w-0">
            <DropdownMenu
              variant="dark"
              value={priorityFilter}
              onChange={setPriorityFilter}
              title="FILTER BY PRIORITY"
              options={[
                { value: 'All',    label: `All Priority (${projects.length})`, shortLabel: 'Priority' },
                { value: 'High',   label: `High (${projects.filter(p => p.priority === 'High').length})` },
                { value: 'Medium', label: `Medium (${projects.filter(p => p.priority === 'Medium').length})` },
                { value: 'Low',    label: `Low (${projects.filter(p => p.priority === 'Low').length})` },
              ]}
              className="flex-1 min-w-0"
              buttonClassName="ctrl-btn h-[30px] px-2.5 w-full"
            />
            <DropdownMenu
              variant="dark"
              value={statusFilter}
              onChange={setStatusFilter}
              title="FILTER BY STATUS"
              options={[
                { value: 'All',       label: `All Status (${projects.length})`,  shortLabel: 'Status' },
                { value: 'Active',    label: `Active (${projects.filter(p => p.status === 'Active' || p.status === 'Overdue').length})` },
                { value: 'Completed', label: `Completed (${projects.filter(p => p.status === 'Completed').length})` },
                { value: 'Overdue',   label: `Overdue (${projects.filter(p => p.status === 'Overdue').length})` },
              ]}
              className="flex-1 min-w-0"
              buttonClassName="ctrl-btn h-[30px] px-2.5 w-full"
            />
            <DropdownMenu
              variant="dark"
              value={groupBy}
              onChange={(v) => setGroupBy(v as typeof groupBy)}
              title="GROUP RECORDS"
              options={[
                { value: 'None',     label: 'Group: None',     shortLabel: 'Group' },
                { value: 'customer', label: 'By Customer',      shortLabel: 'Customer' },
                { value: 'priority', label: 'By Priority',      shortLabel: 'Priority' },
                { value: 'status',   label: 'By Status',        shortLabel: 'Status' },
                { value: 'Phase',    label: 'By Current Phase', shortLabel: 'Phase' },
                { value: 'qty',      label: 'By Quantity',      shortLabel: 'Qty' },
              ]}
              className="flex-1 min-w-0"
              buttonClassName="ctrl-btn h-[30px] px-2.5 w-full"
            />
            <DropdownMenu
              variant="dark"
              value={phaseFilter}
              onChange={setPhaseFilter}
              title="FILTER BY PHASE"
              options={[
                { value: 'All', label: `All Phases (${projects.length})`, shortLabel: 'Phase' },
                ...STAGE_STRUCTURE.map((stage) => {
                  const count = projects.filter(p => {
                    const currentStage = STAGE_STRUCTURE.find(s => {
                      const status = String(p.stages[`${s.name}_status`] || '').toLowerCase();
                      return status !== 'done' && status !== 'completed' && status !== 'n/a' && status !== 'na';
                    });
                    const projectPhase = currentStage ? currentStage.name : 'Completed';
                    return projectPhase === stage.name;
                  }).length;
                  return { value: stage.name, label: `${stage.name} (${count})` };
                }),
                { 
                  value: 'Completed', 
                  label: `Completed (${projects.filter(p => !STAGE_STRUCTURE.find(s => {
                    const status = String(p.stages[`${s.name}_status`] || '').toLowerCase();
                    return status !== 'done' && status !== 'completed' && status !== 'n/a' && status !== 'na';
                  })).length})`, 
                  shortLabel: 'Done' 
                },
              ]}
              className="flex-1 min-w-0"
              buttonClassName="ctrl-btn h-[30px] px-2.5 w-full"
            />
          </div>

          {/* Zoom controls */}
          <div className="no-drag flex items-center gap-1.5 bg-slate-800/40 border border-slate-700/50 rounded-lg p-0.5 shrink-0 h-[30px] px-2.5">
            <span className="text-[7px] font-black text-slate-400 select-none mr-1 tracking-widest uppercase">ZOOM</span>
            <button
              type="button"
              onClick={handleZoomOut}
              disabled={zoomLevel <= 0.7}
              title="ZOOM OUT"
              className="win-btn h-5 w-5 border-none hover:bg-slate-700/60 disabled:opacity-30 disabled:pointer-events-none"
            >
              <Minus className="h-3 w-3" />
            </button>
            <span className="text-[10px] font-black text-slate-300 w-10 text-center select-none font-mono">
              {Math.round(zoomLevel * 100)}%
            </span>
            <button
              type="button"
              onClick={handleZoomIn}
              disabled={zoomLevel >= 1.5}
              title="ZOOM IN"
              className="win-btn h-5 w-5 border-none hover:bg-slate-700/60 disabled:opacity-30 disabled:pointer-events-none"
            >
              <Plus className="h-3 w-3" />
            </button>
            {zoomLevel !== 1.0 && (
              <button
                type="button"
                onClick={handleZoomReset}
                title="RESET ZOOM"
                className="ctrl-btn ctrl-btn-ghost h-5 px-2 border-none hover:bg-slate-700/60 text-[8px] font-extrabold uppercase ml-1"
              >
                Reset
              </button>
            )}
          </div>

          {/* Primary Actions — fixed, never shrink */}
          <button
            id="btn-new-entry"
            title="CREATE NEW ENTRY"
            onClick={openCreateModal}
            className="ctrl-btn ctrl-btn-accent shrink-0 px-4"
          >
            <Plus className="h-3.5 w-3.5" />
            New Entry
          </button>
          <button
            id="btn-settings"
            type="button"
            title="System Settings"
            onClick={() => setIsToolsOpen(true)}
            className="no-drag flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-sky-500/25 bg-sky-500/12 text-sky-300 shadow-sm backdrop-blur-md transition-all duration-200 hover:scale-105 hover:bg-sky-500/20 hover:border-sky-400/50 hover:text-white active:scale-95 cursor-pointer"
          >
            <Settings className="h-4 w-4" />
          </button>
        </div>
      </header>

      {/* ────── MAIN CONTENT (Requirement 4: Minimal Whitespace Padding) ────── */}
      <div className="min-h-0 flex-1 overflow-hidden p-1.5 sm:p-2 bg-slate-900/90">
        <div className="sheet-shell flex flex-col h-full overflow-hidden rounded-xl border border-slate-200 bg-white shadow-[0_4px_24px_rgba(15,23,42,0.06)]">
          <div className="flex flex-1 min-h-0 overflow-hidden relative">
            <div className="sheet-scroll min-w-0 flex-1 overflow-auto" style={{ zoom: zoomLevel }}>
              <table className="min-w-full border-separate border-spacing-0">
              <thead className="text-center">
                <tr className="table-header-row">
                  <th rowSpan={2} className="sticky top-0 left-0 z-30 h-[68px] min-w-[56px] w-[56px] border-b border-r border-slate-600 px-2 py-0 align-middle" style={{background:'linear-gradient(180deg,#111c30 0%,#0d1826 100%)'}}>SR No</th>
                  <th rowSpan={2} className="sticky top-0 left-[56px] z-30 h-[68px] min-w-[210px] border-b border-r border-slate-600 px-4 py-0 align-middle" style={{background:'linear-gradient(180deg,#111c30 0%,#0d1826 100%)'}}>Project Name</th>
                  <th rowSpan={2} className="sticky top-0 left-[266px] z-30 h-[68px] min-w-[170px] border-b border-r border-slate-600 px-4 py-0 align-middle" style={{background:'linear-gradient(180deg,#111c30 0%,#0d1826 100%)'}}>Customer</th>
                  <th rowSpan={2} className="sticky top-0 left-[436px] z-30 h-[68px] min-w-[82px] border-b border-r border-slate-600 px-4 py-0 align-middle" style={{background:'linear-gradient(180deg,#111c30 0%,#0d1826 100%)'}}>Qty</th>
                  <th rowSpan={2} className="sticky top-0 z-20 h-[68px] min-w-[108px] border-b border-r border-slate-600 px-4 py-0 align-middle" style={{background:'linear-gradient(180deg,#111c30 0%,#0d1826 100%)'}}>Priority</th>
                  <th rowSpan={2} className="sticky top-0 z-20 h-[68px] min-w-[200px] border-b border-r border-slate-600 px-4 py-0 align-middle" style={{background:'linear-gradient(180deg,#111c30 0%,#0d1826 100%)'}}>Remarks</th>
                  <th rowSpan={2} className="sticky top-0 z-20 h-[68px] min-w-[112px] border-b border-r border-slate-600 px-4 py-0 align-middle" style={{background:'linear-gradient(180deg,#111c30 0%,#0d1826 100%)'}}>Data Received</th>
                  <th rowSpan={2} className="sticky top-0 z-20 h-[68px] min-w-[112px] border-b border-r border-slate-600 stage-separator-header px-4 py-0 align-middle" style={{background:'linear-gradient(180deg,#111c30 0%,#0d1826 100%)'}}>Data Shared</th>
                  {STAGE_STRUCTURE.map((stage) => (
                    <th key={stage.name} colSpan={stage.sub.length} className="table-stage-header sticky top-0 z-20 h-8 border-b border-r border-slate-600 stage-separator-header px-4 py-0 align-middle" style={{background:'linear-gradient(180deg,#0d1e2e 0%,#0c1826 100%)', minWidth: `${stage.sub.length * 110}px`}}>
                      {stage.name}
                    </th>
                  ))}
                  <th rowSpan={2} className="sticky top-0 z-20 h-[68px] min-w-[84px] border-b border-slate-600 px-4 py-0 align-middle" style={{background:'linear-gradient(180deg,#111c30 0%,#0d1826 100%)'}}>Actions</th>
                </tr>
                <tr className="table-header-row">
                   {STAGE_STRUCTURE.flatMap((stage) =>
                    stage.sub.map((sub, idx) => {
                      const isLast = idx === stage.sub.length - 1;
                      return (
                        <th
                          key={`${stage.name}-${sub}`}
                          className={`sticky top-8 z-20 h-9 min-w-[110px] border-b border-r border-slate-600/70 px-3 py-0 align-middle text-[8px] text-slate-400 ${
                            isLast ? 'stage-separator-header' : ''
                          }`}
                          style={{background:'linear-gradient(180deg,#0d1826 0%,#0b1520 100%)'}}
                        >
                          {(sub as string) === 'Date completion' || (sub as string) === 'Done date' ? 'Done Date' : sub.toUpperCase()}
                        </th>
                      );
                    })
                  )}
                </tr>
              </thead>
              <tbody className="text-sm text-slate-700">
                {Object.entries(groupedProjects).map(([group, items]) => (
                  <React.Fragment key={group}>
                    {groupBy !== 'None' && (
                      <tr className="group-row sticky top-[68px] z-[25]">
                        <td colSpan={9 + stageColumns.length} className="p-0">
                          <div className="sticky left-0 flex items-center gap-3 px-4 py-2 w-fit bg-white/95 backdrop-blur-md">
                            <div className="flex h-8 items-center gap-3 rounded-lg bg-slate-900/10 px-3 py-1.5 ring-1 ring-slate-900/5">
                              <span className="text-[9px] font-black uppercase tracking-[0.2em] text-slate-500 whitespace-nowrap">
                                {groupBy === 'Phase' ? 'Current Department' : groupBy.toUpperCase()}
                              </span>
                              <div className="h-3 w-px bg-slate-300" />
                              <span className="text-[11px] font-black text-slate-900 uppercase tracking-wider">
                                {group}
                              </span>
                            </div>
                            <div className="brand-divider !h-4 !mx-0" />
                            <div className="flex items-center gap-2 rounded-lg bg-emerald-500/10 px-2.5 py-1.5 ring-1 ring-emerald-500/20">
                              <span className="text-[10px] font-black text-emerald-600 uppercase tracking-widest tabular-nums whitespace-nowrap">
                                {items.length} {items.length === 1 ? 'Project' : 'Projects'}
                              </span>
                            </div>

                            {/* WORK DONE INFO — Only for Phase Grouping */}
                            {groupBy === 'Phase' && group !== 'Completed' && (
                              <>
                                <div className="brand-divider !h-4 !mx-2" />
                                <div className="flex items-center gap-2">
                                  <span className="text-[9px] font-black uppercase tracking-[0.15em] text-slate-400">Advanced Progress:</span>
                                  <div className="flex flex-wrap gap-1">
                                    {STAGE_STRUCTURE.filter(s => {
                                      const groupIndex = STAGE_STRUCTURE.findIndex(st => st.name === group);
                                      const stageIndex = STAGE_STRUCTURE.findIndex(st => st.name === s.name);
                                      return stageIndex > groupIndex;
                                    }).map(s => {
                                      const doneCount = items.filter(p => {
                                        const status = String(p.stages[`${s.name}_status`] || '').toLowerCase();
                                        return status === 'done' || status === 'completed' || status === 'n/a' || status === 'na';
                                      }).length;
                                      if (doneCount === 0) return null;
                                      return (
                                        <div key={s.name} className="flex items-center gap-1.5 rounded-md bg-blue-500/10 px-2 py-0.5 ring-1 ring-blue-500/20">
                                          <span className="text-[9px] font-bold text-blue-600 uppercase">{s.name}</span>
                                          <div className="h-2 w-px bg-blue-200" />
                                          <span className="text-[9px] font-black text-blue-700">{doneCount}</span>
                                        </div>
                                      );
                                    })}
                                  </div>
                                </div>
                              </>
                            )}
                          </div>
                        </td>
                      </tr>
                    )}
                    {items.map((project, idx) => (
                      <ProjectRow
                        key={project.id}
                        rowIndex={idx}
                        project={project}
                        selected={selectedProject?.id === project.id}
                        onSelect={handleRowSelect}
                        renderCell={renderEditableTableCell}
                        renderDate={renderSheetDate}
                        openEditModal={openEditModal}
                        handleDelete={handleDelete}
                        onStageNoteEdit={(proj, noteKey) => {
                          setSelectedProject(proj);
                          startCellEdit(proj, noteKey, 'stage', 'sidebar');
                        }}
                        editingCell={editingCell}
                        renderInlineEditor={renderInlineEditor}
                        getEditableCellValue={getEditableCellValue}
                        onStageCellEdit={handleStageTableCellEdit}
                        onOpenAnalytics={(proj) => setAnalyticsProject(proj)}
                        onOpenCustomerDetails={openCustomerModalFor}
                      />
                    ))}
                  </React.Fragment>
                ))}
                {filteredProjects.length === 0 && (
                  <tr>
                    <td colSpan={9 + stageColumns.length} className="empty-state">
                      <ClipboardList className="mx-auto h-10 w-10 text-slate-300" />
                      <h3 className="mt-3 text-lg font-bold text-slate-700">No matching projects</h3>
                      <p className="mt-1 text-sm text-slate-400">Try another filter or import your PMS workbook.</p>
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>

          <AnimatePresence>
            {selectedProject && (
              <motion.div
                initial={{ x: 380, opacity: 0 }}
                animate={{ x: 0, opacity: 1 }}
                exit={{ x: 380, opacity: 0 }}
                transition={MOTION_SIDEBAR}
                className="absolute top-0 right-0 h-full w-full sm:w-[360px] border-l border-slate-200 bg-white shadow-2xl flex flex-col z-30 shrink-0 overflow-hidden"
              >
                <div className="flex flex-col h-full overflow-hidden">
                  <div className="relative border-b border-slate-800 bg-[#0c1322] p-4 text-white shrink-0">
                    <div className="relative flex items-start justify-between gap-2">
                      <div className="space-y-0.5 min-w-0">
                        <p className="text-[8px] font-black uppercase tracking-[0.3em] text-emerald-400">Project Detail</p>
                        <h2 className="text-base font-black text-white tracking-tight leading-tight truncate">
                          <span className="text-[10px] font-bold text-slate-400 tracking-wider">SR {selectedProject.sr_no || 'NA'}</span>
                          <span className="mx-1.5 text-slate-600">|</span>
                          <span className="text-white">{selectedProject.project_name}</span>
                        </h2>
                        <p className="text-[10px] font-semibold text-slate-400 uppercase tracking-wide truncate">
                          {selectedProject.customer || 'NA'}
                        </p>
                      </div>
                      <button
                        onClick={() => setSelectedProject(null)}
                        className="flex h-7 w-7 items-center justify-center rounded-lg border border-white/10 bg-white/5 text-slate-400 hover:bg-white/10 hover:text-white transition-all shrink-0 cursor-pointer"
                      >
                        <X size={15} />
                      </button>
                    </div>

                    <div className="relative mt-2.5 flex gap-1.5">
                      <span className={`inline-flex items-center rounded-md px-2 py-0.5 text-[9px] font-black uppercase ${
                        selectedProject.priority === 'High' ? 'bg-rose-500 text-white' :
                        selectedProject.priority === 'Medium' ? 'bg-amber-400 text-amber-950' : 'bg-slate-700 text-slate-200'
                      }`}>
                        {selectedProject.priority} Priority
                      </span>
                      <span className="inline-flex items-center rounded-md border border-emerald-500/30 bg-emerald-500/10 px-2 py-0.5 text-[9px] font-black uppercase text-emerald-300">
                        {selectedProject.status || 'ACTIVE'}
                      </span>
                    </div>
                  </div>

                  <div className="flex-1 overflow-auto p-3.5 custom-scrollbar bg-slate-50/40 space-y-3" style={{scrollbarWidth:'thin'}}>
                    {/* Master Remarks */}
                    <div className="space-y-1">
                      <p className="text-[8px] font-black uppercase tracking-[0.2em] text-slate-400 px-0.5">Master Remarks</p>
                      {renderSidebarEditableField(selectedProject, 'remarks', 'project', selectedProject.remarks, 'REMARKS', (
                        <div className="min-h-[40px] text-[11px] font-medium leading-normal text-slate-700 italic">
                          {selectedProject.remarks || 'No remarks recorded'}
                        </div>
                      ), 'sidebar-field-card')}
                    </div>

                    {/* Core fields grid */}
                    <div className="grid grid-cols-2 gap-2">
                      {renderSidebarEditableField(selectedProject, 'priority', 'project', selectedProject.priority, 'URGENCY', (
                        <div className="min-w-0">
                          <p className="text-[8px] font-black text-slate-400 mb-1 uppercase tracking-widest">Priority</p>
                          <span className={`badge ${'badge-priority-' + selectedProject.priority.toLowerCase()}`}>{selectedProject.priority}</span>
                        </div>
                      ), 'sidebar-field-card')}

                      {renderSidebarEditableField(selectedProject, 'qty', 'project', String(selectedProject.qty), 'QUANTITY', (
                        <div className="min-w-0">
                          <p className="text-[8px] font-black text-slate-400 mb-0.5 uppercase tracking-widest">Qty</p>
                          <p className="text-lg font-black text-slate-900">{selectedProject.qty}</p>
                        </div>
                      ), 'sidebar-field-card')}

                      {renderSidebarEditableField(selectedProject, 'due_date', 'project', selectedProject.due_date, 'DATA RECEIVED', (
                        <div className="min-w-0">
                          <p className="text-[8px] font-black text-slate-400 mb-1 uppercase tracking-widest">Data Received</p>
                          {renderSheetDate(selectedProject.due_date, '', 'neutral', 'date-pill-violet')}
                        </div>
                      ), 'sidebar-field-card')}

                      {renderSidebarEditableField(selectedProject, 'completion_date', 'project', selectedProject.completion_date, 'DATA SHARED', (
                        <div className="min-w-0">
                          <p className="text-[8px] font-black text-slate-400 mb-1 uppercase tracking-widest">Data Shared</p>
                          {renderSheetDate(selectedProject.completion_date, '', 'neutral', 'date-pill-pink')}
                        </div>
                      ), 'sidebar-field-card')}
                    </div>

                    {/* Stage divider */}
                    <div className="flex items-center gap-2 py-0.5">
                      <div className="h-px flex-1 bg-slate-200" />
                      <span className="text-[8px] font-black uppercase tracking-[0.25em] text-slate-400">Stage Timeline</span>
                      <div className="h-px flex-1 bg-slate-200" />
                    </div>

                    {/* Stage cards */}
                    <div className="space-y-2">
                      {STAGE_STRUCTURE.map((stage, stageIdx) => {
                        const statusKey = `${stage.name}_status`;
                        const remarkKey = `${stage.name}_remarks`;
                        const status = String(selectedProject.stages[statusKey] || 'Pending');
                        const remark = String(selectedProject.stages[remarkKey] || '');
                        const accentColor = SPIS_PALETTE[stageIdx % SPIS_PALETTE.length];

                        return (
                          <div key={stage.name} className="sidebar-stage-card relative overflow-hidden rounded-xl border border-slate-200 bg-white" style={{ borderLeft: `3.5px solid ${accentColor}` }}>
                            <div className="flex items-center justify-between bg-slate-50 px-3 py-2 border-b border-slate-100">
                              <div className="flex items-center gap-1.5">
                                <span className="h-2 w-2 rounded-full shadow-sm" style={{ backgroundColor: accentColor }} />
                                <span className="text-[10px] font-black text-slate-800 uppercase tracking-wider">{stage.name}</span>
                              </div>
                              {renderSidebarEditableField(selectedProject, statusKey, 'stage', status, `${stage.name} STATUS`, (
                                <span className={getStageColorClass(status) + ' min-w-[76px]'}>
                                  {status.toUpperCase()}
                                </span>
                              ), 'cursor-pointer hover:scale-105 transition-transform')}
                            </div>

                            <div className="p-2.5 bg-slate-50/30 border-b border-slate-100">
                              <div className="grid grid-cols-2 gap-2">
                                {stage.sub.filter((s) => s !== 'status').map((sub) => {
                                  const key = `${stage.name}_${sub}`;
                                  let dateVal = String(selectedProject.stages[key] || '');
                                  if (key === 'Dispatch_Remaining Qty') {
                                    const totalQty = selectedProject.qty || 0;
                                    const shippedQty = parseFloat(String(selectedProject.stages['Dispatch_Shipped Qty'] || '0')) || 0;
                                    dateVal = String(Math.max(0, totalQty - shippedQty));
                                  }
                                  const completionValue = String(selectedProject.stages[`${stage.name}_Done date`] || selectedProject.stages[`${stage.name}_Date completion`] || '');
                                  const dueKind = (sub as string) === 'status' ? 'neutral' : ((sub as string) === 'Date completion' || (sub as string) === 'Done date' ? 'completion' : 'due');
                                  const isDateField = isDateFieldKey(key);
                                  const subLabel = sub.replace('Date completion', 'Done Date').replace('Start date', 'Start Date').replace('Due date', 'Due Date').toUpperCase();

                                  let pillContent: React.ReactNode;
                                  if (isDateField) {
                                    if (sub === 'Start date') {
                                      pillContent = renderSheetDate(dateVal, completionValue, 'neutral', 'date-pill-blue');
                                    } else {
                                      pillContent = renderSheetDate(dateVal, completionValue, dueKind === 'completion' ? 'completion' : 'due');
                                    }
                                  } else if (key === 'Dispatch_Remaining Qty') {
                                    const valNum = parseFloat(dateVal) || 0;
                                    const totalQty = selectedProject.qty || 0;
                                    let qtyClass = "qty-pill-rose";
                                    if (valNum === 0) qtyClass = "qty-pill-emerald";
                                    else if (valNum < totalQty) qtyClass = "qty-pill-amber";
                                    pillContent = <span className={qtyClass}>{dateVal}</span>;
                                  } else if (key === 'Dispatch_Shipped Qty') {
                                    pillContent = <span className="qty-pill-amber">{dateVal || '0'}</span>;
                                  } else {
                                    pillContent = <span className="text-[11px] font-bold text-slate-800 bg-slate-100 px-2 py-0.5 rounded border border-slate-200 inline-block">{dateVal || '0'}</span>;
                                  }

                                  return (
                                    <React.Fragment key={key}>
                                      {renderSidebarEditableField(selectedProject, key, 'stage', dateVal, subLabel, (
                                        <div className="min-w-0">
                                          <p className="text-[8px] font-black text-slate-400 mb-0.5 uppercase tracking-widest">{subLabel}</p>
                                          {pillContent}
                                        </div>
                                      ), 'sidebar-field-card')}
                                    </React.Fragment>
                                  );
                                })}
                              </div>
                            </div>

                              <div className="p-3">
                                {renderSidebarEditableField(selectedProject, remarkKey, 'stage', remark, 'STAGE REMARKS', (
                                  <div className="space-y-1">
                                    <div className="flex items-center gap-1.5 opacity-50">
                                      <Tag className="h-3 w-3" />
                                      <span className="text-[8px] font-black uppercase tracking-widest text-slate-500">Notes</span>
                                    </div>
                                    <p className="text-[11px] font-semibold text-slate-600 leading-relaxed italic">
                                      {remark || 'No notes...'}
                                    </p>
                                  </div>
                                ), 'cursor-pointer hover:bg-slate-50 p-2 rounded-lg transition-all border border-transparent hover:border-slate-200')}
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </div>

                  <div className="border-t border-slate-200 bg-white p-4">
                    <div className="flex gap-2">
                      <button
                        onClick={(e) => openEditModal(selectedProject, e)}
                        className="flex-1 flex items-center justify-center gap-2 rounded-lg bg-slate-900 py-2.5 text-[11px] font-black text-white shadow-md transition-all hover:bg-slate-800 hover:scale-[1.02] active:scale-[0.98]"
                      >
                        <Pencil className="h-3.5 w-3.5" /> Edit Row
                      </button>
                      <button
                        onClick={(e) => handleDelete(selectedProject.id!, e)}
                        className="flex h-9 w-9 items-center justify-center rounded-lg border border-rose-200 text-rose-500 hover:bg-rose-600 hover:text-white hover:border-rose-600 transition-all"
                      >
                        <Trash2 className="h-4 w-4" />
                      </button>
                    </div>
                  </div>
                </div>
              </motion.div>
            )}
          </AnimatePresence>
        </div>

        {/* Pagination Controls Footer */}
        {filteredProjects.length > 0 && (
          <div className="relative z-30 flex flex-col sm:flex-row items-center justify-between gap-4 border-t border-slate-200 bg-white/90 px-6 py-1.5 backdrop-blur-md shrink-0">
            {/* Left: Info */}
            <div className="text-[10px] text-slate-500 font-semibold uppercase tracking-wider">
              Showing <span className="font-extrabold text-slate-800">{pageSize === -1 ? 1 : (currentPage - 1) * pageSize + 1}</span> to{' '}
              <span className="font-extrabold text-slate-800">
                {pageSize === -1 ? filteredProjects.length : Math.min(currentPage * pageSize, filteredProjects.length)}
              </span>{' '}
              of <span className="font-extrabold text-slate-800">{filteredProjects.length}</span> entries
            </div>

            {/* Middle: Pagination Buttons */}
            {pageSize !== -1 && filteredProjects.length > pageSize && (
              <div className="flex items-center gap-1">
                <button
                  type="button"
                  disabled={currentPage === 1}
                  onClick={() => setCurrentPage(1)}
                  className="flex h-7 w-7 items-center justify-center rounded-lg border border-slate-200 bg-white text-slate-500 transition-all hover:bg-slate-50 hover:text-slate-800 disabled:opacity-40 disabled:hover:bg-white disabled:hover:text-slate-500"
                  title="First Page"
                >
                  <ChevronsLeft className="h-3.5 w-3.5" />
                </button>
                <button
                  type="button"
                  disabled={currentPage === 1}
                  onClick={() => setCurrentPage(p => Math.max(1, p - 1))}
                  className="flex h-7 w-7 items-center justify-center rounded-lg border border-slate-200 bg-white text-slate-500 transition-all hover:bg-slate-50 hover:text-slate-800 disabled:opacity-40 disabled:hover:bg-white disabled:hover:text-slate-500"
                  title="Previous Page"
                >
                  <ChevronLeft className="h-3.5 w-3.5" />
                </button>

                {/* Simple active page number badge */}
                <div className="flex items-center justify-center px-2.5 h-7 rounded-lg border border-indigo-200 bg-indigo-50/50 text-[10px] font-extrabold text-indigo-600">
                  PAGE {currentPage} OF {Math.ceil(filteredProjects.length / pageSize)}
                </div>

                <button
                  type="button"
                  disabled={currentPage >= Math.ceil(filteredProjects.length / pageSize)}
                  onClick={() => setCurrentPage(p => Math.min(Math.ceil(filteredProjects.length / pageSize), p + 1))}
                  className="flex h-7 w-7 items-center justify-center rounded-lg border border-slate-200 bg-white text-slate-500 transition-all hover:bg-slate-50 hover:text-slate-800 disabled:opacity-40 disabled:hover:bg-white disabled:hover:text-slate-500"
                  title="Next Page"
                >
                  <ChevronRight className="h-3.5 w-3.5" />
                </button>
                <button
                  type="button"
                  disabled={currentPage >= Math.ceil(filteredProjects.length / pageSize)}
                  onClick={() => setCurrentPage(Math.ceil(filteredProjects.length / pageSize))}
                  className="flex h-7 w-7 items-center justify-center rounded-lg border border-slate-200 bg-white text-slate-500 transition-all hover:bg-slate-50 hover:text-slate-800 disabled:opacity-40 disabled:hover:bg-white disabled:hover:text-slate-500"
                  title="Last Page"
                >
                  <ChevronsRight className="h-3.5 w-3.5" />
                </button>
              </div>
            )}

            {/* Right: Page Size Selector */}
            <div className="flex items-center gap-2 relative z-40">
              <span className="text-[10px] font-black uppercase tracking-wider text-slate-400">Rows per page:</span>
              <div className="relative z-50" id="pagesize-dropdown-container">
                <button
                  type="button"
                  onClick={() => setIsPageSizeDropdownOpen(p => !p)}
                  className="flex h-7.5 min-w-[72px] items-center justify-between gap-1.5 rounded-lg border border-slate-300 bg-white px-2.5 text-[11px] font-extrabold text-slate-800 shadow-sm hover:border-indigo-400 hover:bg-indigo-50/30 transition-all cursor-pointer"
                >
                  <span>{pageSize === -1 ? 'All' : pageSize}</span>
                  <ChevronDown className={`h-3.5 w-3.5 text-slate-500 transition-transform duration-200 ${isPageSizeDropdownOpen ? 'rotate-180 text-indigo-600' : ''}`} />
                </button>

                <AnimatePresence>
                  {isPageSizeDropdownOpen && (
                    <motion.div 
                      initial={{ opacity: 0, scale: 0.95, y: 4 }}
                      animate={{ opacity: 1, scale: 1, y: 0 }}
                      exit={{ opacity: 0, scale: 0.95, y: 4 }}
                      transition={{ duration: 0.1, ease: 'easeOut' }}
                      className="absolute right-0 bottom-full mb-2 w-32 rounded-xl border border-slate-200 bg-white p-1.5 shadow-[0_12px_36px_rgba(0,0,0,0.22)] z-[99999] ring-1 ring-slate-900/10 pointer-events-auto will-change-transform"
                      style={{ backgroundColor: '#ffffff', opacity: 1 }}
                    >
                      <div className="px-2 py-1 text-[9px] font-black uppercase tracking-wider text-slate-400 border-b border-slate-100 mb-1">
                        Rows Per Page
                      </div>
                      {([25, 50, 100, 250, -1] as const).map((size) => (
                        <button
                          key={size}
                          type="button"
                          onClick={() => {
                            setPageSize(size);
                            setCurrentPage(1);
                            setIsPageSizeDropdownOpen(false);
                          }}
                          className={`flex w-full items-center justify-between rounded-lg px-2.5 py-1.5 text-xs font-bold transition-all cursor-pointer ${
                            pageSize === size
                              ? 'bg-indigo-50 text-indigo-700 font-black'
                              : 'text-slate-700 hover:bg-slate-100 hover:text-slate-900'
                          }`}
                        >
                          <span>{size === -1 ? 'All' : size}</span>
                          {pageSize === size && <Check className="h-3.5 w-3.5 text-indigo-600 stroke-[2.5]" />}
                        </button>
                      ))}
                    </motion.div>
                  )}
                </AnimatePresence>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>

      <AnimatePresence>
      {isProjectModalOpen && (
        <motion.div
          initial={{ opacity: 0 }}
          animate={{ opacity: 1 }}
          exit={{ opacity: 0 }}
          transition={MOTION_FADE}
          className="dialog-backdrop fixed inset-0 z-[1000] flex items-center justify-center bg-slate-950/55 p-4"
        >
          <motion.div
            initial={{ scale: 0.95, opacity: 0, y: 28 }}
            animate={{ scale: 1, opacity: 1, y: 0 }}
            exit={{ scale: 0.95, opacity: 0, y: 20 }}
            transition={MOTION_PANEL}
            className="dialog-panel max-h-[92vh] w-full max-w-6xl overflow-hidden rounded-[28px] bg-white shadow-[0_35px_120px_rgba(15,23,42,0.25)] flex flex-col"
          >
            <div className="dialog-scroll flex-1 overflow-y-auto p-8">
              <div className="flex items-center justify-between border-b border-slate-200 pb-5"><div><p className="text-xs font-semibold uppercase tracking-[0.24em] text-teal-700">{editingId ? 'Edit Entry' : 'New Entry'}</p><h3 className="mt-2 text-3xl font-semibold tracking-tight text-slate-900">{editingId ? 'Update workbook row' : 'Add workbook row'}</h3></div><button title="Close modal" onClick={() => setIsProjectModalOpen(false)} className="rounded-2xl border border-slate-200 p-3 text-slate-500"><X className="h-5 w-5" /></button></div>
            <form onSubmit={saveProject} className="space-y-8 pt-6">
              <div className="grid gap-4 md:grid-cols-2 lg:grid-cols-4">
                {(['sr_no', 'customer', 'project_name', 'qty', 'priority', 'due_date', 'completion_date'] as const).map((field) => (
                  <label key={field} className="space-y-2">
                    <span className="text-sm font-medium text-slate-700">{FORM_LABELS[field] || field.replaceAll('_', ' ')}</span>
                    {field === 'priority' ? (
                      <DropdownMenu
                        variant="light"
                        value={form.priority}
                        onChange={(value) => setForm((c) => ({ ...c, priority: value }))}
                        options={[
                          { value: 'High', label: 'High' },
                          { value: 'Medium', label: 'Medium' },
                          { value: 'Low', label: 'Low' },
                        ]}
                        className="w-full"
                        buttonClassName="ctrl-input !h-[36px] text-xs font-semibold text-slate-800"
                        menuClassName="w-full"
                      />
                    ) : field === 'customer' ? (
                      <div className="relative">
                        <input
                          value={form.customer}
                          type="text"
                          onChange={(e) => {
                            const val = e.target.value;
                            setForm((c) => ({ ...c, customer: val }));
                            setShowCustomerSuggestions(val.trim().length > 0);
                          }}
                          onBlur={() => setTimeout(() => setShowCustomerSuggestions(false), 200)}
                          className="ctrl-input !h-[36px] text-xs font-semibold text-slate-900 w-full"
                          placeholder="ENTER CUSTOMER NAME"
                        />
                        {showCustomerSuggestions && customerSuggestions.length > 0 && (
                          <div className="absolute left-0 right-0 top-full mt-1 z-[1200] max-h-48 overflow-y-auto rounded-xl border border-slate-200 bg-white shadow-xl py-1">
                            {customerSuggestions.map((name) => (
                              <button
                                key={name}
                                type="button"
                                onMouseDown={() => {
                                  setForm((c) => ({ ...c, customer: name }));
                                  setShowCustomerSuggestions(false);
                                }}
                                className="w-full px-3 py-2 text-left text-xs font-bold text-slate-700 hover:bg-sky-50 hover:text-sky-700 transition-colors border-b border-slate-100 last:border-none flex items-center justify-between cursor-pointer"
                              >
                                <span>{name}</span>
                                <Building2 className="h-3 w-3 text-slate-400" />
                              </button>
                            ))}
                          </div>
                        )}
                      </div>
                    ) : (
                      <input value={form[field]} type={field === 'qty' ? 'number' : field.includes('date') ? 'date' : 'text'} onChange={(e) => setForm((c) => ({ ...c, [field]: e.target.value }))} className="ctrl-input !h-[36px] text-xs font-semibold text-slate-900" />
                    )}
                  </label>
                ))}
              </div>
              <div className="space-y-4 pt-2">
                <div className="flex items-center gap-3">
                  <div className="h-px flex-1 bg-slate-200" />
                  <span className="text-[11px] font-black uppercase tracking-[0.25em] text-slate-400">Production Stage Timeline</span>
                  <div className="h-px flex-1 bg-slate-200" />
                </div>

                <div className="grid gap-4 md:grid-cols-2 lg:grid-cols-3">
                  {STAGE_STRUCTURE.map((stage, stageIdx) => {
                    const accentColor = SPIS_PALETTE[stageIdx % SPIS_PALETTE.length];
                    return (
                      <div
                        key={stage.name}
                        className="relative overflow-hidden rounded-2xl border border-slate-200 bg-white p-4 shadow-sm transition-all hover:shadow-md hover:border-slate-300"
                        style={{ borderTop: `4px solid ${accentColor}` }}
                      >
                        <div className="flex items-center justify-between pb-3 mb-3 border-b border-slate-100">
                          <div className="flex items-center gap-2 min-w-0">
                            <span
                              className="h-2.5 w-2.5 rounded-full shrink-0 shadow-sm"
                              style={{ backgroundColor: accentColor }}
                            />
                            <span className="text-xs font-black uppercase tracking-wider text-slate-800 truncate">
                              {stage.name}
                            </span>
                          </div>
                          <span
                            className="text-[9px] font-bold px-2 py-0.5 rounded-full uppercase tracking-wider text-white shrink-0"
                            style={{ backgroundColor: accentColor, opacity: 0.9 }}
                          >
                            Stage {stageIdx + 1}
                          </span>
                        </div>

                        <div className="space-y-3">
                          {stage.sub.map((sub) => {
                            const key = `${stage.name}_${sub}`;
                            const subLabel = sub.replace('Date completion', 'Done Date').replace('Start date', 'Start Date').replace('Due date', 'Due Date');
                            const isStatus = sub === 'status';

                            return (
                              <div key={key} className="space-y-1">
                                <span className="block text-[10px] font-black uppercase tracking-widest text-slate-400">
                                  {subLabel}
                                </span>
                                {isStatus ? (
                                  <DropdownMenu
                                    variant="light"
                                    value={form.stages[key] || 'Pending'}
                                    onChange={(value) => setForm((c) => ({ ...c, stages: { ...c.stages, [key]: value } }))}
                                    options={[
                                      { value: 'Pending', label: 'Pending' },
                                      { value: 'Running', label: 'Running' },
                                      { value: 'On Hold', label: 'On Hold' },
                                      { value: 'Done', label: 'Done' },
                                      { value: 'N/A', label: 'N/A' },
                                    ]}
                                    className="w-full"
                                    buttonClassName="h-9 w-full rounded-xl border border-slate-200 bg-slate-50 px-3 text-xs font-bold text-slate-800 text-left"
                                    menuClassName="w-full"
                                  />
                                ) : (
                                  <input
                                    value={
                                      key === 'Dispatch_Remaining Qty'
                                        ? String(Math.max(0, (Number(form.qty) || 0) - (Number(form.stages['Dispatch_Shipped Qty']) || 0)))
                                        : (form.stages[key] || '')
                                    }
                                    type={key.includes('date') || key.includes('Date') ? 'date' : 'text'}
                                    readOnly={key === 'Dispatch_Remaining Qty'}
                                    onChange={(e) => setForm((c) => ({ ...c, stages: { ...c.stages, [key]: e.target.value } }))}
                                    className={`h-9 w-full rounded-xl border border-slate-200 bg-slate-50 px-3 text-xs font-medium text-slate-800 outline-none transition-all focus:border-teal-500 focus:bg-white focus:ring-2 focus:ring-teal-500/10 ${
                                      key === 'Dispatch_Remaining Qty' ? 'cursor-not-allowed text-slate-400 font-bold opacity-70' : ''
                                    }`}
                                    placeholder={key.includes('Due date') ? 'due date' : key.includes('Date completion') ? 'completion date' : ''}
                                  />
                                )}
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
              <label className="space-y-2"><span className="text-sm font-medium text-slate-700">Remarks</span><textarea value={form.remarks} onChange={(e) => setForm((c) => ({ ...c, remarks: e.target.value }))} rows={4} className="w-full rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3" /></label>
              <div className="flex justify-end gap-3"><button type="button" onClick={() => setIsProjectModalOpen(false)} className="rounded-2xl border border-slate-200 px-5 py-3 text-sm font-semibold text-slate-600">Cancel</button><button type="submit" className="rounded-2xl bg-slate-950 px-6 py-3 text-sm font-semibold text-white">{editingId ? 'Update Entry' : 'Add Entry'}</button></div>
            </form>
            </div>
          </motion.div>
        </motion.div>
      )}
      </AnimatePresence>

      <AnimatePresence>
      {isImportOpen && (
        <motion.div
          initial={{ opacity: 0 }}
          animate={{ opacity: 1 }}
          exit={{ opacity: 0 }}
          transition={MOTION_FADE}
          className="dialog-backdrop fixed inset-0 z-[1000] flex items-center justify-center bg-slate-950/55 p-4"
        >
          <motion.div
            initial={{ scale: 0.95, opacity: 0, y: 28 }}
            animate={{ scale: 1, opacity: 1, y: 0 }}
            exit={{ scale: 0.95, opacity: 0, y: 20 }}
            transition={MOTION_PANEL}
            className="dialog-panel w-full max-w-4xl rounded-[28px] bg-white p-8 shadow-[0_35px_120px_rgba(15,23,42,0.25)]"
          >
            <div className="flex items-center justify-between"><div><p className="text-xs font-semibold uppercase tracking-[0.24em] text-teal-700">Excel Import</p><h3 className="mt-2 text-3xl font-semibold tracking-tight text-slate-900">Import PMS workbook</h3></div><button title="Close import dialog" onClick={() => setIsImportOpen(false)} className="rounded-2xl border border-slate-200 p-3 text-slate-500"><X className="h-5 w-5" /></button></div>
            {!importRows.length ? (
              <label className="mt-8 flex cursor-pointer flex-col items-center justify-center rounded-[28px] border-2 border-dashed border-slate-300 bg-slate-50 px-6 py-20 text-center">
                <Upload className="h-10 w-10 text-teal-700" />
                <h4 className="mt-4 text-xl font-semibold text-slate-900">Choose `pms.xlsm`</h4>
                <p className="mt-2 text-sm text-slate-500">The importer reads the exact 2-row PMS header structure.</p>
                <input type="file" accept=".xlsx,.xls,.xlsm" className="hidden" onChange={handleImportFile} />
              </label>
            ) : (
              <div className="mt-8 space-y-6">
                <div className="grid gap-4 md:grid-cols-2">{CORE_FIELDS.map((field) => <label key={field.key} className="space-y-2"><span className="text-sm font-medium text-slate-700">{field.label}</span><DropdownMenu variant="light" value={mappings[field.key] || ''} onChange={(value) => setMappings((c) => ({ ...c, [field.key]: value }))} placeholder="Select source column" options={importHeaders.map((h) => ({ value: h, label: h }))} className="w-full" buttonClassName="h-12 w-full rounded-2xl border border-slate-200 bg-slate-50 px-4 text-slate-700" menuClassName="w-full" /></label>)}</div>
                <div className="flex items-center justify-between rounded-2xl bg-slate-950 px-5 py-4 text-white"><span className="text-sm font-medium">{importRows.length} rows ready for import</span><button onClick={confirmImport} className="rounded-2xl bg-teal-400 px-5 py-2 text-sm font-semibold text-slate-950">Confirm Import</button></div>
              </div>
            )}
          </motion.div>
        </motion.div>
      )}
      </AnimatePresence>

      <AnimatePresence>
      {isToolsOpen && (
        <motion.div
          initial={{ opacity: 0 }}
          animate={{ opacity: 1 }}
          exit={{ opacity: 0 }}
          transition={MOTION_FADE}
          className="dialog-backdrop fixed inset-0 z-[1000] flex items-center justify-center bg-slate-950/55 p-4"
        >
          <motion.div
            initial={{ scale: 0.95, opacity: 0, y: 28 }}
            animate={{ scale: 1, opacity: 1, y: 0 }}
            exit={{ scale: 0.95, opacity: 0, y: 20 }}
            transition={MOTION_PANEL}
            className="dialog-panel w-full max-w-xl rounded-[32px] bg-white shadow-[0_40px_130px_rgba(15,23,42,0.28)] overflow-hidden flex flex-col max-h-[90vh]"
          >
            {/* Modal Header */}
            <div className="px-8 pt-8 pb-4 border-b border-slate-100 shrink-0 bg-white z-10">
              <div className="flex items-center justify-between mb-6">
                <div>
                  <div className="flex items-center gap-2 mb-1">
                    <div className="h-5 w-1 rounded-full bg-emerald-500" />
                    <p className="text-[10px] font-black uppercase tracking-[0.3em] text-emerald-600">Command Center</p>
                  </div>
                  <h3 className="text-3xl font-black tracking-tight text-slate-900 leading-none">System Settings</h3>
                </div>
                <button title="Close settings" onClick={() => setIsToolsOpen(false)} className="rounded-2xl border border-slate-200 p-3 text-slate-400 hover:bg-slate-50 hover:text-slate-600 transition-all">
                  <X className="h-6 w-6" />
                </button>
              </div>

              {/* Tab Navigation */}
              <div className="flex gap-1 bg-slate-100 p-1.5 rounded-2xl">
                <button 
                  onClick={() => setSettingsTab('General')}
                  className={`flex-1 flex items-center justify-center gap-2 py-2.5 rounded-xl text-[11px] font-black uppercase tracking-widest transition-all ${
                    settingsTab === 'General' ? 'bg-white text-slate-900 shadow-sm' : 'text-slate-400 hover:text-slate-600'
                  }`}
                >
                  <Settings size={14} /> General
                </button>
                <button 
                  onClick={() => setSettingsTab('Logs')}
                  className={`flex-1 flex items-center justify-center gap-2 py-2.5 rounded-xl text-[11px] font-black uppercase tracking-widest transition-all ${
                    settingsTab === 'Logs' ? 'bg-white text-slate-900 shadow-sm' : 'text-slate-400 hover:text-slate-600'
                  }`}
                >
                  <ClipboardList size={14} /> Change Log
                </button>
              </div>
            </div>

            {/* Scrollable Content */}
            <div className="flex-1 overflow-y-auto px-8 py-6 custom-scrollbar bg-slate-50/30">
              {settingsTab === 'General' ? (
                <div className="space-y-6">
                  {/* Operator Settings */}
                  <div className="group/card rounded-2xl border border-slate-200 bg-white p-4 transition-all hover:border-emerald-200 shadow-sm">
                    <div className="flex items-center gap-3 mb-3">
                      <div className="flex h-8 w-8 items-center justify-center rounded-lg bg-slate-900 text-white shadow">
                        <User className="h-4 w-4" />
                      </div>
                      <div>
                        <label className="block text-[10px] font-black text-slate-400 uppercase tracking-widest">Operator Identity</label>
                        <h4 className="text-xs font-bold text-slate-900 uppercase">Profile Settings</h4>
                      </div>
                    </div>
                    <div className="relative">
                      <input
                        type="text"
                        value={currentUser}
                        onChange={(e) => {
                          const val = e.target.value.toUpperCase();
                          setCurrentUser(val);
                          localStorage.setItem('pms-current-user', val);
                        }}
                        placeholder="ENTER OPERATOR NAME"
                        className="ctrl-input !h-[34px] text-xs font-bold text-slate-900"
                      />
                    </div>
                  </div>

                  {/* Permanent Default Zoom Setting (Requirement 2) */}
                  <div className="group/card rounded-2xl border border-slate-200 bg-white p-4 transition-all hover:border-emerald-200 shadow-sm">
                    <div className="flex items-center gap-3 mb-3">
                      <div className="flex h-8 w-8 items-center justify-center rounded-lg bg-emerald-600 text-white shadow">
                        <Search className="h-4 w-4" />
                      </div>
                      <div>
                        <label className="block text-[10px] font-black text-slate-400 uppercase tracking-widest">Display Preferences</label>
                        <h4 className="text-xs font-bold text-slate-900 uppercase">Default App Zoom (Permanent)</h4>
                      </div>
                    </div>
                    <div className="flex items-center gap-1.5">
                      {[0.7, 0.8, 0.9, 1.0, 1.1, 1.25].map((z) => (
                        <button
                          key={z}
                          type="button"
                          onClick={() => {
                            setPermanentDefaultZoom(z);
                            localStorage.setItem('pms-permanent-default-zoom', String(z));
                            setZoomLevel(z);
                          }}
                          className={`flex-1 py-1.5 rounded-lg text-[10px] font-bold transition-all cursor-pointer ${
                            permanentDefaultZoom === z
                              ? 'bg-emerald-600 text-white shadow-md'
                              : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                          }`}
                        >
                          {Math.round(z * 100)}%
                        </button>
                      ))}
                    </div>
                    <p className="mt-2 text-[10px] font-medium text-slate-500 italic">
                      App starts with this permanent zoom level when opened. Temporary zoom on home controls current session only.
                    </p>
                  </div>

                  {/* Automation Settings */}
                  <div className="group/card rounded-2xl border border-slate-200 bg-white p-4 transition-all hover:border-sky-200 shadow-sm">
                    <div className="flex items-center gap-3 mb-3">
                      <div className="flex h-8 w-8 items-center justify-center rounded-lg bg-sky-600 text-white shadow">
                        <Clock className="h-4 w-4" />
                      </div>
                      <div>
                        <label className="block text-[10px] font-black text-slate-400 uppercase tracking-widest">Notification Engine</label>
                        <h4 className="text-xs font-bold text-slate-900 uppercase">Snooze Duration</h4>
                      </div>
                    </div>
                    
                    <div className="grid grid-cols-2 gap-3">
                      <div className="space-y-1">
                        <label className="text-[9px] font-bold uppercase text-slate-400 pl-1 tracking-wider">Hours</label>
                        <input
                          type="number"
                          min="0"
                          max="24"
                          value={alertSnoozeHours}
                          onChange={(e) => {
                            const val = Math.max(0, Math.min(24, Number(e.target.value)));
                            setAlertSnoozeHours(val);
                            localStorage.setItem('traceability-alert-snooze-hours', String(val));
                          }}
                          className="ctrl-input !h-[34px] text-xs font-bold text-slate-900"
                        />
                      </div>
                      <div className="space-y-1">
                        <label className="text-[9px] font-bold uppercase text-slate-400 pl-1 tracking-wider">Minutes</label>
                        <input
                          type="number"
                          min="0"
                          max="59"
                          value={alertSnoozeMinutes}
                          onChange={(e) => {
                            const val = Math.max(0, Math.min(59, Number(e.target.value)));
                            setAlertSnoozeMinutes(val);
                            localStorage.setItem('traceability-alert-snooze-minutes', String(val));
                          }}
                          className="ctrl-input !h-[34px] text-xs font-bold text-slate-900"
                        />
                      </div>
                    </div>
                  </div>

                  {/* Operations Section */}
                  <div className="space-y-3">
                    <label className="block text-[10px] font-black uppercase tracking-[0.3em] text-slate-400 px-1">System Operations</label>
                    
                    <button onClick={() => { setIsToolsOpen(false); setIsImportOpen(true); }} className="w-full flex items-center justify-between rounded-2xl border border-slate-200 bg-white p-5 transition-all hover:bg-emerald-50 hover:border-emerald-200 group/btn">
                      <div className="flex items-center gap-4">
                        <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-emerald-100 text-emerald-600 transition-colors group-hover/btn:bg-emerald-500 group-hover/btn:text-white">
                          <Upload className="h-5 w-5" />
                        </div>
                        <div className="text-left">
                          <span className="block text-sm font-black text-slate-900 uppercase">Import PMS Workbook</span>
                          <span className="text-[11px] text-slate-500 font-medium">Synchronize app data with pms.xlsm file.</span>
                        </div>
                      </div>
                      <ChevronRight className="h-5 w-5 text-slate-300 transition-transform group-hover/btn:translate-x-1" />
                    </button>

                    <button onClick={() => { setIsToolsOpen(false); exportReport(); }} className="w-full flex items-center justify-between rounded-2xl border border-slate-200 bg-white p-5 transition-all hover:bg-sky-50 hover:border-sky-200 group/btn">
                      <div className="flex items-center gap-4">
                        <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-sky-100 text-sky-600 transition-colors group-hover/btn:bg-sky-500 group-hover/btn:text-white">
                          <Search className="h-5 w-5" />
                        </div>
                        <div className="text-left">
                          <span className="block text-sm font-black text-slate-900 uppercase">Export Current View</span>
                          <span className="text-[11px] text-slate-500 font-medium">Generate an Excel report from current filters.</span>
                        </div>
                      </div>
                      <ChevronRight className="h-5 w-5 text-slate-300 transition-transform group-hover/btn:translate-x-1" />
                    </button>

                    <button onClick={() => { setIsToolsOpen(false); handleClearAll(); }} className="w-full flex items-center justify-between rounded-2xl border border-rose-100 bg-rose-50/50 p-5 transition-all hover:bg-rose-100 hover:border-rose-200 group/btn">
                      <div className="flex items-center gap-4">
                        <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-rose-100 text-rose-600 transition-colors group-hover/btn:bg-rose-500 group-hover/btn:text-white">
                          <Trash2 className="h-5 w-5" />
                        </div>
                        <div className="text-left">
                          <span className="block text-sm font-black text-rose-700 uppercase">Wipe All Records</span>
                          <span className="text-[11px] text-rose-600 opacity-70 font-medium font-semibold">Permanently delete all project data.</span>
                        </div>
                      </div>
                      <AlertCircle className="h-5 w-5 text-rose-300 transition-transform group-hover/btn:rotate-12" />
                    </button>
                  </div>
                </div>
              ) : (
                <ChangeLog />
              )}
            </div>

            {/* Modal Footer */}
            <div className="p-6 bg-slate-50/80 border-t border-slate-100 flex justify-end shrink-0">
               <button onClick={() => setIsToolsOpen(false)} className="rounded-2xl bg-slate-900 px-8 py-3.5 text-sm font-black text-white shadow-xl shadow-slate-900/20 active:scale-95 transition-all">
                  Close Settings
               </button>
            </div>
          </motion.div>
        </motion.div>
      )}
      </AnimatePresence>

      <AnimatePresence>
        {isDueAlertOpen && visibleDueAlerts.length > 0 && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            transition={MOTION_FADE}
            className="dialog-backdrop fixed inset-0 z-[55] flex items-center justify-center bg-slate-950/75 p-4"
          >
            <motion.div
              initial={{ scale: 0.94, opacity: 0, y: 16 }}
              animate={{ scale: 1, opacity: 1, y: 0 }}
              exit={{ scale: 0.94, opacity: 0, y: 12 }}
              transition={MOTION_PANEL}
              className="dialog-panel w-full max-w-2xl rounded-3xl border border-slate-200 bg-white p-6 shadow-2xl"
            >
              <div className="flex items-start justify-between gap-3 border-b border-slate-100 pb-3">
                <div>
                  <p className="text-[10px] font-black uppercase tracking-[0.2em] text-rose-600">Deadline Notification Engine</p>
                  <h3 className="text-xl font-black tracking-tight text-slate-900 mt-0.5">Due Date Alerts</h3>
                  <p className="text-xs font-medium text-slate-500 mt-1">{visibleDueAlerts.length} stage{visibleDueAlerts.length === 1 ? '' : 's'} are due soon or overdue across {visibleDueAlertGroups.length} project{visibleDueAlertGroups.length === 1 ? '' : 's'}.</p>
                </div>
                <button title="Remind later for all alerts" onClick={() => handleRemindDueAlerts()} className="rounded-xl border border-slate-200 p-2 text-slate-400 hover:bg-slate-50 hover:text-slate-700 transition-all cursor-pointer">
                  <X className="h-4 w-4" />
                </button>
              </div>

              <div className="mt-4 max-h-[50vh] space-y-2.5 dialog-scroll pr-1">
                {slicedLaunchAlertGroups.map(({ groupKey, alerts }) => {
                  const lead = alerts[0];
                  return (
                    <div key={groupKey} className="rounded-2xl border border-slate-200 bg-slate-50/70 p-3.5 space-y-2">
                      <div className="flex flex-wrap items-center justify-between gap-2">
                        <div className="min-w-0">
                          <p className="text-[10px] font-black uppercase tracking-[0.18em] text-slate-400">{lead.projectName || 'UNTITLED PROJECT'}</p>
                          <h4 className="mt-0.5 text-xs font-bold text-slate-900">SR {lead.srNo} | {lead.customer || 'NO CUSTOMER'} | {alerts.length} ALERT{alerts.length === 1 ? '' : 'S'}</h4>
                        </div>
                        <span className={`inline-flex rounded-md border px-2 py-0.5 text-[9px] font-black uppercase ${dueAlertTone(lead.severity)}`}>{formatDueLabel(Math.min(...alerts.map((item) => item.daysLeft)))}</span>
                      </div>
                      <div className="space-y-1.5">
                        {alerts.map((alert) => (
                          <div key={alert.key} className="rounded-xl border border-slate-200 bg-white p-2.5">
                            <div className="flex flex-wrap items-center justify-between gap-2">
                              <div>
                                <p className="text-[10px] font-bold uppercase tracking-[0.14em] text-slate-600">{alert.stageName}</p>
                                <p className="mt-0.5 text-[10px] font-semibold text-slate-400">{alert.deadlineLabel} {alert.deadlineDate}</p>
                              </div>
                              <span className={`inline-flex rounded-md border px-2 py-0.5 text-[9px] font-black uppercase ${dueAlertTone(alert.severity)}`}>{formatDueLabel(alert.daysLeft)}</span>
                            </div>
                          </div>
                        ))}
                      </div>
                      <div className="flex flex-wrap gap-2 pt-1">
                        <button title="Remind later for all stages in this project" onClick={() => handleRemindDueAlerts(alerts.map((alert) => alert.key))} className="rounded-lg border border-slate-200 bg-white px-2.5 py-1 text-[10px] font-bold text-slate-600 hover:bg-slate-100 cursor-pointer">Remind in {alertSnoozeHours}h {alertSnoozeMinutes}m</button>
                        <button title="Acknowledge all stages in this project" onClick={() => handleAcknowledgeDueAlerts(alerts.map((alert) => alert.key))} className="rounded-lg bg-slate-900 px-2.5 py-1 text-[10px] font-bold text-white hover:bg-slate-800 cursor-pointer">Acknowledged</button>
                      </div>
                    </div>
                  );
                })}

                {visibleDueAlertGroups.length > visibleLaunchAlertCount && (
                  <div className="flex justify-center pt-2 pb-1">
                    <button
                      type="button"
                      onClick={() => setVisibleLaunchAlertCount(prev => prev + 15)}
                      className="flex items-center gap-2 rounded-xl border border-rose-200 bg-rose-50 hover:bg-rose-100 px-6 py-2 text-xs font-black text-rose-700 transition-all cursor-pointer"
                    >
                      Load More Alerts ({visibleDueAlertGroups.length - visibleLaunchAlertCount} remaining)
                    </button>
                  </div>
                )}
              </div>

              <div className="mt-4 flex flex-wrap justify-end gap-2 border-t border-slate-100 pt-3">
                <button
                  title="Remind later for all current alerts"
                  onClick={() => handleRemindDueAlerts()}
                  className="rounded-xl border border-slate-200 px-4 py-2 text-xs font-bold text-slate-600 hover:bg-slate-50 cursor-pointer"
                >
                  Remind in {alertSnoozeHours}h {alertSnoozeMinutes > 0 ? `${alertSnoozeMinutes}m` : ''}
                </button>
                <button
                  title="Acknowledge all current alerts"
                  onClick={() => handleAcknowledgeDueAlerts()}
                  className="rounded-xl bg-slate-900 px-5 py-2 text-xs font-black text-white hover:bg-slate-800 cursor-pointer"
                >
                  Acknowledged
                </button>
              </div>

            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      <AnimatePresence>
        {statusModalType && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 z-[60] flex items-center justify-center bg-slate-950/75 p-4 backdrop-blur-md"
          >
            <motion.div
              initial={{ scale: 0.95, opacity: 0, y: 16 }}
              animate={{ scale: 1, opacity: 1, y: 0 }}
              exit={{ scale: 0.95, opacity: 0, y: 12 }}
              className="w-full max-w-xl rounded-3xl border border-slate-200 bg-white shadow-2xl flex flex-col max-h-[85vh] overflow-hidden text-slate-900"
            >
              <div className="p-5 border-b border-slate-100 flex items-center justify-between shrink-0 bg-white">
                <div className="flex items-center gap-3">
                  <div className={`flex h-10 w-10 items-center justify-center rounded-xl text-white shadow-md ${
                    statusModalType === 'Total' ? 'bg-slate-900' :
                    statusModalType === 'Active' ? 'bg-emerald-500' : 
                    statusModalType === 'Alerts' ? 'bg-rose-500' : 'bg-sky-500'
                  }`}>
                    {statusModalType === 'Total' ? <ClipboardList size={20} strokeWidth={2.5} /> : 
                     statusModalType === 'Active' ? <Clock size={20} strokeWidth={2.5} /> : 
                     statusModalType === 'Alerts' ? <ShieldAlert size={20} strokeWidth={2.5} /> : <Check size={20} strokeWidth={2.5} />}
                  </div>
                  <div>
                    <h3 className="text-lg font-black tracking-tight text-slate-900 leading-none">{statusModalType === 'Alerts' ? 'DEADLINE ALERTS' : statusModalType.toUpperCase() + ' PROJECTS'}</h3>
                    <p className="text-[11px] text-slate-500 font-bold uppercase tracking-wider mt-1">
                       {statusModalType === 'Alerts' ? 'Critical Action Required' : 'Matching'} <span className={`${statusModalType === 'Alerts' ? 'text-rose-600' : 'text-slate-900'} font-black`}>{
                      statusModalType === 'Total' ? stats.total :
                      statusModalType === 'Active' ? stats.active : 
                      statusModalType === 'Alerts' ? dueAlertCount : stats.completed
                    } RECORDS</span>
                    </p>
                  </div>
                </div>
                <button onClick={() => setStatusModalType(null)} className="rounded-xl border border-slate-200 p-2 text-slate-400 hover:bg-slate-50 hover:text-slate-700 transition-all cursor-pointer">
                  <X className="h-4 w-4" />
                </button>
              </div>

              {/* Compact Search bar */}
              <div className="px-5 py-3 border-b border-slate-100 bg-slate-50/50 shrink-0">
                <div className="relative">
                  <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-3.5 w-3.5 text-slate-400" />
                  <input
                    type="text"
                    placeholder="Search by name, customer, sr no..."
                    value={modalQuery}
                    onChange={(e) => {
                      setModalQuery(e.target.value);
                      setModalVisibleCount(30);
                    }}
                    className="w-full h-9 pl-9 pr-3 rounded-xl border border-slate-200 bg-white text-xs font-semibold text-slate-800 outline-none focus:border-sky-500 transition-all placeholder:text-slate-400"
                  />
                </div>
              </div>
 
              <div className="flex-1 dialog-scroll custom-scrollbar bg-slate-50/50">
                {(() => {
                  const items = projects.filter(p => {
                    let matchesType = false;
                    if (statusModalType === 'Total') matchesType = true;
                    else if (statusModalType === 'Active') matchesType = p.status !== 'Completed';
                    else if (statusModalType === 'Done') matchesType = p.status === 'Completed';
                    else if (statusModalType === 'Alerts') {
                      matchesType = allDueAlerts.some(a => a.projectId === p.id);
                    }
                    
                    if (!matchesType) return false;

                    const q = modalQuery.toLowerCase().trim();
                    if (!q) return true;

                    const haystack = [p.sr_no, p.customer, p.project_name, p.remarks].join(' ').toLowerCase();
                    return haystack.includes(q);
                  });
 
                  if (items.length === 0) {
                    return (
                      <div className="flex flex-col items-center justify-center py-24 opacity-30 text-slate-400">
                        <Search size={64} strokeWidth={1} className="mb-6" />
                        <p className="text-xl font-black uppercase tracking-widest">No matching records</p>
                      </div>
                    );
                  }
 
                  const slicedItems = items.slice(0, modalVisibleCount);

                  // Group by Customer
                  const groups: Record<string, Project[]> = {};
                  slicedItems.forEach(p => {
                    const c = p.customer || 'Unassigned Customer';
                    if (!groups[c]) groups[c] = [];
                    groups[c].push(p);
                  });
 
                  return (
                    <>
                      {Object.entries(groups)
                        .sort(([a], [b]) => a.localeCompare(b))
                        .map(([customer, groupProjects]) => (
                          <div key={customer} className="mb-6 last:mb-0 text-left">
                            <div className="sticky top-0 bg-white/98 px-10 py-4 z-10 border-b border-slate-200 flex items-center justify-between shadow-sm">
                              <h5 className="text-[12px] font-black uppercase tracking-[0.18em] text-slate-900">
                                 {customer}
                              </h5>
                              <div className="flex items-center gap-3">
                                <span className="text-[10px] font-black text-slate-500 uppercase tracking-widest px-2.5 py-1 rounded-lg bg-slate-100 border border-slate-200">
                                   {groupProjects.length} PROJECTS
                                </span>
                              </div>
                            </div>
                            <div className="grid gap-2.5 px-10 py-6">
                              {groupProjects.map(p => (
                                <button
                                  key={p.id}
                                  onClick={() => {
                                    setSelectedProject(p);
                                    setStatusModalType(null);
                                  }}
                                  className="group flex items-center justify-between p-4.5 rounded-2xl border border-slate-200 bg-white hover:border-slate-400 hover:shadow-xl hover:shadow-slate-200/40 active:scale-[0.99] transition-all text-left"
                                >
                                  <div className="min-w-0 pr-4">
                                    <div className="flex items-center gap-4">
                                      <div className={`h-2.5 w-2.5 rounded-full shadow-sm ${
                                        statusModalType === 'Alerts' ? 'bg-rose-500 animate-pulse' :
                                        p.status === 'Completed' ? 'bg-sky-400' :
                                        p.priority === 'High' ? 'bg-rose-500' : 'bg-emerald-400'
                                      }`} />
                                      <span className="text-[10px] font-bold text-slate-400 tracking-tighter uppercase tabular-nums">SR {normalizeSrNo(p.sr_no)}</span>
                                      <h4 className="text-[14px] font-black text-slate-800 truncate uppercase mt-0.5 tracking-tight group-hover:text-slate-950 transition-colors">{p.project_name}</h4>
                                    </div>
                                    <div className="mt-2 flex items-center gap-4 pl-6.5">
                                      <span className={`text-[9.5px] font-black uppercase tracking-wider px-2.5 py-1 rounded-lg ${
                                        p.priority === 'High' ? 'bg-rose-50 text-rose-600 border border-rose-100' :
                                        p.priority === 'Medium' ? 'bg-amber-50 text-amber-600 border border-amber-100' : 'bg-slate-100 text-slate-500 border border-slate-200'
                                      }`}>
                                        {p.priority}
                                      </span>
                                      <div className="h-4 w-px bg-slate-200" />
                                      <span className="text-[12px] font-black text-slate-600 flex items-center gap-1.5 font-mono">
                                        <ClipboardList size={12} className="text-slate-400" strokeWidth={2.5} />
                                        {p.qty} <span className="opacity-40 text-[9px]">QTY</span>
                                      </span>
                                    </div>
                                  </div>
                                  <div className="h-10 w-10 rounded-2xl bg-slate-50 flex items-center justify-center text-slate-300 group-hover:bg-slate-900 group-hover:text-white transition-all shadow-inner">
                                    <ChevronRight className="h-5 w-5 group-hover:translate-x-0.5 transition-transform" />
                                  </div>
                                </button>
                              ))}
                            </div>
                          </div>
                        ))}

                      {items.length > modalVisibleCount && (
                        <div className="flex justify-center p-8 bg-white border-t border-slate-100 shadow-inner">
                          <button
                            type="button"
                            onClick={() => setModalVisibleCount(prev => prev + 40)}
                            className="flex items-center gap-2 rounded-2xl border border-indigo-200 bg-indigo-50/50 hover:bg-indigo-100/60 px-8 py-3.5 text-xs font-black text-indigo-600 transition-all shadow-md active:scale-95"
                          >
                            Load More Entries ({items.length - modalVisibleCount} remaining)
                          </button>
                        </div>
                      )}
                    </>
                  );
                })()}
              </div>

              <div className="p-10 bg-slate-50/80 border-t border-slate-100 flex justify-end shrink-0 backdrop-blur-md">
                <button
                  onClick={() => setStatusModalType(null)}
                  className="rounded-2xl bg-slate-900 px-10 py-4 text-sm font-black text-white shadow-2xl shadow-slate-900/30 active:scale-95 transition-all hover:bg-slate-800"
                >
                  Return to Dashboard
                </button>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      <AppDialog
        state={dialogState}
        onConfirm={handleDialogConfirm}
        onClose={() => setDialogState(s => ({ ...s, open: false }))}
      />
      <AnimatePresence>
        {isTodoOpen && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            transition={MOTION_FADE}
            className="dialog-backdrop fixed inset-0 z-[60] flex items-center justify-center bg-slate-950/80 p-4"
          >
            <motion.div
              initial={{ scale: 0.95, opacity: 0, y: 24 }}
              animate={{ scale: 1, opacity: 1, y: 0 }}
              exit={{ scale: 0.95, opacity: 0, y: 20 }}
              transition={{ type: 'spring', damping: 30, stiffness: 300, mass: 0.8 }}
              className="dialog-panel w-full max-w-2xl rounded-[44px] border border-white/10 bg-white p-10 shadow-[0_60px_200px_rgba(15,23,42,0.5)] flex flex-col max-h-[90vh]"
            >
              <div className="flex items-center justify-between border-b border-slate-100 pb-6">
                <div className="flex items-center gap-4">
                  <div className="flex h-12 w-12 items-center justify-center rounded-2xl bg-indigo-50 text-indigo-600">
                    <ListChecks className="h-6 w-6" />
                  </div>
                  <div>
                    <h3 className="text-2xl font-bold tracking-tight text-slate-900">Task Command Center</h3>
                    <p className="text-sm text-slate-500">{todos.filter(t => !t.completed).length} active missions remaining</p>
                  </div>
                </div>
                <button onClick={() => setIsTodoOpen(false)} className="rounded-2xl border border-slate-200 p-3 text-slate-400 hover:bg-slate-50 hover:text-slate-600 transition-all">
                  <X className="h-5 w-5" />
                </button>
              </div>

              <div className="mt-8">
                <div className="flex flex-col gap-4 rounded-3xl bg-slate-50 p-4 border border-slate-100">
                  <div className="flex gap-2">
                    <input
                      value={newTodoTask}
                      onChange={(e) => setNewTodoTask(e.target.value)}
                      onKeyDown={(e) => e.key === 'Enter' && addTodo()}
                      placeholder="Capture a new strategic task..."
                      className="flex-1 rounded-2xl border border-slate-200 bg-white px-5 py-3 text-sm outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500"
                    />
                    <button
                      onClick={addTodo}
                      disabled={!newTodoTask.trim()}
                      className={`rounded-2xl px-6 font-semibold text-white shadow-lg transition-all flex items-center gap-2 ${editingTodoId ? 'bg-amber-500 hover:bg-amber-600 shadow-amber-100' : 'bg-indigo-600 hover:bg-indigo-700 shadow-indigo-100'} disabled:opacity-50`}
                    >
                      {editingTodoId ? <Pencil className="h-4 w-4" /> : <Plus className="h-4 w-4" />}
                      {editingTodoId ? 'Update Mission' : 'Add Task'}
                    </button>
                    {editingTodoId && (
                      <button
                        onClick={() => {
                          setEditingTodoId(null);
                          setNewTodoTask('');
                          setNewTodoDueDate('');
                          setNewTodoDueTime('');
                        }}
                        className="rounded-2xl border border-slate-200 px-4 text-slate-400 hover:bg-slate-50 transition-all"
                      >
                        Cancel
                      </button>
                    )}
                  </div>
                  <div className="space-y-3">
                    {/* Row 1: Identity */}
                    <div className="flex gap-2">
                      <DropdownMenu
                        variant="light"
                        value={newTodoPriority}
                        onChange={(v) => setNewTodoPriority(v as any)}
                        title="SET PRIORITY"
                        options={[
                          { value: 'High', label: 'PRIORITY: HIGH' },
                          { value: 'Medium', label: 'PRIORITY: MEDIUM' },
                          { value: 'Low', label: 'PRIORITY: LOW' },
                        ]}
                        className="flex-1 min-w-0"
                        buttonClassName="h-8 w-full rounded-xl bg-white border border-slate-200 px-3 text-xs font-bold text-slate-600 uppercase"
                        icon={<AlertCircle className="h-3.5 w-3.5 shrink-0" />}
                      />
                      <DropdownMenu
                        variant="light"
                        value={newTodoCategory}
                        onChange={(v) => setNewTodoCategory(v as any)}
                        title="SET CATEGORY"
                        options={[
                          { value: 'Production', label: 'PRODUCTION' },
                          { value: 'QC', label: 'QC' },
                          { value: 'Testing', label: 'TESTING' },
                          { value: 'Procurement', label: 'PROCUREMENT' },
                          { value: 'Dispatch', label: 'DISPATCH' },
                          { value: 'Meeting', label: 'MEETING' },
                          { value: 'Misc', label: 'MISC' },
                        ]}
                        className="flex-1 min-w-0"
                        buttonClassName="h-8 w-full rounded-xl bg-white border border-slate-200 px-3 text-xs font-bold text-slate-600 uppercase"
                        icon={<Tag className="h-3.5 w-3.5 shrink-0" />}
                      />
                    </div>
                    {/* Row 2: Timing */}
                    <div className="flex gap-2">
                      <div className="flex-1 flex items-center gap-2 px-3 h-8 rounded-xl bg-white border border-slate-200 min-w-0">
                        <Calendar className="h-3.5 w-3.5 shrink-0 text-slate-400" />
                        <input type="date" value={newTodoDueDate} title="SELECT DUE DATE" onClick={(e) => (e.target as any).showPicker?.()} onChange={(e) => setNewTodoDueDate(e.target.value)} className="flex-1 text-xs font-semibold text-slate-600 bg-transparent outline-none cursor-pointer uppercase min-w-0" />
                      </div>
                      <div className="flex-1 flex items-center gap-2 px-3 h-8 rounded-xl bg-white border border-slate-200 min-w-0">
                        <Clock className="h-3.5 w-3.5 shrink-0 text-slate-400" />
                        <input type="time" value={newTodoDueTime} title="SELECT DUE TIME" onClick={(e) => (e.target as any).showPicker?.()} onChange={(e) => setNewTodoDueTime(e.target.value)} className="flex-1 text-xs font-semibold text-slate-600 bg-transparent outline-none cursor-pointer uppercase min-w-0" />
                      </div>
                    </div>
                  </div>
                </div>
              </div>

              <div 
                onScroll={handleTodoScroll}
                className="mt-8 flex-1 dialog-scroll pr-2 space-y-3 custom-scrollbar overflow-y-auto max-h-[50vh]"
              >
                {todos.length === 0 ? (
                  <div className="flex flex-col items-center justify-center py-12 opacity-40">
                    <ListChecks className="h-16 w-16 mb-4" />
                    <p className="text-lg font-medium">Your command list is clear</p>
                  </div>
                ) : (
                  <>
                    {todos.slice(0, visibleTodoCount).map(todo => (
                      <div key={todo.id} className={`group flex items-center justify-between rounded-2xl p-4 transition-all duration-300 border ${todo.completed ? 'bg-slate-50 border-transparent opacity-60' : 'bg-white border-slate-100 hover:border-indigo-100 hover:shadow-md'}`}>
                        <div className="flex items-center gap-4 flex-1 min-w-0">
                          <button
                            onClick={() => toggleTodo(todo)}
                            className={`flex h-6 w-6 shrink-0 items-center justify-center rounded-lg border transition-all ${todo.completed ? 'bg-indigo-600 border-indigo-600 text-white' : 'border-slate-300 hover:border-indigo-500'}`}
                          >
                            {todo.completed && <Check className="h-4 w-4" />}
                          </button>
                          <div className="min-w-0 flex-1">
                            <p className={`text-sm font-semibold truncate ${todo.completed ? 'text-slate-500 line-through' : 'text-slate-900'}`}>{todo.task.toUpperCase()}</p>
                            <div className="mt-1.5 flex items-center gap-3">
                              <span className={`text-[9px] font-bold uppercase tracking-wider px-2 py-0.5 rounded-lg ${todo.priority === 'High' ? 'bg-rose-50 text-rose-600' :
                                  todo.priority === 'Medium' ? 'bg-amber-50 text-amber-600' : 'bg-slate-50 text-slate-500'
                                }`}>
                                {todo.priority}
                              </span>
                              <span className="text-[9px] font-bold uppercase tracking-wider text-slate-400 bg-white border border-slate-100 px-2 py-0.5 rounded-lg">{todo.category}</span>
                              {todo.due_date && (
                                <span className="flex items-center gap-1 text-[9px] font-bold text-indigo-500 bg-indigo-50/50 px-2 py-0.5 rounded-lg">
                                  <Calendar className="h-2.5 w-2.5" />
                                  {new Date(todo.due_date).toLocaleDateString()}
                                  {todo.due_time && ` @ ${todo.due_time}`}
                                </span>
                              )}
                              {todo.due_time && !todo.completed && (
                                <Bell className="h-2.5 w-2.5 text-amber-500 animate-pulse" />
                              )}
                            </div>
                          </div>
                        </div>
                        <div className="flex items-center gap-2">
                          <button onClick={() => startEditTodo(todo)} className="h-9 w-9 flex items-center justify-center rounded-xl bg-slate-100 text-slate-400 opacity-0 group-hover:opacity-100 hover:bg-amber-50 hover:text-amber-600 transition-all">
                            <Pencil className="h-4 w-4" />
                          </button>
                          <button onClick={() => deleteTodo(todo.id)} className="h-9 w-9 flex items-center justify-center rounded-xl bg-slate-100 text-slate-400 opacity-0 group-hover:opacity-100 hover:bg-rose-50 hover:text-rose-600 transition-all">
                            <Trash2 className="h-4 w-4" />
                          </button>
                        </div>
                      </div>
                    ))}

                    {/* Scroll Pagination Footer Indicator */}
                    {visibleTodoCount < todos.length && (
                      <div className="pt-2 pb-1 text-center">
                        <button
                          type="button"
                          onClick={() => setVisibleTodoCount(prev => Math.min(todos.length, prev + 10))}
                          className="rounded-xl border border-indigo-200 bg-indigo-50 px-4 py-1.5 text-xs font-bold text-indigo-700 hover:bg-indigo-100 transition-all cursor-pointer shadow-sm"
                        >
                          Load More Tasks ({visibleTodoCount} of {todos.length} shown)
                        </button>
                      </div>
                    )}
                  </>
                )}
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* ────── STATS MODAL ────── */}
      <AnimatePresence>
        {isStatsOpen && (() => {
          // ── Computed Stats ──────────────────────────────────────────────
          const activeProjects = projects.filter(p => p.status !== 'Completed');
          const completedProjects = projects.filter(p => p.status === 'Completed');

          // 1. Phase Distribution — count projects by their "deepest active stage"
          const phaseCount: Record<string, number> = {};
          STAGE_STRUCTURE.forEach(s => { phaseCount[s.name] = 0; });
          activeProjects.forEach(p => {
            let deepest: string = STAGE_STRUCTURE[0].name;
            STAGE_STRUCTURE.forEach(s => {
              const st = String(p.stages[`${s.name}_status`] || '').toLowerCase();
              if (st === 'running' || st === 'on hold') deepest = s.name;
            });
            phaseCount[deepest] = (phaseCount[deepest] || 0) + 1;
          });
          const totalPhase = activeProjects.length || 1;

          // 2. Priority Load
          const priorityCounts = {
            High:   projects.filter(p => p.priority === 'High').length,
            Medium: projects.filter(p => p.priority === 'Medium').length,
            Low:    projects.filter(p => p.priority === 'Low').length,
          };
          const priorityTotal = projects.length || 1;

          // 3. Stage Aging — stages Running for too long (no completion date, days since project received)
          type AgingRow = { project: string; stage: string; sr: string; days: number };
          const agingRows: AgingRow[] = [];
          activeProjects.forEach(p => {
            STAGE_STRUCTURE.forEach(s => {
              const st = String(p.stages[`${s.name}_status`] || '').toLowerCase();
              if (st !== 'running') return;
              const completion = p.stages[`${s.name}_Date completion`];
              if (completion) return;
              const received = p.due_date ? parseFlexibleDate(p.due_date) : null;
              if (!received) return;
              const days = Math.floor((Date.now() - received.getTime()) / 86400000);
              if (days > 0) agingRows.push({ project: p.project_name, stage: s.name, sr: p.sr_no, days });
            });
          });
          agingRows.sort((a, b) => b.days - a.days);

          // 4. Overdue by Customer
          const overdueByCustomer: Record<string, number> = {};
          allDueAlerts.forEach(a => {
            const c = a.customer || 'Unknown';
            overdueByCustomer[c] = (overdueByCustomer[c] || 0) + 1;
          });
          const overdueEntries = Object.entries(overdueByCustomer).sort((a, b) => b[1] - a[1]).slice(0, 6);
          const maxOverdue = overdueEntries[0]?.[1] || 1;

          // 5. Avg Lead Time
          const leadTimes: number[] = [];
          completedProjects.forEach(p => {
            const r = p.due_date ? parseFlexibleDate(p.due_date) : null;
            const c = p.completion_date ? parseFlexibleDate(p.completion_date) : null;
            if (r && c && c > r) leadTimes.push((c.getTime() - r.getTime()) / 86400000);
          });
          const avgLeadTime = leadTimes.length ? Math.round(leadTimes.reduce((a, b) => a + b, 0) / leadTimes.length) : null;

          // 6. On-Time Delivery (OTD)
          let onTime = 0;
          completedProjects.forEach(p => {
            const dueVal = p.stages['Dispatch_Due date'] || p.stages['Dispatch/FG_Due date'] || '';
            const doneVal = p.stages['Dispatch_Done date'] || p.stages['Dispatch/FG_Date completion'] || '';
            const due = dueVal ? parseFlexibleDate(dueVal) : null;
            const done = doneVal ? parseFlexibleDate(doneVal) : null;
            if (due && done && done <= due) onTime++;
          });
          const otdPct = completedProjects.length ? Math.round((onTime / completedProjects.length) * 100) : null;

          // 7. Velocity by Phase — avg days between plan date and completion for done stages
          const velocityByPhase: Record<string, number[]> = {};
          STAGE_STRUCTURE.forEach(s => { velocityByPhase[s.name] = []; });
          projects.forEach(p => {
            STAGE_STRUCTURE.forEach(s => {
              const planVal = p.stages[`${s.name}_Due date`] || (s.name === 'SMT' || s.name === 'TH' ? p.stages[`${s.name}_Plan given date`] : '') || '';
              const plan = planVal ? parseFlexibleDate(planVal) : null;
              
              const doneVal = p.stages[`${s.name}_Done date`] || p.stages[`${s.name}_Date completion`] || '';
              const done = doneVal ? parseFlexibleDate(doneVal) : null;
              if (plan && done) {
                const diff = (done.getTime() - plan.getTime()) / 86400000;
                velocityByPhase[s.name].push(diff);
              }
            });
          });
          const velocityEntries = STAGE_STRUCTURE.map(s => ({
            name: s.name,
            avg: velocityByPhase[s.name].length
              ? Math.round(velocityByPhase[s.name].reduce((a, b) => a + b, 0) / velocityByPhase[s.name].length)
              : null,
            count: velocityByPhase[s.name].length,
          })).filter(e => e.avg !== null);
          const maxVelocity = Math.max(...velocityEntries.map(e => Math.abs(e.avg!)), 1);

          // SVG Donut helpers
          const DONUT_R = 52, DONUT_CX = 64, DONUT_CY = 64, DONUT_STROKE = 18;
          const circumference = 2 * Math.PI * DONUT_R;
          // Use SPIS_PALETTE for consistency with ProjectsHub.tsx phase distribution
          const PHASE_COLORS = SPIS_PALETTE;
          let donutOffset = 0;

          return (
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              transition={{ duration: 0.22, ease: [0.16, 1, 0.3, 1] }}
              className="dialog-backdrop fixed inset-0 z-[70] flex items-center justify-center bg-slate-950/80 p-4"
              onClick={() => setIsStatsOpen(false)}
            >
              <motion.div
                initial={{ scale: 0.94, opacity: 0, y: 32 }}
                animate={{ scale: 1, opacity: 1, y: 0 }}
                exit={{ scale: 0.94, opacity: 0, y: 24 }}
              transition={MOTION_PANEL}
                className="dialog-panel w-full max-w-5xl rounded-[44px] border border-slate-200/60 bg-white shadow-[0_80px_240px_rgba(15,23,42,0.45)] flex flex-col max-h-[92vh] overflow-hidden"
                onClick={e => e.stopPropagation()}
              >
                {/* Modal Header */}
                <div className="px-10 pt-10 pb-7 border-b border-slate-100 flex items-center justify-between shrink-0 bg-white">
                  <div className="flex items-center gap-5">
                    <div className="flex h-14 w-14 items-center justify-center rounded-2xl bg-slate-900 text-emerald-400 shadow-xl shadow-slate-900/20">
                      <BarChart2 size={26} strokeWidth={2.5} />
                    </div>
                    <div>
                      <p className="text-[10px] font-black uppercase tracking-[0.3em] text-emerald-600">Production Intelligence</p>
                      <h3 className="text-2xl font-black tracking-tight text-slate-900 leading-none mt-1">Analytics Dashboard</h3>
                    </div>
                  </div>
                  <button onClick={() => setIsStatsOpen(false)} className="rounded-2xl border border-slate-200 p-3.5 text-slate-400 hover:bg-slate-50 hover:text-slate-700 hover:rotate-90 transition-all duration-300">
                    <X className="h-5 w-5" />
                  </button>
                </div>

                {/* Scrollable Content */}
                <div className="dialog-scroll flex-1 bg-slate-50/40 px-10 py-8">
                  <div className="grid grid-cols-1 gap-6">

                    {/* ROW 1: KPI Strip */}
                    <div className="grid grid-cols-4 gap-4">
                      {[
                        { label: 'Total Projects', value: projects.length, color: 'text-slate-800', bg: 'bg-slate-900', icon: <ClipboardList size={16} /> },
                        { label: 'Active', value: activeProjects.length, color: 'text-emerald-700', bg: 'bg-emerald-500', icon: <Clock size={16} /> },
                        { label: 'Completed', value: completedProjects.length, color: 'text-sky-700', bg: 'bg-sky-500', icon: <Check size={16} /> },
                        { label: 'Deadline Alerts', value: allDueAlerts.length, color: 'text-rose-700', bg: 'bg-rose-500', icon: <ShieldAlert size={16} /> },
                      ].map(kpi => (
                        <div key={kpi.label} className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
                          <div className={`inline-flex h-8 w-8 items-center justify-center rounded-xl ${kpi.bg} text-white mb-3`}>{kpi.icon}</div>
                          <p className="text-[10px] font-black uppercase tracking-widest text-slate-400">{kpi.label}</p>
                          <p className={`text-3xl font-black mt-1 ${kpi.color}`}>{kpi.value}</p>
                        </div>
                      ))}
                    </div>

                    {/* ROW 2: Donut + Priority Load */}
                    <div className="grid grid-cols-5 gap-4">
                      {/* Phase Distribution Donut */}
                      <div className="col-span-3 rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
                        <div className="flex items-center mb-5">
                          <p className="text-[11px] font-black uppercase tracking-[0.2em] text-slate-400">Phase Distribution</p>
                          <InfoTip text="Real-time count of projects at their deepest active production stage. Shows you where the physical volume of work is currently located in the factory." />
                        </div>
                        <div className="flex items-center gap-6">
                          <svg width="128" height="128" viewBox="0 0 128 128" className="shrink-0">
                            {Object.entries(phaseCount).map(([phaseName, count], i) => {
                              const pct = count / totalPhase;
                              const dash = circumference * pct;
                              const gap = circumference - dash;
                              const segment = (
                                <circle
                                  key={phaseName}
                                  cx={DONUT_CX} cy={DONUT_CY} r={DONUT_R}
                                  fill="none"
                                  stroke={PHASE_COLORS[i % PHASE_COLORS.length]}
                                  strokeWidth={DONUT_STROKE}
                                  strokeDasharray={`${dash} ${gap}`}
                                  strokeDashoffset={-donutOffset}
                                  strokeLinecap="butt"
                                  style={{ transform: 'rotate(-90deg)', transformOrigin: '64px 64px' }}
                                />
                              );
                              donutOffset += dash;
                              return segment;
                            })}
                            <text x="64" y="60" textAnchor="middle" className="text-xl font-black fill-slate-900" style={{fontSize:'18px', fontWeight:900, fill:'#0f172a'}}>{activeProjects.length}</text>
                            <text x="64" y="76" textAnchor="middle" style={{fontSize:'8px', fontWeight:700, fill:'#94a3b8', letterSpacing:'0.1em', textTransform:'uppercase'}}>ACTIVE</text>
                          </svg>
                          <div className="grid grid-cols-2 gap-x-6 gap-y-2 flex-1 min-w-0">
                            {Object.entries(phaseCount).map(([phaseName, count], i) => (
                              <div key={phaseName} className="flex items-center gap-2 min-w-0">
                                <div className="h-2 w-2 rounded-full shrink-0" style={{ background: PHASE_COLORS[i % PHASE_COLORS.length] }} />
                                <span className="text-[9px] font-bold text-slate-500 truncate uppercase tracking-tight">{phaseName.split('/')[0].trim()}</span>
                                <span className="ml-auto text-[10px] font-black text-slate-800">{count}</span>
                              </div>
                            ))}
                          </div>
                        </div>
                      </div>

                      <div className="col-span-2 rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
                        <div className="flex items-center mb-5">
                          <p className="text-[11px] font-black uppercase tracking-[0.2em] text-slate-400">Priority Load</p>
                          <InfoTip text="Breakdown of the current workload by urgency. Helps factory managers balance the intensity between High-Priority rushes and stable production." />
                        </div>
                        <div className="space-y-4">
                          {[
                            { label: 'High', count: priorityCounts.High, color: 'bg-rose-500', text: 'text-rose-600' },
                            { label: 'Medium', count: priorityCounts.Medium, color: 'bg-amber-400', text: 'text-amber-600' },
                            { label: 'Low', count: priorityCounts.Low, color: 'bg-slate-400', text: 'text-slate-500' },
                          ].map(p => (
                            <div key={p.label}>
                              <div className="flex justify-between mb-1.5">
                                <span className={`text-[10px] font-black uppercase tracking-widest ${p.text}`}>{p.label}</span>
                                <span className="text-[11px] font-black text-slate-700">{p.count} <span className="text-slate-400 font-bold">({Math.round(p.count/priorityTotal*100)}%)</span></span>
                              </div>
                              <div className="h-2 rounded-full bg-slate-100 overflow-hidden">
                                <div className={`h-full rounded-full ${p.color} transition-all duration-700`} style={{ width: `${(p.count/priorityTotal)*100}%` }} />
                              </div>
                            </div>
                          ))}
                        </div>
                        {avgLeadTime !== null && (
                          <div className="mt-5 pt-4 border-t border-slate-100">
                            <p className="text-[10px] font-black uppercase tracking-widest text-slate-400">Avg Lead Time</p>
                            <p className="text-2xl font-black text-slate-900 mt-0.5">{avgLeadTime} <span className="text-sm font-bold text-slate-400">days</span></p>
                          </div>
                        )}
                        {otdPct !== null && (
                          <div className="mt-3">
                            <p className="text-[10px] font-black uppercase tracking-widest text-slate-400">On-Time Delivery</p>
                            <p className={`text-2xl font-black mt-0.5 ${otdPct >= 80 ? 'text-emerald-600' : otdPct >= 50 ? 'text-amber-500' : 'text-rose-600'}`}>{otdPct}%</p>
                          </div>
                        )}
                      </div>
                    </div>

                    {/* ROW 3: Overdue Heatmap + Stage Aging */}
                    <div className="grid grid-cols-2 gap-4">
                      {/* Overdue by Customer */}
                      <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm flex flex-col h-[320px]">
                        <div className="flex items-center mb-5 shrink-0">
                          <p className="text-[11px] font-black uppercase tracking-[0.2em] text-slate-400">Overdue Alerts by Customer</p>
                          <InfoTip text="Counts active stage delays grouped by client. Highlights which user accounts are facing the highest systemic risk of late delivery." />
                        </div>
                        {overdueEntries.length === 0 ? (
                          <div className="flex-1 flex flex-col items-center justify-center text-slate-300">
                            <Check size={32} strokeWidth={1.5} />
                            <p className="text-xs font-bold mt-2 uppercase tracking-widest">No Active Alerts</p>
                          </div>
                        ) : (
                          <div className="flex-1 overflow-auto custom-scrollbar pr-2 space-y-4">
                            {Object.entries(overdueByCustomer)
                              .sort((a, b) => b[1] - a[1])
                              .map(([customer, count]) => (
                                <div key={customer} className="group">
                                  <div className="flex justify-between mb-1.5">
                                    <span className="text-[10px] font-black text-slate-700 uppercase tracking-tight truncate max-w-[70%] group-hover:text-slate-900 transition-colors">{customer}</span>
                                    <span className="text-[10px] font-black text-rose-600 bg-rose-50 px-2 py-0.5 rounded-md border border-rose-100">{count} alert{count > 1 ? 's' : ''}</span>
                                  </div>
                                  <div className="h-1.5 rounded-full bg-slate-100 overflow-hidden">
                                    <div className="h-full rounded-full bg-rose-400 transition-all duration-700 w-0 group-in-view" style={{ width: `${(count/maxOverdue)*100}%` }} />
                                  </div>
                                </div>
                              ))}
                          </div>
                        )}
                      </div>

                       {/* Stage Aging */}
                      <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm flex flex-col h-[320px]">
                        <div className="flex items-center mb-5 shrink-0">
                          <p className="text-[11px] font-black uppercase tracking-[0.2em] text-slate-400">Stage Aging (Running)</p>
                          <InfoTip text="Calculates total days elapsed since the project was received for stages currently in 'Running' status. Red (>14d) signals a critical bottleneck." />
                        </div>
                        {agingRows.length === 0 ? (
                          <div className="flex-1 flex flex-col items-center justify-center text-slate-300">
                             <Check size={32} strokeWidth={1.5} />
                             <p className="text-xs font-bold mt-2 uppercase tracking-widest">Pipeline Clear</p>
                           </div>
                         ) : (
                          <div className="flex-1 overflow-auto custom-scrollbar pr-2 space-y-2">
                            {agingRows.map((row, i) => (
                               <div key={i} className="flex items-center justify-between gap-3 p-2.5 rounded-xl bg-slate-50 border border-slate-100">
                                 <div className="min-w-0 flex-1">
                                   <p className="text-[10px] font-black text-slate-800 truncate uppercase">{row.project}</p>
                                  <p className="text-[9px] font-bold text-slate-400 uppercase tracking-wide mt-0.5">SR {row.sr} · {row.stage}</p>
                                </div>
                                <span className={`text-[10px] font-black shrink-0 px-2 py-1 rounded-lg border ${row.days > 14 ? 'bg-rose-50 text-rose-600 border-rose-100' : row.days > 7 ? 'bg-amber-50 text-amber-600 border-amber-100' : 'bg-slate-100 text-slate-500 border-slate-200'}`}>
                                  {row.days}d
                                </span>
                              </div>
                            ))}
                          </div>
                        )}
                      </div>
                    </div>

                    {/* ROW 4: Accuracy vs. Due Dates */}
                    {velocityEntries.length > 0 && (
                      <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
                        <div className="flex items-center justify-between mb-5">
                          <div className="flex items-center">
                            <p className="text-[11px] font-black uppercase tracking-[0.2em] text-slate-400">Accuracy vs. Due Dates</p>
                            <InfoTip text="Historical performance audit. Measures the average gap between planned targets and actual completion. Left (Green) = Early, Right (Red) = Overdue." />
                          </div>
                          <span className="text-[9px] font-bold text-slate-300 uppercase tracking-widest">Target: Stage Due Dates</span>
                        </div>
                        
                        <div className="relative h-32 flex items-end justify-center gap-6 px-10">
                          {/* Centered Baseline */}
                          <div className="absolute top-[60%] left-10 right-10 h-px bg-slate-100 z-0" />
                          
                          {velocityEntries.map(e => {
                            const isNeg = (e.avg ?? 0) < 0;
                            const barH = Math.max(Math.round((Math.abs(e.avg!) / maxVelocity) * 60), 4);
                            return (
                              <div key={e.name} className="relative z-10 flex flex-col items-center w-14 shrink-0 group" title={`${e.name}: ${e.avg}d avg from Due Date (${e.count} projects)`}>
                                {/* Positive Bar (Late) */}
                                {!isNeg && (
                                  <div 
                                    className={`w-full rounded-t-lg ${(e.avg ?? 0) > 10 ? 'bg-rose-400' : 'bg-sky-400'} transition-all group-hover:brightness-90`}
                                    style={{ height: `${barH}px`, marginBottom: '13px' }}
                                  />
                                )}
                                
                                <span className={`text-[8px] font-black absolute ${isNeg ? 'top-[65%]' : 'bottom-[42%]'} ${isNeg ? 'text-emerald-600' : (e.avg ?? 0) > 10 ? 'text-rose-600' : 'text-slate-500'}`}>
                                  {isNeg ? '' : '+'}{e.avg}D
                                </span>

                                {/* Negative Bar (Early) */}
                                {isNeg && (
                                  <div 
                                    className="w-full rounded-b-lg bg-emerald-400 transition-all group-hover:brightness-90"
                                    style={{ height: `${barH}px`, marginTop: '13px' }}
                                  />
                                )}
                                
                                <span className="absolute -bottom-6 text-[8px] font-bold text-slate-400 uppercase text-center leading-tight w-full truncate">
                                  {e.name.split('/')[0].split(' ')[0]}
                                </span>
                              </div>
                            );
                          })}
                        </div>
                        
                        <div className="mt-10 flex items-center gap-4 text-[9px] font-bold text-slate-400 justify-center">
                          <span className="flex items-center gap-1.5"><span className="h-2 w-2 rounded-full bg-emerald-400 inline-block" /> Early (-D)</span>
                          <span className="flex items-center gap-1.5"><span className="h-2 w-2 rounded-full bg-sky-400 inline-block" /> Late (+D)</span>
                          <span className="flex items-center gap-1.5"><span className="h-2 w-2 rounded-full bg-rose-400 inline-block" /> Critical Late (+10D)</span>
                        </div>
                      </div>
                    )}

                  </div>
                </div>

                {/* Modal Footer */}
                <div className="px-10 py-6 bg-white border-t border-slate-100 flex justify-between items-center shrink-0">
                  <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Based on {projects.length} records · Live data</p>
                  <button onClick={() => setIsStatsOpen(false)} className="rounded-2xl bg-slate-900 px-8 py-3 text-sm font-black text-white shadow-xl shadow-slate-900/20 active:scale-95 transition-all hover:bg-slate-800">
                    Close
                  </button>
                </div>
              </motion.div>
            </motion.div>
          );
        })()}
      </AnimatePresence>

      {/* ────── ULTRA PREMIUM DEVELOPER CREDITS MODAL ────── */}
      <AnimatePresence>
        {isDeveloperModalOpen && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            transition={{ duration: 0.15, ease: 'easeOut' }}
            className="dialog-backdrop fixed inset-0 z-[120] flex items-center justify-center bg-slate-950/75 p-4 backdrop-blur-md"
            onClick={() => setIsDeveloperModalOpen(false)}
          >
            <motion.div
              initial={{ scale: 0.94, opacity: 0, y: 12 }}
              animate={{ scale: 1, opacity: 1, y: 0 }}
              exit={{ scale: 0.94, opacity: 0, y: 8 }}
              transition={{ duration: 0.15, ease: 'easeOut' }}
              onClick={(e) => e.stopPropagation()}
              className="relative w-full max-w-sm overflow-hidden rounded-3xl border border-white/15 bg-slate-900/95 p-7 shadow-2xl text-white will-change-transform"
            >
              <button
                title="Close"
                onClick={() => setIsDeveloperModalOpen(false)}
                className="absolute right-4 top-4 rounded-xl border border-white/10 bg-white/5 p-1.5 text-slate-400 hover:bg-white/10 hover:text-white transition-all cursor-pointer"
              >
                <X className="h-4 w-4" />
              </button>

              <div className="flex flex-col items-center text-center space-y-4">
                {/* App Emblem - Borderless (Requirement 7) */}
                <div className="flex h-16 w-16 items-center justify-center p-1">
                  <img src={appLogo} alt="App Logo" className="h-full w-full object-contain" />
                </div>

                <div className="space-y-1">
                  <h2 className="text-2xl font-black tracking-tight text-white">
                    MEET MISTRY
                  </h2>
                  <p className="text-xs font-bold text-emerald-400 uppercase tracking-wider">
                    Full-Stack Developer
                  </p>
                </div>

                <div className="w-full space-y-2.5 pt-1">
                  <div className="rounded-xl border border-white/10 bg-white/[0.04] p-3.5 text-left space-y-1.5">
                    <div className="flex items-center justify-between text-[10px] font-black uppercase tracking-wider text-slate-400">
                      <span>Product Suite</span>
                      <span className="text-emerald-400 font-extrabold">v1.4 Enterprise Edition</span>
                    </div>
                    <p className="text-[11px] leading-relaxed text-slate-300 font-medium">
                      Designed and engineered with high-throughput Excel synchronization, real-time stage traceability analytics, and desktop-grade workflow controls.
                    </p>
                  </div>

                  <div className="flex flex-wrap items-center justify-center gap-1.5 pt-1">
                    {['React 18', 'TypeScript', 'Tailwind CSS', 'Vite', 'Node.js', 'ExcelJS', 'Framer Motion'].map((tech) => (
                      <span key={tech} className="rounded-lg border border-white/10 bg-white/5 px-2.5 py-0.5 text-[9px] font-extrabold tracking-wider text-slate-300">
                        {tech}
                      </span>
                    ))}
                  </div>
                </div>

                <button
                  type="button"
                  onClick={() => setIsDeveloperModalOpen(false)}
                  className="w-full h-10 rounded-xl bg-gradient-to-r from-emerald-500 to-teal-600 text-xs font-black uppercase tracking-widest text-white shadow-lg shadow-emerald-500/20 hover:brightness-110 active:scale-[0.98] transition-all cursor-pointer"
                >
                  CLOSE
                </button>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>
      {/* ────── PROJECT ANALYTICS & GANTT CHART MODAL ────── */}
      <ProjectAnalyticsModal
        project={analyticsProject}
        onClose={() => setAnalyticsProject(null)}
        stageStructure={STAGE_STRUCTURE}
      />

      {/* ────── CUSTOMER DATABASE MODAL (Requirement 8) ────── */}
      <CustomerModal
        isOpen={isCustomerModalOpen}
        onClose={() => {
          setIsCustomerModalOpen(false);
          setCustomerModalSearch('');
        }}
        projects={projects}
        initialSearch={customerModalSearch}
      />
    </div>
  );
}
