import { useMemo } from 'react';
import { motion, AnimatePresence } from 'framer-motion';
import { X, AlertTriangle, BarChart3, ShieldCheck, Activity, Zap, Layers } from 'lucide-react';

interface StageStructureItem {
  name: string;
  sub: readonly string[];
  display?: string;
}

interface Project {
  id?: number | string;
  sr_no: string;
  customer: string;
  project_name: string;
  qty: number;
  priority: string;
  due_date: string;
  completion_date: string;
  remarks: string;
  status?: string;
  stages: Record<string, any>;
  created_at?: string;
}

interface ProjectAnalyticsModalProps {
  project: Project | null;
  onClose: () => void;
  stageStructure: readonly StageStructureItem[];
}

const MOTION_EASE: [number, number, number, number] = [0.22, 1, 0.36, 1];

function parseDate(dateStr: string | null | undefined): Date | null {
  if (!dateStr || dateStr === '---' || dateStr.trim() === '') return null;
  
  const parts = dateStr.trim().split('-');
  if (parts.length === 3) {
    if (parts[0].length === 4) {
      const d = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
      return isNaN(d.getTime()) ? null : d;
    } else if (parts[2].length === 4) {
      const d = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
      return isNaN(d.getTime()) ? null : d;
    }
  }
  
  const d = new Date(dateStr);
  return isNaN(d.getTime()) ? null : d;
}

function formatDateDisplay(d: Date | null): string {
  if (!d) return 'Not Set';
  return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
}

function getDaysDifference(start: Date, end: Date): number {
  const diffTime = Math.abs(end.getTime() - start.getTime());
  return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
}

export default function ProjectAnalyticsModal({ project, onClose, stageStructure }: ProjectAnalyticsModalProps) {
  if (!project) return null;

  const stageAnalytics = useMemo(() => {
    let completedCount = 0;
    let runningCount = 0;
    let pendingCount = 0;
    let overdueCount = 0;
    let naCount = 0;

    let earliestDate: Date | null = null;
    let latestDate: Date | null = null;

    const stagesData = stageStructure.map((stg) => {
      const statusRaw = String(project.stages[`${stg.name}_status`] || 'Pending').trim();
      const statusLower = statusRaw.toLowerCase();

      let startDateStr = '';
      let dueDateStr = '';
      let doneDateStr = '';

      for (const sub of stg.sub) {
        const key = `${stg.name}_${sub}`;
        const val = String(project.stages[key] || '');
        const subLower = sub.toLowerCase();

        if (subLower.includes('start') || subLower.includes('plan')) {
          startDateStr = val;
        } else if (subLower.includes('due')) {
          dueDateStr = val;
        } else if (subLower.includes('done') || subLower.includes('completion')) {
          doneDateStr = val;
        }
      }

      const startDate = parseDate(startDateStr);
      const dueDate = parseDate(dueDateStr);
      const doneDate = parseDate(doneDateStr);

      [startDate, dueDate, doneDate].forEach((d) => {
        if (d) {
          if (!earliestDate || d < earliestDate) earliestDate = d;
          if (!latestDate || d > latestDate) latestDate = d;
        }
      });

      let durationDays = 0;
      if (startDate && doneDate) {
        durationDays = getDaysDifference(startDate, doneDate);
      } else if (startDate && dueDate) {
        durationDays = getDaysDifference(startDate, dueDate);
      } else if (startDate) {
        durationDays = getDaysDifference(startDate, new Date());
      }

      let isOverdue = false;
      let delayDays = 0;
      if (dueDate) {
        const compareTarget = doneDate || new Date();
        if (compareTarget > dueDate && statusLower !== 'done' && statusLower !== 'n/a' && statusLower !== 'na') {
          isOverdue = true;
          delayDays = getDaysDifference(dueDate, compareTarget);
        }
      }

      let normalizedStatus: 'DONE' | 'RUNNING' | 'PENDING' | 'OVERDUE' | 'N/A' | 'ON HOLD' = 'PENDING';
      if (statusLower === 'done' || statusLower === 'completed') {
        normalizedStatus = 'DONE';
        completedCount++;
      } else if (statusLower === 'n/a' || statusLower === 'na') {
        normalizedStatus = 'N/A';
        naCount++;
      } else if (statusLower === 'running') {
        normalizedStatus = isOverdue ? 'OVERDUE' : 'RUNNING';
        if (isOverdue) overdueCount++; else runningCount++;
      } else if (statusLower === 'on hold') {
        normalizedStatus = 'ON HOLD';
        pendingCount++;
      } else {
        normalizedStatus = isOverdue ? 'OVERDUE' : 'PENDING';
        if (isOverdue) overdueCount++; else pendingCount++;
      }

      return {
        name: stg.name,
        status: normalizedStatus,
        startDate,
        dueDate,
        doneDate,
        durationDays,
        isOverdue,
        delayDays,
        note: String(project.stages[`${stg.name}_note`] || '')
      };
    });

    const activeStagesCount = stageStructure.length - naCount;
    const progressPercent = activeStagesCount > 0 
      ? Math.min(100, Math.round((completedCount / activeStagesCount) * 100))
      : 0;

    // Metric 1: Current Department & Phase Age
    const currentActiveStage = stagesData.find(s => s.status !== 'DONE' && s.status !== 'N/A');
    const currentDepartment = currentActiveStage ? currentActiveStage.name : 'Completed All Stages';

    let phaseAgeDays = 0;
    if (currentActiveStage && currentActiveStage.startDate) {
      phaseAgeDays = getDaysDifference(currentActiveStage.startDate, new Date());
    }

    // Metric 3: Average Stage Velocity
    const completedWithDuration = stagesData.filter(s => s.status === 'DONE' && s.durationDays > 0);
    const avgStageSpeedDays = completedWithDuration.length > 0
      ? (completedWithDuration.reduce((acc, s) => acc + s.durationDays, 0) / completedWithDuration.length).toFixed(1)
      : null;

    return {
      stages: stagesData,
      progressPercent,
      completedCount,
      runningCount,
      pendingCount,
      overdueCount,
      naCount,
      currentDepartment,
      phaseAgeDays,
      avgStageSpeedDays,
      earliestDate: earliestDate || new Date(),
      latestDate: latestDate || new Date()
    };
  }, [project, stageStructure]);

  const totalProjectDays = useMemo(() => {
    const start = stageAnalytics.earliestDate;
    const end = stageAnalytics.latestDate;
    const diff = getDaysDifference(start, end);
    return Math.max(14, diff + 4);
  }, [stageAnalytics]);

  return (
    <AnimatePresence>
      <div className="dialog-backdrop fixed inset-0 z-[5000] flex items-center justify-center bg-slate-950/80 p-4 sm:p-6 backdrop-blur-md overflow-y-auto">
        <motion.div
          initial={{ scale: 0.94, opacity: 0, y: 16 }}
          animate={{ scale: 1, opacity: 1, y: 0 }}
          exit={{ scale: 0.94, opacity: 0, y: 12 }}
          transition={{ duration: 0.2, ease: MOTION_EASE }}
          className="relative w-full max-w-5xl my-auto overflow-hidden rounded-3xl border border-slate-700/60 bg-[#0b1626] shadow-2xl text-slate-100 flex flex-col max-h-[90vh]"
        >
          {/* ────── MODAL HEADER ────── */}
          <div className="relative border-b border-slate-800 bg-gradient-to-r from-[#0d1c30] via-[#0f243d] to-[#0d1c30] px-6 sm:px-8 py-5 shrink-0">
            <div className="flex items-start justify-between gap-4">
              <div className="space-y-1 min-w-0">
                <div className="flex items-center gap-2 flex-wrap">
                  <span className="inline-flex items-center gap-1.5 rounded-full bg-emerald-500/10 px-3 py-0.5 text-[10px] font-black uppercase tracking-widest text-emerald-400 border border-emerald-500/20">
                    <Activity className="h-3 w-3" /> Live Analytics
                  </span>
                  <span className="rounded-md bg-slate-800 px-2.5 py-0.5 text-[10px] font-bold text-slate-400 uppercase tracking-wider">
                    SR: {project.sr_no || 'NA'}
                  </span>
                  <span className="rounded-md bg-sky-500/10 px-2.5 py-0.5 text-[10px] font-bold text-sky-400 uppercase tracking-wider border border-sky-500/20">
                    Qty: {project.qty}
                  </span>
                  <span className="rounded-md bg-amber-500/10 px-2.5 py-0.5 text-[10px] font-bold text-amber-400 uppercase tracking-wider border border-amber-500/20">
                    Priority: {project.priority || 'Normal'}
                  </span>
                </div>
                <h2 className="text-2xl font-black text-white tracking-tight flex items-center gap-3">
                  <span>{project.project_name}</span>
                  <span className="text-sm font-semibold text-slate-400">({project.customer || 'No Customer'})</span>
                </h2>
              </div>

              <button
                type="button"
                onClick={onClose}
                className="rounded-xl border border-white/10 bg-white/5 p-2 text-slate-400 hover:bg-white/10 hover:text-white transition-all cursor-pointer shrink-0"
                title="Close Analytics"
              >
                <X className="h-5 w-5" />
              </button>
            </div>

            {/* ────── EXECUTIVE METRICS BAR ────── */}
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 mt-5">
              {/* Progress Metric */}
              <div className="rounded-2xl border border-white/5 bg-slate-900/60 p-3.5 backdrop-blur-sm">
                <p className="text-[10px] font-bold uppercase tracking-widest text-slate-400">Overall Progress</p>
                <div className="flex items-baseline justify-between mt-1">
                  <span className="text-xl font-black text-emerald-400">{stageAnalytics.progressPercent}%</span>
                  <span className="text-[11px] font-semibold text-slate-400">{stageAnalytics.completedCount}/{stageStructure.length - stageAnalytics.naCount} Stages</span>
                </div>
                <div className="mt-2 h-1.5 w-full overflow-hidden rounded-full bg-slate-800">
                  <div 
                    className="h-full rounded-full bg-gradient-to-r from-emerald-500 to-teal-400 transition-all duration-500" 
                    style={{ width: `${stageAnalytics.progressPercent}%` }}
                  />
                </div>
              </div>

              {/* Metric 1: Current Department & Phase Age */}
              <div className="rounded-2xl border border-white/5 bg-slate-900/60 p-3.5 backdrop-blur-sm">
                <p className="text-[10px] font-bold uppercase tracking-widest text-slate-400 flex items-center gap-1">
                  <Layers className="h-3 w-3 text-sky-400" /> Current Department
                </p>
                <p className="text-sm font-black text-sky-300 mt-1 truncate">
                  {stageAnalytics.currentDepartment}
                </p>
                <p className="text-[10px] font-semibold text-slate-400 mt-0.5">
                  {stageAnalytics.phaseAgeDays > 0 ? `${stageAnalytics.phaseAgeDays} Days in Phase` : 'Stage In-Progress'}
                </p>
              </div>

              {/* Metric 3: Average Stage Speed */}
              <div className="rounded-2xl border border-white/5 bg-slate-900/60 p-3.5 backdrop-blur-sm">
                <p className="text-[10px] font-bold uppercase tracking-widest text-slate-400 flex items-center gap-1">
                  <Zap className="h-3 w-3 text-amber-400" /> Avg Stage Speed
                </p>
                <p className="text-sm font-black text-amber-300 mt-1">
                  {stageAnalytics.avgStageSpeedDays ? `${stageAnalytics.avgStageSpeedDays} Days / Stage` : 'N/A'}
                </p>
                <p className="text-[10px] font-semibold text-slate-400 mt-0.5">
                  Based on {stageAnalytics.completedCount} Completed Stages
                </p>
              </div>

              {/* SLA Health */}
              <div className="rounded-2xl border border-white/5 bg-slate-900/60 p-3.5 backdrop-blur-sm">
                <p className="text-[10px] font-bold uppercase tracking-widest text-slate-400">SLA & Delay Health</p>
                <div className="flex items-center gap-1.5 mt-1.5">
                  {stageAnalytics.overdueCount > 0 ? (
                    <>
                      <AlertTriangle className="h-4 w-4 text-rose-400 shrink-0" />
                      <span className="text-sm font-black text-rose-400">{stageAnalytics.overdueCount} Delayed Stages</span>
                    </>
                  ) : (
                    <>
                      <ShieldCheck className="h-4 w-4 text-emerald-400 shrink-0" />
                      <span className="text-sm font-black text-emerald-400">100% On Schedule</span>
                    </>
                  )}
                </div>
              </div>
            </div>
          </div>

          {/* ────── MODAL BODY: GANTT CHART & STAGE BREAKDOWN ────── */}
          <div className="flex-1 overflow-y-auto p-6 sm:p-8 space-y-6">
            <div className="flex items-center justify-between">
              <div>
                <h3 className="text-base font-bold text-white flex items-center gap-2">
                  <BarChart3 className="h-4 w-4 text-emerald-400" />
                  <span>Manufacturing Stage Gantt Timeline</span>
                </h3>
                <p className="text-xs text-slate-400">Visual schedule breakdown based on recorded start, due, and completion dates.</p>
              </div>
              <div className="flex items-center gap-3 text-[11px] font-semibold text-slate-400">
                <span className="flex items-center gap-1.5"><span className="h-2.5 w-2.5 rounded-full bg-emerald-500" /> Done</span>
                <span className="flex items-center gap-1.5"><span className="h-2.5 w-2.5 rounded-full bg-sky-500" /> Running</span>
                <span className="flex items-center gap-1.5"><span className="h-2.5 w-2.5 rounded-full bg-amber-500" /> Pending</span>
                <span className="flex items-center gap-1.5"><span className="h-2.5 w-2.5 rounded-full bg-rose-500" /> Overdue</span>
              </div>
            </div>

            <div className="space-y-3">
              {stageAnalytics.stages.map((stg, idx) => {
                let badgeBg = 'bg-slate-800 text-slate-400 border-slate-700';
                let barBg = 'bg-slate-800';
                let statusLabel = 'PENDING';

                if (stg.status === 'DONE') {
                  badgeBg = 'bg-emerald-500/10 text-emerald-400 border-emerald-500/20';
                  barBg = 'bg-gradient-to-r from-emerald-600 to-emerald-400';
                  statusLabel = 'DONE';
                } else if (stg.status === 'RUNNING') {
                  badgeBg = 'bg-sky-500/10 text-sky-400 border-sky-500/20';
                  barBg = 'bg-gradient-to-r from-sky-600 to-sky-400';
                  statusLabel = 'RUNNING';
                } else if (stg.status === 'OVERDUE') {
                  badgeBg = 'bg-rose-500/10 text-rose-400 border-rose-500/20';
                  barBg = 'bg-gradient-to-r from-rose-600 to-rose-400';
                  statusLabel = 'OVERDUE';
                } else if (stg.status === 'N/A') {
                  badgeBg = 'bg-slate-800/50 text-slate-500 border-slate-700/50';
                  barBg = 'bg-slate-800/40';
                  statusLabel = 'N/A';
                } else if (stg.status === 'ON HOLD') {
                  badgeBg = 'bg-rose-500/10 text-rose-300 border-rose-500/20';
                  barBg = 'bg-rose-900/60';
                  statusLabel = 'ON HOLD';
                }

                const startOffsetDays = stg.startDate 
                  ? getDaysDifference(stageAnalytics.earliestDate, stg.startDate)
                  : (idx * 2);
                
                const leftPercent = Math.min(85, Math.max(0, (startOffsetDays / totalProjectDays) * 100));
                const widthPercent = Math.min(100 - leftPercent, Math.max(12, ((stg.durationDays || 3) / totalProjectDays) * 100));

                return (
                  <div 
                    key={stg.name} 
                    className="group rounded-2xl border border-slate-800 bg-slate-900/40 p-4 transition-all duration-200 hover:border-slate-700 hover:bg-slate-900/80"
                  >
                    <div className="flex flex-col md:flex-row md:items-center justify-between gap-3">
                      <div className="min-w-[220px] shrink-0 space-y-1">
                        <div className="flex items-center gap-2">
                          <span className="text-xs font-black text-slate-500">#{idx + 1}</span>
                          <span className="text-sm font-bold text-white tracking-tight">{stg.name}</span>
                          <span className={`inline-flex items-center rounded-md border px-2 py-0.5 text-[9px] font-black tracking-wider uppercase ${badgeBg}`}>
                            {statusLabel}
                          </span>
                        </div>

                        <div className="flex items-center gap-3 text-[11px] font-medium text-slate-400">
                          <span>Start: <strong className="text-slate-200">{formatDateDisplay(stg.startDate)}</strong></span>
                          <span>•</span>
                          <span>Due: <strong className="text-slate-200">{formatDateDisplay(stg.dueDate)}</strong></span>
                          {stg.doneDate && (
                            <>
                              <span>•</span>
                              <span>Done: <strong className="text-emerald-300">{formatDateDisplay(stg.doneDate)}</strong></span>
                            </>
                          )}
                        </div>
                      </div>

                      <div className="flex-1 min-w-0 space-y-1">
                        <div className="relative h-6 w-full rounded-xl bg-slate-950/60 p-1 border border-slate-800 overflow-hidden">
                          <div 
                            className={`h-full rounded-lg ${barBg} transition-all duration-500 shadow-md relative flex items-center justify-end px-2`}
                            style={{ 
                              marginLeft: `${leftPercent}%`, 
                              width: `${widthPercent}%` 
                            }}
                          >
                            <span className="text-[9px] font-black text-white drop-shadow whitespace-nowrap">
                              {stg.durationDays > 0 ? `${stg.durationDays}d` : ''}
                            </span>
                          </div>
                        </div>

                        <div className="flex items-center justify-between text-[11px]">
                          <span className="text-slate-400 truncate max-w-[300px]">
                            {stg.note ? `Note: ${stg.note}` : ''}
                          </span>
                          {stg.isOverdue && stg.delayDays > 0 && (
                            <span className="text-rose-400 font-bold flex items-center gap-1">
                              <AlertTriangle className="h-3 w-3" /> Delayed by {stg.delayDays} days
                            </span>
                          )}
                        </div>
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>

          {/* ────── MODAL FOOTER ────── */}
          <div className="border-t border-slate-800 bg-[#0d1c30] px-6 sm:px-8 py-4 flex items-center justify-between shrink-0">
            <p className="text-xs font-semibold text-slate-400">
              Project Remarks: <span className="text-slate-200">{project.remarks || 'No remarks added.'}</span>
            </p>
            <button
              type="button"
              onClick={onClose}
              className="rounded-xl bg-emerald-500 px-6 py-2.5 text-xs font-black text-slate-950 shadow-lg shadow-emerald-500/20 hover:bg-emerald-400 active:scale-95 transition-all cursor-pointer"
            >
              Close Window
            </button>
          </div>
        </motion.div>
      </div>
    </AnimatePresence>
  );
}
