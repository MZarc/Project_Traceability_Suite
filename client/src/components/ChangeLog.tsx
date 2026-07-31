import React from 'react';
import { CHANGELOG_DATA } from '../config/changelog';
import { GitCommit, Milestone, Terminal, ChevronRight, Zap } from 'lucide-react';

const ChangeLog: React.FC = () => {
  return (
    <div className="space-y-10 py-2">
      {CHANGELOG_DATA.map((entry, idx) => (
        <div key={entry.version} className="relative pl-10 group">
          {/* Timeline Connector */}
          {idx !== CHANGELOG_DATA.length - 1 && (
            <div className="absolute left-[19px] top-10 bottom-[-40px] w-px bg-slate-200 group-hover:bg-emerald-200 transition-colors" />
          )}
          
          {/* Timeline Dot */}
          <div className={`absolute left-0 top-1.5 flex h-10 w-10 items-center justify-center rounded-full border-4 border-white shadow-sm transition-transform group-hover:scale-110 ${
            entry.type === 'major' ? 'bg-slate-900 text-white' : 
            entry.type === 'minor' ? 'bg-emerald-500 text-white' : 'bg-sky-500 text-white'
          }`}>
            {entry.type === 'major' ? <Milestone size={16} /> : <Zap size={16} />}
          </div>

          <div className="space-y-4">
            {/* Header */}
            <div>
              <div className="flex items-center gap-3 mb-1">
                <span className={`text-[10px] font-black uppercase tracking-[0.2em] px-2 py-0.5 rounded ${
                  entry.type === 'major' ? 'bg-slate-900 text-white' : 'bg-emerald-100 text-emerald-700'
                }`}>
                  v{entry.version}
                </span>
                <span className="text-[11px] font-bold text-slate-400 uppercase tracking-widest">{entry.date}</span>
              </div>
              <h4 className="text-lg font-black text-slate-900 tracking-tight">{entry.summary}</h4>
            </div>

            {/* Highlights Grid */}
            <div className="grid gap-3 md:grid-cols-1">
              <div className="rounded-2xl border border-slate-100 bg-white p-5 shadow-sm transition-all group-hover:border-emerald-100 group-hover:shadow-md">
                <div className="flex items-center gap-2 mb-4">
                  <div className="flex h-6 w-6 items-center justify-center rounded-lg bg-emerald-50 text-emerald-600">
                    <Milestone size={12} />
                  </div>
                  <span className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">Key Highlights</span>
                </div>
                <ul className="space-y-2.5">
                  {entry.highlights.map((h, i) => (
                    <li key={i} className="flex items-start gap-3 text-[12px] font-semibold text-slate-600 leading-relaxed">
                      <ChevronRight size={14} className="mt-0.5 text-emerald-500 shrink-0" />
                      {h}
                    </li>
                  ))}
                </ul>
              </div>

              {/* Technical Breakdown */}
              <div className="rounded-2xl border border-slate-100 bg-slate-50/50 p-5 shadow-sm transition-all group-hover:bg-slate-50">
                <div className="flex items-center gap-2 mb-4">
                  <div className="flex h-6 w-6 items-center justify-center rounded-lg bg-slate-200 text-slate-600">
                    <Terminal size={12} />
                  </div>
                  <span className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">Technical Breakdown</span>
                </div>
                <ul className="space-y-2.5">
                  {entry.technical.map((t, i) => (
                    <li key={i} className="flex items-start gap-3 text-[11px] font-bold text-slate-500/80 leading-relaxed font-mono">
                      <GitCommit size={14} className="mt-0.5 text-slate-300 shrink-0" />
                      {t}
                    </li>
                  ))}
                </ul>
              </div>
            </div>
          </div>
        </div>
      ))}

      {/* End of Line */}
      <div className="flex items-center justify-center py-6 opacity-20">
        <div className="h-px w-12 bg-slate-300" />
        <span className="mx-4 text-[10px] font-black uppercase tracking-[0.4em] text-slate-400 whitespace-nowrap">End of History</span>
        <div className="h-px w-12 bg-slate-300" />
      </div>
    </div>
  );
};

export default ChangeLog;
