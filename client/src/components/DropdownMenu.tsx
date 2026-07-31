import { useEffect, useRef, useState, ReactNode } from 'react';
import { ChevronDown, Check } from 'lucide-react';
import { motion, AnimatePresence } from 'framer-motion';

type DropdownOption = {
  value: string;
  label: string;
  shortLabel?: string;
};

type DropdownMenuProps = {
  value: string;
  options: DropdownOption[];
  onChange: (value: string) => void;
  placeholder?: string;
  title?: string;
  className?: string;
  buttonClassName?: string;
  menuClassName?: string;
  variant?: 'dark' | 'light';
  icon?: ReactNode;
  initialOpen?: boolean;
  showActiveBadge?: boolean;
};

const variantStyles = {
  dark: {
    button: 'border-white/10 bg-white/[0.06] text-white hover:bg-white/[0.1] hover:border-white/20',
    menu: 'border-slate-700/80 bg-[#0d1626] text-white shadow-[0_16px_48px_rgba(0,0,0,0.7)]',
    option: 'text-slate-400 hover:bg-white/[0.06] hover:text-white',
    optionActive: 'bg-emerald-500/20 text-emerald-300',
    chevron: 'text-slate-400',
    check: 'text-emerald-400',
    divider: 'border-white/5',
  },
  light: {
    button: 'border-slate-200 bg-slate-50 text-slate-800 hover:bg-slate-100 hover:border-slate-300',
    menu: 'border-slate-200 bg-white text-slate-800 shadow-[0_12px_40px_rgba(15,23,42,0.2)]',
    option: 'text-slate-600 hover:bg-slate-50 hover:text-slate-900',
    optionActive: 'bg-emerald-50 text-emerald-700',
    chevron: 'text-slate-400',
    check: 'text-emerald-600',
    divider: 'border-slate-100',
  },
} as const;

// Single source of truth for the typography style
const typographyStyle = {
  fontSize: '10px',
  fontWeight: '700',
  letterSpacing: '0.08em',
  textTransform: 'uppercase' as const,
  fontFamily: '"Montserrat", sans-serif'
};

export default function DropdownMenu({
  value,
  options,
  onChange,
  placeholder = 'Select',
  title,
  className = '',
  buttonClassName = '',
  menuClassName = '',
  variant = 'light',
  icon,
  initialOpen = false,
  showActiveBadge = false,
}: DropdownMenuProps) {
  const [open, setOpen] = useState(initialOpen);
  const rootRef = useRef<HTMLDivElement>(null);
  const styles = variantStyles[variant];
  const selected = options.find((option) => option.value === value);

  useEffect(() => {
    setOpen(initialOpen);
  }, [initialOpen]);

  const displayLabel = selected
    ? (selected.shortLabel ?? selected.label)
    : placeholder;

  useEffect(() => {
    const handlePointerDown = (event: MouseEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) {
        setOpen(false);
      }
    };
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') setOpen(false);
    };
    document.addEventListener('click', handlePointerDown);
    document.addEventListener('keydown', handleKeyDown);
    return () => {
      document.removeEventListener('click', handlePointerDown);
      document.removeEventListener('keydown', handleKeyDown);
    };
  }, []);

  return (
    <div
      ref={rootRef}
      className={`relative ${className}`}
      style={{ overflow: 'visible' }}
      onClick={(e) => e.stopPropagation()}
      onDoubleClick={(e) => e.stopPropagation()}
    >
      <button
        type="button"
        onClick={() => setOpen((current) => !current)}
        title={title || selected?.label || placeholder}
        aria-expanded={open}
        className={`flex w-full items-center justify-between gap-1.5 rounded-lg border outline-none transition-all cursor-pointer ${styles.button} ${buttonClassName}`}
        style={{ ...typographyStyle, height: '30px', padding: '0 10px' }}
      >
        <div className="flex items-center gap-2 min-w-0 flex-1">
          {icon && <span className="shrink-0 opacity-70">{icon}</span>}
          <span className="truncate">{displayLabel}</span>
          {showActiveBadge && selected?.value !== 'All' && selected?.value !== 'None' && (
            <span className={`ml-auto text-[8px] px-1.5 py-0.5 rounded-md ${variant === 'dark' ? 'bg-white/10 text-white/60' : 'bg-slate-900/10 text-slate-500'}`}>
              ACTIVE
            </span>
          )}
        </div>
        <ChevronDown
          className={`h-3 w-3 shrink-0 transition-transform duration-200 ${styles.chevron} ${open ? 'rotate-180' : ''}`}
        />
      </button>

      <AnimatePresence>
        {open && (
          <motion.div
            initial={{ opacity: 0, scale: 0.96, y: -4 }}
            animate={{ opacity: 1, scale: 1, y: 0 }}
            exit={{ opacity: 0, scale: 0.96, y: -4 }}
            transition={{ duration: 0.1, ease: 'easeOut' }}
            className={`dropdown-panel absolute left-0 top-full z-[9999] mt-1 w-full rounded-lg border p-1 ${styles.menu} ${menuClassName} will-change-transform`}
            style={{ 
              boxShadow: variant === 'dark' ? '0 16px 48px rgba(0,0,0,0.7), 0 0 0 1px rgba(255,255,255,0.05) inset' : '0 16px 48px rgba(15,23,42,0.25), 0 0 0 1px rgba(15,23,42,0.08)',
              backgroundColor: variant === 'dark' ? '#0d1626' : '#ffffff',
              opacity: 1,
            }}
          >
            <div className="max-h-48 overflow-y-auto px-0.5 custom-dropdown-scroll">
              {options.map((option) => {
                const active = option.value === value;
                return (
                  <button
                    key={option.value}
                    type="button"
                    onClick={() => {
                      onChange(option.value);
                      setOpen(false);
                    }}
                    title={option.label}
                    className={`flex w-full items-center justify-between gap-2 rounded-md px-2.5 py-1.5 text-left transition-colors cursor-pointer ${
                      active ? styles.optionActive : styles.option
                    }`}
                    style={typographyStyle}
                  >
                    <span className="truncate flex-1">{option.label}</span>
                    {active && <Check className={`h-2.5 w-2.5 shrink-0 ${styles.check}`} />}
                  </button>
                );
              })}
            </div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}
