export interface ChangelogEntry {
  version: string;
  date: string;
  summary: string;
  highlights: string[];
  technical: string[];
  type: 'major' | 'minor' | 'patch';
}

export const CHANGELOG_DATA: ChangelogEntry[] = [
  {
    version: '1.3.0',
    date: '2026-07-28',
    type: 'major',
    summary: 'Project Traceability Suite 1.3 - Customer Database Integration & UI Engine Refinement',
    highlights: [
      'Added complete Customer Database system with customer details drawer, search by project name/SR number, and lightweight pagination.',
      'Upgraded Contact Persons system with 2x2 grid cards for Name, Phone, Email, and Designation/Notes.',
      'Refined topbar header controls with compact icon+text buttons for To-Do (with red alert counter), Customers, and Stats.',
      'Enhanced table grid row 1 stage note tooltips with seamless hover bridge and elevated stacking context.',
      'Streamlined Project Details right sidebar drawer with compact typography, 360px width, and optimized stage cards.',
      'Implemented direct customer redirect from table grid rows straight into the Customer Details view modal.',
      'Added full multi-contact formatting for Customer Database Excel sheet exports.'
    ],
    technical: [
      'Integrated SQLite JSON contacts column support with full backwards compatibility.',
      'Added letter-threshold conditional trigger (q.length > 0) for Customer auto-complete suggestions.',
      'Implemented lightweight 5-item project pagination inside Customer Details drawer to prevent DOM lag.',
      'Elevated hovered stage cell stacking context (z-30 / z-40) to prevent row clipping.',
      'Added complete modal state reset on CustomerModal close (!isOpen) to prevent lingering search state.'
    ]
  },
  {
    version: '1.2.0',
    date: '2026-04-23',
    type: 'major',
    summary: 'Industrial Dashboard Refinement & Release 1.2',
    highlights: [
      'Transitioned to official Version 1.2 release.',
      'Implemented high-fidelity "Live Statistics" dashboard modal.',
      'Added interactive "Stage Cards" with real-time status counts.',
      'Refined Industrial Dashboard layout for better column alignment.',
      'Optimized 32-column grid visibility and interaction.'
    ],
    technical: [
      'Synchronized versioning across package.json and Electron build configs.',
      'Implemented dynamic Stats Engine with filter-aware project counting.',
      'Added "Live View" modals for Total, Active, and Overdue projects.',
      'Enhanced responsive scaling for widescreen industrial displays.'
    ]
  },
  {
    version: '1.1.0',
    date: '2026-04-20',
    type: 'minor',
    summary: 'Instant-Paint Splash & Performance Engine',
    highlights: [
      'Added instant-loading splash screen with no white flicker.',
      'Improved application startup speed significantly.',
      'Added system tray "Hide-to-Tray" architecture for background persistence.',
      'Restored high-fidelity branding across the boot sequence.'
    ],
    technical: [
      'Implemented background matching for Electron BrowserWindow to eliminate white paint.',
      'Added mandatory 3-second splash duration for brand consistency.',
      'Implemented Tray icon support with native context menus.',
      'Refactored window control hit-test logic for better responsiveness.'
    ]
  },
  {
    version: '1.0.0',
    date: '2026-04-08',
    type: 'major',
    summary: 'Initial Production Release - Traceability System',
    highlights: [
      'Launched core Project Traceability System.',
      'Implemented full CRUD operations for project management.',
      'Added Excel (PMS workbook) import/export capabilities.',
      'Implemented real-time stage tracking and status badges.'
    ],
    technical: [
      'Built with React (Frontend), Node.js (Backend), and Electron (Desktop Wrapper).',
      'Implemented SQLite database for local-first data persistence.',
      'Created custom 32-column grid engine for high-density data display.',
      'Integrated Lucide-React for industrial-grade iconography.'
    ]
  }
];
