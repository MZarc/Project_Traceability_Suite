<div align="center">
  <img src="client/public/favicon.png" width="128" height="128" alt="Project Traceability Suite Logo" />
  <h1>Project Traceability Suite</h1>
  <p><strong>Enterprise-grade desktop traceability suite for real-time Excel synchronization, stage aging analytics, and industrial workflow monitoring.</strong></p>

  <p>
    <a href="#-tech-stack"><img src="https://img.shields.io/badge/Electron-30%2B-47848F?style=for-the-badge&logo=electron&logoColor=white" alt="Electron"></a>
    <a href="#-tech-stack"><img src="https://img.shields.io/badge/React-18.3-61DAFB?style=for-the-badge&logo=react&logoColor=black" alt="React"></a>
    <a href="#-tech-stack"><img src="https://img.shields.io/badge/TypeScript-5.6-3178C6?style=for-the-badge&logo=typescript&logoColor=white" alt="TypeScript"></a>
    <a href="#-tech-stack"><img src="https://img.shields.io/badge/Node.js-Express-339933?style=for-the-badge&logo=node.js&logoColor=white" alt="Node.js"></a>
    <a href="#-tech-stack"><img src="https://img.shields.io/badge/SQLite-3-003B57?style=for-the-badge&logo=sqlite&logoColor=white" alt="SQLite"></a>
    <a href="LICENSE"><img src="https://img.shields.io/badge/License-MIT-emerald?style=for-the-badge" alt="License"></a>
  </p>
</div>

---

## 📌 Overview

**Project Traceability Suite** is a high-performance Windows 11 desktop application designed to solve complex industrial project tracking, stage bottleneck detection, and daily Excel report synchronization. Built on top of **Electron**, **React 18**, **Express**, and **SQLite**, it combines local-first database performance with interactive analytics dashboards and system tray desktop controls.

---

## ✨ Key Features

- 📊 **Real-Time Stage Aging Analytics**: Track running vs. completed stage durations, flag bottleneck projects (>14 days in a single stage), and monitor target due date precision.
- 📁 **High-Throughput Excel Synchronization**: Import, map, and export complex Excel project logs with high-performance async parsing powered by ExcelJS.
- ⚡ **Local-First SQLite Engine**: Enterprise reliability with transactional integrity, fast offline database queries, and automatic background data persistence in `%AppData%`.
- 🖥️ **Windows Desktop Native Integration**: Native frameless windowing, system tray minimizes, progress notifications, and background alert triggers.
- 🎨 **Modern Fluent Light UI**: Engineered with dynamic glassmorphism, responsive grid resizers, stage progress timelines, and custom scrollbars.

---

## 🛠️ Architecture & Tech Stack

```mermaid
graph TD
    subgraph Desktop Shell (Electron Main)
        A[Electron 30+ Main Process] --> B[System Tray & IPC Bridge]
        A --> C[Native Window Manager]
    end

    subgraph User Interface (Client)
        D[React 18 + TypeScript] --> E[Vite Dev & Build Pipeline]
        D --> F[Tailwind CSS v4 + Lucide Icons]
        D --> G[TanStack Table v8 Excel Grid]
    end

    subgraph Backend API Engine
        H[Node.js + Express REST API] --> I[better-sqlite3 Engine]
        H --> J[ExcelJS Import/Export Parser]
        I --> K[(SQLite Local DB in %AppData%)]
    end

    C <-->|Preload IPC| D
    D <-->|HTTP / REST| H
```

| Layer | Technologies |
| :--- | :--- |
| **Desktop Shell** | Electron 30+, Preload IPC, Windows Native Shell Controls |
| **Frontend UI** | React 18, TypeScript, Vite, TanStack Table v8, Tailwind CSS v4, Lucide Icons |
| **Backend Engine** | Node.js, Express, `better-sqlite3`, ExcelJS |
| **Storage** | SQLite 3 (stored locally in `%AppData%` under desktop profile) |

---

## 📁 Repository Structure

```plaintext
project_traceability_system/
├── backend/                  # Express REST API & SQLite database engine
│   ├── package.json          # Backend dependencies (express, better-sqlite3, exceljs)
│   └── server.js             # API routes & database controller
├── client/                   # Modern React 18 frontend dashboard
│   ├── public/               # Static assets & favicon icons
│   ├── src/                  # React components, stage aging tools & stores
│   ├── package.json          # Frontend dependencies (React, Vite, TanStack)
│   └── vite.config.ts        # Vite build configuration
├── electron/                 # Electron main process & desktop wrapper
│   ├── icon.ico / icon.png   # Application branded icons
│   ├── main.js               # Main desktop window lifecycle & IPC registration
│   ├── preload.js            # Secure IPC renderer bridge
│   └── package.json          # Electron builder & packaging settings
├── package.json              # Root monorepo workspace & orchestration scripts
├── .gitignore                # Production root ignore rules
├── LICENSE                   # MIT License
└── README.md                 # Project documentation
```

---

## 🚀 Getting Started

### Prerequisites

- **Node.js**: `v18.0.0` or higher
- **npm**: `v9.0.0` or higher
- **OS**: Windows 10/11 (for desktop installer packaging)

### Installation

1. **Clone the repository**:
   ```bash
   git clone https://github.com/MZarc/Project_Traceability_Suite.git
   cd Project_Traceability_Suite
   ```

2. **Install all workspace dependencies**:
   ```bash
   npm run install:all
   ```

3. **Rebuild native SQLite bindings (if prompted)**:
   ```bash
   npm run rebuild
   ```

---

## 💻 Development & Packaging

### Run in Development Mode
Launches the Client (Vite server), Backend (Express API), and Electron Desktop Shell concurrently:
```bash
npm run dev
```

### Build & Package Executable Installer
Compiles the client web build, synchronizes assets to Electron, and generates a standalone Windows Setup installer in `electron/dist/`:
```bash
npm run package
```

---

## 📋 Workspace Commands Summary

| Command | Action |
| :--- | :--- |
| `npm run dev` | Runs client, backend server, and Electron app concurrently in dev mode |
| `npm run build` | Builds the client React application for production |
| `npm run package` | Builds client assets and packages the standalone Windows Electron executable |
| `npm run install:all` | Installs dependencies across root, client, backend, and electron workspaces |
| `npm run rebuild` | Rebuilds `better-sqlite3` native binaries for local environment |

---

## 🤝 Contributing & License

Developed and maintained by **Meet Mistry** ([MZarc](https://github.com/MZarc)).

Distributed under the **MIT License**. See [`LICENSE`](LICENSE) for details.
