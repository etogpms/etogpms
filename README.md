# 🏗️ MWSS Project Monitoring and Forms Management System (PMFMS)

A comprehensive, state-of-the-art management information system (MIS) designed for the **Engineering and Technical Operations Group (ETOG)** and other key departments of the **Metropolitan Waterworks and Sewerage System (MWSS)**. 

This platform streamlines water infrastructure project tracking, deepwell operations, dam elevation monitoring, document registries, and digitizes standard performance evaluation forms (OPCR/IPCR) with native mobile wrapping.

---

## 🚀 Key Features

### 📊 1. Executive Dashboard & Analytics
* **Real-time Stats:** Track total construction projects, deepwell statuses, and YTD water production at a glance.
* **Environmental Metrics:** View current reservoir levels for **Angat**, **Ipo**, and **La Mesa** dams, including historical high/low records.
* **Interactive GIS Mapping:** Leaflet-based map visualizing water source coordinates across NCR, Cavite, and Rizal.
* **Data Visualization:** Clean, responsive SVG/Canvas-based charts showcasing inflows, net supply, and monthly production by concessionaires (Manila Water / Maynilad).

### 🦺 2. Construction Project Tracker
* Detailed records of water infrastructure, pipeline installations, and facility upgrades.
* Track implementing agencies, contractors, target completion dates, and overall percentage of accomplishment.
* Real-time search, pagination, and multi-criteria filters (Agency, Status, Progress).

### 💧 3. Deepwell Monitoring & Reporting
* Log municipal deepwell operations, rated yields, permit numbers, and operational statuses.
* Generate and export automated Word (`.docx`) and Excel (`.xlsx`) production summaries using templating engines.
* Searchable grid sorting deepwell nodes by production averages.

### 📈 4. Service Update Reports & Treatment Plants
* Log water treatment plant production data against capacity benchmarks.
* Highlight low production anomalies automatically (alerts when production falls below 80% of averages).
* Interactive graphs showing inflows vs. net water supply.

### 📋 5. OPCR & IPCR Digitization
* Digital modules for **Office Performance Commitment and Review (OPCR)** and **Individual Performance Commitment and Review (IPCR)** forms.
* Mathematical formula engines built-in to auto-calculate Strategic, Core, and Support functions.
* Clean inputs, evaluation reviews, and printing templates.

### 🛠️ 6. Office Utilities & Collaboration
* **Calendar of Activities:** Centralized event organizer for construction inspections, public hearings, and meetings.
* **Product Presentations:** Registry for technical presentations and product demos by industry suppliers.
* **Document Registry:** Log, categorize, and query official letters, memoranda, and technical reports.
* **Personal Tasks:** A Kanban-style personal task planner (Work-related, Admin, Personal) with status and priority tags.
* **Real-time Messaging:** In-app chat messaging system enabling cross-department communication.

---

## 🛠️ Tech Stack & Libraries

* **Core UI & Logic:** HTML5, Modern Vanilla JavaScript (ES6+ Modules), CSS3 (Modern Glassmorphic UI / custom components), Bootstrap 5.
* **Mapping/GIS:** Leaflet JS.
* **Charts/Analytics:** Chart.js & React ChartJS wrapper.
* **Document Generation:** `xlsx` (SheetJS), `docxtemplater`, and `pizzip` for word doc template compilation.
* **Mobile Engine:** **CapacitorJS** (allowing deployment to native iOS and Android packages).
* **Database & Hosting:** Firebase Firestore & Firebase Hosting.

---

## 📁 Repository Structure

```text
├── index.html                  # Main application dashboard layout
├── ipcr-form.html              # Individual Performance Review module
├── opcr-form.html              # Office Performance Review module
├── deepwell_summary.html       # Report generator layout
├── capacitor.config.ts         # Capacitor hybrid mobile build settings
├── package.json                # NPM scripts and project dependencies
├── styles.css                  # Classic UI styles
├── style-modern.css            # Sleek, premium dark/glassmorphic interface
├── js/
│   ├── utils.js                # Core utility functions (formatting, dates)
│   ├── docreg.js               # Document Registry engine
│   ├── services/               # Firebase hooks (projects, users, messaging)
│   └── features/               # Modules (calendar, deepwells, ipcr, chat, etc.)
├── tools/                      # Build tools, asset copiers, and scripts
├── www/                        # Output folder compiled for mobile/web deployment
└── assets/                     # Logos, templates, and static resources
```

---

## 💻 Local Development Setup

### Prerequisites
Make sure you have [Node.js](https://nodejs.org/) (v16 or higher) installed.

### 1. Installation
Clone the repository and install the dependencies:
```bash
npm install
```

### 2. Run the Development Server
Launch the local web server to serve the application:
```bash
npm run dev-server
```
This runs the application locally on `http://localhost:8080`.

### 3. Sync Native Mobile Assets (Capacitor)
If you are developing for iOS or Android:
```bash
# Sync web directory assets to native platforms
npm run cap-sync

# Open the native project in Xcode
npx cap open ios
```

---

## 🌐 Deployment

This application is configured for Firebase Hosting. To deploy it live:

```bash
# Compile latest assets (runs tools/copy-assets.js to build www)
npm run build-web

# Deploy to Firebase Hosting
firebase deploy
```

---
*Created and maintained by the Engineering and Technical Operations Group (ETOG), MWSS.*
