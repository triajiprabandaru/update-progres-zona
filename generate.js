#!/usr/bin/env node
/**
 * Dashboard Progres Kavling Pemakaman 25HA - Generator
 * 
 * Cara pakai:
 *   node generate.js                    → baca data.xlsx, output dashboard.html
 *   node generate.js input.xlsx         → baca input.xlsx, output dashboard.html
 *   node generate.js input.xlsx out.html → baca input.xlsx, output out.html
 */

const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// =====================================================================
// CONFIG
// =====================================================================
const args = process.argv.slice(2);
const INPUT_FILE = args[0] || 'data.xlsx';
const OUTPUT_FILE = args[1] || 'dashboard.html';

// =====================================================================
// UTILITY FUNCTIONS
// =====================================================================
function excelDateToStr(serial) {
  if (!serial || typeof serial !== 'number' || serial < 1000) return '';
  const d = new Date((serial - 25569) * 86400000);
  const dd = d.getUTCDate().toString().padStart(2, '0');
  const mm = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'][d.getUTCMonth()];
  return `${dd} ${mm} ${d.getUTCFullYear()}`;
}

function cellVal(ws, r, c) {
  const addr = XLSX.utils.encode_cell({ r, c });
  const cell = ws[addr];
  return cell ? cell.v : null;
}

function num(v) { return (v === null || v === undefined || v === '') ? 0 : Number(v) || 0; }
function str(v) { return (v === null || v === undefined) ? '' : String(v).trim(); }

// =====================================================================
// DATA PROCESSING
// =====================================================================
function processProgresZona(ws) {
  const data = [];
  for (let r = 5; r <= 46; r++) {
    const zona = str(cellVal(ws, r, 1));
    if (!zona) continue;
    data.push({
      zona,
      type: str(cellVal(ws, r, 2)),
      accMakam: str(cellVal(ws, r, 3)),
      stokRencana: num(cellVal(ws, r, 4)),
      stokSPK: num(cellVal(ws, r, 5)),
      terjual: num(cellVal(ws, r, 6)),
      sisaStok: num(cellVal(ws, r, 7)),
      funeralReady: num(cellVal(ws, r, 8)),
      pctReady: num(cellVal(ws, r, 9)),
      status: str(cellVal(ws, r, 10)),
      kontraktor: str(cellVal(ws, r, 11)),
      tglMulai: excelDateToStr(cellVal(ws, r, 12)),
      tglSelesai: excelDateToStr(cellVal(ws, r, 13)),
      aktualProgress: num(cellVal(ws, r, 14)),
      deviasi: cellVal(ws, r, 15),
      keterangan: str(cellVal(ws, r, 16)),
    });
  }
  return data;
}

function processTotalByType(ws) {
  // Detail type (rows 4-16 = index 3-15)
  const detail = [];
  for (let r = 3; r <= 15; r++) {
    const code = str(cellVal(ws, r, 0));
    if (!code) continue;
    detail.push({
      code,
      name: str(cellVal(ws, r, 1)),
      stokRencana: num(cellVal(ws, r, 2)),
      stokSPK: num(cellVal(ws, r, 3)),
      terjual: num(cellVal(ws, r, 4)),
      sisaStok: num(cellVal(ws, r, 5)),
      funeralReady: num(cellVal(ws, r, 6)),
      pct: num(cellVal(ws, r, 7)),
    });
  }

  // Summary category (rows 21-26 = index 20-25)
  const summary = [];
  for (let r = 20; r <= 25; r++) {
    const type = str(cellVal(ws, r, 0));
    if (!type) continue;
    summary.push({
      type,
      stokRencana: num(cellVal(ws, r, 1)),
      stokSPK: num(cellVal(ws, r, 2)),
      terjual: num(cellVal(ws, r, 3)),
      sisaStok: num(cellVal(ws, r, 4)),
      funeralReady: num(cellVal(ws, r, 5)),
      pct: num(cellVal(ws, r, 6)),
      target: num(cellVal(ws, r, 7)),
      selisih: num(cellVal(ws, r, 8)),
    });
  }
  return { detail, summary };
}

function processSiteplan(ws) {
  // 3 report columns: 25 Sep, 15 Sep, 8 Sep
  const reports = [];
  const dates = [
    { col: 2, dateStr: str(cellVal(ws, 2, 2)), fotoCol: 4, statusCol: 5, zonaCol: 3 },
    { col: 7, dateStr: str(cellVal(ws, 2, 7)), fotoCol: 9, statusCol: 10, zonaCol: 8 },
    { col: 12, dateStr: str(cellVal(ws, 2, 12)) || '8 September 2026', fotoCol: 14, statusCol: 15, zonaCol: 13 },
  ];

  for (const d of dates) {
    const items = [];
    for (let r = 5; r <= 17; r++) {
      const zona = str(cellVal(ws, r, d.zonaCol));
      if (!zona) continue;
      items.push({
        zona,
        foto: str(cellVal(ws, r, d.fotoCol)),
        status: str(cellVal(ws, r, d.statusCol)),
      });
    }
    if (items.length) {
      reports.push({ date: d.dateStr, items });
    }
  }
  return reports;
}

function processChangelog(ws) {
  const entries = [];
  for (let r = 3; r <= 30; r++) {
    const date = cellVal(ws, r, 0);
    const desc = str(cellVal(ws, r, 1));
    if (!desc) continue;
    entries.push({ date: str(date), desc });
  }
  return entries;
}

function calculateSummary(zones) {
  const s = {
    totalRencana: 0, totalSPK: 0, totalTerjual: 0,
    totalSisa: 0, totalReady: 0,
    statusCounts: {}, zonaProgress: [],
  };

  const zonaMap = new Map();
  for (const z of zones) {
    s.totalRencana += z.stokRencana;
    s.totalSPK += z.stokSPK;
    s.totalTerjual += z.terjual;
    s.totalSisa += z.sisaStok;
    s.totalReady += z.funeralReady;

    const statusKey = z.status.startsWith('2') ? 'tender'
      : z.status.startsWith('5') ? 'pengerjaan'
      : z.status.startsWith('6') ? 'selesai' : 'lainnya';
    s.statusCounts[statusKey] = (s.statusCounts[statusKey] || 0) + 1;

    if (!zonaMap.has(z.zona)) {
      zonaMap.set(z.zona, {
        zona: z.zona,
        status: z.status,
        aktualProgress: z.aktualProgress,
        deviasi: z.deviasi,
        kontraktor: z.kontraktor,
        totalRencana: 0, totalSPK: 0, totalTerjual: 0, totalReady: 0,
      });
    }
    const zm = zonaMap.get(z.zona);
    zm.totalRencana += z.stokRencana;
    zm.totalSPK += z.stokSPK;
    zm.totalTerjual += z.terjual;
    zm.totalReady += z.funeralReady;
  }

  s.pctReady = s.totalSPK > 0 ? ((s.totalReady / s.totalSPK) * 100).toFixed(1) : '0';
  s.zonaProgress = Array.from(zonaMap.values())
    .filter(z => z.status.startsWith('5') || z.status.startsWith('6'))
    .sort((a, b) => b.aktualProgress - a.aktualProgress);

  return s;
}

// =====================================================================
// HTML GENERATION
// =====================================================================
function generateHTML(data) {
  const { zones, types, siteplan, changelog, summary, reportDate } = data;

  const zonesJSON = JSON.stringify(zones);
  const typeSummaryJSON = JSON.stringify(types.summary);
  const zonaProgressJSON = JSON.stringify(summary.zonaProgress);
  const siteplanJSON = JSON.stringify(siteplan);
  const changelogJSON = JSON.stringify(changelog);

  const spkPctOfRencana = summary.totalRencana > 0 
    ? ((summary.totalSPK / summary.totalRencana) * 100).toFixed(1) 
    : '0';
  const terjualPctOfSPK = summary.totalSPK > 0 
    ? ((summary.totalTerjual / summary.totalSPK) * 100).toFixed(1) 
    : '0';

  return `<!DOCTYPE html>
<html lang="id" class="dark">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Dashboard Progres Kavling Pemakaman 25HA</title>

    <!-- Open Sans Font -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link
        href="https://fonts.googleapis.com/css2?family=Open+Sans:ital,wght@0,300..800;1,300..800&family=JetBrains+Mono:wght@400;500;600;700&display=swap"
        rel="stylesheet">

    <!-- Tailwind CSS v4 Browser CDN -->
    <script src="https://unpkg.com/@tailwindcss/browser@4"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>

    <style type="text/tailwindcss">
        @import "tailwindcss";

@custom-variant dark (&:is(.dark *));

:root {
  --background: oklch(1.0000 0 0);
  --foreground: oklch(0.1884 0.0128 248.5103);
  --card: oklch(0.9784 0.0011 197.1387);
  --card-foreground: oklch(0.1884 0.0128 248.5103);
  --popover: oklch(1.0000 0 0);
  --popover-foreground: oklch(0.1884 0.0128 248.5103);
  --primary: oklch(0.6723 0.1606 244.9955);
  --primary-foreground: oklch(1.0000 0 0);
  --secondary: oklch(0.1884 0.0128 248.5103);
  --secondary-foreground: oklch(1.0000 0 0);
  --muted: oklch(0.9222 0.0013 286.3737);
  --muted-foreground: oklch(0.48 0.0128 248.51);
  --accent: oklch(0.9392 0.0166 250.8453);
  --accent-foreground: oklch(0.6723 0.1606 244.9955);
  --destructive: oklch(0.6188 0.2376 25.7658);
  --destructive-foreground: oklch(1.0000 0 0);
  --border: oklch(83.902% 0.0001 271.152);
  --input: oklch(0.9809 0.0025 228.7836);
  --ring: oklch(0.6818 0.1584 243.3540);
  --chart-1: oklch(0.6723 0.1606 244.9955);
  --chart-2: oklch(0.6907 0.1554 160.3454);
  --chart-3: oklch(0.8214 0.1600 82.5337);
  --chart-4: oklch(0.7064 0.1822 151.7125);
  --chart-5: oklch(0.5919 0.2186 10.5826);
  --sidebar: oklch(0.9784 0.0011 197.1387);
  --sidebar-foreground: oklch(0.1884 0.0128 248.5103);
  --sidebar-primary: oklch(0.6723 0.1606 244.9955);
  --sidebar-primary-foreground: oklch(1.0000 0 0);
  --sidebar-accent: oklch(0.9392 0.0166 250.8453);
  --sidebar-accent-foreground: oklch(0.6723 0.1606 244.9955);
  --sidebar-border: oklch(0.9271 0.0101 238.5177);
  --sidebar-ring: oklch(0.6818 0.1584 243.3540);
  --font-sans: 'Open Sans', system-ui, sans-serif;
  --font-serif: Georgia, serif;
  --font-mono: 'JetBrains Mono', Menlo, monospace;
  --radius: 1.3rem;
  --shadow-x: 0px;
  --shadow-y: 2px;
  --shadow-blur: 0px;
  --shadow-spread: 0px;
  --shadow-opacity: 0;
  --shadow-color: rgba(29,161,242,0.15);
  --shadow-2xs: 0px 2px 0px 0px hsl(202.8169 89.1213% 53.1373% / 0.00);
  --shadow-xs: 0px 2px 0px 0px hsl(202.8169 89.1213% 53.1373% / 0.00);
  --shadow-sm: 0px 2px 0px 0px hsl(202.8169 89.1213% 53.1373% / 0.00), 0px 1px 2px -1px hsl(202.8169 89.1213% 53.1373% / 0.00);
  --shadow: 0px 2px 0px 0px hsl(202.8169 89.1213% 53.1373% / 0.00), 0px 1px 2px -1px hsl(202.8169 89.1213% 53.1373% / 0.00);
  --shadow-md: 0px 2px 0px 0px hsl(202.8169 89.1213% 53.1373% / 0.00), 0px 2px 4px -1px hsl(202.8169 89.1213% 53.1373% / 0.00);
  --shadow-lg: 0px 2px 0px 0px hsl(202.8169 89.1213% 53.1373% / 0.00), 0px 4px 6px -1px hsl(202.8169 89.1213% 53.1373% / 0.00);
  --shadow-xl: 0px 2px 0px 0px hsl(202.8169 89.1213% 53.1373% / 0.00), 0px 8px 10px -1px hsl(202.8169 89.1213% 53.1373% / 0.00);
  --shadow-2xl: 0px 2px 0px 0px hsl(202.8169 89.1213% 53.1373% / 0.00);
  --tracking-normal: 0em;
  --spacing: 0.25rem;
}

.dark {
  --background: oklch(0 0 0);
  --foreground: oklch(0.9328 0.0025 228.7857);
  --card: oklch(0.2097 0.0080 274.5332);
  --card-foreground: oklch(0.8853 0 0);
  --popover: oklch(0 0 0);
  --popover-foreground: oklch(0.9328 0.0025 228.7857);
  --primary: oklch(0.6692 0.1607 245.0110);
  --primary-foreground: oklch(1.0000 0 0);
  --secondary: oklch(0.9622 0.0035 219.5331);
  --secondary-foreground: oklch(0.1884 0.0128 248.5103);
  --muted: oklch(0.2090 0 0);
  --muted-foreground: oklch(0.5637 0.0078 247.9662);
  --accent: oklch(0.1928 0.0331 242.5459);
  --accent-foreground: oklch(0.6692 0.1607 245.0110);
  --destructive: oklch(0.6188 0.2376 25.7658);
  --destructive-foreground: oklch(1.0000 0 0);
  --border: oklch(0.2674 0.0047 248.0045);
  --input: oklch(0.3020 0.0288 244.8244);
  --ring: oklch(0.6818 0.1584 243.3540);
  --chart-1: oklch(0.6723 0.1606 244.9955);
  --chart-2: oklch(0.6907 0.1554 160.3454);
  --chart-3: oklch(0.8214 0.1600 82.5337);
  --chart-4: oklch(0.7064 0.1822 151.7125);
  --chart-5: oklch(0.5919 0.2186 10.5826);
  --sidebar: oklch(0.2097 0.0080 274.5332);
  --sidebar-foreground: oklch(0.8853 0 0);
  --sidebar-primary: oklch(0.6818 0.1584 243.3540);
  --sidebar-primary-foreground: oklch(1.0000 0 0);
  --sidebar-accent: oklch(0.1928 0.0331 242.5459);
  --sidebar-accent-foreground: oklch(0.6692 0.1607 245.0110);
  --sidebar-border: oklch(0.3795 0.0220 240.5943);
  --sidebar-ring: oklch(0.6818 0.1584 243.3540);
  --font-sans: 'Open Sans', system-ui, sans-serif;
  --font-serif: Georgia, serif;
  --font-mono: 'JetBrains Mono', Menlo, monospace;
  --radius: 1.3rem;
  --shadow-x: 0px;
  --shadow-y: 2px;
  --shadow-blur: 0px;
  --shadow-spread: 0px;
  --shadow-opacity: 0;
  --shadow-color: rgba(29,161,242,0.25);
  --shadow-2xs: 0px 2px 0px 0px hsl(202.8169 89.1213% 53.1373% / 0.00);
  --shadow-xs: 0px 2px 0px 0px hsl(202.8169 89.1213% 53.1373% / 0.00);
  --shadow-sm: 0px 2px 0px 0px hsl(202.8169 89.1213% 53.1373% / 0.00), 0px 1px 2px -1px hsl(202.8169 89.1213% 53.1373% / 0.00);
  --shadow: 0px 2px 0px 0px hsl(202.8169 89.1213% 53.1373% / 0.00), 0px 1px 2px -1px hsl(202.8169 89.1213% 53.1373% / 0.00);
  --shadow-md: 0px 2px 0px 0px hsl(202.8169 89.1213% 53.1373% / 0.00), 0px 2px 4px -1px hsl(202.8169 89.1213% 53.1373% / 0.00);
  --shadow-lg: 0px 2px 0px 0px hsl(202.8169 89.1213% 53.1373% / 0.00), 0px 4px 6px -1px hsl(202.8169 89.1213% 53.1373% / 0.00);
  --shadow-xl: 0px 2px 0px 0px hsl(202.8169 89.1213% 53.1373% / 0.00), 0px 8px 10px -1px hsl(202.8169 89.1213% 53.1373% / 0.00);
  --shadow-2xl: 0px 2px 0px 0px hsl(202.8169 89.1213% 53.1373% / 0.00);
}

@theme inline {
  --color-background: var(--background);
  --color-foreground: var(--foreground);
  --color-card: var(--card);
  --color-card-foreground: var(--card-foreground);
  --color-popover: var(--popover);
  --color-popover-foreground: var(--popover-foreground);
  --color-primary: var(--primary);
  --color-primary-foreground: var(--primary-foreground);
  --color-secondary: var(--secondary);
  --color-secondary-foreground: var(--secondary-foreground);
  --color-muted: var(--muted);
  --color-muted-foreground: var(--muted-foreground);
  --color-accent: var(--accent);
  --color-accent-foreground: var(--accent-foreground);
  --color-destructive: var(--destructive);
  --color-destructive-foreground: var(--destructive-foreground);
  --color-border: var(--border);
  --color-input: var(--input);
  --color-ring: var(--ring);
  --color-chart-1: var(--chart-1);
  --color-chart-2: var(--chart-2);
  --color-chart-3: var(--chart-3);
  --color-chart-4: var(--chart-4);
  --color-chart-5: var(--chart-5);
  --color-sidebar: var(--sidebar);
  --color-sidebar-foreground: var(--sidebar-foreground);
  --color-sidebar-primary: var(--sidebar-primary);
  --color-sidebar-primary-foreground: var(--sidebar-primary-foreground);
  --color-sidebar-accent: var(--sidebar-accent);
  --color-sidebar-accent-foreground: var(--sidebar-accent-foreground);
  --color-sidebar-border: var(--sidebar-border);
  --color-sidebar-ring: var(--sidebar-ring);

  --font-sans: var(--font-sans);
  --font-mono: var(--font-mono);
  --font-serif: var(--font-serif);

  --radius-sm: calc(var(--radius) - 6px);
  --radius-md: calc(var(--radius) - 3px);
  --radius-lg: var(--radius);
  --radius-xl: calc(var(--radius) + 4px);
  --radius-2xl: calc(var(--radius) + 8px);

  --shadow-2xs: var(--shadow-2xs);
  --shadow-xs: var(--shadow-xs);
  --shadow-sm: var(--shadow-sm);
  --shadow: var(--shadow);
  --shadow-md: var(--shadow-md);
  --shadow-lg: var(--shadow-lg);
  --shadow-xl: var(--shadow-xl);
  --shadow-2xl: var(--shadow-2xl);
}

@layer base {
  * {
    border-color: var(--border);
  }
  body {
    background-color: var(--background);
    color: var(--foreground);
    font-family: var(--font-sans);
    letter-spacing: var(--tracking-normal);
  }
}

.scrollbar-thin::-webkit-scrollbar { width: 6px; height: 6px; }
.scrollbar-thin::-webkit-scrollbar-thumb { background: var(--border); border-radius: 9999px; }
.progress-bar { transition: width 0.7s cubic-bezier(0.16, 1, 0.3, 1); }
tr.zona-row { transition: background-color 0.15s ease; }
tr.zona-row:hover { background-color: var(--accent); }
.tab-btn.active {
  background-color: var(--primary);
  color: var(--primary-foreground);
  border-color: var(--primary);
}
</style>
</head>

<body class="min-h-screen antialiased selection:bg-primary selection:text-primary-foreground">

    <!-- ============ HEADER ============ -->
    <header class="border-b border-border bg-card/80 backdrop-blur-md sticky top-0 z-30">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-4">
            <div class="flex items-center justify-between gap-4">
                <div class="flex items-center gap-3.5">
                    <div
                        class="w-11 h-11 rounded-[var(--radius)] bg-primary text-primary-foreground flex items-center justify-center shadow-sm">
                        <svg xmlns="http://www.w3.org/2000/svg" width="22" height="22" viewBox="0 0 24 24" fill="none"
                            stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                            <line x1="3" x2="21" y1="22" y2="22" />
                            <line x1="6" x2="6" y1="18" y2="11" />
                            <line x1="10" x2="10" y1="18" y2="11" />
                            <line x1="14" x2="14" y1="18" y2="11" />
                            <line x1="18" x2="18" y1="18" y2="11" />
                            <polygon points="12 2 20 7 4 7" />
                        </svg>
                    </div>
                    <div>
                        <div class="flex items-center gap-2">
                            <h1 class="text-lg sm:text-xl font-bold tracking-tight text-foreground">Progres Kavling
                                Pemakaman</h1>
                            <span
                                class="text-xs font-mono font-semibold px-2.5 py-0.5 rounded-full border border-border bg-accent text-accent-foreground">25
                                HA</span>
                        </div>
                        <p class="text-xs text-muted-foreground mt-0.5 font-medium">Monitoring operasional, ketersediaan
                            unit &amp; SPK fisik</p>
                    </div>
                </div>

                <!-- Right Toolbar & Theme Toggle -->
                <div class="flex items-center gap-3">
                    <div class="hidden sm:flex items-center gap-2.5 text-xs font-mono border-r border-border pr-3.5">
                        <span class="flex items-center gap-1.5 text-muted-foreground font-medium">
                            <span class="w-2 h-2 rounded-full bg-emerald-500"></span> ${reportDate}
                        </span>
                        <span class="text-muted-foreground/40">|</span>
                        <span class="text-muted-foreground" id="genDate"></span>
                    </div>

                    <button id="themeToggle"
                        class="w-9 h-9 rounded-[calc(var(--radius)-4px)] border border-border bg-card flex items-center justify-center text-foreground hover:bg-accent transition-colors cursor-pointer"
                        title="Toggle Light/Dark Theme">
                        <svg id="moonIcon" class="w-4 h-4 hidden dark:block text-primary"
                            xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor"
                            stroke-width="2">
                            <path d="M12 3a6 6 0 0 0 9 9 9 9 0 1 1-9-9Z" />
                        </svg>
                        <svg id="sunIcon" class="w-4 h-4 block dark:hidden text-amber-500"
                            xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor"
                            stroke-width="2">
                            <circle cx="12" cy="12" r="4" />
                            <path d="M12 2v2" />
                            <path d="M12 20v2" />
                            <path d="m4.93 4.93 1.41 1.41" />
                            <path d="m17.66 17.66 1.41 1.41" />
                            <path d="M2 12h2" />
                            <path d="M20 12h2" />
                            <path d="m6.34 17.66-1.41 1.41" />
                            <path d="m19.07 4.93-1.41 1.41" />
                        </svg>
                    </button>
                </div>
            </div>
        </div>
    </header>

    <main class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-6">

        <!-- ============ KPI CARDS ============ -->
        <section class="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-3.5">
            <div
                class="bg-card rounded-[var(--radius)] border border-border p-4 shadow-sm hover:border-primary/40 transition-colors">
                <div class="flex items-center justify-between text-muted-foreground mb-1">
                    <span class="text-[11px] font-bold uppercase tracking-wider">Rencana</span>
                    <svg class="w-4 h-4 text-primary" viewBox="0 0 24 24" fill="none" stroke="currentColor"
                        stroke-width="2">
                        <rect width="7" height="7" x="3" y="3" rx="1" />
                        <rect width="7" height="7" x="14" y="3" rx="1" />
                        <rect width="7" height="7" x="14" y="14" rx="1" />
                        <rect width="7" height="7" x="3" y="14" rx="1" />
                    </svg>
                </div>
                <p class="text-2xl font-bold font-mono tracking-tight text-foreground">${summary.totalRencana.toLocaleString('id-ID')}</p>
                <p class="text-[11px] text-muted-foreground mt-0.5 font-medium">Total kavling</p>
            </div>

            <div
                class="bg-card rounded-[var(--radius)] border border-border p-4 shadow-sm hover:border-primary/40 transition-colors">
                <div class="flex items-center justify-between text-muted-foreground mb-1">
                    <span class="text-[11px] font-bold uppercase tracking-wider">SPK</span>
                    <svg class="w-4 h-4 text-primary" viewBox="0 0 24 24" fill="none" stroke="currentColor"
                        stroke-width="2">
                        <path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2" />
                        <path d="M12 11h4" />
                        <path d="M12 16h4" />
                    </svg>
                </div>
                <p class="text-2xl font-bold font-mono tracking-tight text-foreground">${summary.totalSPK.toLocaleString('id-ID')}</p>
                <p class="text-[11px] text-muted-foreground mt-0.5 font-medium">${spkPctOfRencana}% dari rencana</p>
            </div>

            <div
                class="bg-card rounded-[var(--radius)] border border-border p-4 shadow-sm hover:border-primary/40 transition-colors">
                <div class="flex items-center justify-between text-muted-foreground mb-1">
                    <span class="text-[11px] font-bold uppercase tracking-wider">Terjual</span>
                    <svg class="w-4 h-4 text-primary" viewBox="0 0 24 24" fill="none" stroke="currentColor"
                        stroke-width="2">
                        <circle cx="12" cy="12" r="10" />
                        <path d="M16 8h-6a2 2 0 1 0 0 4h4a2 2 0 1 1 0 4H8" />
                        <path d="M12 18V6" />
                    </svg>
                </div>
                <p class="text-2xl font-bold font-mono tracking-tight text-foreground">${summary.totalTerjual.toLocaleString('id-ID')}</p>
                <p class="text-[11px] text-muted-foreground mt-0.5 font-medium">${terjualPctOfSPK}% dari SPK</p>
            </div>

            <div
                class="bg-card rounded-[var(--radius)] border border-border p-4 shadow-sm hover:border-primary/40 transition-colors">
                <div class="flex items-center justify-between text-muted-foreground mb-1">
                    <span class="text-[11px] font-bold uppercase tracking-wider">Sisa Stok</span>
                    <svg class="w-4 h-4 text-primary" viewBox="0 0 24 24" fill="none" stroke="currentColor"
                        stroke-width="2">
                        <path
                            d="M21 8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16Z" />
                        <path d="M12 22V12" />
                    </svg>
                </div>
                <p class="text-2xl font-bold font-mono tracking-tight text-foreground">${summary.totalSisa.toLocaleString('id-ID')}</p>
                <p class="text-[11px] text-muted-foreground mt-0.5 font-medium">Belum terjual</p>
            </div>

            <div
                class="bg-card rounded-[var(--radius)] border border-border p-4 shadow-sm hover:border-primary/40 transition-colors">
                <div class="flex items-center justify-between text-muted-foreground mb-1">
                    <span class="text-[11px] font-bold uppercase tracking-wider">Funeral Ready</span>
                    <svg class="w-4 h-4 text-emerald-500" viewBox="0 0 24 24" fill="none" stroke="currentColor"
                        stroke-width="2">
                        <circle cx="12" cy="12" r="10" />
                        <path d="m9 12 2 2 4-4" />
                    </svg>
                </div>
                <p class="text-2xl font-bold font-mono tracking-tight text-emerald-500">${summary.totalReady.toLocaleString('id-ID')}</p>
                <p class="text-[11px] text-muted-foreground mt-0.5 font-medium">Siap pakai</p>
            </div>

            <!-- Primary Featured Card -->
            <div class="bg-primary text-primary-foreground rounded-[var(--radius)] border border-primary p-4 shadow-sm">
                <div class="flex items-center justify-between opacity-90 mb-1">
                    <span class="text-[11px] font-bold uppercase tracking-wider">% Ready</span>
                    <svg class="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <polyline points="22 7 13.5 15.5 8.5 10.5 2 17" />
                        <polyline points="16 7 22 7 22 13" />
                    </svg>
                </div>
                <p class="text-2xl font-bold font-mono tracking-tight">${summary.pctReady}%</p>
                <p class="text-[11px] opacity-80 mt-0.5 font-medium">dari total SPK</p>
            </div>
        </section>

        <!-- ============ STATUS & CHARTS ============ -->
        <section class="grid grid-cols-1 lg:grid-cols-12 gap-5">
            <!-- Status Overview -->
            <div
                class="lg:col-span-3 bg-card rounded-[var(--radius)] border border-border p-5 flex flex-col justify-between shadow-sm">
                <div>
                    <div class="flex items-center justify-between border-b border-border pb-3 mb-3">
                        <h3 class="text-xs font-bold text-muted-foreground uppercase tracking-wider">Status Zona</h3>
                        <span
                            class="text-xs font-mono font-bold px-2 py-0.5 rounded-full bg-accent text-accent-foreground">${zones.length}
                            Zona</span>
                    </div>
                    <div class="space-y-2.5">
                        <div class="flex items-center justify-between text-xs">
                            <span class="flex items-center gap-2 text-muted-foreground font-medium">
                                <span class="w-2.5 h-2.5 rounded-full bg-emerald-500"></span> Selesai
                            </span>
                            <span class="font-mono font-bold text-foreground">${summary.statusCounts.selesai || 0}</span>
                        </div>
                        <div class="flex items-center justify-between text-xs">
                            <span class="flex items-center gap-2 text-muted-foreground font-medium">
                                <span class="w-2.5 h-2.5 rounded-full bg-primary"></span> Pengerjaan
                            </span>
                            <span class="font-mono font-bold text-foreground">${summary.statusCounts.pengerjaan || 0}</span>
                        </div>
                        <div class="flex items-center justify-between text-xs">
                            <span class="flex items-center gap-2 text-muted-foreground font-medium">
                                <span class="w-2.5 h-2.5 rounded-full bg-muted-foreground/60"></span> Tender
                            </span>
                            <span class="font-mono font-bold text-foreground">${summary.statusCounts.tender || 0}</span>
                        </div>
                    </div>
                </div>

                <div class="border-t border-border pt-4 mt-4">
                    <div class="flex justify-between text-xs mb-1.5 font-medium">
                        <span class="text-muted-foreground">Penyerapan SPK</span>
                        <span class="font-mono font-bold text-foreground">${terjualPctOfSPK}%</span>
                    </div>
                    <div class="w-full bg-muted rounded-full h-2 overflow-hidden">
                        <div class="bg-primary h-full rounded-full progress-bar" style="width: ${terjualPctOfSPK}%"></div>
                    </div>
                    <div class="flex justify-between text-[11px] text-muted-foreground font-mono mt-1.5">
                        <span>${summary.totalTerjual.toLocaleString('id-ID')} terjual</span>
                        <span>${summary.totalSisa.toLocaleString('id-ID')} sisa</span>
                    </div>
                </div>
            </div>

            <!-- Chart Distribusi -->
            <div
                class="lg:col-span-4 bg-card rounded-[var(--radius)] border border-border p-5 shadow-sm flex flex-col justify-between">
                <div class="flex items-center justify-between mb-3 border-b border-border pb-2.5">
                    <h3 class="text-xs font-bold text-muted-foreground uppercase tracking-wider">Distribusi SPK per Tipe
                    </h3>
                    <span class="text-[11px] text-muted-foreground font-mono">${summary.totalSPK.toLocaleString('id-ID')} SPK</span>
                </div>
                <div class="flex items-center gap-4">
                    <div class="w-36 h-36 flex-shrink-0 relative">
                        <canvas id="chartStokType"></canvas>
                    </div>
                    <div id="chartStokLegend" class="flex-1 space-y-1.5 text-xs font-mono min-w-0"></div>
                </div>
            </div>

            <!-- Progress per Zona (Bar List) -->
            <div class="lg:col-span-5 bg-card rounded-[var(--radius)] border border-border p-5 shadow-sm">
                <div class="flex items-center justify-between mb-3 border-b border-border pb-2.5">
                    <h3 class="text-xs font-bold text-muted-foreground uppercase tracking-wider">Progres Aktual Zona
                        Aktif</h3>
                    <span class="text-[11px] text-muted-foreground font-mono">Progress • Deviasi</span>
                </div>
                <div id="progressBars" class="space-y-1.5 max-h-[175px] overflow-y-auto scrollbar-thin pr-1 text-xs">
                </div>
            </div>
        </section>

        <!-- ============ DETAIL TABLE ============ -->
        <section class="bg-card rounded-[var(--radius)] border border-border overflow-hidden shadow-sm">
            <div
                class="px-5 py-4 border-b border-border flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 bg-card">
                <div>
                    <h2 class="text-sm font-bold text-foreground tracking-tight">Detail Progres per Zona</h2>
                    <p class="text-xs text-muted-foreground">Monitoring SPK, fisik makam, kontraktor &amp; deviasi
                        lapangan</p>
                </div>
                <div class="flex gap-1.5 flex-wrap">
                    <button
                        class="tab-btn active px-3.5 py-1.5 rounded-full text-xs font-semibold border border-border transition-colors cursor-pointer"
                        data-filter="all">Semua</button>
                    <button
                        class="tab-btn px-3.5 py-1.5 rounded-full text-xs font-semibold border border-border text-muted-foreground hover:text-foreground hover:bg-accent transition-colors cursor-pointer"
                        data-filter="pengerjaan">Pengerjaan</button>
                    <button
                        class="tab-btn px-3.5 py-1.5 rounded-full text-xs font-semibold border border-border text-muted-foreground hover:text-foreground hover:bg-accent transition-colors cursor-pointer"
                        data-filter="tender">Tender</button>
                    <button
                        class="tab-btn px-3.5 py-1.5 rounded-full text-xs font-semibold border border-border text-muted-foreground hover:text-foreground hover:bg-accent transition-colors cursor-pointer"
                        data-filter="selesai">Selesai</button>
                </div>
            </div>
            <div class="overflow-x-auto scrollbar-thin">
                <table class="w-full text-xs text-left" id="zonaTable">
                    <thead
                        class="bg-muted/40 border-b border-border text-muted-foreground font-mono uppercase text-[11px]">
                        <tr>
                            <th class="px-4 py-3 cursor-pointer hover:text-foreground font-semibold" data-sort="zona">
                                Zona ↕</th>
                            <th class="px-4 py-3 font-semibold">Type</th>
                            <th class="px-4 py-3 text-right cursor-pointer hover:text-foreground font-semibold"
                                data-sort="stokSPK">SPK ↕</th>
                            <th class="px-4 py-3 text-right cursor-pointer hover:text-foreground font-semibold"
                                data-sort="terjual">Terjual ↕</th>
                            <th class="px-4 py-3 text-right font-semibold">Sisa</th>
                            <th class="px-4 py-3 text-right cursor-pointer hover:text-foreground font-semibold"
                                data-sort="funeralReady">Ready ↕</th>
                            <th class="px-4 py-3 min-w-[130px] font-semibold">% Ready</th>
                            <th class="px-4 py-3 font-semibold">Status</th>
                            <th class="px-4 py-3 font-semibold">Kontraktor</th>
                            <th class="px-4 py-3 min-w-[130px] cursor-pointer hover:text-foreground font-semibold"
                                data-sort="aktualProgress">Progress ↕</th>
                            <th class="px-4 py-3 font-semibold">Deviasi</th>
                            <th class="px-4 py-3 font-semibold">Jadwal</th>
                            <th class="px-4 py-3 font-semibold">Siteplan</th>
                        </tr>
                    </thead>
                    <tbody id="zonaTableBody" class="divide-y divide-border"></tbody>
                </table>
            </div>
        </section>

        <!-- ============ REKAP PER TIPE ============ -->
        <section class="bg-card rounded-[var(--radius)] border border-border overflow-hidden shadow-sm">
            <div class="px-5 py-4 border-b border-border">
                <h2 class="text-sm font-bold text-foreground tracking-tight">Rekapitulasi per Tipe Kavling</h2>
                <p class="text-xs text-muted-foreground">Evaluasi target ketersediaan vs realisasi unit siap pakai</p>
            </div>
            <div class="overflow-x-auto scrollbar-thin">
                <table class="w-full text-xs text-left">
                    <thead
                        class="bg-muted/40 border-b border-border text-muted-foreground font-mono uppercase text-[11px]">
                        <tr>
                            <th class="px-4 py-3 font-semibold">Tipe Kavling</th>
                            <th class="px-4 py-3 text-right font-semibold">Rencana</th>
                            <th class="px-4 py-3 text-right font-semibold">SPK</th>
                            <th class="px-4 py-3 text-right font-semibold">Terjual</th>
                            <th class="px-4 py-3 text-right font-semibold">Sisa Stok</th>
                            <th class="px-4 py-3 text-right font-semibold">Funeral Ready</th>
                            <th class="px-4 py-3 min-w-[130px] font-semibold">% Ready</th>
                            <th class="px-4 py-3 text-right font-semibold">Target</th>
                            <th class="px-4 py-3 text-right font-semibold">Selisih</th>
                        </tr>
                    </thead>
                    <tbody id="rekapBody" class="divide-y divide-border"></tbody>
                </table>
            </div>
        </section>

        <!-- ============ LAPORAN FOTO & SITEPLAN ============ -->
        <section class="bg-card rounded-[var(--radius)] border border-border overflow-hidden shadow-sm">
            <div class="px-5 py-4 border-b border-border">
                <h2 class="text-sm font-bold text-foreground tracking-tight">Laporan Foto &amp; Siteplan Lapangan</h2>
                <p class="text-xs text-muted-foreground">Arsip dokumentasi berkala dan plotting fisik makam</p>
            </div>
            <div class="p-5">
                <div class="flex gap-2 mb-4 flex-wrap" id="siteplanTabs"></div>
                <div id="siteplanContent" class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-3.5">
                </div>
            </div>
        </section>

        <!-- ============ CHANGELOG ============ -->
        <section class="bg-card rounded-[var(--radius)] border border-border overflow-hidden shadow-sm">
            <div class="px-5 py-4 border-b border-border">
                <h2 class="text-sm font-bold text-foreground tracking-tight">Changelog &amp; Riwayat Revisi</h2>
                <p class="text-xs text-muted-foreground">Log penyesuaian lot makam, pergeseran batas struktur &amp;
                    approval</p>
            </div>
            <div class="p-5" id="changelogContent"></div>
        </section>

    </main>

    <!-- ============ FOOTER ============ -->
    <footer class="border-t border-border mt-12 py-6 bg-card text-center text-xs text-muted-foreground font-mono">
        <p>Dashboard Progres Kavling Pemakaman 25HA &bull; OKLCH Electric Blue Theme</p>
        <p class="mt-1">Untuk update data: edit file Excel, jalankan <code
                class="px-2 py-0.5 rounded-full bg-muted text-foreground border border-border">node generate.js</code>
        </p>
    </footer>

    <!-- ============ JAVASCRIPT ============ -->
    <script>
        (function () {
            // === THEME TOGGLE (Sync with class .dark on <html>) ===
            const themeToggle = document.getElementById('themeToggle');
            themeToggle.addEventListener('click', () => {
                const isDark = document.documentElement.classList.toggle('dark');
                localStorage.setItem('theme', isDark ? 'dark' : 'light');
                updateChartTheme();
            });

            if (localStorage.getItem('theme') === 'light') {
                document.documentElement.classList.remove('dark');
            }

            // === DATA SOURCE ===
            const ZONES = ${zonesJSON};
            const TYPE_SUMMARY = ${typeSummaryJSON};
            const ZONA_PROGRESS = ${zonaProgressJSON};
            const SITEPLAN = ${siteplanJSON};
            const CHANGELOG = ${changelogJSON};

            // === GENERATED TIMESTAMP ===
            document.getElementById('genDate').textContent = new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' });

            // === HELPERS ===
            function pctNum(v) {
                if (v === null || v === undefined) return 0;
                const n = Number(v);
                if (isNaN(n)) return 0;
                return n <= 1 && n >= 0 ? n * 100 : n;
            }

            // === STATUS BADGE ===
            function statusBadge(s) {
                if (s.startsWith('6')) return '<span class="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-[11px] font-mono font-semibold border border-emerald-500/20 bg-emerald-500/10 text-emerald-500"><span class="w-1.5 h-1.5 rounded-full bg-emerald-500"></span>Selesai</span>';
                if (s.startsWith('5')) return '<span class="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-[11px] font-mono font-semibold border border-primary/20 bg-primary/10 text-primary"><span class="w-1.5 h-1.5 rounded-full bg-primary animate-pulse"></span>Pengerjaan</span>';
                if (s.startsWith('2')) return '<span class="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-[11px] font-mono font-semibold border border-border bg-muted/60 text-muted-foreground"><span class="w-1.5 h-1.5 rounded-full bg-muted-foreground/50"></span>Tender</span>';
                return '<span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-[11px] font-mono border border-border bg-secondary text-muted-foreground">' + s + '</span>';
            }

            // === PROGRESS BAR ===
            function progressBar(pct, small) {
                const p = Math.min(100, Math.max(0, pct));
                const h = small ? 'h-1.5' : 'h-2';
                return '<div class="flex items-center gap-2 font-mono"><div class="flex-1 bg-muted rounded-full ' + h + ' overflow-hidden"><div class="progress-bar bg-primary ' + h + ' rounded-full" style="width:' + p + '%"></div></div><span class="text-[11px] text-muted-foreground font-semibold min-w-[38px] text-right">' + p.toFixed(1) + '%</span></div>';
            }

            // === DEVIASI BADGE ===
            function devBadge(v) {
                if (v === null || v === undefined || v === '') return '<span class="text-muted-foreground font-mono">-</span>';
                const n = Number(v);
                if (isNaN(n)) return '<span class="text-muted-foreground font-mono">-</span>';
                if (n > 0) return '<span class="font-mono text-emerald-500 font-bold">+' + n.toFixed(1) + '%</span>';
                if (n < 0) return '<span class="font-mono text-destructive font-bold">' + n.toFixed(1) + '%</span>';
                return '<span class="text-muted-foreground font-mono">0.0%</span>';
            }

            // ===================================================================
            // RENDER TABLE
            // ===================================================================
            let currentFilter = 'all';
            let currentSort = { key: null, asc: true };

            function renderTable() {
                let filtered = ZONES;
                if (currentFilter === 'pengerjaan') filtered = ZONES.filter(z => z.status.startsWith('5'));
                else if (currentFilter === 'tender') filtered = ZONES.filter(z => z.status.startsWith('2'));
                else if (currentFilter === 'selesai') filtered = ZONES.filter(z => z.status.startsWith('6'));

                if (currentSort.key) {
                    filtered = [...filtered].sort((a, b) => {
                        let va = a[currentSort.key], vb = b[currentSort.key];
                        if (typeof va === 'string') { va = va.toLowerCase(); vb = (vb || '').toLowerCase(); }
                        if (va < vb) return currentSort.asc ? -1 : 1;
                        if (va > vb) return currentSort.asc ? 1 : -1;
                        return 0;
                    });
                }

                const tbody = document.getElementById('zonaTableBody');
                tbody.innerHTML = filtered.map(z => {
                    const readyPct = pctNum(z.pctReady);
                    const siteplanLink = z.accMakam && z.accMakam.startsWith('http')
                        ? '<a href="' + z.accMakam + '" target="_blank" class="text-primary hover:underline font-mono inline-flex items-center gap-1 text-[11px] font-semibold">Buka ↗</a>'
                        : '<span class="text-muted-foreground font-mono">-</span>';

                    return '<tr class="zona-row">'
                        + '<td class="px-4 py-3 font-semibold text-foreground whitespace-nowrap">' + z.zona + '</td>'
                        + '<td class="px-4 py-3"><span class="px-2 py-0.5 rounded-full text-[11px] font-mono font-semibold border border-border bg-accent text-accent-foreground">' + z.type + '</span></td>'
                        + '<td class="px-4 py-3 text-right font-mono text-foreground font-semibold">' + z.stokSPK.toLocaleString() + '</td>'
                        + '<td class="px-4 py-3 text-right font-mono text-muted-foreground">' + z.terjual.toLocaleString() + '</td>'
                        + '<td class="px-4 py-3 text-right font-mono text-muted-foreground">' + z.sisaStok.toLocaleString() + '</td>'
                        + '<td class="px-4 py-3 text-right font-mono font-bold text-emerald-500">' + z.funeralReady.toLocaleString() + '</td>'
                        + '<td class="px-4 py-3">' + progressBar(readyPct, true) + '</td>'
                        + '<td class="px-4 py-3 whitespace-nowrap">' + statusBadge(z.status) + '</td>'
                        + '<td class="px-4 py-3 text-muted-foreground font-mono">' + (z.kontraktor || '-') + '</td>'
                        + '<td class="px-4 py-3">' + (z.aktualProgress > 0 ? progressBar(z.aktualProgress, false) : '<span class="font-mono text-muted-foreground">-</span>') + '</td>'
                        + '<td class="px-4 py-3 whitespace-nowrap">' + devBadge(z.deviasi) + '</td>'
                        + '<td class="px-4 py-3 font-mono text-muted-foreground whitespace-nowrap text-[11px]">' + (z.tglMulai ? z.tglMulai + ' — ' + z.tglSelesai : '-') + '</td>'
                        + '<td class="px-4 py-3">' + siteplanLink + '</td>'
                        + '</tr>';
                }).join('');
            }

            // Filter Buttons
            document.querySelectorAll('.tab-btn').forEach(btn => {
                btn.addEventListener('click', () => {
                    document.querySelectorAll('.tab-btn').forEach(b => {
                        b.classList.remove('active');
                        b.classList.add('text-muted-foreground');
                    });
                    btn.classList.add('active');
                    btn.classList.remove('text-muted-foreground');
                    currentFilter = btn.dataset.filter;
                    renderTable();
                });
            });

            // Sort Headers
            document.querySelectorAll('[data-sort]').forEach(th => {
                th.addEventListener('click', () => {
                    const key = th.dataset.sort;
                    if (currentSort.key === key) currentSort.asc = !currentSort.asc;
                    else { currentSort.key = key; currentSort.asc = true; }
                    renderTable();
                });
            });

            renderTable();

            // ===================================================================
            // REKAP TABLE
            // ===================================================================
            const rekapBody = document.getElementById('rekapBody');
            rekapBody.innerHTML = TYPE_SUMMARY.map(t => {
                const pct = pctNum(t.pct);
                const selisihColor = t.selisih < 0 ? 'text-destructive font-bold' : t.selisih > 0 ? 'text-emerald-500 font-bold' : 'text-muted-foreground';
                return '<tr class="zona-row">'
                    + '<td class="px-4 py-3 font-semibold text-foreground">' + t.type + '</td>'
                    + '<td class="px-4 py-3 text-right font-mono text-muted-foreground">' + t.stokRencana.toLocaleString() + '</td>'
                    + '<td class="px-4 py-3 text-right font-mono text-foreground font-semibold">' + t.stokSPK.toLocaleString() + '</td>'
                    + '<td class="px-4 py-3 text-right font-mono text-muted-foreground">' + t.terjual.toLocaleString() + '</td>'
                    + '<td class="px-4 py-3 text-right font-mono text-muted-foreground">' + t.sisaStok.toLocaleString() + '</td>'
                    + '<td class="px-4 py-3 text-right font-mono text-emerald-500 font-bold">' + t.funeralReady.toLocaleString() + '</td>'
                    + '<td class="px-4 py-3">' + progressBar(pct, true) + '</td>'
                    + '<td class="px-4 py-3 text-right font-mono text-muted-foreground">' + (t.target || '-') + '</td>'
                    + '<td class="px-4 py-3 text-right font-mono ' + selisihColor + '">' + (t.selisih > 0 ? '+' : '') + (t.selisih || '-') + '</td>'
                    + '</tr>';
            }).join('');

            // ===================================================================
            // SITEPLAN & FOTO
            // ===================================================================
            const siteplanTabs = document.getElementById('siteplanTabs');
            const siteplanContent = document.getElementById('siteplanContent');

            SITEPLAN.forEach((report, idx) => {
                const btn = document.createElement('button');
                btn.className = 'px-3.5 py-1.5 rounded-full text-xs font-mono font-semibold border transition-all cursor-pointer '
                    + (idx === 0 ? 'bg-primary text-primary-foreground border-primary shadow-sm' : 'bg-card text-muted-foreground border-border hover:text-foreground hover:bg-accent');
                btn.textContent = report.date || 'Laporan ' + (idx + 1);
                btn.addEventListener('click', () => {
                    siteplanTabs.querySelectorAll('button').forEach(b => {
                        b.className = 'px-3.5 py-1.5 rounded-full text-xs font-mono font-semibold border transition-all cursor-pointer bg-card text-muted-foreground border-border hover:text-foreground hover:bg-accent';
                    });
                    btn.className = 'px-3.5 py-1.5 rounded-full text-xs font-mono font-semibold border transition-all cursor-pointer bg-primary text-primary-foreground border-primary shadow-sm';
                    renderSiteplan(idx);
                });
                siteplanTabs.appendChild(btn);
            });

            function renderSiteplan(idx) {
                const report = SITEPLAN[idx];
                if (!report) { siteplanContent.innerHTML = '<p class="text-muted-foreground text-xs font-mono">Tidak ada data.</p>'; return; }
                siteplanContent.innerHTML = report.items.map(item => {
                    const fotoLink = item.foto && item.foto.startsWith('http')
                        ? '<a href="' + item.foto + '" target="_blank" class="inline-flex items-center justify-between text-xs font-mono text-foreground hover:text-primary border border-border rounded-[calc(var(--radius)-6px)] px-3 py-1.5 bg-accent/30 hover:bg-accent transition-colors font-medium"><span>Dokumentasi Foto</span><span>↗</span></a>'
                        : '';
                    const statusLink = item.status && item.status.startsWith('http')
                        ? '<a href="' + item.status + '" target="_blank" class="inline-flex items-center justify-between text-xs font-mono text-foreground hover:text-primary border border-border rounded-[calc(var(--radius)-6px)] px-3 py-1.5 bg-accent/30 hover:bg-accent transition-colors font-medium"><span>Siteplan PDF</span><span>↗</span></a>'
                        : (item.status ? '<span class="text-xs font-mono text-muted-foreground">' + item.status + '</span>' : '');
                    return '<div class="border border-border rounded-[calc(var(--radius)-4px)] p-3.5 bg-card hover:border-primary/40 transition-colors shadow-2xs">'
                        + '<h4 class="font-bold text-foreground text-xs mb-2.5 pb-2 border-b border-border">' + item.zona + '</h4>'
                        + '<div class="flex flex-col gap-1.5">' + fotoLink + statusLink + '</div>'
                        + '</div>';
                }).join('');
            }
            if (SITEPLAN.length) renderSiteplan(0);

            // ===================================================================
            // CHANGELOG (Timeline)
            // ===================================================================
            const changelogEl = document.getElementById('changelogContent');
            const grouped = {};
            CHANGELOG.forEach(e => {
                if (!grouped[e.date]) grouped[e.date] = [];
                grouped[e.date].push(e.desc);
            });
            changelogEl.innerHTML = '<div class="space-y-4">' + Object.entries(grouped).map(([date, items]) =>
                '<div class="flex gap-4">'
                + '<div class="flex flex-col items-center"><div class="w-2.5 h-2.5 rounded-full bg-primary mt-1.5 ring-4 ring-primary/20"></div><div class="w-px flex-1 bg-border mt-1"></div></div>'
                + '<div class="pb-3 flex-1">'
                + '<span class="font-mono text-xs font-bold text-foreground bg-accent px-2.5 py-0.5 rounded-full border border-border">' + date + '</span>'
                + '<ul class="mt-2 space-y-1">' + items.map(i => '<li class="text-xs font-mono text-muted-foreground flex items-start gap-2"><span class="text-primary font-bold">&rsaquo;</span>' + i + '</li>').join('') + '</ul>'
                + '</div></div>'
            ).join('') + '</div>';

            // ===================================================================
            // CHART (Sinkron dengan Token OKLCH)
            // ===================================================================
            // OKLCH Chart colors mapped to standard CSS
            const chartColors = [
                '#388bfd', // chart-1 (Electric primary blue)
                '#2da44e', // chart-2 (Emerald)
                '#d29922', // chart-3 (Amber)
                '#3fb950', // chart-4 (Bright Mint)
                '#f85149', // chart-5 (Rose Red)
                '#a371f7'  // chart-6 (Lilac)
            ];

            let stokChart;
            function initChart() {
                const ctx = document.getElementById('chartStokType');
                stokChart = new Chart(ctx, {
                    type: 'doughnut',
                    data: {
                        labels: TYPE_SUMMARY.map(t => t.type),
                        datasets: [{
                            data: TYPE_SUMMARY.map(t => t.stokSPK),
                            backgroundColor: chartColors,
                            borderWidth: 2,
                            borderColor: 'transparent',
                            borderRadius: 4,
                        }]
                    },
                    options: {
                        cutout: '70%',
                        responsive: true,
                        maintainAspectRatio: false,
                        plugins: {
                            legend: { display: false },
                            tooltip: {
                                backgroundColor: '#0d1117',
                                titleColor: '#ffffff',
                                bodyColor: '#8b949e',
                                borderColor: '#30363d',
                                borderWidth: 1,
                                padding: 9,
                                cornerRadius: 8,
                                callbacks: {
                                    label: function (ctx) {
                                        const t = TYPE_SUMMARY[ctx.dataIndex];
                                        return ' ' + t.stokSPK.toLocaleString() + ' SPK (' + t.terjual.toLocaleString() + ' terjual)';
                                    }
                                }
                            }
                        }
                    }
                });
            }
            initChart();

            function updateChartTheme() {
                if (!stokChart) return;
                stokChart.update();
            }

            // Custom Chart Legend
            const legendEl = document.getElementById('chartStokLegend');
            const totalSPK = TYPE_SUMMARY.reduce((s, t) => s + t.stokSPK, 0);
            legendEl.innerHTML = TYPE_SUMMARY.map((t, i) => {
                const pct = totalSPK > 0 ? (t.stokSPK / totalSPK * 100).toFixed(0) : 0;
                return '<div class="flex items-center justify-between gap-1 text-[11px]">'
                    + '<div class="flex items-center gap-1.5 min-w-0">'
                    + '<span class="w-2 h-2 rounded-full flex-shrink-0" style="background:' + chartColors[i] + '"></span>'
                    + '<span class="text-muted-foreground truncate font-medium">' + t.type + '</span>'
                    + '</div>'
                    + '<span class="font-bold text-foreground">' + t.stokSPK + ' <span class="text-muted-foreground font-normal">(' + pct + '%)</span></span>'
                    + '</div>';
            }).join('');

            // Progress per Zona List
            const progData = ZONA_PROGRESS.filter(z => z.aktualProgress > 0);
            const progEl = document.getElementById('progressBars');
            if (progEl) {
                progEl.innerHTML = progData.map(z => {
                    const p = Math.min(100, Math.max(0, z.aktualProgress));
                    const d = Number(z.deviasi);
                    const isAhead = !isNaN(d) && d >= 0;
                    const devBadge = isNaN(d) ? '' : (
                        isAhead
                            ? '<span class="text-[10px] font-mono text-emerald-500 font-bold">+' + d.toFixed(1) + '%</span>'
                            : '<span class="text-[10px] font-mono text-destructive font-bold">' + d.toFixed(1) + '%</span>'
                    );
                    return '<div class="flex items-center gap-2 py-1.5 border-b border-border/50 last:border-0 hover:bg-accent/40 rounded-lg px-2 transition-colors">'
                        + '<div class="w-28 truncate font-semibold text-foreground" title="' + z.zona + '">' + z.zona + '</div>'
                        + '<div class="flex-1 min-w-[70px] bg-muted rounded-full h-1.5 overflow-hidden">'
                        + '<div class="h-full bg-primary rounded-full progress-bar" style="width:' + p + '%"></div>'
                        + '</div>'
                        + '<span class="font-mono text-[11px] text-muted-foreground w-11 text-right font-medium">' + p.toFixed(1) + '%</span>'
                        + '<div class="w-12 text-right">' + devBadge + '</div>'
                        + '</div>';
                }).join('');
            }

        })();
    </script>
</body>

</html>`;
}

// =====================================================================
// MAIN
// =====================================================================
function main() {
  const inputPath = path.resolve(INPUT_FILE);
  const outputPath = path.resolve(OUTPUT_FILE);

  console.log('📖 Membaca file Excel:', inputPath);
  if (!fs.existsSync(inputPath)) {
    console.error('❌ File tidak ditemukan:', inputPath);
    console.error('   Letakkan file Excel di folder ini atau berikan path sebagai argumen.');
    console.error('   Contoh: node generate.js "C:\\path\\to\\data.xlsx"');
    process.exit(1);
  }

  const wb = XLSX.readFile(inputPath);
  console.log('   Sheet ditemukan:', wb.SheetNames.join(', '));

  // Process each sheet
  const zones = processProgresZona(wb.Sheets[wb.SheetNames[0]]);
  console.log('   ✓ Progres Zona:', zones.length, 'entri');

  const types = processTotalByType(wb.Sheets[wb.SheetNames[1]]);
  console.log('   ✓ Total by Type:', types.detail.length, 'detail,', types.summary.length, 'kategori');

  const siteplan = processSiteplan(wb.Sheets[wb.SheetNames[2]]);
  console.log('   ✓ Siteplan:', siteplan.length, 'laporan');

  const changelog = processChangelog(wb.Sheets[wb.SheetNames[3]]);
  console.log('   ✓ Changelog:', changelog.length, 'entri');

  const summary = calculateSummary(zones);

  // Get report date from cell C1
  const dateCell = cellVal(wb.Sheets[wb.SheetNames[0]], 0, 2);
  const reportDate = dateCell ? String(dateCell) : new Date().toLocaleDateString('id-ID');

  console.log('\n📊 Ringkasan:');
  console.log('   Total Rencana:', summary.totalRencana);
  console.log('   Total SPK:', summary.totalSPK);
  console.log('   Terjual:', summary.totalTerjual);
  console.log('   Sisa Stok:', summary.totalSisa);
  console.log('   Funeral Ready:', summary.totalReady);
  console.log('   % Ready:', summary.pctReady + '%');

  // Generate HTML
  console.log('\n🎨 Generating dashboard...');
  const html = generateHTML({ zones, types, siteplan, changelog, summary, reportDate });

  fs.writeFileSync(outputPath, html, 'utf-8');
  console.log('✅ Dashboard berhasil di-generate!');
  console.log('   Output:', outputPath);
  console.log('\n   Buka file di browser untuk melihat dashboard.');
  console.log('   Untuk update: edit Excel → jalankan "node generate.js" lagi.');
}

main();
