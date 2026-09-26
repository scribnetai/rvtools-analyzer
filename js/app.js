/* RVTools Analyzer — 100% client-side. No data ever leaves this page.
   Parsing: SheetJS (vendored in lib/). State: in-memory only, never persisted. */
'use strict';

/* ================= Utilities ================= */
const $ = (id) => document.getElementById(id);
const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));

function parseNum(v) {
  if (v == null || v === '') return 0;
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  const n = parseFloat(String(v).replace(/[^0-9.\-]/g, ''));
  return isFinite(n) ? n : 0;
}
function parsePct(v) {
  if (v == null || v === '') return null;
  if (typeof v === 'number') return v <= 1 && v > 0 ? v * 100 : v; // tolerate 0.45 fractions
  const s = String(v).trim();
  const n = parseFloat(s.replace(/[^0-9.\-]/g, ''));
  if (!isFinite(n)) return null;
  if (!/%/.test(s) && n > 0 && n <= 1) return n * 100;
  return n;
}
function pick(row, aliases) {
  for (const a of aliases) {
    if (row[a] !== undefined && row[a] !== null && String(row[a]).trim() !== '') return row[a];
  }
  return '';
}
const fmtInt = (n) => Math.round(n || 0).toLocaleString('en-US');
function fmtMB(mb) {
  mb = mb || 0;
  if (mb >= 1048576) return (mb / 1048576).toFixed(2) + ' TB';
  if (mb >= 1024) return (mb / 1024).toFixed(1) + ' GB';
  return fmtInt(mb) + ' MB';
}
function fmtMHz(mhz) {
  mhz = mhz || 0;
  if (mhz >= 1000) return (mhz / 1000).toFixed(2) + ' GHz';
  return fmtInt(mhz) + ' MHz';
}
const fmtPct = (n) => (n == null || !isFinite(n)) ? '—' : n.toFixed(1) + '%';
function fmtDate(v) {
  if (!v) return '—';
  const d = new Date(v);
  return isNaN(d) ? esc(String(v)) : d.toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' });
}
function daysAgo(v) {
  if (!v) return null;
  const d = new Date(v).getTime();
  if (isNaN(d)) return null;
  return Math.floor((Date.now() - d) / 86400000);
}
// Simple seeded RNG for stable demo data
function lcg(seed) {
  let s = seed >>> 0;
  return () => { s = (s * 1664525 + 1013904223) >>> 0; return s / 4294967296; };
}

/* ================= Column alias maps (modern + legacy RVTools) ================= */
const COL = {
  vm: {
    name: ['VM'], power: ['Powerstate', 'Power state'], template: ['Template'],
    cpus: ['CPUs', 'Num CPUs'], memMB: ['Memory', 'Memory MB', 'Memory MiB'],
    provMB: ['Provisioned MB', 'Provisioned MiB'], usedMB: ['In Use MB', 'In use MiB', 'Used MB', 'Used MiB'],
    unsharedMB: ['Unshared MB', 'Unshared MiB'],
    os: ['OS according to the configuration file', 'OS'], dc: ['Datacenter', 'DC'],
    cluster: ['Cluster'], host: ['Host', 'ESX host'],
    hw: ['HW version', 'Hardware version'], created: ['Creation date'],
    snaps: ['No. of Snapshots', 'Number of Snapshots', '# Snapshots'],
    uuid: ['VM UUID', 'UUID (instance)'],
  },
  host: {
    name: ['Host'], dc: ['Datacenter'], cluster: ['Cluster'],
    sockets: ['# CPU', '#CPU', 'Sockets', 'Num CPU'],
    coresPerCpu: ['Cores per CPU', 'Cores/CPU', 'Cores per cpu'],
    cores: ['# Cores', '#Cores', 'Total Cores'],
    cpuModel: ['CPU Model', 'Processor Type'], cpuMHz: ['Speed', 'CPU MHz', 'Max MHz'],
    cpuPct: ['CPU usage %', '% CPU'], memMB: ['# Memory', 'Memory MB', 'Memory size', 'Total Memory'],
    memPct: ['Memory usage %', '% Memory'], vms: ['# VMs'],
    version: ['ESX Version', 'Version'], vendor: ['Vendor'], model: ['Model'],
    maint: ['in Maintenance Mode', 'Maintenance Mode'],
  },
  ds: {
    name: ['Name', 'Datastore'], type: ['Type'],
    capMB: ['Capacity MiB', 'Capacity MB'], provMB: ['Provisioned MiB', 'Provisioned MB'],
    usedMB: ['In Use MiB', 'In Use MB'],
  },
  snap: {
    vm: ['VM'], name: ['Name', 'Snapshot'], created: ['Date created', 'Created'],
    sizeMB: ['Size MiB (total)', 'Size MB', 'Size MiB'],
  },
  disk: {
    vm: ['VM'], label: ['Disk', 'Label'], capMB: ['Capacity MiB', 'Capacity MB'],
    thin: ['Thin', 'Thin provisioned', 'Thin Provisioned'],
  },
  tools: { vm: ['VM'], status: ['Tools Status', 'Status'], version: ['Tools Version', 'Version'] },
  vcpu: { vm: ['VM'], vcpus: ['# vCPUs', 'vCPUs', 'Num vCPU'], reservation: ['Reservation'], limit: ['Limit'] },
  vmem: { vm: ['VM'], sizeMB: ['Size MiB', 'Size MB'], reservation: ['Reservation'] },
};
const TAB_NAMES = {
  vInfo: ['vInfo'], vHost: ['vHost'], vCluster: ['vCluster'], vDatastore: ['vDatastore'],
  vDisk: ['vDisk'], vSnapshot: ['vSnapshot', 'vSnapshots'], vTools: ['vTools'],
  vCPU: ['vCPU'], vMemory: ['vMemory'], vNetwork: ['vNetwork'],
};

/* ================= Workbook parsing ================= */
function findSheet(wb, candidates) {
  const lower = wb.SheetNames.map((s) => s.toLowerCase());
  for (const c of candidates) {
    const i = lower.indexOf(c.toLowerCase());
    if (i >= 0) return wb.SheetNames[i];
  }
  return null;
}
// sheet_to_json with header-row detection (RVTools sometimes has title rows above headers)
function sheetRows(ws) {
  if (!ws) return [];
  const asArray = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
  if (!asArray.length) return [];
  let headerIdx = 0;
  for (let i = 0; i < Math.min(asArray.length, 5); i++) {
    const joined = asArray[i].join(' ').toLowerCase();
    if (/(^|\s)(vm|host|name|cluster|datastore|snapshot)(s|\s|$)/.test(' ' + joined + ' ')) { headerIdx = i; break; }
  }
  return XLSX.utils.sheet_to_json(ws, { defval: '', raw: true, range: headerIdx });
}

function normPower(v) {
  const s = String(v || '').toLowerCase().replace(/[\s_-]/g, '');
  if (s === 'poweredon') return 'on';
  if (s === 'poweredoff') return 'off';
  if (s === 'suspended') return 'suspended';
  return 'unknown';
}
function normVMs(rows) {
  return rows.map((r) => ({
    name: String(pick(r, COL.vm.name) || 'unknown'),
    power: normPower(pick(r, COL.vm.power)),
    template: /true|yes/i.test(String(pick(r, COL.vm.template))),
    cpus: parseNum(pick(r, COL.vm.cpus)),
    memMB: parseNum(pick(r, COL.vm.memMB)),
    provMB: parseNum(pick(r, COL.vm.provMB)),
    usedMB: parseNum(pick(r, COL.vm.usedMB)),
    unsharedMB: parseNum(pick(r, COL.vm.unsharedMB)),
    os: String(pick(r, COL.vm.os)),
    dc: String(pick(r, COL.vm.dc) || '—'),
    cluster: String(pick(r, COL.vm.cluster) || ''),
    host: String(pick(r, COL.vm.host) || ''),
    hw: String(pick(r, COL.vm.hw)),
    created: pick(r, COL.vm.created),
    snaps: parseNum(pick(r, COL.vm.snaps)),
    uuid: String(pick(r, COL.vm.uuid)),
  }));
}
function normHosts(rows) {
  return rows.map((r) => {
    const sockets = parseNum(pick(r, COL.host.sockets));
    let coresPerCpu = parseNum(pick(r, COL.host.coresPerCpu));
    let cores = parseNum(pick(r, COL.host.cores));
    if (!coresPerCpu && sockets && cores) coresPerCpu = cores / sockets;
    if (!cores && sockets && coresPerCpu) cores = sockets * coresPerCpu;
    return {
      name: String(pick(r, COL.host.name) || 'unknown'),
      dc: String(pick(r, COL.host.dc) || '—'),
      cluster: String(pick(r, COL.host.cluster) || ''),
      sockets, coresPerCpu, cores,
      cpuModel: String(pick(r, COL.host.cpuModel)),
      cpuMHz: parseNum(pick(r, COL.host.cpuMHz)),
      cpuPct: parsePct(pick(r, COL.host.cpuPct)),
      memMB: parseNum(pick(r, COL.host.memMB)),
      memPct: parsePct(pick(r, COL.host.memPct)),
      vmCount: parseNum(pick(r, COL.host.vms)),
      version: String(pick(r, COL.host.version)),
      vendor: String(pick(r, COL.host.vendor)),
      model: String(pick(r, COL.host.model)),
      maint: /true|yes/i.test(String(pick(r, COL.host.maint))),
    };
  });
}
function normDatastores(rows) {
  return rows.map((r) => {
    const capMB = parseNum(pick(r, COL.ds.capMB));
    const usedMB = parseNum(pick(r, COL.ds.usedMB));
    return {
      name: String(pick(r, COL.ds.name) || 'unknown'),
      type: String(pick(r, COL.ds.type) || '—'),
      capMB, provMB: parseNum(pick(r, COL.ds.provMB)), usedMB,
      freeMB: Math.max(0, capMB - usedMB),
    };
  });
}
function normSnapshots(rows) {
  return rows.map((r) => ({
    vm: String(pick(r, COL.snap.vm)), name: String(pick(r, COL.snap.name)),
    created: pick(r, COL.snap.created), sizeMB: parseNum(pick(r, COL.snap.sizeMB)),
  }));
}
function normDisks(rows) {
  return rows.map((r) => ({
    vm: String(pick(r, COL.disk.vm)), label: String(pick(r, COL.disk.label)),
    capMB: parseNum(pick(r, COL.disk.capMB)),
    thin: /true|yes/i.test(String(pick(r, COL.disk.thin))),
  }));
}
function normTools(rows) {
  const m = {};
  rows.forEach((r) => { m[String(pick(r, COL.tools.vm))] = { status: String(pick(r, COL.tools.status)), version: String(pick(r, COL.tools.version)) }; });
  return m;
}

/* ================= Demo data generator (synthetic, in-memory only) ================= */
function genDemoData() {
  const rnd = lcg(20260926);
  const ri = (a, b) => a + Math.floor(rnd() * (b - a + 1));
  const pickOne = (arr) => arr[Math.floor(rnd() * arr.length)];

  const clusters = [
    { name: 'PROD-Cluster-01', sockets: 2, coresPerCpu: 16, cpuModel: 'Intel Xeon Gold 6248R', cpuMHz: 3000, memMB: 786432, hosts: 6, vms: 74 },
    { name: 'PROD-Cluster-02', sockets: 2, coresPerCpu: 10, cpuModel: 'Intel Xeon Silver 4210R', cpuMHz: 2400, memMB: 393216, hosts: 4, vms: 52 },
    { name: 'DEV-Cluster-01', sockets: 2, coresPerCpu: 8, cpuModel: 'Intel Xeon Silver 4110', cpuMHz: 2100, memMB: 262144, hosts: 3, vms: 38 },
  ];
  const oss = ['Microsoft Windows Server 2022 (64-bit)', 'Ubuntu Linux (64-bit)', 'Red Hat Enterprise Linux 9 (64-bit)', 'Microsoft Windows Server 2019 (64-bit)', 'Debian GNU/Linux 12 (64-bit)'];
  const roles = ['WEB', 'APP', 'DB', 'API', 'FILE', 'DC', 'MON', 'BKUP', 'JUMP', 'PROXY'];
  const dcs = ['DC-East'];

  const hosts = [], vms = [], snapshots = [], disks = [];
  const datastores = [
    { name: 'PROD-DS-01', type: 'VMFS', capMB: 40 * 1048576, usedMB: 0, provMB: 0 },
    { name: 'PROD-DS-02', type: 'VMFS', capMB: 40 * 1048576, usedMB: 0, provMB: 0 },
    { name: 'PROD-DS-03', type: 'vSAN', capMB: 60 * 1048576, usedMB: 0, provMB: 0 },
    { name: 'DEV-DS-01', type: 'NFS', capMB: 25 * 1048576, usedMB: 0, provMB: 0 },
    { name: 'ISO-DS', type: 'VMFS', capMB: 2 * 1048576, usedMB: 0, provMB: 0 },
  ];
  const toolsMap = {};
  let vmSeq = 0;

  clusters.forEach((c, ci) => {
    for (let h = 1; h <= c.hosts; h++) {
      const hn = `esx-${String(ci * 10 + h).padStart(2, '0')}.lab.local`;
      const cores = c.sockets * c.coresPerCpu;
      hosts.push({
        name: hn, dc: dcs[0], cluster: c.name,
        sockets: c.sockets, coresPerCpu: c.coresPerCpu, cores,
        cpuModel: c.cpuModel, cpuMHz: c.cpuMHz,
        cpuPct: ri(28, 68), memPct: ri(45, 82),
        memMB: c.memMB, vmCount: 0,
        version: (ci === 2 && h === 3) ? '7.0.3' : '8.0.3',
        vendor: 'Dell Inc.', model: 'PowerEdge R740',
        maint: false,
      });
    }
    for (let v = 0; v < c.vms; v++) {
      vmSeq++;
      const role = pickOne(roles);
      const name = `${role}-${String(vmSeq).padStart(3, '0')}`;
      const roll = rnd();
      const power = roll < 0.82 ? 'on' : roll < 0.93 ? 'off' : 'suspended';
      const template = rnd() < 0.04;
      const cpus = pickOne([2, 2, 4, 4, 4, 8, 8, 16]);
      const memMB = pickOne([4096, 8192, 8192, 16384, 16384, 32768, 65536]);
      const provMB = Math.round((ri(40, 400) * 1024 + (cpus >= 8 ? ri(100, 600) * 1024 : 0)) / 1024) * 1024;
      const usedMB = Math.round(provMB * (0.35 + rnd() * 0.5));
      const host = hosts[hosts.length - c.hosts + ri(0, c.hosts - 1)];
      host.vmCount++;
      const vm = {
        name, power: template ? 'off' : power, template,
        cpus, memMB, provMB, usedMB, unsharedMB: Math.round(usedMB * 0.9),
        os: pickOne(oss), dc: dcs[0], cluster: c.name, host: host.name,
        hw: pickOne(['vmx-19', 'vmx-19', 'vmx-19', 'vmx-18', 'vmx-14']),
        created: new Date(Date.now() - ri(30, 1200) * 86400000).toISOString(),
        snaps: 0, uuid: 'demo-' + vmSeq,
      };
      vms.push(vm);
      disks.push({ vm: name, label: 'Hard disk 1', capMB: provMB, thin: rnd() < 0.7 });
      const ds = datastores[ci === 2 ? 3 : ri(0, 2)];
      ds.usedMB += usedMB; ds.provMB += provMB;
      toolsMap[name] = { status: rnd() < 0.9 ? 'toolsOk' : 'toolsOld', version: '12388' };
      if (!template && power === 'on' && rnd() < 0.12) {
        const age = ri(2, 120);
        snapshots.push({ vm: name, name: 'Pre-patch snapshot', created: new Date(Date.now() - age * 86400000).toISOString(), sizeMB: ri(2, 60) * 1024 });
        vm.snaps = 1;
      }
    }
  });
  datastores.forEach((d) => { d.freeMB = Math.max(0, d.capMB - d.usedMB); });
  datastores[1].usedMB = Math.round(datastores[1].capMB * 0.93); // one hot datastore for the demo
  datastores[1].freeMB = datastores[1].capMB - datastores[1].usedMB;
  datastores[1].provMB = Math.round(datastores[1].capMB * 1.35); // over-provisioned

  return {
    fileName: 'demo-environment.xlsx (synthetic)',
    demo: true,
    tabs: { vInfo: true, vHost: true, vCluster: false, vDatastore: true, vDisk: true, vSnapshot: true, vTools: true },
    counts: { vms: vms.length, hosts: hosts.length, datastores: datastores.length, snapshots: snapshots.length, disks: disks.length },
    vms, hosts, datastores, snapshots, disks, toolsMap,
  };
}

/* ================= Analyzer ================= */
const LICENSE_MIN_CORES_PER_SOCKET = 16;

function buildModel(parsed) {
  const { vms, hosts, datastores, snapshots, disks } = parsed;

  // host -> cluster lookup
  const hostCluster = {};
  hosts.forEach((h) => { hostCluster[h.name] = h.cluster || '(standalone)'; });

  // Cluster rollups
  const cmap = {};
  const ensure = (name) => {
    if (!cmap[name]) cmap[name] = {
      name, dc: '—', hosts: [], vms: [], sockets: 0, cores: 0, licenseCores: 0,
      cpuMHz: 0, memMB: 0, vcpu: 0, vramMB: 0, provMB: 0, usedMB: 0,
      poweredOn: 0, templates: 0,
    };
    return cmap[name];
  };
  hosts.forEach((h) => {
    const c = ensure(h.cluster || '(standalone)');
    c.hosts.push(h);
    if (h.dc !== '—') c.dc = h.dc;
    c.sockets += h.sockets; c.cores += h.cores;
    c.licenseCores += h.sockets * Math.max(h.coresPerCpu || 0, LICENSE_MIN_CORES_PER_SOCKET);
    c.cpuMHz += h.sockets * h.coresPerCpu * h.cpuMHz;
    c.memMB += h.memMB;
  });
  vms.forEach((vm) => {
    const cname = vm.cluster || hostCluster[vm.host] || '(standalone)';
    const c = ensure(cname);
    c.vms.push(vm);
    if (vm.dc !== '—') c.dc = vm.dc;
    if (vm.power === 'on') { c.vcpu += vm.cpus; c.vramMB += vm.memMB; c.poweredOn++; }
    if (vm.template) c.templates++;
    c.provMB += vm.provMB; c.usedMB += vm.usedMB;
  });
  const clusters = Object.values(cmap).sort((a, b) => b.vms.length - a.vms.length);
  clusters.forEach((c) => {
    c.phantom = c.licenseCores - c.cores;
    c.vcpuPerCore = c.cores ? c.vcpu / c.cores : 0;
    c.memOvercommit = c.memMB ? c.vramMB / c.memMB : 0;
    const cpuPcts = c.hosts.map((h) => h.cpuPct).filter((x) => x != null);
    const memPcts = c.hosts.map((h) => h.memPct).filter((x) => x != null);
    c.avgCpuPct = cpuPcts.length ? cpuPcts.reduce((a, b) => a + b, 0) / cpuPcts.length : null;
    c.avgMemPct = memPcts.length ? memPcts.reduce((a, b) => a + b, 0) / memPcts.length : null;
  });

  // Totals
  const on = vms.filter((v) => v.power === 'on');
  const totals = {
    vms: vms.length, poweredOn: on.length,
    poweredOff: vms.filter((v) => v.power === 'off' && !v.template).length,
    suspended: vms.filter((v) => v.power === 'suspended').length,
    templates: vms.filter((v) => v.template).length,
    hosts: hosts.length, clusters: clusters.length,
    sockets: hosts.reduce((a, h) => a + h.sockets, 0),
    cores: hosts.reduce((a, h) => a + h.cores, 0),
    licenseCores: hosts.reduce((a, h) => a + h.sockets * Math.max(h.coresPerCpu || 0, LICENSE_MIN_CORES_PER_SOCKET), 0),
    vcpu: on.reduce((a, v) => a + v.cpus, 0),
    vramMB: on.reduce((a, v) => a + v.memMB, 0),
    hostMemMB: hosts.reduce((a, h) => a + h.memMB, 0),
    provMB: vms.reduce((a, v) => a + v.provMB, 0),
    usedMB: vms.reduce((a, v) => a + v.usedMB, 0),
    dsCapMB: datastores.reduce((a, d) => a + d.capMB, 0),
    dsUsedMB: datastores.reduce((a, d) => a + d.usedMB, 0),
    dsProvMB: datastores.reduce((a, d) => a + d.provMB, 0),
    snapCount: snapshots.length,
    snapMB: snapshots.reduce((a, s) => a + s.sizeMB, 0),
    poweredOffProvMB: vms.filter((v) => v.power === 'off' && !v.template).reduce((a, v) => a + v.provMB, 0),
  };
  totals.phantom = totals.licenseCores - totals.cores;
  totals.vcpuPerCore = totals.cores ? totals.vcpu / totals.cores : 0;

  const thinDisks = disks.filter((d) => d.thin).length;
  const thinMB = disks.filter((d) => d.thin).reduce((a, d) => a + d.capMB, 0);

  // Findings — plain-English, SE-ready
  const findings = [];
  const F = (sev, title, detail) => findings.push({ sev, title, detail });

  clusters.forEach((c) => {
    if (c.phantom > 0) {
      const pct = c.licenseCores ? Math.round((c.phantom / c.licenseCores) * 100) : 0;
      F('warn', `${c.name}: ${fmtInt(c.phantom)} phantom cores (${pct}% of license footprint)`,
        `Hosts in this cluster bill at the 16-core-per-socket minimum, so you're licensing ${fmtInt(c.licenseCores)} cores for ${fmtInt(c.cores)} physical. Consolidating onto higher-density hosts (≥16 cores/socket) removes the waste without losing capacity.`);
    }
    if (c.vcpuPerCore > 8) F('warn', `${c.name}: high vCPU overcommit (${c.vcpuPerCore.toFixed(1)}:1)`,
      `Allocated vCPUs are ${c.vcpuPerCore.toFixed(1)}× physical cores. Fine for light workloads, risky for CPU-heavy ones — validate with performance data before adding VMs.`);
    if (c.memOvercommit > 1.25) F('warn', `${c.name}: memory overcommitted at ${(c.memOvercommit * 100).toFixed(0)}%`,
      `Powered-on VMs are allocated ${(c.memOvercommit * 100).toFixed(0)}% of physical RAM. Watch ballooning/swapping on these hosts.`);
  });
  datastores.forEach((d) => {
    if (!d.capMB) return;
    const usedPct = (d.usedMB / d.capMB) * 100;
    const provPct = (d.provMB / d.capMB) * 100;
    if (usedPct >= 90) F('crit', `Datastore ${d.name} is ${usedPct.toFixed(0)}% full`,
      `Only ${fmtMB(d.freeMB)} free of ${fmtMB(d.capMB)}. Snapshot growth or a single large provisioning could tip it over — act before it hits 95%.`);
    else if (usedPct >= 80) F('warn', `Datastore ${d.name} is ${usedPct.toFixed(0)}% used`, `${fmtMB(d.freeMB)} free of ${fmtMB(d.capMB)}. Plan expansion or rebalance.`);
    if (provPct > 110) F('warn', `Datastore ${d.name} is over-provisioned at ${provPct.toFixed(0)}%`,
      `Thin provisioning has allocated ${fmtMB(d.provMB)} against ${fmtMB(d.capMB)} of capacity. Safe until usage catches up — then it breaks fast.`);
  });
  if (totals.snapCount) {
    const old = snapshots.filter((s) => { const d = daysAgo(s.created); return d != null && d > 30; });
    F(old.length ? 'warn' : 'info', `${fmtInt(totals.snapCount)} snapshots consuming ${fmtMB(totals.snapMB)}`,
      old.length ? `${old.length} are older than 30 days — snapshots are not backups; they hurt performance and complicate restores. Top candidates listed under Virtual Machines.` : 'All snapshots are under 30 days old. Keep it that way — review monthly.');
  }
  if (totals.poweredOff > 0) F('info', `${fmtInt(totals.poweredOff)} powered-off VMs holding ${fmtMB(totals.poweredOffProvMB)}`,
    'Powered-off VMs still consume licensed storage and backup capacity. Confirm what\'s decommissioned vs. seasonal, then reclaim or archive.');
  const bigVMs = vms.filter((v) => v.power === 'on' && v.cpus >= 8);
  if (bigVMs.length) F('info', `${bigVMs.length} powered-on VMs with 8+ vCPUs`,
    'Wide VMs are the usual over-provisioning suspects (and hurt NUMA locality). Spot-check their actual CPU utilization — most can be right-sized down.');
  const versions = [...new Set(hosts.map((h) => h.version).filter(Boolean))];
  if (versions.length > 1) F('info', `ESXi version sprawl: ${versions.join(', ')}`,
    'Mixed build levels complicate patching and vMotion compatibility (EVC). Align to a single target build.');
  const oldHw = vms.filter((v) => { const m = String(v.hw).match(/(\d+)/); return m && parseInt(m[1]) < 19; });
  if (oldHw.length) F('info', `${fmtInt(oldHw.length)} VMs on older virtual hardware`,
    'Virtual hardware below vmx-19 misses out on newer vSphere features and security. Upgrade during the next maintenance window.');
  if (disks.length) {
    const thinPct = Math.round((thinDisks / disks.length) * 100);
    F('info', `${thinPct}% of virtual disks are thin-provisioned (${fmtInt(thinDisks)} of ${fmtInt(disks.length)})`,
      `Thin disks represent ${fmtMB(thinMB)} of allocated capacity — efficient, but monitor datastore headroom since usage can grow silently.`);
  }
  const maintHosts = hosts.filter((h) => h.maint);
  if (maintHosts.length) F('warn', `${maintHosts.length} host(s) in maintenance mode`, maintHosts.map((h) => h.name).join(', '));

  const sevRank = { crit: 0, warn: 1, info: 2 };
  findings.sort((a, b) => sevRank[a.sev] - sevRank[b.sev]);

  return { clusters, totals, findings, thinDisks, thinMB, versions, hostCluster };
}

/* ================= SVG chart helpers ================= */
function donutSVG(segments, size = 170, thickness = 26) {
  const total = segments.reduce((a, s) => a + s.value, 0) || 1;
  const r = (size - thickness) / 2, cx = size / 2, cy = size / 2;
  const circ = 2 * Math.PI * r;
  let offset = 0;
  const arcs = segments.map((s) => {
    const frac = s.value / total;
    const el = `<circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="${s.color}" stroke-width="${thickness}" stroke-dasharray="${(frac * circ).toFixed(1)} ${circ.toFixed(1)}" stroke-dashoffset="${(-offset * circ).toFixed(1)}" transform="rotate(-90 ${cx} ${cy})"/>`;
    offset += frac;
    return el;
  }).join('');
  return `<svg width="${size}" height="${size}" viewBox="0 0 ${size} ${size}">${arcs}<text x="${cx}" y="${cy - 4}" text-anchor="middle" fill="#e9edf3" font-size="20" font-weight="800">${esc(segments[0].centerTop || '')}</text><text x="${cx}" y="${cy + 16}" text-anchor="middle" fill="#9aa4b4" font-size="12">${esc(segments[0].centerSub || '')}</text></svg>`;
}
function barRow(label, pct, color, val) {
  const p = Math.max(0, Math.min(100, pct || 0));
  return `<div class="bar-row"><span class="bar-label" title="${esc(label)}">${esc(label)}</span><div class="bar-track"><div class="bar-fill ${color}" style="width:${p.toFixed(1)}%"></div></div><span class="bar-val">${esc(val)}</span></div>`;
}
// Sortable table builder
function sortableTable(headers, rows, tableId) {
  // headers: [{k,label,num}], rows: array of arrays (display HTML), sortKeys: parallel array of raw values
  const th = headers.map((h, i) => `<th data-i="${i}" data-num="${h.num ? 1 : 0}" style="cursor:pointer">${esc(h.label)} <span class="sort-ind"></span></th>`).join('');
  const tb = rows.map((r) => `<tr>${r.cells.map((c, i) => `<td class="${headers[i].num ? 'num' : ''}">${c}</td>`).join('')}</tr>`).join('');
  return `<div style="overflow-x:auto"><table class="data" id="${tableId}"><thead><tr>${th}</tr></thead><tbody>${tb}</tbody></table></div>`;
}
function wireSort(tableId, rows, headers, rerender) {
  const tbl = $(tableId);
  if (!tbl) return;
  let state = { i: -1, dir: 1 };
  tbl.querySelectorAll('th').forEach((th) => {
    th.addEventListener('click', () => {
      const i = parseInt(th.dataset.i), num = th.dataset.num === '1';
      state = { i, dir: state.i === i ? -state.dir : 1 };
      const sorted = [...rows].sort((a, b) => {
        const x = a.keys[i], y = b.keys[i];
        const cmp = num ? (x - y) : String(x).localeCompare(String(y));
        return cmp * state.dir;
      });
      tbl.querySelector('tbody').innerHTML = sorted.map((r) => `<tr>${r.cells.map((c, j) => `<td class="${headers[j].num ? 'num' : ''}">${c}</td>`).join('')}</tr>`).join('');
      tbl.querySelectorAll('.sort-ind').forEach((s) => s.textContent = '');
      th.querySelector('.sort-ind').textContent = state.dir === 1 ? '▲' : '▼';
    });
  });
}

/* ================= Render: Summary ================= */
function renderSummary(model, parsed) {
  const t = model.totals;
  const el = $('tab-summary');
  const dsUsedPct = t.dsCapMB ? (t.dsUsedMB / t.dsCapMB) * 100 : null;
  el.innerHTML = `
    <div class="stat-grid">
      <div class="stat"><div class="v blue">${fmtInt(t.clusters)}</div><div class="l">Clusters</div></div>
      <div class="stat"><div class="v blue">${fmtInt(t.hosts)}</div><div class="l">ESXi hosts</div></div>
      <div class="stat"><div class="v green">${fmtInt(t.poweredOn)}</div><div class="l">VMs powered on</div></div>
      <div class="stat"><div class="v">${fmtInt(t.vms)}</div><div class="l">Total VMs</div></div>
      <div class="stat"><div class="v purple">${fmtInt(t.vcpu)}</div><div class="l">Allocated vCPUs</div></div>
      <div class="stat"><div class="v purple">${t.vcpuPerCore.toFixed(1)}:1</div><div class="l">vCPU : pCore</div></div>
      <div class="stat"><div class="v amber">${fmtInt(t.cores)}</div><div class="l">Physical cores</div></div>
      <div class="stat"><div class="v ${t.phantom > 0 ? 'red' : 'green'}">${fmtInt(t.licenseCores)}</div><div class="l">License cores required</div></div>
      <div class="stat"><div class="v">${fmtMB(t.provMB)}</div><div class="l">Provisioned storage</div></div>
      <div class="stat"><div class="v">${fmtMB(t.dsCapMB)}</div><div class="l">Datastore capacity</div></div>
    </div>

    <div class="panel"><h3>🎯 Key findings <span class="sub">auto-generated from your export</span></h3>
      ${model.findings.length ? model.findings.map((f) => `
        <div class="finding ${f.sev}">
          <span class="sev">${f.sev === 'crit' ? '🔴' : f.sev === 'warn' ? '🟡' : '🔵'}</span>
          <div><strong>${esc(f.title)}</strong><p>${esc(f.detail)}</p></div>
        </div>`).join('') : '<p class="muted">No findings — either a very tidy environment or limited tabs in the export.</p>'}
    </div>

    <div class="grid2">
      <div class="panel"><h3>💾 Datastore capacity</h3>
        <div class="donut-wrap">
          ${donutSVG([
            { value: t.dsUsedMB, color: '#4f8cff', centerTop: dsUsedPct != null ? dsUsedPct.toFixed(0) + '%' : '—', centerSub: 'used' },
            { value: Math.max(0, t.dsCapMB - t.dsUsedMB), color: '#1e2635' },
          ])}
          <div class="legend">
            <div><span class="sw" style="background:#4f8cff"></span>Used — ${fmtMB(t.dsUsedMB)}</div>
            <div><span class="sw" style="background:#1e2635;border:1px solid #2a3446"></span>Free — ${fmtMB(t.dsCapMB - t.dsUsedMB)}</div>
            <div><span class="sw" style="background:#f5a623"></span>Provisioned — ${fmtMB(t.dsProvMB)}</div>
          </div>
        </div>
        <p class="note">Provisioned can exceed capacity with thin provisioning — that's normal until usage catches up.</p>
      </div>
      <div class="panel"><h3>🖥️ Cluster host utilization <span class="sub">live at export time</span></h3>
        ${model.clusters.map((c) => barRow(c.name + ' · CPU', c.avgCpuPct, c.avgCpuPct > 80 ? 'red' : c.avgCpuPct > 60 ? 'amber' : 'green', fmtPct(c.avgCpuPct))).join('') || '<p class="empty-note">No vHost tab in export.</p>'}
        <div style="height:10px"></div>
        ${model.clusters.map((c) => barRow(c.name + ' · MEM', c.avgMemPct, c.avgMemPct > 85 ? 'red' : c.avgMemPct > 70 ? 'amber' : 'green', fmtPct(c.avgMemPct))).join('')}
      </div>
    </div>

    <div class="panel"><h3>📦 VM power states</h3>
      <div class="donut-wrap">
        ${donutSVG([
          { value: t.poweredOn, color: '#3ddc84', centerTop: fmtInt(t.poweredOn), centerSub: 'powered on' },
          { value: t.poweredOff, color: '#ff6b6b' },
          { value: t.suspended, color: '#f5a623' },
          { value: t.templates, color: '#b388ff' },
        ])}
        <div class="legend">
          <div><span class="sw" style="background:#3ddc84"></span>Powered on — ${fmtInt(t.poweredOn)}</div>
          <div><span class="sw" style="background:#ff6b6b"></span>Powered off — ${fmtInt(t.poweredOff)}</div>
          <div><span class="sw" style="background:#f5a623"></span>Suspended — ${fmtInt(t.suspended)}</div>
          <div><span class="sw" style="background:#b388ff"></span>Templates — ${fmtInt(t.templates)}</div>
        </div>
      </div>
    </div>`;
}

/* ================= Render: Clusters ================= */
function renderClusters(model) {
  const el = $('tab-clusters');
  el.innerHTML = model.clusters.map((c) => `
    <div class="panel">
      <h3>🗂️ ${esc(c.name)} <span class="sub">${esc(c.dc)} · ${c.hosts.length} hosts · ${c.vms.length} VMs (${c.poweredOn} on)</span></h3>
      <div class="stat-grid">
        <div class="stat"><div class="v">${fmtInt(c.sockets)}</div><div class="l">Sockets</div></div>
        <div class="stat"><div class="v">${fmtInt(c.cores)}</div><div class="l">Physical cores</div></div>
        <div class="stat"><div class="v ${c.phantom > 0 ? 'red' : 'green'}">${fmtInt(c.licenseCores)}</div><div class="l">License cores</div></div>
        <div class="stat"><div class="v">${fmtMHz(c.cpuMHz)}</div><div class="l">Total CPU</div></div>
        <div class="stat"><div class="v">${fmtMB(c.memMB)}</div><div class="l">Total RAM</div></div>
        <div class="stat"><div class="v">${c.vcpuPerCore.toFixed(1)}:1</div><div class="l">vCPU:pCore</div></div>
        <div class="stat"><div class="v">${fmtMB(c.provMB)}</div><div class="l">VM provisioned</div></div>
        <div class="stat"><div class="v">${fmtMB(c.usedMB)}</div><div class="l">VM used</div></div>
      </div>
      ${barRow('Avg host CPU', c.avgCpuPct, 'blue', fmtPct(c.avgCpuPct))}
      ${barRow('Avg host memory', c.avgMemPct, 'purple', fmtPct(c.avgMemPct))}
      ${barRow('Memory allocation (overcommit)', Math.min(100, c.memOvercommit * 100), c.memOvercommit > 1.25 ? 'amber' : 'green', (c.memOvercommit * 100).toFixed(0) + '%')}
    </div>`).join('') || '<div class="panel"><p class="empty-note">No cluster data — the export needs vHost or vInfo with cluster info.</p></div>';
}

/* ================= Render: Licensing ================= */
function renderLicensing(model) {
  const t = model.totals;
  const el = $('tab-licensing');
  const rows = model.clusters.map((c) => {
    const perSocket = c.hosts.length ? Math.round(c.cores / Math.max(1, c.sockets)) : 0;
    const wastePct = c.licenseCores ? (c.phantom / c.licenseCores) * 100 : 0;
    return {
      keys: [c.name, c.hosts.length, c.sockets, perSocket, c.cores, c.licenseCores, c.phantom, wastePct],
      cells: [`<strong>${esc(c.name)}</strong>`, c.hosts.length, c.sockets, perSocket, fmtInt(c.cores), `<strong>${fmtInt(c.licenseCores)}</strong>`,
        `<span style="color:${c.phantom > 0 ? 'var(--red)' : 'var(--green)'}">${fmtInt(c.phantom)}</span>`,
        `${wastePct.toFixed(0)}%`],
    };
  });
  const headers = [
    { k: 'cluster', label: 'Cluster' }, { k: 'hosts', label: 'Hosts', num: 1 },
    { k: 'sockets', label: 'Sockets', num: 1 }, { k: 'cps', label: 'Cores / socket', num: 1 },
    { k: 'cores', label: 'Physical cores', num: 1 }, { k: 'lic', label: 'License cores', num: 1 },
    { k: 'phantom', label: 'Phantom cores', num: 1 }, { k: 'waste', label: 'Waste', num: 1 },
  ];
  el.innerHTML = `
    <div class="panel">
      <h3>🧮 Per-core licensing impact <span class="sub">Broadcom model · 16-core minimum per socket</span></h3>
      <p class="muted">Every socket is billed at <strong>max(cores per socket, 16)</strong>. Sockets under 16 cores generate <strong style="color:var(--red)">phantom cores</strong> — licenses paid for but physically nonexistent. Click any column to sort.</p>
      ${sortableTable(headers, rows, 'licTable')}
      <div class="stat-grid" style="margin-top:20px">
        <div class="stat"><div class="v">${fmtInt(t.sockets)}</div><div class="l">Total sockets</div></div>
        <div class="stat"><div class="v">${fmtInt(t.cores)}</div><div class="l">Physical cores</div></div>
        <div class="stat"><div class="v amber">${fmtInt(t.licenseCores)}</div><div class="l">Cores to license</div></div>
        <div class="stat"><div class="v ${t.phantom > 0 ? 'red' : 'green'}">${fmtInt(t.phantom)}</div><div class="l">Phantom cores</div></div>
      </div>
    </div>
    <div class="panel"><h3>📉 Where the waste lives</h3>
      ${model.clusters.filter((c) => c.phantom > 0).map((c) => {
        const wastePct = c.licenseCores ? (c.phantom / c.licenseCores) * 100 : 0;
        return barRow(`${c.name} — ${fmtInt(c.phantom)} phantom of ${fmtInt(c.licenseCores)}`, wastePct, 'red', wastePct.toFixed(0) + '% waste');
      }).join('') || '<p class="muted">✅ No phantom cores — every socket has ≥16 cores. This environment is already license-efficient.</p>'}
      <p class="note">SE talk track: phantom cores are the fastest way to cut a Broadcom renewal — consolidating two 2×8-core hosts (64 license cores for 32 physical) onto one 2×32-core host (64 for 64) halves the waste while keeping compute.</p>
    </div>
    <div class="callout">⚠️ Indicative math for scoping conversations — editions, bundles (VVF/VCF), and partner pricing change the dollars. Always validate against an official Broadcom quote.</div>`;
  wireSort('licTable', rows, headers);
}

/* ================= Render: Hosts ================= */
function renderHosts(model, parsed) {
  const el = $('tab-hosts');
  const rows = parsed.hosts.map((h) => {
    const lic = h.sockets * Math.max(h.coresPerCpu || 0, LICENSE_MIN_CORES_PER_SOCKET);
    return {
      keys: [h.name, h.cluster || '—', h.sockets, h.coresPerCpu, h.cores, lic, h.cpuMHz * h.cores, h.memMB, h.cpuPct || -1, h.memPct || -1, h.vmCount, h.version],
      cells: [`<strong>${esc(h.name)}</strong>${h.maint ? ' <span class="pill-tag warn">maint</span>' : ''}`,
        esc(h.cluster || '—'), h.sockets, h.coresPerCpu, fmtInt(h.cores), fmtInt(lic),
        fmtMHz(h.cpuMHz * h.cores), fmtMB(h.memMB), fmtPct(h.cpuPct), fmtPct(h.memPct),
        fmtInt(h.vmCount), esc(h.version || '—')],
    };
  });
  const headers = [
    { k: 'h', label: 'Host' }, { k: 'c', label: 'Cluster' }, { k: 's', label: 'Sockets', num: 1 },
    { k: 'cps', label: 'Cores/socket', num: 1 }, { k: 'cores', label: 'Cores', num: 1 },
    { k: 'lic', label: 'License cores', num: 1 }, { k: 'ghz', label: 'Total CPU', num: 1 },
    { k: 'mem', label: 'RAM', num: 1 }, { k: 'cpu', label: 'CPU %', num: 1 }, { k: 'memp', label: 'Mem %', num: 1 },
    { k: 'vms', label: 'VMs', num: 1 }, { k: 'v', label: 'ESXi' },
  ];
  el.innerHTML = `
    <div class="panel"><h3>🖥️ Host inventory <span class="sub">${parsed.hosts.length} hosts · click any column to sort</span></h3>
      ${parsed.hosts.length ? sortableTable(headers, rows, 'hostTable') : '<p class="empty-note">No vHost tab in this export.</p>'}
      <p class="note">CPU % / Mem % are point-in-time at export. License cores = sockets × max(cores/socket, 16).</p>
    </div>`;
  wireSort('hostTable', rows, headers);
}

/* ================= Render: Storage ================= */
function renderStorage(model, parsed) {
  const el = $('tab-storage');
  const dsRows = parsed.datastores.map((d) => {
    const usedPct = d.capMB ? (d.usedMB / d.capMB) * 100 : 0;
    const provPct = d.capMB ? (d.provMB / d.capMB) * 100 : 0;
    const flag = usedPct >= 90 ? '<span class="pill-tag off">critical</span>' : usedPct >= 80 ? '<span class="pill-tag warn">watch</span>' : provPct > 110 ? '<span class="pill-tag warn">over-prov</span>' : '<span class="pill-tag on">healthy</span>';
    return {
      keys: [d.name, d.type, d.capMB, d.usedMB, d.freeMB, usedPct, d.provMB, provPct],
      cells: [`<strong>${esc(d.name)}</strong>`, esc(d.type), fmtMB(d.capMB), fmtMB(d.usedMB), fmtMB(d.freeMB), usedPct.toFixed(1) + '%', fmtMB(d.provMB), provPct.toFixed(0) + '%', flag],
    };
  });
  const dsHeaders = [
    { k: 'n', label: 'Datastore' }, { k: 't', label: 'Type' }, { k: 'cap', label: 'Capacity', num: 1 },
    { k: 'used', label: 'Used', num: 1 }, { k: 'free', label: 'Free', num: 1 }, { k: 'up', label: 'Used %', num: 1 },
    { k: 'prov', label: 'Provisioned', num: 1 }, { k: 'pp', label: 'Prov %', num: 1 }, { k: 'f', label: 'Health' },
  ];
  const topVMs = [...parsed.vms].sort((a, b) => b.provMB - a.provMB).slice(0, 15).map((v) => {
    const eff = v.provMB ? (v.usedMB / v.provMB) * 100 : 0;
    return { keys: [v.name, v.cluster, v.provMB, v.usedMB, eff], cells: [esc(v.name), esc(v.cluster || '—'), fmtMB(v.provMB), fmtMB(v.usedMB), eff.toFixed(0) + '%'] };
  });
  const thinMB = parsed.disks.filter((d) => d.thin).reduce((a, d) => a + d.capMB, 0);
  const thickMB = parsed.disks.filter((d) => !d.thin).reduce((a, d) => a + d.capMB, 0);
  el.innerHTML = `
    <div class="panel"><h3>💾 Datastores <span class="sub">click any column to sort</span></h3>
      ${parsed.datastores.length ? sortableTable(dsHeaders, dsRows, 'dsTable') : '<p class="empty-note">No vDatastore tab in this export.</p>'}
    </div>
    <div class="grid2">
      <div class="panel"><h3>🥧 Thin vs thick <span class="sub">by allocated capacity</span></h3>
        ${parsed.disks.length ? `<div class="donut-wrap">
          ${donutSVG([
            { value: thinMB, color: '#4f8cff', centerTop: fmtMB(thinMB), centerSub: 'thin' },
            { value: thickMB, color: '#f5a623' },
          ])}
          <div class="legend">
            <div><span class="sw" style="background:#4f8cff"></span>Thin — ${fmtMB(thinMB)} (${parsed.disks.filter((d) => d.thin).length} disks)</div>
            <div><span class="sw" style="background:#f5a623"></span>Thick — ${fmtMB(thickMB)} (${parsed.disks.filter((d) => !d.thin).length} disks)</div>
          </div></div>` : '<p class="empty-note">No vDisk tab in this export.</p>'}
      </div>
      <div class="panel"><h3>🐘 Top VMs by provisioned storage</h3>
        ${topVMs.length ? `<div style="overflow-x:auto"><table class="data"><thead><tr><th>VM</th><th>Cluster</th><th class="num">Provisioned</th><th class="num">Used</th><th class="num">Efficiency</th></tr></thead><tbody>${topVMs.map((r) => `<tr>${r.cells.map((c, i) => `<td class="${i >= 2 ? 'num' : ''}">${c}</td>`).join('')}</tr>`).join('')}</tbody></table></div>` : '<p class="empty-note">No VM data.</p>'}
      </div>
    </div>`;
  wireSort('dsTable', dsRows, dsHeaders);
}

/* ================= Render: VMs ================= */
function renderVMs(model, parsed) {
  const el = $('tab-vms');
  const clusters = [...new Set(parsed.vms.map((v) => v.cluster || '—'))].sort();
  el.innerHTML = `
    <div class="panel"><h3>🖧 Virtual machines <span class="sub">${fmtInt(parsed.vms.length)} total</span></h3>
      <div class="toolbar">
        <input id="vmSearch" placeholder="🔍 Search VMs…">
        <select id="vmPower"><option value="">All power states</option><option value="on">Powered on</option><option value="off">Powered off</option><option value="suspended">Suspended</option></select>
        <select id="vmCluster"><option value="">All clusters</option>${clusters.map((c) => `<option>${esc(c)}</option>`).join('')}</select>
        <label class="muted small" style="align-self:center"><input type="checkbox" id="vmTemplates"> hide templates</label>
      </div>
      <div id="vmTableWrap"></div>
    </div>
    <div class="grid2">
      <div class="panel"><h3>💤 Powered-off VMs <span class="sub">reclaim candidates</span></h3><div id="offList"></div></div>
      <div class="panel"><h3>📸 Snapshots <span class="sub">${parsed.snapshots.length} total</span></h3><div id="snapList"></div></div>
    </div>
    <div class="panel"><h3>🔍 Right-size review candidates <span class="sub">structural flags — validate with performance data</span></h3><div id="sizeList"></div></div>`;

  const drawTable = () => {
    const q = $('vmSearch').value.toLowerCase(), pw = $('vmPower').value, cl = $('vmCluster').value, hideTpl = $('vmTemplates').checked;
    const list = parsed.vms.filter((v) =>
      (!q || v.name.toLowerCase().includes(q) || v.os.toLowerCase().includes(q)) &&
      (!pw || v.power === pw) && (!cl || (v.cluster || '—') === cl) && !(hideTpl && v.template));
    const tag = (v) => v.template ? '<span class="pill-tag tpl">template</span>' : v.power === 'on' ? '<span class="pill-tag on">on</span>' : v.power === 'off' ? '<span class="pill-tag off">off</span>' : '<span class="pill-tag warn">susp</span>';
    $('vmTableWrap').innerHTML = `<p class="muted small">${fmtInt(list.length)} VMs shown</p>
      <div style="overflow-x:auto;max-height:520px;overflow-y:auto"><table class="data"><thead><tr><th>VM</th><th>State</th><th>Cluster</th><th>Host</th><th class="num">vCPU</th><th class="num">RAM</th><th class="num">Prov</th><th class="num">Used</th><th>OS</th></tr></thead>
      <tbody>${list.slice(0, 500).map((v) => `<tr><td><strong>${esc(v.name)}</strong></td><td>${tag(v)}</td><td>${esc(v.cluster || '—')}</td><td>${esc(v.host || '—')}</td><td class="num">${v.cpus}</td><td class="num">${fmtMB(v.memMB)}</td><td class="num">${fmtMB(v.provMB)}</td><td class="num">${fmtMB(v.usedMB)}</td><td>${esc(v.os.split('(')[0].trim())}</td></tr>`).join('')}</tbody></table></div>
      ${list.length > 500 ? '<p class="note">Showing first 500 — refine your search to narrow it down.</p>' : ''}`;
  };
  ['vmSearch', 'vmPower', 'vmCluster', 'vmTemplates'].forEach((id) => $(id).addEventListener('input', drawTable));
  drawTable();

  const off = parsed.vms.filter((v) => v.power === 'off' && !v.template).sort((a, b) => b.provMB - a.provMB).slice(0, 12);
  $('offList').innerHTML = off.length ? `<table class="data"><thead><tr><th>VM</th><th class="num">Provisioned</th><th>Created</th></tr></thead><tbody>${off.map((v) => `<tr><td>${esc(v.name)}</td><td class="num">${fmtMB(v.provMB)}</td><td>${fmtDate(v.created)}</td></tr>`).join('')}</tbody></table>` : '<p class="empty-note">None — tidy.</p>';

  const snaps = [...parsed.snapshots].sort((a, b) => b.sizeMB - a.sizeMB).slice(0, 12);
  $('snapList').innerHTML = snaps.length ? `<table class="data"><thead><tr><th>VM</th><th>Snapshot</th><th class="num">Size</th><th>Age</th></tr></thead><tbody>${snaps.map((s) => { const d = daysAgo(s.created); return `<tr><td>${esc(s.vm)}</td><td>${esc(s.name) || '—'}</td><td class="num">${fmtMB(s.sizeMB)}</td><td>${d == null ? '—' : d + 'd'}${d != null && d > 30 ? ' <span class="pill-tag warn">old</span>' : ''}</td></tr>`; }).join('')}</tbody></table>` : '<p class="empty-note">No snapshots found.</p>';

  const wide = parsed.vms.filter((v) => v.power === 'on' && v.cpus >= 8).sort((a, b) => b.cpus - a.cpus).slice(0, 12);
  const fat = parsed.vms.filter((v) => v.power === 'on' && v.memMB >= 65536).sort((a, b) => b.memMB - a.memMB).slice(0, 12);
  $('sizeList').innerHTML = `
    <div class="grid2">
      <div><h4 class="muted">Wide VMs (8+ vCPU)</h4>${wide.length ? `<table class="data"><thead><tr><th>VM</th><th class="num">vCPU</th><th class="num">RAM</th></tr></thead><tbody>${wide.map((v) => `<tr><td>${esc(v.name)}</td><td class="num">${v.cpus}</td><td class="num">${fmtMB(v.memMB)}</td></tr>`).join('')}</tbody></table>` : '<p class="empty-note">None.</p>'}</div>
      <div><h4 class="muted">Large-memory VMs (64GB+)</h4>${fat.length ? `<table class="data"><thead><tr><th>VM</th><th class="num">RAM</th><th class="num">vCPU</th></tr></thead><tbody>${fat.map((v) => `<tr><td>${esc(v.name)}</td><td class="num">${fmtMB(v.memMB)}</td><td class="num">${v.cpus}</td></tr>`).join('')}</tbody></table>` : '<p class="empty-note">None.</p>'}</div>
    </div>
    <p class="note">RVTools exports are point-in-time inventory — they don't include performance history. Treat these as a shortlist for vROps/Aria Operations validation, not verdicts.</p>`;
}

/* ================= Render: Report tab ================= */
function renderReportTab(model, parsed) {
  const el = $('tab-report');
  el.innerHTML = `
    <div class="panel">
      <h3>📄 Briefing report</h3>
      <p class="muted">Generate a standalone HTML briefing from this analysis — self-contained, no external dependencies, safe to email to your team or the customer. It contains only the analysis of the file you loaded.</p>
      <div class="dash-actions" style="margin:16px 0">
        <button class="btn primary" id="dlReportBtn2">⬇ Download HTML report</button>
        <button class="btn ghost" id="printBtn2">🖨 Print / save as PDF</button>
      </div>
      <h4>What's inside — every dashboard section</h4>
      <ul class="muted">
        <li>Executive summary and all findings</li>
        <li>Per-cluster breakdown and licensing analysis</li>
        <li>Full host inventory and datastore tables</li>
        <li>Thin vs thick storage, top VMs by provisioned</li>
        <li>Full VM inventory, OS mix, Tools status</li>
        <li>Reclaim list (all), snapshots (all), right-size candidates (all)</li>
        <li>Methodology and licensing disclaimer</li>
      </ul>
    </div>
    <div class="panel">
      <h3>🔒 Session &amp; privacy</h3>
      <p class="muted">This analysis exists only in this page's memory. Nothing has been uploaded, stored, or logged — the app makes zero network requests with your data.</p>
      <button class="btn danger-ghost" id="clearBtn2">✕ Clear session data now</button>
    </div>`;
  $('dlReportBtn2').addEventListener('click', downloadReport);
  $('printBtn2').addEventListener('click', () => window.print());
  $('clearBtn2').addEventListener('click', clearSession);
}

function buildReportHTML(model, parsed) {
  const t = model.totals;
  const gen = new Date().toLocaleString('en-US');
  const css = `body{font-family:-apple-system,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;color:#1a1f28;max-width:1000px;margin:0 auto;padding:40px 24px;line-height:1.55}h1{font-size:1.9rem;margin-bottom:4px}h2{font-size:1.3rem;border-bottom:2px solid #4f8cff;padding-bottom:6px;margin-top:38px}h3{font-size:1.05rem;margin-top:24px}.meta{color:#5b6472;font-size:.9rem;margin-bottom:24px}.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin:18px 0}.stat{border:1px solid #dfe4ec;border-radius:10px;padding:12px 14px}.stat .v{font-size:1.4rem;font-weight:800}.stat .l{font-size:.8rem;color:#5b6472}table{width:100%;border-collapse:collapse;font-size:.86rem;margin:12px 0}th{text-align:left;background:#f2f5fa;padding:8px 10px;font-size:.72rem;text-transform:uppercase;letter-spacing:.04em}td{padding:8px 10px;border-bottom:1px solid #e6ebf2;vertical-align:top}.num{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}.finding{border:1px solid #dfe4ec;border-left:4px solid #4f8cff;border-radius:8px;padding:12px 16px;margin:10px 0}.finding.warn{border-left-color:#f5a623}.finding.crit{border-left-color:#e5484d}.finding strong{display:block;margin-bottom:4px}.finding p{margin:0;color:#3c4452;font-size:.92rem}.bad{color:#c92a2a;font-weight:700}.good{color:#1e7e34;font-weight:700}.toc{background:#f2f5fa;border-radius:10px;padding:16px 22px;margin:20px 0}.toc a{color:#2b5fc7;text-decoration:none}.toc li{margin:4px 0}.disclaimer{background:#fff8e8;border:1px solid #f5d48a;border-radius:10px;padding:14px 18px;font-size:.9rem;margin-top:28px}.note{font-size:.85rem;color:#5b6472}@media print{.noprint{display:none}tr{page-break-inside:avoid}h2{page-break-after:avoid}}`;

  const findingHTML = model.findings.map((f) => `<div class="finding ${f.sev}"><strong>${esc(f.title)}</strong><p>${esc(f.detail)}</p></div>`).join('');

  const clusterRows = model.clusters.map((c) => {
    const cps = c.sockets ? Math.round(c.cores / c.sockets) : 0;
    return `<tr><td><strong>${esc(c.name)}</strong><br><span class="note">${esc(c.dc)} · ${c.hosts.length} hosts</span></td><td class="num">${fmtInt(c.sockets)}</td><td class="num">${cps}</td><td class="num">${fmtInt(c.cores)}</td><td class="num"><strong>${fmtInt(c.licenseCores)}</strong></td><td class="num ${c.phantom > 0 ? 'bad' : 'good'}">${fmtInt(c.phantom)}</td><td class="num">${c.vms.length} (${c.poweredOn} on)</td><td class="num">${c.vcpuPerCore.toFixed(1)}:1</td><td class="num">${fmtPct(c.avgCpuPct)} / ${fmtPct(c.avgMemPct)}</td><td class="num">${fmtMB(c.provMB)}</td></tr>`;
  }).join('');

  const licRows = model.clusters.map((c) => {
    const cps = c.sockets ? Math.round(c.cores / c.sockets) : 0;
    const waste = c.licenseCores ? (c.phantom / c.licenseCores * 100).toFixed(0) + '%' : '—';
    return `<tr><td><strong>${esc(c.name)}</strong></td><td class="num">${c.hosts.length}</td><td class="num">${fmtInt(c.sockets)}</td><td class="num">${cps}</td><td class="num">${fmtInt(c.cores)}</td><td class="num"><strong>${fmtInt(c.licenseCores)}</strong></td><td class="num ${c.phantom > 0 ? 'bad' : 'good'}">${fmtInt(c.phantom)}</td><td class="num">${waste}</td></tr>`;
  }).join('');

  const hostRows = parsed.hosts.map((h) => {
    const lic = h.sockets * Math.max(h.coresPerCpu || 0, LICENSE_MIN_CORES_PER_SOCKET);
    return `<tr><td><strong>${esc(h.name)}</strong>${h.maint ? ' (maint)' : ''}</td><td>${esc(h.cluster || '—')}</td><td class="num">${h.sockets}</td><td class="num">${h.coresPerCpu}</td><td class="num">${fmtInt(h.cores)}</td><td class="num">${fmtInt(lic)}</td><td class="num">${fmtMHz(h.cpuMHz * h.cores)}</td><td class="num">${fmtMB(h.memMB)}</td><td class="num">${fmtPct(h.cpuPct)}</td><td class="num">${fmtPct(h.memPct)}</td><td class="num">${fmtInt(h.vmCount)}</td><td>${esc(h.version || '—')}</td></tr>`;
  }).join('');

  const dsRows = parsed.datastores.map((d) => {
    const p = d.capMB ? (d.usedMB / d.capMB * 100).toFixed(1) + '%' : '—';
    return `<tr><td><strong>${esc(d.name)}</strong></td><td>${esc(d.type)}</td><td class="num">${fmtMB(d.capMB)}</td><td class="num">${fmtMB(d.usedMB)}</td><td class="num">${fmtMB(d.freeMB)}</td><td class="num">${p}</td><td class="num">${fmtMB(d.provMB)}</td></tr>`;
  }).join('');

  const thinMB = parsed.disks.filter((d) => d.thin).reduce((a, d) => a + d.capMB, 0);
  const thickMB = parsed.disks.filter((d) => !d.thin).reduce((a, d) => a + d.capMB, 0);
  const thinN = parsed.disks.filter((d) => d.thin).length;

  const topVMs = [...parsed.vms].sort((a, b) => b.provMB - a.provMB).slice(0, 15).map((v) => {
    const eff = v.provMB ? (v.usedMB / v.provMB * 100).toFixed(0) + '%' : '—';
    return `<tr><td>${esc(v.name)}</td><td>${esc(v.cluster || '—')}</td><td class="num">${fmtMB(v.provMB)}</td><td class="num">${fmtMB(v.usedMB)}</td><td class="num">${eff}</td></tr>`;
  }).join('');

  const offVMs = parsed.vms.filter((v) => v.power === 'off' && !v.template).sort((a, b) => b.provMB - a.provMB).map((v) =>
    `<tr><td>${esc(v.name)}</td><td>${esc(v.cluster || '—')}</td><td class="num">${v.cpus}</td><td class="num">${fmtMB(v.memMB)}</td><td class="num">${fmtMB(v.provMB)}</td><td>${fmtDate(v.created)}</td></tr>`).join('');

  const snapRows = [...parsed.snapshots].sort((a, b) => b.sizeMB - a.sizeMB).map((s) => {
    const d = daysAgo(s.created);
    return `<tr><td>${esc(s.vm)}</td><td>${esc(s.name) || '—'}</td><td class="num">${fmtMB(s.sizeMB)}</td><td>${d == null ? '—' : d + ' days'}${d != null && d > 30 ? ' <span class="bad">old</span>' : ''}</td></tr>`;
  }).join('');

  const wideVMs = parsed.vms.filter((v) => v.power === 'on' && v.cpus >= 8).sort((a, b) => b.cpus - a.cpus).map((v) =>
    `<tr><td>${esc(v.name)}</td><td>${esc(v.cluster || '—')}</td><td class="num">${v.cpus}</td><td class="num">${fmtMB(v.memMB)}</td><td class="num">${fmtMB(v.provMB)}</td><td>${esc(v.os.split('(')[0].trim())}</td></tr>`).join('');
  const fatVMs = parsed.vms.filter((v) => v.power === 'on' && v.memMB >= 65536).sort((a, b) => b.memMB - a.memMB).map((v) =>
    `<tr><td>${esc(v.name)}</td><td>${esc(v.cluster || '—')}</td><td class="num">${fmtMB(v.memMB)}</td><td class="num">${v.cpus}</td><td class="num">${fmtMB(v.provMB)}</td><td>${esc(v.os.split('(')[0].trim())}</td></tr>`).join('');


  const clusterHostLists = model.clusters.map((c) => `<p><strong>${esc(c.name)}</strong> <span class="note">${c.hosts.length} hosts</span><br><span class="note">${c.hosts.map((h) => esc(h.name)).join(', ') || '—'}</span></p>`).join('');

  const osCounts = {};
  parsed.vms.forEach((v) => { const k = (v.os || 'unknown').split('(')[0].trim() || 'unknown'; osCounts[k] = (osCounts[k] || 0) + 1; });
  const osRows = Object.entries(osCounts).sort((a, b) => b[1] - a[1]).slice(0, 12).map(([os, n]) => `<tr><td>${esc(os)}</td><td class="num">${fmtInt(n)}</td><td class="num">${(n / parsed.vms.length * 100).toFixed(1)}%</td></tr>`).join('');

  const toolEntries = Object.values(parsed.tools || {});
  const toolStatusRows = toolEntries.length ? Object.entries(toolEntries.reduce((m, x) => { const k = x.status || 'unknown'; m[k] = (m[k] || 0) + 1; return m; }, {})).sort((a, b) => b[1] - a[1]).map(([st, n]) => `<tr><td>${esc(st)}</td><td class="num">${fmtInt(n)}</td><td class="num">${(n / toolEntries.length * 100).toFixed(1)}%</td></tr>`).join('') : '';

  const allVMRows = [...parsed.vms].sort((a, b) => (a.cluster || '').localeCompare(b.cluster || '') || a.name.localeCompare(b.name)).map((v) => {
    const st = v.template ? 'template' : v.power;
    return `<tr><td><strong>${esc(v.name)}</strong></td><td>${st}</td><td>${esc(v.cluster || '—')}</td><td>${esc(v.host || '—')}</td><td class="num">${v.cpus}</td><td class="num">${fmtMB(v.memMB)}</td><td class="num">${fmtMB(v.provMB)}</td><td class="num">${fmtMB(v.usedMB)}</td><td>${esc((v.os || '').split('(')[0].trim())}</td><td>${esc(v.hw || '—')}</td><td>${fmtDate(v.created)}</td></tr>`;
  }).join('');

  return `<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>RVTools Environment Briefing — ${esc(parsed.fileName)}</title><style>${css}</style></head><body>
<h1>VMware Environment Briefing</h1>
<div class="meta">Generated ${esc(gen)} from <strong>${esc(parsed.fileName)}</strong> · RVTools Analyzer — full analysis, all sections. Source data never leaves the browser.</div>
<div class="toc"><strong>Contents</strong><ol>
<li><a href="#s1">Executive summary</a></li><li><a href="#s2">Key findings</a></li><li><a href="#s3">Cluster breakdown</a></li><li><a href="#s4">Licensing analysis</a></li><li><a href="#s5">Host inventory</a></li><li><a href="#s6">Storage</a></li><li><a href="#s7">Virtual machines</a> (OS mix, Tools status, full inventory, reclaim, snapshots, right-size)</li><li><a href="#s8">Methodology</a></li>
</ol></div>

<h2 id="s1">1. Executive summary</h2>
<div class="grid">
<div class="stat"><div class="v">${fmtInt(t.clusters)}</div><div class="l">Clusters</div></div>
<div class="stat"><div class="v">${fmtInt(t.hosts)}</div><div class="l">ESXi hosts</div></div>
<div class="stat"><div class="v">${fmtInt(t.poweredOn)} / ${fmtInt(t.vms)}</div><div class="l">VMs on / total</div></div>
<div class="stat"><div class="v">${fmtInt(t.sockets)}</div><div class="l">CPU sockets</div></div>
<div class="stat"><div class="v">${fmtInt(t.cores)}</div><div class="l">Physical cores</div></div>
<div class="stat"><div class="v">${fmtInt(t.licenseCores)}</div><div class="l">License cores required</div></div>
<div class="stat"><div class="v">${fmtInt(t.phantom)}</div><div class="l">Phantom cores</div></div>
<div class="stat"><div class="v">${t.vcpuPerCore.toFixed(1)}:1</div><div class="l">vCPU : pCore</div></div>
<div class="stat"><div class="v">${fmtMB(t.provMB)}</div><div class="l">VM provisioned storage</div></div>
<div class="stat"><div class="v">${fmtMB(t.dsCapMB)}</div><div class="l">Datastore capacity</div></div>
<div class="stat"><div class="v">${fmtInt(t.snapCount)}</div><div class="l">Snapshots (${fmtMB(t.snapMB)})</div></div>
<div class="stat"><div class="v">${fmtInt(t.templates)}</div><div class="l">Templates</div></div>
</div>

<h2 id="s2">2. Key findings</h2>${findingHTML || '<p>No findings.</p>'}

<h2 id="s3">3. Cluster breakdown</h2>
<table><thead><tr><th>Cluster</th><th class="num">Sockets</th><th class="num">Cores/sock</th><th class="num">Cores</th><th class="num">License cores</th><th class="num">Phantom</th><th class="num">VMs</th><th class="num">vCPU:pCore</th><th class="num">CPU% / Mem%</th><th class="num">Prov storage</th></tr></thead><tbody>${clusterRows || '<tr><td colspan="10">No cluster data.</td></tr>'}</tbody></table>
<p class="note">CPU% / Mem% are point-in-time host utilization at export. vCPU:pCore counts powered-on VMs only.</p>
<h3>Hosts per cluster</h3>
${clusterHostLists}

<h2 id="s4">4. Licensing analysis</h2>
<p>Broadcom licenses vSphere per <strong>physical core</strong> with a <strong>minimum of 16 cores per CPU socket</strong>: <code>license cores = sockets × max(cores per socket, 16)</code>. The shortfall is <strong>phantom cores</strong> — paid for, unusable.</p>
<table><thead><tr><th>Cluster</th><th class="num">Hosts</th><th class="num">Sockets</th><th class="num">Cores/socket</th><th class="num">Physical</th><th class="num">License cores</th><th class="num">Phantom</th><th class="num">Waste</th></tr></thead><tbody>${licRows}</tbody></table>
<div class="grid">
<div class="stat"><div class="v">${fmtInt(t.sockets)}</div><div class="l">Total sockets</div></div>
<div class="stat"><div class="v">${fmtInt(t.cores)}</div><div class="l">Physical cores</div></div>
<div class="stat"><div class="v">${fmtInt(t.licenseCores)}</div><div class="l">Cores to license</div></div>
<div class="stat"><div class="v">${fmtInt(t.phantom)}</div><div class="l">Phantom cores</div></div>
</div>

<h2 id="s5">5. Host inventory</h2>
${parsed.hosts.length ? `<table><thead><tr><th>Host</th><th>Cluster</th><th class="num">Sockets</th><th class="num">Cores/sock</th><th class="num">Cores</th><th class="num">Lic cores</th><th class="num">Total CPU</th><th class="num">RAM</th><th class="num">CPU%</th><th class="num">Mem%</th><th class="num">VMs</th><th>ESXi</th></tr></thead><tbody>${hostRows}</tbody></table>` : '<p class="note">No vHost tab in export.</p>'}

<h2 id="s6">6. Storage</h2>
<h3>Datastores</h3>
${parsed.datastores.length ? `<table><thead><tr><th>Datastore</th><th>Type</th><th class="num">Capacity</th><th class="num">Used</th><th class="num">Free</th><th class="num">Used %</th><th class="num">Provisioned</th></tr></thead><tbody>${dsRows}</tbody></table>` : '<p class="note">No vDatastore tab in export.</p>'}
<h3>Thin vs thick</h3>
${parsed.disks.length ? `<p>${fmtInt(thinN)} of ${fmtInt(parsed.disks.length)} virtual disks are thin-provisioned — <strong>${fmtMB(thinMB)}</strong> thin vs <strong>${fmtMB(thickMB)}</strong> thick by allocated capacity.</p>` : '<p class="note">No vDisk tab in export.</p>'}
<h3>Top 15 VMs by provisioned storage</h3>
<table><thead><tr><th>VM</th><th>Cluster</th><th class="num">Provisioned</th><th class="num">Used</th><th class="num">Efficiency</th></tr></thead><tbody>${topVMs}</tbody></table>

<h2 id="s7">7. Virtual machines</h2>
<p><strong>${fmtInt(t.poweredOn)}</strong> powered on · <strong>${fmtInt(t.poweredOff)}</strong> powered off · <strong>${fmtInt(t.suspended)}</strong> suspended · <strong>${fmtInt(t.templates)}</strong> templates · <strong>${fmtInt(t.vcpu)}</strong> allocated vCPUs · <strong>${fmtMB(t.vramMB)}</strong> allocated RAM (powered-on).</p>
<h3>Guest OS mix (top 12)</h3>
<table><thead><tr><th>Operating system</th><th class="num">VMs</th><th class="num">Share</th></tr></thead><tbody>${osRows}</tbody></table>
${toolStatusRows ? `<h3>VMware Tools status</h3><table><thead><tr><th>Status</th><th class="num">VMs</th><th class="num">Share</th></tr></thead><tbody>${toolStatusRows}</tbody></table>` : ''}
<h3>Full VM inventory (${fmtInt(parsed.vms.length)} VMs)</h3>
<table><thead><tr><th>VM</th><th>State</th><th>Cluster</th><th>Host</th><th class="num">vCPU</th><th class="num">RAM</th><th class="num">Prov</th><th class="num">Used</th><th>OS</th><th>HW</th><th>Created</th></tr></thead><tbody>${allVMRows}</tbody></table>
<h3>Powered-off VMs — reclaim candidates (all)</h3>
${offVMs ? `<table><thead><tr><th>VM</th><th>Cluster</th><th class="num">vCPU</th><th class="num">RAM</th><th class="num">Provisioned</th><th>Created</th></tr></thead><tbody>${offVMs}</tbody></table><p class="note">Holding <strong>${fmtMB(t.poweredOffProvMB)}</strong> of provisioned storage. Confirm decommissioned vs seasonal, then reclaim or archive.</p>` : '<p class="note">None — tidy.</p>'}
<h3>Snapshots (all)</h3>
${snapRows ? `<table><thead><tr><th>VM</th><th>Snapshot</th><th class="num">Size</th><th>Age</th></tr></thead><tbody>${snapRows}</tbody></table>` : '<p class="note">No snapshots found.</p>'}
<h3>Right-size review candidates</h3>
<p class="note">Structural flags from point-in-time inventory — validate against performance history (vROps / Aria Operations) before acting.</p>
<h3>Wide VMs (8+ vCPU)</h3>
${wideVMs ? `<table><thead><tr><th>VM</th><th>Cluster</th><th class="num">vCPU</th><th class="num">RAM</th><th class="num">Provisioned</th><th>OS</th></tr></thead><tbody>${wideVMs}</tbody></table>` : '<p class="note">None.</p>'}
<h3>Large-memory VMs (64 GB+)</h3>
${fatVMs ? `<table><thead><tr><th>VM</th><th>Cluster</th><th class="num">RAM</th><th class="num">vCPU</th><th class="num">Provisioned</th><th>OS</th></tr></thead><tbody>${fatVMs}</tbody></table>` : '<p class="note">None.</p>'}

<h2 id="s8">8. Methodology</h2>
<p>Licensing math follows Broadcom's per-core subscription model: every CPU socket requires <strong>max(physical cores per socket, 16)</strong> core licenses; the difference is reported as phantom cores. Utilization figures are point-in-time at export. Right-sizing flags are structural candidates — validate against performance history before acting. Snapshot ages derive from the export's creation timestamps.</p>
<div class="disclaimer">⚠️ <strong>Indicative analysis, not a quote.</strong> Editions, bundles (VVF/VCF), vSAN entitlements, and partner pricing affect real licensing cost. Validate all figures against an official Broadcom quote before committing to purchases.</div>
</body></html>`;
}

function downloadReport() {
  if (!APP.model || !APP.parsed) return;
  const html = buildReportHTML(APP.model, APP.parsed);
  const blob = new Blob([html], { type: 'text/html' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'rvtools-briefing-' + new Date().toISOString().slice(0, 10) + '.html';
  document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
}

/* ================= Scenario modeler (calculator-style, physgun-inspired) ================= */
function renderScenario(model) {
  const t = model.totals;
  const wrap = $('scenarioWrap');
  if (!wrap) return;
  const get = (id) => parseInt($(id).value);
  const sockets = get('scSockets'), cps = get('scCores'), n = get('scHosts');
  const newLic = n * sockets * Math.max(cps, LICENSE_MIN_CORES_PER_SOCKET);
  const newPhys = n * sockets * cps;
  const delta = t.licenseCores - newLic;
  const needed = Math.max(1, Math.ceil(t.cores / (sockets * cps)));
  $('scCoresVal').textContent = cps + ' cores';
  $('scSocketsVal').textContent = sockets;
  $('scHostsVal').textContent = n + ' hosts';
  $('scResult').innerHTML = `
    <div class="stat-grid">
      <div class="stat"><div class="v">${fmtInt(t.licenseCores)}</div><div class="l">License cores today</div></div>
      <div class="stat"><div class="v ${delta >= 0 ? 'green' : 'red'}">${fmtInt(newLic)}</div><div class="l">License cores in scenario</div></div>
      <div class="stat"><div class="v ${delta >= 0 ? 'green' : 'red'}">${delta >= 0 ? '−' : '+'}${fmtInt(Math.abs(delta))}</div><div class="l">${delta >= 0 ? 'Cores saved' : 'Extra cores'}</div></div>
      <div class="stat"><div class="v">${fmtInt(newPhys)}</div><div class="l">Physical cores in scenario</div></div>
    </div>
    <p class="muted">${delta >= 0
      ? `This refresh cuts <strong style="color:var(--green)">${fmtInt(delta)} license cores</strong> — mostly by eliminating phantom cores on sub-16-core sockets.`
      : `This scenario needs <strong style="color:var(--red)">${fmtInt(-delta)} more</strong> license cores than today — the density doesn't pay for the host count.`}
      To simply match today's ${fmtInt(t.cores)} physical cores you'd need <strong>${needed}</strong> of these hosts.</p>
    <p class="note">Estimates only — capacity, workload performance, and vSAN/storage needs decide the real host count. Validate against an official quote.</p>`;
}
function scenarioHTML() {
  return `
    <div class="panel" id="scenarioWrap">
      <h3>🔬 Refresh scenario modeler <span class="sub">what if we consolidated?</span></h3>
      <p class="muted">Drag the sliders — the math updates live. The fastest way to kill phantom cores is fewer, denser hosts.</p>
      <div class="grid2">
        <div>
          <label class="muted small">Sockets per new host: <strong id="scSocketsVal">2</strong></label>
          <input type="range" id="scSockets" min="1" max="4" step="1" value="2" style="width:100%">
        </div>
        <div>
          <label class="muted small">Cores per socket: <strong id="scCoresVal">32 cores</strong></label>
          <input type="range" id="scCores" min="8" max="64" step="2" value="32" style="width:100%">
        </div>
      </div>
      <div style="margin-top:12px">
        <label class="muted small">Replacement hosts: <strong id="scHostsVal">6 hosts</strong></label>
        <input type="range" id="scHosts" min="1" max="24" step="1" value="6" style="width:100%">
      </div>
      <div id="scResult" style="margin-top:16px"></div>
    </div>`;
}

/* ================= App state & wiring ================= */
const APP = { parsed: null, model: null };

function setStatus(msg) {
  const s = $('parseStatus');
  s.hidden = false; s.innerHTML = msg;
}
function showError(title, hint, sheets) {
  const e = $('fileError');
  e.hidden = false;
  e.innerHTML = `<strong>${esc(title)}</strong><br>${esc(hint)}${sheets ? `<br><span class="muted small">Sheets found: ${esc(sheets.join(', '))}</span>` : ''}`;
}

function parseWorkbook(wb, fileName) {
  const sn = (k) => findSheet(wb, TAB_NAMES[k]);
  const vInfoSheet = sn('vInfo');
  if (!vInfoSheet) {
    showError('That doesn\'t look like an RVTools export.', 'No vInfo tab found — export via File → Export all to Excel in RVTools.', wb.SheetNames);
    return null;
  }
  const get = (k, norm) => { const s = sn(k); return s ? norm(sheetRows(wb.Sheets[s])) : []; };
  const vms = normVMs(get('vInfo', (r) => r));
  if (!vms.length) { showError('The vInfo tab is empty.', 'The export parsed but contained no VM rows.', wb.SheetNames); return null; }
  const tabs = {};
  Object.keys(TAB_NAMES).forEach((k) => { tabs[k] = !!sn(k); });
  return {
    fileName, demo: false, tabs,
    counts: { vms: vms.length },
    vms,
    hosts: normHosts(get('vHost', (r) => r)),
    datastores: normDatastores(get('vDatastore', (r) => r)),
    snapshots: normSnapshots(get('vSnapshot', (r) => r)),
    disks: normDisks(get('vDisk', (r) => r)),
    toolsMap: normTools(get('vTools', (r) => r)),
  };
}

async function handleFile(file) {
  $('fileError').hidden = true;
  if (!/\.(xlsx|xls)$/i.test(file.name)) { showError('Unsupported file type.', 'Please drop an RVTools Excel export (.xlsx or .xls).'); return; }
  setStatus('📖 Reading ' + esc(file.name) + '…');
  try {
    const buf = await file.arrayBuffer();
    setStatus('⚙️ Parsing workbook locally…');
    await new Promise((r) => setTimeout(r, 30)); // let UI paint
    const wb = XLSX.read(buf, { type: 'array' });
    const parsed = parseWorkbook(wb, file.name);
    if (!parsed) { $('parseStatus').hidden = true; return; }
    boot(parsed);
  } catch (err) {
    console.error(err);
    showError('Couldn\'t parse that file.', 'The workbook appears to be corrupt or password-protected. Try re-exporting from RVTools.');
    $('parseStatus').hidden = true;
  }
}

function boot(parsed) {
  APP.parsed = parsed;
  APP.model = buildModel(parsed);
  $('parseStatus').hidden = true;
  $('landing').hidden = true;
  $('dashboard').hidden = false;
  $('dashFileName').textContent = parsed.fileName;
  const bits = [];
  if (parsed.tabs.vInfo) bits.push(fmtInt(parsed.vms.length) + ' VMs');
  if (parsed.hosts.length) bits.push(fmtInt(parsed.hosts.length) + ' hosts');
  if (parsed.datastores.length) bits.push(fmtInt(parsed.datastores.length) + ' datastores');
  $('dashMeta').textContent = bits.join(' · ') + (parsed.demo ? ' · demo data' : '');
  renderSummary(APP.model, parsed);
  renderClusters(APP.model);
  // licensing tab gets the table + scenario modeler
  renderLicensing(APP.model);
  $('tab-licensing').insertAdjacentHTML('beforeend', scenarioHTML());
  ['scSockets', 'scCores', 'scHosts'].forEach((id) => $(id).addEventListener('input', () => renderScenario(APP.model)));
  // default host count: enough dense hosts to cover current cores
  $('scHosts').value = Math.max(1, Math.min(24, Math.ceil(APP.model.totals.cores / 64)));
  renderScenario(APP.model);
  renderHosts(APP.model, parsed);
  renderStorage(APP.model, parsed);
  renderVMs(APP.model, parsed);
  renderReportTab(APP.model, parsed);
  switchTab('summary');
  window.scrollTo({ top: 0 });
  document.querySelector('#dashboard').scrollIntoView();
}

function switchTab(name) {
  document.querySelectorAll('#tabs button').forEach((b) => b.classList.toggle('active', b.dataset.tab === name));
  document.querySelectorAll('.tab-panel').forEach((p) => { p.hidden = p.id !== 'tab-' + name; });
}

function clearSession() {
  // Wipe everything: in-memory state only ever existed here.
  APP.parsed = null; APP.model = null;
  $('fileInput').value = '';
  $('fileError').hidden = true; $('parseStatus').hidden = true;
  $('dashboard').hidden = true; $('landing').hidden = false;
  document.querySelectorAll('.tab-panel').forEach((p) => { p.innerHTML = ''; });
  window.scrollTo({ top: 0 });
  setStatus('✅ Session cleared — no data was ever stored, uploaded, or retained.');
  setTimeout(() => { $('parseStatus').hidden = true; }, 4000);
}

/* ================= Changelog ================= */
function renderChangelog() {
  const body = $('changelog-body');
  if (!body) return;
  fetch('CHANGELOG.md', { cache: 'no-store' })
    .then((res) => { if (!res.ok) throw new Error('bad status'); return res.text(); })
    .then((md) => {
      let html = '', inList = false;
      const closeList = () => { if (inList) { html += '</ul>'; inList = false; } };
      for (const line of md.split('\n')) {
        if (line.startsWith('## ')) { closeList(); html += '<h4>' + esc(line.slice(3).trim()) + '</h4>'; }
        else if (line.startsWith('- ')) { if (!inList) { html += '<ul>'; inList = true; } html += '<li>' + esc(line.slice(2).trim()) + '</li>'; }
        else if (line.trim() === '' || line.startsWith('# ')) { closeList(); }
        else { closeList(); html += '<p>' + esc(line.trim()) + '</p>'; }
      }
      closeList();
      body.innerHTML = html;
    })
    .catch(() => { body.innerHTML = "<p class='muted'>Changelog unavailable.</p>"; });
}

function wireApp() {
  const dz = $('dropzone'), fi = $('fileInput');
  dz.addEventListener('click', () => fi.click());
  dz.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') fi.click(); });
  fi.addEventListener('change', () => { if (fi.files[0]) handleFile(fi.files[0]); });
  ['dragenter', 'dragover'].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.add('drag'); }));
  ['dragleave', 'drop'].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.remove('drag'); }));
  dz.addEventListener('drop', (e) => { const f = e.dataTransfer.files[0]; if (f) handleFile(f); });

  $('demoBtn').addEventListener('click', () => { setStatus('🧪 Generating synthetic demo environment…'); setTimeout(() => boot(genDemoData()), 60); });

  document.querySelectorAll('#tabs button').forEach((b) => b.addEventListener('click', () => switchTab(b.dataset.tab)));

  $('newFileBtn').addEventListener('click', () => { $('dashboard').hidden = true; $('landing').hidden = false; window.scrollTo({ top: 0 }); document.querySelector('#upload').scrollIntoView({ behavior: 'smooth' }); });
  $('clearBtn').addEventListener('click', clearSession);
  $('dlReportBtn').addEventListener('click', downloadReport);
  $('printBtn').addEventListener('click', () => window.print());
  $('brandHome').addEventListener('click', (e) => { e.preventDefault(); $('dashboard').hidden = true; $('landing').hidden = false; window.scrollTo({ top: 0 }); });
  renderChangelog();
}

document.addEventListener('DOMContentLoaded', wireApp);
