param([string]$EventType)

$BasePath = "C:\Timesheets"
$CsvPath = "$BasePath\WorkTimesheet.csv"

# Create the Timesheets folder if it doesn't exist
if (!(Test-Path $BasePath)) {
    New-Item -ItemType Directory -Force -Path $BasePath | Out-Null
}

# 1. LOG TO CSV
if (!(Test-Path $CsvPath)) {
    "Date,Time,Event" | Out-File $CsvPath -Encoding utf8
}

$Now = Get-Date

# Hibernate/sleep: the END task only really runs once the PC wakes up, so it would be stamped
# with the wake time. Back-date it to when Windows actually went to sleep (Kernel-Power 42).
$SkipWrite = $false
if ($EventType -eq 'END') {
    $last = (Get-Content $CsvPath -Tail 1) -split ','
    # Several triggers can fire END for one hibernate/lock; ignore it if the last row is already an END
    if ($last.Count -ge 3 -and $last[2].Trim() -eq 'END') { $SkipWrite = $true }
    if ($last.Count -ge 3 -and $last[2].Trim() -eq 'START') {
        $lastStart = [datetime]"$($last[0]) $($last[1])"
        $sleep = Get-WinEvent -FilterHashtable @{ LogName = 'System'; ProviderName = 'Microsoft-Windows-Kernel-Power'; Id = 42; StartTime = $lastStart } -MaxEvents 1 -ErrorAction SilentlyContinue
        if ($sleep -and ($Now - $sleep.TimeCreated).TotalMinutes -gt 1) { $Now = $sleep.TimeCreated }
    }
}

# Restart/shutdown: the END task often gets killed before it writes. If a START arrives while the
# last row is still a START, back-fill the END at the first shutdown/sleep after that START.
# shortcut: crashes / power loss leave no shutdown event (only 6008 at next boot), so no END is back-filled there.
if ($EventType -eq 'START' -and (Test-Path $CsvPath)) {
    $last = (Get-Content $CsvPath -Tail 1) -split ','
    if ($last.Count -ge 3 -and $last[2].Trim() -eq 'START') {
        $lastStart = [datetime]"$($last[0]) $($last[1])"
        # 1074 = shutdown/restart requested, 6006 = event log stopped (clean shutdown), 42 = sleep
        $stop = Get-WinEvent -FilterHashtable @{ LogName = 'System'; Id = 1074, 6006, 42; StartTime = $lastStart } -ErrorAction SilentlyContinue |
            Where-Object { $_.TimeCreated -gt $lastStart } | Sort-Object TimeCreated | Select-Object -First 1
        if ($stop) { "$($stop.TimeCreated.ToString('yyyy-MM-dd,HH:mm:ss')),END" | Out-File $CsvPath -Append -Encoding utf8 }
    }
}

$DateStr = $Now.ToString("yyyy-MM-dd")
$TimeStr = $Now.ToString("HH:mm:ss")

# Only write to CSV if an event was passed
if ($EventType -and -not $SkipWrite) {
    "$DateStr,$TimeStr,$EventType" | Out-File $CsvPath -Append -Encoding utf8
}

# 2. GENERATE WEEKLY HTML FILES
if (!(Test-Path $CsvPath)) { exit }
$Data = Import-Csv $CsvPath

$HtmlHead = @"
<!DOCTYPE html>
<html>
<head>
<title>Weekly Timesheet</title>
<style>
    body { font-family: "Segoe UI", Arial, sans-serif; background-color: #f4f6f8; margin: 0; padding: 20px; color: #333; }
    .container { max-width: 800px; margin: auto; background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
    h1 { margin-top: 0; border-bottom: 2px solid #0078D7; padding-bottom: 10px; }
    h2 { color: #555; margin-top: 10px; font-size: 1.1em; border-left: 4px solid #0078D7; padding-left: 10px; display: flex; justify-content: space-between; align-items: center; background: #f8fafc; padding: 10px; border-radius: 0 4px 4px 0;}
    .week-total { color: #0078D7; font-weight: bold; font-size: 1.1em; }
    table { border-collapse: collapse; width: 100%; margin-bottom: 10px; }
    th, td { padding: 10px; border-bottom: 1px solid #eee; text-align: left; }
    th { background-color: #f9fafb; font-weight: 600; }
    tr:hover { background-color: #f1f5f9; }
    .day-header { background-color: #e2e8f0 !important; }
    .day-total { color: #333; font-weight: bold; }
    .START { color: #107C10; font-weight: bold; background-color: #e6f4ea; padding: 4px 8px; border-radius: 4px; font-size: 0.85em; }
    .END { color: #D13438; font-weight: bold; background-color: #fde7e9; padding: 4px 8px; border-radius: 4px; font-size: 0.85em; }
</style>
</head>
<body>
<div class="container">
<h1>Weekly Timesheet Log</h1>
"@

# Prepare data and figure out the week dates
$ProcessedData = foreach ($row in $Data) {
    if ([string]::IsNullOrWhiteSpace($row.Date) -or [string]::IsNullOrWhiteSpace($row.Time)) { continue }
    $dateObj = [datetime]::ParseExact($row.Date, 'yyyy-MM-dd', $null)
    $dayOfWeek = [int]$dateObj.DayOfWeek
    $offset = if ($dayOfWeek -eq 0) { -6 } else { 1 - $dayOfWeek }
    $weekStart = $dateObj.AddDays($offset).ToString('yyyy-MM-dd')
    
    [PSCustomObject]@{
        Date = $row.Date
        Time = $row.Time
        Event = $row.Event
        WeekOf = $weekStart
        DateTimeObj = [datetime]("$($row.Date) $($row.Time)")
    }
}

# Group everything by Week (Newest weeks first)
$GroupedWeeks = $ProcessedData | Sort-Object DateTimeObj | Group-Object WeekOf | Sort-Object Name -Descending

foreach ($week in $GroupedWeeks) {
    $WeekTotal = [TimeSpan]::Zero
    $OpenStart = $null
    $WeekHtml = "<table><tr><th>Date</th><th>Time</th><th>Event</th></tr>"
    
    # Group by Day within the week (Newest days first)
    $GroupedDays = $week.Group | Group-Object Date | Sort-Object Name -Descending
    
    foreach ($day in $GroupedDays) {
        $DayTotal = [TimeSpan]::Zero
        $CurrentStart = $null
        $DayHtml = ""
        
        # Sort events chronologically to calculate the time between them
        $DayEvents = $day.Group | Sort-Object DateTimeObj
        
        foreach ($evt in $DayEvents) {
            if ($evt.Event -eq 'START' -and $null -eq $CurrentStart) {
                $CurrentStart = $evt.DateTimeObj
            } elseif ($evt.Event -eq 'END' -and $null -ne $CurrentStart) {
                $DayTotal += ($evt.DateTimeObj - $CurrentStart)
                $CurrentStart = $null
            }
            $DayHtml += "<tr><td width='30%'>$($evt.Date)</td><td width='30%'>$($evt.Time)</td><td><span class='$($evt.Event)'>$($evt.Event)</span></td></tr>"
        }
        
        $WeekTotal += $DayTotal
        $DayTotalFormatted = "{0}h {1:D2}m" -f [math]::Floor($DayTotal.TotalHours), $DayTotal.Minutes

        # Open session today (START without END): let the page keep counting while it's open
        $DayLive = ""
        if ($null -ne $CurrentStart -and $day.Name -eq $DateStr) {
            $OpenStart = $CurrentStart.ToString('yyyy-MM-ddTHH:mm:ss')
            $DayLive = " data-base='$([int]$DayTotal.TotalSeconds)' data-start='$OpenStart'"
        }

        # Add a grey header row for each day showing that day's total
        $WeekHtml += "<tr class='day-header'><td colspan='2'><strong>$($day.Name)</strong></td><td class='day-total'>Total: <span class='live'$DayLive>$DayTotalFormatted</span></td></tr>"
        $WeekHtml += $DayHtml
    }
    $WeekHtml += "</table>"
    
    # Render the Week header with the grand total
    $WeekTotalFormatted = "{0}h {1:D2}m" -f [math]::Floor($WeekTotal.TotalHours), $WeekTotal.Minutes
    $WeekLive = ""
    if ($OpenStart) { $WeekLive = " data-base='$([int]$WeekTotal.TotalSeconds)' data-start='$OpenStart'" }
    $HtmlBody = "<h2><span>Week starting Monday, $($week.Name)</span> <span class='week-total'>Week Total: <span class='live'$WeekLive>$WeekTotalFormatted</span></span></h2>"
    $HtmlBody += $WeekHtml

    $HtmlFoot = @"
</div>
<script>
// Elements with data-start have a session still open: base seconds + time since START
function tick() {
  document.querySelectorAll('.live[data-start]').forEach(function (el) {
    var s = Number(el.dataset.base) + Math.max(0, (Date.now() - new Date(el.dataset.start).getTime()) / 1000);
    var m = Math.floor(s / 60);
    el.textContent = Math.floor(m / 60) + 'h ' + String(m % 60).padStart(2, '0') + 'm';
  });
}
tick();
setInterval(tick, 15000);
</script>
</body>
</html>
"@

    # Construct the filename using the Monday date of this specific week
    $WeeklyHtmlPath = "$BasePath\Timesheet_$($week.Name).html"
    
    # Combine and write the HTML file for this week
    $FinalHtml = $HtmlHead + $HtmlBody + $HtmlFoot
    $FinalHtml | Out-File $WeeklyHtmlPath -Encoding utf8
}

# 3. GENERATE WEEKS.HTML (list of weekly pages)
$IndexHtml = @"
<!DOCTYPE html>
<html>
<head>
<title>Timesheet Directory</title>
<style>
    body { font-family: "Segoe UI", Arial, sans-serif; background-color: #f4f6f8; margin: 0; padding: 20px; color: #333; }
    .container { max-width: 800px; margin: auto; background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
    h1 { border-bottom: 2px solid #0078D7; padding-bottom: 10px; margin-top: 0; }
    ul { list-style-type: none; padding: 0; }
    li { margin: 12px 0; padding: 10px; background: #f8fafc; border-radius: 4px; }
    a { text-decoration: none; color: #0078D7; font-weight: bold; font-size: 1.1em; display: block; }
    a:hover { color: #005a9e; }
</style>
</head>
<body>
<div class="container">
<h1>Timesheet Directory</h1>
<p><a href='./index.html'>&larr; Back to dashboard</a></p>
<ul>
"@

foreach ($week in $GroupedWeeks) {
    $filename = "Timesheet_$($week.Name).html"
    $IndexHtml += "<li><a href='./$filename'>Week starting $($week.Name)</a></li>"
}

$IndexHtml += "</ul></div></body></html>"
$IndexHtml | Out-File "$BasePath\weeks.html" -Encoding utf8


# 4. GENERATE INDEX.HTML = DASHBOARD (GitHub Pages landing page) (all stats computed in the browser from the embedded sessions)
# Pair START/END across the whole log (not per day) so a session crossing midnight is kept;
# the dashboard splits it between the two days
$Sessions = & {
    $cs = $null
    foreach ($evt in ($ProcessedData | Sort-Object DateTimeObj)) {
        if ($evt.Event -eq 'START' -and $null -eq $cs) { $cs = $evt.DateTimeObj }
        elseif ($evt.Event -eq 'END' -and $null -ne $cs) {
            [PSCustomObject]@{ s = $cs.ToString('yyyy-MM-ddTHH:mm:ss'); e = $evt.DateTimeObj.ToString('yyyy-MM-ddTHH:mm:ss') }
            $cs = $null
        }
    }
    # Still clocked in (started < 24h ago): end = null, the page counts it live. Older dangling STARTs are ignored.
    if ($null -ne $cs -and ((Get-Date) - $cs).TotalHours -lt 24) { [PSCustomObject]@{ s = $cs.ToString('yyyy-MM-ddTHH:mm:ss'); e = $null } }
}
$SessionsJson = if ($Sessions) { ConvertTo-Json @($Sessions) -Compress } else { '[]' }

$DashTemplate = @'
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Time Dashboard</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
:root {
  --bg: #0b0f17; --panel: #121826; --panel2: #182032; --border: #232c40;
  --text: #e6eaf2; --muted: #8b95a9; --accent: #4f8cff; --accent2: #8b5cf6;
  --good: #22c55e; --bad: #f43f5e; --warn: #f59e0b;
}
* { box-sizing: border-box; }
body { margin: 0; background: var(--bg); color: var(--text); font: 14px/1.45 "Segoe UI", system-ui, sans-serif; }
.wrap { max-width: 1280px; margin: auto; padding: 20px 16px 40px; }
header { display: flex; flex-wrap: wrap; gap: 12px; align-items: center; justify-content: space-between; margin-bottom: 16px; }
h1 { font-size: 22px; margin: 0; font-weight: 600; }
h1 small { color: var(--muted); font-weight: 400; font-size: 12px; margin-left: 8px; }
a.link { color: var(--accent); text-decoration: none; font-size: 13px; }
.bar { display: flex; flex-wrap: wrap; gap: 8px; align-items: center; background: var(--panel); border: 1px solid var(--border); border-radius: 12px; padding: 10px; margin-bottom: 16px; }
.presets { display: flex; flex-wrap: wrap; gap: 4px; }
.presets button, .btn { background: var(--panel2); color: var(--muted); border: 1px solid var(--border); border-radius: 8px; padding: 6px 10px; cursor: pointer; font: inherit; font-size: 13px; }
.presets button:hover, .btn:hover { color: var(--text); }
.presets button.on { background: var(--accent); border-color: var(--accent); color: #fff; }
.sep { flex: 1; }
label.f { display: flex; align-items: center; gap: 6px; color: var(--muted); font-size: 13px; }
input[type=date], input[type=number] { background: var(--panel2); color: var(--text); border: 1px solid var(--border); border-radius: 8px; padding: 5px 8px; font: inherit; color-scheme: dark; }
input[type=number] { width: 64px; }
.kpis { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 12px; margin-bottom: 16px; }
.kpi { background: var(--panel); border: 1px solid var(--border); border-radius: 12px; padding: 14px 16px; }
.kpi .l { color: var(--muted); font-size: 12px; text-transform: uppercase; letter-spacing: .05em; }
.kpi .v { font-size: 28px; font-weight: 600; margin-top: 4px; font-variant-numeric: tabular-nums; }
.kpi .s { color: var(--muted); font-size: 12px; margin-top: 2px; }
.kpi.hero { border-color: transparent; background: linear-gradient(135deg, #1b2a4a, #2a1d4a); }
.pos { color: var(--good); } .neg { color: var(--bad); }
.grid { display: grid; grid-template-columns: 2fr 1fr; gap: 12px; margin-bottom: 12px; }
@media (max-width: 900px) { .grid { grid-template-columns: 1fr; } }
.card { background: var(--panel); border: 1px solid var(--border); border-radius: 12px; padding: 14px 16px; min-width: 0; }
.card h3 { margin: 0 0 10px; font-size: 14px; font-weight: 600; color: var(--muted); }
.chart { position: relative; height: 280px; }
.live-dot { display: inline-block; width: 8px; height: 8px; border-radius: 50%; background: var(--good); margin-right: 6px; box-shadow: 0 0 0 0 rgba(34,197,94,.6); animation: pulse 2s infinite; }
@keyframes pulse { 70% { box-shadow: 0 0 0 8px rgba(34,197,94,0); } 100% { box-shadow: 0 0 0 0 rgba(34,197,94,0); } }
table { width: 100%; border-collapse: collapse; font-variant-numeric: tabular-nums; }
th, td { padding: 8px 10px; text-align: left; border-bottom: 1px solid var(--border); white-space: nowrap; }
th { color: var(--muted); font-weight: 500; font-size: 12px; text-transform: uppercase; cursor: pointer; user-select: none; }
tbody tr:hover { background: var(--panel2); }
td.r, th.r { text-align: right; }
.tbl { overflow-x: auto; max-height: 460px; overflow-y: auto; }
.tbl thead th { position: sticky; top: 0; background: var(--panel); }
.weekend { color: var(--muted); }
.empty { color: var(--muted); padding: 30px; text-align: center; }
.minibar { height: 6px; border-radius: 3px; background: var(--panel2); min-width: 80px; overflow: hidden; }
.minibar i { display: block; height: 100%; background: var(--accent); }
</style>
</head>
<body>
<div class="wrap">
  <header>
    <h1>Time Dashboard <small>updated __GENERATED__</small></h1>
    <a class="link" href="./weeks.html">Weekly timesheets &rarr;</a>
  </header>

  <div class="bar">
    <div class="presets" id="presets"></div>
    <span class="sep"></span>
    <label class="f">From <input type="date" id="from"></label>
    <label class="f">To <input type="date" id="to"></label>
    <label class="f" title="Expected hours per week, spread evenly Mon-Fri">Target h/week <input type="number" id="target" min="0" max="168" step="0.5"></label>
  </div>

  <div class="kpis" id="kpis"></div>

  <div class="grid">
    <div class="card"><h3>Hours per day vs target</h3><div class="chart"><canvas id="cDaily"></canvas></div></div>
    <div class="card"><h3>Average by weekday</h3><div class="chart"><canvas id="cWeekday"></canvas></div></div>
  </div>
  <div class="grid">
    <div class="card"><h3>Running balance (overtime &uarr; / owed &darr;)</h3><div class="chart"><canvas id="cBal"></canvas></div></div>
    <div class="card"><h3>Weekly totals</h3><div class="chart"><canvas id="cWeek"></canvas></div></div>
  </div>

  <div class="card">
    <h3>Daily detail</h3>
    <div class="tbl"><table id="tbl"></table></div>
  </div>
</div>

<script>
const SESSIONS = __SESSIONS__;
const DAY = 864e5, WD = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
const $ = id => document.getElementById(id);
const store = { get(k, d) { try { const v = localStorage.getItem('td.' + k); return v === null ? d : JSON.parse(v); } catch { return d; } },
                set(k, v) { try { localStorage.setItem('td.' + k, JSON.stringify(v)); } catch {} } };

// ---- dates (local time, keys as yyyy-mm-dd)
const key = d => d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
const parse = k => { const [y, m, d] = k.split('-').map(Number); return new Date(y, m - 1, d); };
const addDays = (d, n) => new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
const monday = d => addDays(d, -((d.getDay() + 6) % 7));
const today = () => { const n = new Date(); return new Date(n.getFullYear(), n.getMonth(), n.getDate()); };
const isWork = d => d.getDay() > 0 && d.getDay() < 6;
const fmtH = h => { const s = h < 0 ? '-' : ''; const m = Math.round(Math.abs(h) * 60); return s + Math.floor(m / 60) + 'h ' + String(m % 60).padStart(2, '0') + 'm'; };
const fmtSigned = h => (h > 0 ? '+' : '') + fmtH(h);
const fmtClock = d => d ? d.toTimeString().slice(0, 5) : '\u2014';
const firstData = SESSIONS.length ? parse(SESSIONS[0].s.slice(0, 10)) : today();

const PRESETS = [
  ['Today', () => [today(), today()]],
  ['Yesterday', () => [addDays(today(), -1), addDays(today(), -1)]],
  ['This week', () => [monday(today()), addDays(monday(today()), 6)]],
  ['Last week', () => [addDays(monday(today()), -7), addDays(monday(today()), -1)]],
  ['Last 7 days', () => [addDays(today(), -6), today()]],
  ['This month', () => { const t = today(); return [new Date(t.getFullYear(), t.getMonth(), 1), new Date(t.getFullYear(), t.getMonth() + 1, 0)]; }],
  ['Last month', () => { const t = today(); return [new Date(t.getFullYear(), t.getMonth() - 1, 1), new Date(t.getFullYear(), t.getMonth(), 0)]; }],
  ['Last 30 days', () => [addDays(today(), -29), today()]],
  ['Quarter', () => { const t = today(), q = Math.floor(t.getMonth() / 3) * 3; return [new Date(t.getFullYear(), q, 1), new Date(t.getFullYear(), q + 3, 0)]; }],
  ['This year', () => { const y = today().getFullYear(); return [new Date(y, 0, 1), new Date(y, 11, 31)]; }],
  ['All time', () => [firstData, today()]],
];

// ---- per-day worked hours (open session counts up to now)
function perDay() {
  const m = {}, first = {}, last = {}, open = {};
  for (const x of SESSIONS) {
    const S = new Date(x.s), E = x.e ? new Date(x.e) : new Date();
    // Split sessions that cross midnight into one piece per calendar day
    for (let s = S; s < E; ) {
      const mid = addDays(new Date(s.getFullYear(), s.getMonth(), s.getDate()), 1), e = E < mid ? E : mid, k = key(s);
      m[k] = (m[k] || 0) + (e - s) / 36e5;
      if (!first[k] || s < first[k]) first[k] = s;
      if (!last[k] || e > last[k]) last[k] = e;
      if (!x.e && e === E) open[k] = true;
      s = e;
    }
  }
  return { m, first, last, open };
}

let state = { preset: store.get('preset', 'This week'), from: null, to: null, week: store.get('weekTarget', 40), sort: ['date', -1] };
const charts = {};

function init() {
  $('presets').innerHTML = PRESETS.map(([n]) => `<button data-p="${n}">${n}</button>`).join('');
  $('presets').onclick = e => { const p = e.target.dataset.p; if (p) applyPreset(p); };
  $('from').onchange = $('to').onchange = () => { state.preset = null; store.set('preset', null); readRange(); render(); };
  $('target').value = state.week;
  $('target').oninput = () => { state.week = Math.max(0, +$('target').value || 0); store.set('weekTarget', state.week); render(); };
  Chart.defaults.color = '#8b95a9'; Chart.defaults.borderColor = '#232c40'; Chart.defaults.font.family = '"Segoe UI", system-ui, sans-serif';
  const p = state.preset && PRESETS.find(x => x[0] === state.preset) ? state.preset : null;
  if (p) applyPreset(p); else { const [f, t] = PRESETS[2][1](); $('from').value = store.get('from', key(f)); $('to').value = store.get('to', key(t)); readRange(); render(); }
  setInterval(() => { if (SESSIONS.some(s => !s.e)) render(true); }, 30000);
  // The embedded SESSIONS go stale in a tab left open (an open session would keep counting past its END), so re-fetch the page
  setInterval(() => location.reload(), 5 * 60e3);
}
function applyPreset(p) {
  const [f, t] = PRESETS.find(x => x[0] === p)[1]();
  state.preset = p; store.set('preset', p);
  $('from').value = key(f); $('to').value = key(t); readRange(); render();
}
function readRange() {
  let f = $('from').value ? parse($('from').value) : today(), t = $('to').value ? parse($('to').value) : today();
  if (f > t) [f, t] = [t, f];
  state.from = f; state.to = t; store.set('from', key(f)); store.set('to', key(t));
}

function render(liveOnly) {
  document.querySelectorAll('#presets button').forEach(b => b.classList.toggle('on', b.dataset.p === state.preset));
  const { m, first, last, open } = perDay(), T = state.week / 5, now = today();
  // Expected only counts days up to today (future days in the range don't count against you)
  const end = state.to < now ? state.to : now;
  // Today isn't over yet: it can add overtime but not count as a short day
  const expOf = (d, w) => !isWork(d) ? 0 : d < now ? T : Math.min(T, w);
  const days = [];
  for (let d = state.from; d <= end; d = addDays(d, 1)) {
    const k = key(d), w = m[k] || 0, exp = expOf(d, w);
    days.push({ k, d, w, exp, delta: w - exp, first: first[k], last: last[k], open: !!open[k] });
  }
  const worked = days.reduce((a, x) => a + x.w, 0), expected = days.reduce((a, x) => a + x.exp, 0), bal = worked - expected;
  const daysWorked = days.filter(x => x.w > 0).length;
  const ot = days.reduce((a, x) => a + Math.max(0, x.delta), 0), owedDays = days.reduce((a, x) => a + Math.max(0, -x.delta), 0);
  const todayK = key(now), todayW = m[todayK] || 0, isOpen = !!open[todayK];

  // All-time balance, independent of range
  let allBal = 0; for (let d = firstData; d <= now; d = addDays(d, 1)) { const w = m[key(d)] || 0; allBal += w - expOf(d, w); }

  const remain = Math.max(0, (isWork(now) ? T : 0) - todayW);
  const leaveAt = isOpen && remain > 0 ? new Date(Date.now() + remain * 36e5) : null;
  $('kpis').innerHTML = [
    kpi('Balance (range)', fmtSigned(bal), bal >= 0 ? 'pos' : 'neg', bal >= 0 ? 'overtime in selected range' : 'you owe in selected range', true),
    kpi('Worked', fmtH(worked), '', `of ${fmtH(expected)} expected`),
    kpi('Overtime days total', fmtH(ot), 'pos', 'sum of hours above target'),
    kpi('Short days total', fmtH(owedDays), 'neg', 'sum of hours below target'),
    kpi('Avg / day worked', fmtH(daysWorked ? worked / daysWorked : 0), '', `${daysWorked} day${daysWorked === 1 ? '' : 's'} worked`),
    kpi(`${isOpen ? '<span class="live-dot"></span>' : ''}Today`, fmtH(todayW), '', leaveAt ? `target reached at ${fmtClock(leaveAt)}` : (remain > 0 ? `${fmtH(remain)} left to target` : 'target reached')),
    kpi('All-time balance', fmtSigned(allBal), allBal >= 0 ? 'pos' : 'neg', `since ${key(firstData)}`),
  ].join('');

  drawDaily(days, T, liveOnly); drawBal(days, liveOnly); drawWeekday(days, T, liveOnly); drawWeek(days, liveOnly);
  drawTable(days);
}
const kpi = (l, v, cls, s, hero) => `<div class="kpi${hero ? ' hero' : ''}"><div class="l">${l}</div><div class="v ${cls}">${v}</div><div class="s">${s}</div></div>`;

function upsert(id, cfg, liveOnly) {
  if (charts[id]) { charts[id].data = cfg.data; charts[id].options = cfg.options; charts[id].update(liveOnly ? 'none' : undefined); }
  else charts[id] = new Chart($(id), cfg);
}
const tipH = { callbacks: { label: c => `${c.dataset.label}: ${fmtH(c.parsed.y)}` } };
const baseOpts = (extra = {}) => ({ responsive: true, maintainAspectRatio: false, interaction: { mode: 'index', intersect: false },
  plugins: { legend: { labels: { boxWidth: 10, boxHeight: 10 } }, tooltip: tipH }, scales: { x: { grid: { display: false } }, y: { beginAtZero: true, ticks: { callback: v => +v.toFixed(2) + 'h' } } }, ...extra });

function drawDaily(days, T, l) {
  upsert('cDaily', { type: 'bar', data: {
    labels: days.map(x => (days.length <= 14 ? WD[x.d.getDay()] + ' ' : '') + x.k.slice(5)),
    datasets: [
      { label: 'Worked', data: days.map(x => +x.w.toFixed(2)), borderRadius: 4, maxBarThickness: 36,
        backgroundColor: days.map(x => x.w === 0 ? '#2a3348' : x.delta >= 0 ? '#22c55e' : x.exp === 0 ? '#8b5cf6' : '#4f8cff') },
      { type: 'line', label: 'Target', data: days.map(x => x.exp), borderColor: '#f59e0b', borderDash: [5, 4], pointRadius: 0, stepped: 'middle', borderWidth: 1.5 },
    ] }, options: baseOpts() }, l);
}
function drawBal(days, l) {
  let c = 0; const pts = days.map(x => +(c += x.delta).toFixed(2));
  upsert('cBal', { type: 'line', data: { labels: days.map(x => x.k.slice(5)), datasets: [{ label: 'Balance', data: pts, tension: .25, pointRadius: days.length > 40 ? 0 : 2,
    borderColor: '#8b5cf6', fill: { target: 'origin', above: 'rgba(34,197,94,.18)', below: 'rgba(244,63,94,.18)' } }] },
    options: baseOpts({ plugins: { legend: { display: false }, tooltip: { callbacks: { label: c => 'Balance: ' + fmtSigned(c.parsed.y) } } }, scales: { x: { grid: { display: false } }, y: { ticks: { callback: v => fmtH(v) } } } }) }, l);
}
function drawWeekday(days, T, l) {
  const sum = [0,0,0,0,0,0,0], n = [0,0,0,0,0,0,0];
  days.forEach(x => { if (x.w > 0) { sum[x.d.getDay()] += x.w; n[x.d.getDay()]++; } });
  const order = [1,2,3,4,5,6,0];
  upsert('cWeekday', { type: 'bar', data: { labels: order.map(i => WD[i]), datasets: [
    { label: 'Avg worked', data: order.map(i => n[i] ? +(sum[i] / n[i]).toFixed(2) : 0), backgroundColor: '#4f8cff', borderRadius: 4 },
    { type: 'line', label: 'Target', data: order.map(i => i > 0 && i < 6 ? T : 0), borderColor: '#f59e0b', borderDash: [5, 4], pointRadius: 0, borderWidth: 1.5 },
  ] }, options: baseOpts() }, l);
}
function drawWeek(days, l) {
  const w = {}; days.forEach(x => { const k = key(monday(x.d)); (w[k] ||= { h: 0, e: 0 }); w[k].h += x.w; w[k].e += x.exp; });
  const ks = Object.keys(w).sort();
  upsert('cWeek', { type: 'bar', data: { labels: ks.map(k => k.slice(5)), datasets: [
    { label: 'Worked', data: ks.map(k => +w[k].h.toFixed(2)), backgroundColor: ks.map(k => w[k].h >= w[k].e ? '#22c55e' : '#4f8cff'), borderRadius: 4, maxBarThickness: 40 },
    { type: 'line', label: 'Expected', data: ks.map(k => w[k].e), borderColor: '#f59e0b', borderDash: [5, 4], pointRadius: 2, borderWidth: 1.5 },
  ] }, options: baseOpts() }, l);
}

function drawTable(days) {
  const [col, dir] = state.sort;
  const rows = days.filter(x => x.w > 0 || x.exp > 0).slice().sort((a, b) => {
    const v = { date: x => x.k, worked: x => x.w, delta: x => x.delta, first: x => x.first ? x.first.getHours() * 60 + x.first.getMinutes() : -1, last: x => x.last ? x.last.getHours() * 60 + x.last.getMinutes() : -1 }[col];
    return (v(a) > v(b) ? 1 : v(a) < v(b) ? -1 : 0) * dir;
  });
  const max = Math.max(state.week / 5, ...days.map(x => x.w), 1);
  const th = (c, t, r) => `<th class="${r ? 'r' : ''}" data-c="${c}">${t}${col === c ? (dir > 0 ? ' \u25B2' : ' \u25BC') : ''}</th>`;
  $('tbl').innerHTML = rows.length ? `<thead><tr>${th('date', 'Date')}<th>Day</th>${th('first', 'First in')}${th('last', 'Last out')}${th('worked', 'Worked', 1)}<th></th>${th('delta', 'Over / owed', 1)}</tr></thead><tbody>` +
    rows.map(x => `<tr class="${isWork(x.d) ? '' : 'weekend'}"><td>${x.k}</td><td>${WD[x.d.getDay()]}</td><td>${fmtClock(x.first)}</td><td>${x.open ? '<span class="live-dot"></span>now' : fmtClock(x.last)}</td>
      <td class="r">${fmtH(x.w)}</td><td><div class="minibar"><i style="width:${(x.w / max * 100).toFixed(1)}%"></i></div></td>
      <td class="r ${x.delta >= 0 ? 'pos' : 'neg'}">${fmtSigned(x.delta)}</td></tr>`).join('') + '</tbody>'
    : '<tbody><tr><td class="empty">No data in this range</td></tr></tbody>';
  $('tbl').querySelectorAll('th[data-c]').forEach(h => h.onclick = () => { const c = h.dataset.c; state.sort = [c, state.sort[0] === c ? -state.sort[1] : -1]; drawTable(days); });
}

init();
</script>
</body>
</html>
'@
$DashTemplate.Replace('__SESSIONS__', $SessionsJson).Replace('__GENERATED__', (Get-Date).ToString('yyyy-MM-dd HH:mm')) | Out-File "$BasePath\index.html" -Encoding utf8

# 5. AUTO-SYNC TO GITHUB
Set-Location $BasePath
if (Test-Path "$BasePath\.git") {
    # Wait 10 seconds for Wi-Fi to connect after waking up
    Start-Sleep -Seconds 10
    
    # The scheduled task runs as a different user whose PATH has no git, so use the full path
    $Git = "C:\Users\DMarier\AppData\Local\Programs\Git\cmd\git.exe"
    # safe.directory: this user doesn't own the repo folder; author hardcoded so no profile is needed
    $G = @("-c", "safe.directory=C:/Timesheets", "-c", "user.name=Timesheet Bot", "-c", "user.email=timesheet@local", "-c", "credential.interactive=false")
    # Never wait for a sign-in window (the task runs hidden): fail fast and log instead
    $env:GIT_TERMINAL_PROMPT = "0"; $env:GCM_INTERACTIVE = "never"

    # Backup of this script (the live copy stays in C:\Scripts); refreshed every run so it never goes stale
    Copy-Item $PSCommandPath "$BasePath\TimeLogger.ps1" -Force
    $out = & $Git @G add . 2>&1
    $out += & $Git @G commit -m "Auto-sync timesheet" --quiet 2>&1
    # Pick up edits made on GitHub (e.g. README) first, or every push gets rejected as non-fast-forward
    $out += & $Git @G pull --rebase --quiet origin main 2>&1
    # On a conflict, don't leave the repo stuck mid-rebase; the push below then fails and gets logged
    if ($LASTEXITCODE -ne 0) { $out += & $Git @G rebase --abort 2>&1 }
    $out += & $Git @G push origin main --quiet 2>&1
    if ($LASTEXITCODE -ne 0) {
        # Native commands don't throw, so check the exit code and log the real output
        "$(Get-Date -Format s) push failed:`n$($out | Out-String)" | Out-File "$BasePath\git_error.log" -Append
    }
}
