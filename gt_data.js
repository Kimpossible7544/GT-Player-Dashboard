// gt_data.js — fetches and parses the GT Excel file from Dropbox
// Serves gt_standings.html, gt_dashboard.html, and gt_projections.html
// Depends on SheetJS (xlsx.full.min.js) already loaded on the page.

// Returns true if the week has started (start date has arrived).
function hasWeekStarted(weekLabel) {
  const parts = weekLabel.split(" - ");
  if (parts.length < 2) return true;

  const startPart = parts[0].trim();
  const now = new Date();
  const year = now.getFullYear();

  let startDate = new Date(`${startPart}, ${year}`);
  if (isNaN(startDate.getTime())) return true;

  if (startDate - now > 180 * 24 * 60 * 60 * 1000) {
    startDate = new Date(`${startPart}, ${year - 1}`);
  }

  return now >= startDate;
}

// Returns true if the week's Friday end date has fully passed.
function isWeekComplete(weekLabel) {
  const parts = weekLabel.split(" - ");
  if (parts.length < 2) return true;

  const endPart = parts[1].trim();
  const now = new Date();
  const year = now.getFullYear();

  let endDate = new Date(`${endPart}, ${year}`);
  if (isNaN(endDate.getTime())) return true;

  if (endDate - now > 180 * 24 * 60 * 60 * 1000) {
    endDate = new Date(`${endPart}, ${year - 1}`);
  }

  const completionDate = new Date(endDate);
  completionDate.setDate(completionDate.getDate() + 1);

  return now >= completionDate;
}

// Season start date — all charts and data begin here.
const SEASON_START = new Date("July 13, 2026");

// Returns true if the week begins on or after the season start date.
function isOnOrAfterSeasonStart(weekLabel) {
  const parts = weekLabel.split(" - ");
  if (parts.length < 2) return true;

  const startPart = parts[0].trim();
  const now = new Date();
  const year = now.getFullYear();

  let startDate = new Date(`${startPart}, ${year}`);
  if (isNaN(startDate.getTime())) return true;

  if (startDate - now > 180 * 24 * 60 * 60 * 1000) {
    startDate = new Date(`${startPart}, ${year - 1}`);
  }

  return startDate >= SEASON_START;
}

// Parses a power value into a float expressed in MILLIONS.
// Values are normally entered in millions ("182.6 M", "182.6", 157.7), but a
// few cells hold the full raw count (e.g. 161660000). Anything >= 100000 is
// treated as a raw count and divided by 1,000,000 so every value is in millions.
function parsePower(val) {
  if (!val && val !== 0) return null;
  const str = String(val).replace(/Mil/gi, "").replace(/M/gi, "").replace(/\s/g, "").trim();
  let num = parseFloat(str);
  if (isNaN(num)) return null;
  if (Math.abs(num) >= 100000) num = num / 1000000;
  return num;
}

// Formats a date-or-string cell into an MM/DD/YYYY string.
function fmtCellDate(v) {
  if (v instanceof Date) {
    return String(v.getMonth() + 1).padStart(2, "0") + "/" +
           String(v.getDate()).padStart(2, "0") + "/" + v.getFullYear();
  }
  return String(v);
}

// Parses one Arena Power sheet row into a growth object.
// Column layout (0-indexed):
//   0=Name, 1=Date, 2=Level, 3=Arena Power, 4=HQ Power,
//   5=Δ Arena session, 6=Δ HQ session, 7=Δ Arena overall, 8=Δ HQ overall,
//   9=Level note
//   History groups of 4 starting at col 10 (date,level,arena,HQ), oldest furthest right.
function parseArenaRow(row) {
  const currentArena = parsePower(row[3]);
  const currentHQ    = parsePower(row[4]);
  const currentLevel = row[2] || null;

  let firstArena = null;
  let firstHQ    = null;
  let firstLevel = null;
  let baselineDate = null;

  for (let c = 12; c < row.length; c += 4) {
    const v = parsePower(row[c]);
    if (v !== null) firstArena = v;
  }
  for (let c = 13; c < row.length; c += 4) {
    const v = parsePower(row[c]);
    if (v !== null) firstHQ = v;
  }
  for (let c = 11; c < row.length; c += 4) {
    const v = row[c];
    if (v !== null && v !== "") firstLevel = v;
  }
  for (let c = 10; c < row.length; c += 4) {
    const v = row[c];
    if (v !== null && v !== "") baselineDate = fmtCellDate(v);
  }

  // If no history yet, first = current
  if (firstArena === null) firstArena = currentArena;
  if (firstHQ    === null) firstHQ    = currentHQ;
  if (firstLevel === null) firstLevel = currentLevel;
  if (baselineDate === null && row[1]) baselineDate = fmtCellDate(row[1]);

  return {
    currentLevel,
    currentArena,
    currentHQ,
    firstLevel,
    firstArena,
    firstHQ,
    baselineDate,
    deltaArenaSession:  parsePower(row[5]),
    deltaHQSession:     parsePower(row[6]),
    deltaArenaOverall:  parsePower(row[7]),
    deltaHQOverall:     parsePower(row[8]),
    levelNote:          row[9] || null
  };
}

// Formats a date-or-string cell into an MM/DD/YYYY string.
function fmtCellDate(v) {
  if (v instanceof Date) {
    return String(v.getMonth() + 1).padStart(2, "0") + "/" +
           String(v.getDate()).padStart(2, "0") + "/" + v.getFullYear();
  }
  return String(v);
}

// Parses one Arena Power sheet row into a growth object.
// Column layout (0-indexed):
//   0=Name, 1=Date, 2=Level, 3=Arena Power, 4=HQ Power,
//   5=Δ Arena session, 6=Δ HQ session, 7=Δ Arena overall, 8=Δ HQ overall,
//   9=Level note
//   History groups of 4 starting at col 10 (date,level,arena,HQ), oldest furthest right.
function parseArenaRow(row) {
  const currentArena = parsePower(row[3]);
  const currentHQ    = parsePower(row[4]);
  const currentLevel = row[2] || null;

  let firstArena = null;
  let firstHQ    = null;
  let firstLevel = null;
  let baselineDate = null;

  for (let c = 12; c < row.length; c += 4) {
    const v = parsePower(row[c]);
    if (v !== null) firstArena = v;
  }
  for (let c = 13; c < row.length; c += 4) {
    const v = parsePower(row[c]);
    if (v !== null) firstHQ = v;
  }
  for (let c = 11; c < row.length; c += 4) {
    const v = row[c];
    if (v !== null && v !== "") firstLevel = v;
  }
  for (let c = 10; c < row.length; c += 4) {
    const v = row[c];
    if (v !== null && v !== "") baselineDate = fmtCellDate(v);
  }

  // If no history yet, first = current
  if (firstArena === null) firstArena = currentArena;
  if (firstHQ    === null) firstHQ    = currentHQ;
  if (firstLevel === null) firstLevel = currentLevel;
  if (baselineDate === null && row[1]) baselineDate = fmtCellDate(row[1]);

  return {
    currentLevel,
    currentArena,
    currentHQ,
    currentDate:        fmtCellDate(row[1]),
    firstLevel,
    firstArena,
    firstHQ,
    baselineDate,
    deltaArenaSession:  parsePower(row[5]),
    deltaHQSession:     parsePower(row[6]),
    deltaArenaOverall:  parsePower(row[7]),
    deltaHQOverall:     parsePower(row[8]),
    levelNote:          row[9] || null
  };
}

async function fetchWithRetry(url, options = {}, retries = 3, baseDelay = 800) {
  let lastErr;
  let delay = baseDelay;

  for (let i = 0; i < retries; i++) {
    const cacheBuster = (url.includes("?") ? "&" : "?") + "cb=" + Date.now();
    const fullUrl = url + cacheBuster;
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 60000);

    try {
      const resp = await fetch(fullUrl, { ...options, cache: "no-store", signal: controller.signal });
      clearTimeout(timeoutId);

      if (resp.ok) return resp;

      if (resp.status >= 400 && resp.status < 500 && resp.status !== 429) {
        throw new Error(`Dropbox returned ${resp.status} ${resp.statusText}. Make sure the file is shared as 'Anyone with the link'.`);
      }
      throw new Error(`Dropbox returned ${resp.status} ${resp.statusText}.`);
    } catch (err) {
      clearTimeout(timeoutId);
      if (err.name === "AbortError") {
        lastErr = new Error("Dropbox took longer than 60 seconds to respond.");
      } else {
        lastErr = err;
      }
      if (i < retries - 1) {
        console.warn(`[GT] Fetch attempt ${i + 1}/${retries} failed: ${lastErr.message}; retrying in ${delay}ms`);
        await new Promise(resolve => setTimeout(resolve, delay));
        delay *= 2;
      }
    }
  }

  throw lastErr;
}

async function loadGTData() {

  // =========================================================
  // DROPBOX FILE LOCATION
  // =========================================================
  const DROPBOX_URL =
    "https://dl.dropboxusercontent.com/scl/fi/twt9mo3vvvngnx2cgrht6/GTStatsFINAL.xlsm" +
    "?rlkey=u2zj5dscvonfqvtgd0opaapvo&raw=1";

  // Previous team's workbook — used only to pull Arena/HQ power history for
  // players who were on BOTH teams (matched by roster ID) so their growth
  // spans both. See CROSS_TEAM_PLAYER_IDS below.
  const WPX_DROPBOX_URL =
    "https://dl.dropboxusercontent.com/scl/fi/dx7xgqmjshf8hciso3uya/WPXStatsFinal.xlsm" +
    "?rlkey=oyw14lm3fod48uxygykdnixar&raw=1";

  // Roster IDs (stable across teams) whose WPX Arena/HQ power history should be
  // merged into their GT growth card and shown as a legacy WPX History card.
  // Leave this array empty to auto-detect every ID that exists in both rosters.
  const CROSS_TEAM_PLAYER_IDS = [];

  const PLAYER_TRACKING_SHEET = "Player Tracking";
  const WEEK_SETTINGS_SHEET   = "Week Settings";
  const ARENA_POWER_SHEET     = "Arena Power";

  const DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];

  const DAILY_GOAL  = 6000000;
  const WEEKLY_GOAL = 20000000;

  // =========================================================
  // VERIFY SHEETJS EXISTS
  // =========================================================
  if (typeof XLSX === "undefined") {
    throw new Error(
      "SheetJS (XLSX) is not loaded. " +
      "Make sure xlsx.full.min.js loads BEFORE gt_data.js"
    );
  }

  // =========================================================
  // FETCH EXCEL FILE
  // =========================================================
  let response;
  try {
    response = await fetchWithRetry(DROPBOX_URL);
  } catch (err) {
    console.error("[GT] Fetch failed after retries:", err);
    throw new Error("Could not reach Dropbox after multiple attempts. Details: " + err.message);
  }

  if (!response.ok) {
    throw new Error(
      `Dropbox returned ${response.status} ${response.statusText}. ` +
      "Make sure the file is shared as 'Anyone with the link'."
    );
  }

  const buffer = await response.arrayBuffer();

  if (!buffer || buffer.byteLength < 1000) {
    throw new Error("Downloaded file is invalid or too small.");
  }

  // =========================================================
  // PARSE EXCEL FILE
  // =========================================================
  let workbook;
  try {
    workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  } catch (err) {
    console.error("[GT] XLSX parse failed:", err);
    throw new Error("SheetJS could not parse the Excel file: " + err.message);
  }

  // =========================================================
  // VERIFY PLAYER TRACKING SHEET EXISTS
  // =========================================================
  if (!workbook.SheetNames.includes(PLAYER_TRACKING_SHEET)) {
    throw new Error(
      `Sheet "${PLAYER_TRACKING_SHEET}" not found.\n` +
      `Available sheets: ${workbook.SheetNames.join(", ")}`
    );
  }

  // =========================================================
  // READ PLAYER TRACKING SHEET
  // =========================================================
  const ptRows = XLSX.utils.sheet_to_json(
    workbook.Sheets[PLAYER_TRACKING_SHEET],
    { header: 1, defval: null }
  );

  if (ptRows.length < 2) {
    throw new Error("Player Tracking sheet is empty.");
  }

  const headers = ptRows[0];

  // =========================================================
  // PARSE WEEK COLUMNS
  // =========================================================
  const weekLabels    = [];
  const weekTotalCols = [];
  const weekRankCols  = [];

  let weekStartCol = -1;
  for (let c = 0; c < headers.length; c++) {
    const h = headers[c];
    if (h && typeof h === "string" && h.match(/Weekly Total \(.+\)/)) {
      weekStartCol = c;
      break;
    }
  }

  if (weekStartCol < 0) throw new Error("No 'Weekly Total' columns found.");

  for (let c = weekStartCol; c < headers.length; c += 2) {
    const h = headers[c];
    if (!h || typeof h !== "string") break;
    const match = h.match(/Weekly Total \((.+)\)/);
    if (!match) break;
    const label = match[1];
    if (!hasWeekStarted(label)) {
      console.log(`[GT] Skipping future week: ${label}`);
      continue;
    }
    if (!isOnOrAfterSeasonStart(label)) {
      console.log(`[GT] Skipping pre-season week: ${label}`);
      continue;
    }
    weekLabels.push(label);
    weekTotalCols.push(c);
    weekRankCols.push(c + 1);
  }

  if (weekLabels.length === 0) throw new Error("No valid weeks found.");

  const currentWeekLabel = weekLabels.find(label => !isWeekComplete(label)) || null;

  weekLabels.reverse();
  weekTotalCols.reverse();
  weekRankCols.reverse();

  // =========================================================
  // BUILD PLAYER OBJECTS
  // =========================================================
  const colOf = (label) => headers.indexOf(label);
  const COL_JOIN_DATE   = 1; // Column B — player join date
  const COL_TOTAL       = colOf("Overall Total");
  const COL_PUSH_TOTAL  = colOf("Pushing Total");
  const COL_RANK        = colOf("Overall Rank");
  const COL_PUSH_RANK   = colOf("Pushing Rank");
  const COL_AVG         = colOf("Overall Weekly Average");
  const COL_MISS_DAILY  = colOf("Overall Missed Daily Goals");
  const COL_MISS_WEEKLY = colOf("Overall Missed Weekly Goals");

  const players = {};

  for (let r = 1; r < ptRows.length; r++) {
    const row  = ptRows[r];
    const name = row[0];
    if (!name) continue;

    // Parse join date from column B
    let joinDate = null;
    const rawDate = row[COL_JOIN_DATE];
    if (rawDate instanceof Date) {
      joinDate = String(rawDate.getMonth()+1).padStart(2,"0") + "/" +
                 String(rawDate.getDate()).padStart(2,"0") + "/" + rawDate.getFullYear();
    } else if (rawDate) {
      joinDate = String(rawDate);
    }

    players[name] = {
      name,
      joinDate,
      totalScore:   (COL_TOTAL      >= 0 ? row[COL_TOTAL]       : row[1]) || 0,
      pushingTotal: (COL_PUSH_TOTAL  >= 0 ? row[COL_PUSH_TOTAL]  : 0)     || 0,
      overallRank:  (COL_RANK        >= 0 ? row[COL_RANK]        : null)   ||
                    (COL_PUSH_RANK   >= 0 ? row[COL_PUSH_RANK]   : null),
      pushingRank:  (COL_PUSH_RANK   >= 0 ? row[COL_PUSH_RANK]   : null),
      weeklyAvg:    (COL_AVG         >= 0 ? row[COL_AVG]         : row[3]) || 0,
      missedDaily:  (COL_MISS_DAILY  >= 0 ? row[COL_MISS_DAILY]  : row[4]) || 0,
      missedWeekly: (COL_MISS_WEEKLY >= 0 ? row[COL_MISS_WEEKLY] : row[5]) || 0,
      weeks: weekTotalCols.map((col, i) => ({
        label:      weekLabels[i],
        score:      row[col]            || 0,
        rank:       row[weekRankCols[i]] || null,
        inProgress: weekLabels[i] === currentWeekLabel
      })),
      // growth populated below from Arena Power sheet
      growth: null
    };
  }

  if (Object.keys(players).length === 0) throw new Error("No player data found.");

  // =========================================================
  // READ WEEK SETTINGS SHEET
  // =========================================================
  const notPushingWeeks = new Set();
  const serverHelpers   = new Map();

  if (workbook.SheetNames.includes(WEEK_SETTINGS_SHEET)) {
    const wsRows = XLSX.utils.sheet_to_json(
      workbook.Sheets[WEEK_SETTINGS_SHEET],
      { header: 1, defval: null }
    );

    for (let r = 1; r < wsRows.length; r++) {
      const weekLabel = wsRows[r][0];
      const pushing   = String(wsRows[r][1] || "").trim().toUpperCase();

      if (weekLabel && pushing === "N") {
        notPushingWeeks.add(String(weekLabel).trim());
      }

      const helpPlayers = String(wsRows[r][2] || "").trim();
      if (weekLabel && helpPlayers) {
        helpPlayers.split(",").forEach(name => {
          const trimmed = name.trim();
          if (trimmed) serverHelpers.set(trimmed + "|||" + String(weekLabel).trim(), true);
        });
      }
    }
  }

  Object.values(players).forEach(player => {
    player.weeks.forEach(week => {
      week.pushing    = !notPushingWeeks.has(week.label);
      const key       = player.name + "|||" + week.label;
      week.serverHelp = serverHelpers.has(key);
    });
  });

  // =========================================================
  // BUILD DAILY DATA
  // =========================================================
  const dailyData = { weekOrder: weekLabels, players: {} };

  Object.keys(players).forEach(name => {
    dailyData.players[name] = { Monday: [], Tuesday: [], Wednesday: [], Thursday: [], Friday: [], rank: [] };
  });

  for (const weekLabel of weekLabels) {
    if (!workbook.SheetNames.includes(weekLabel)) {
      Object.keys(players).forEach(name => {
        DAYS.forEach(day => dailyData.players[name][day].push(0));
        dailyData.players[name].rank.push(null);
      });
      continue;
    }

    const weekRows = XLSX.utils.sheet_to_json(
      workbook.Sheets[weekLabel],
      { header: 1, defval: null }
    );

    const rankCol = weekRows[0] ? weekRows[0].findIndex(h => String(h || "").trim().toLowerCase() === "rank") : -1;

    const byPlayer = {};
    for (let r = 1; r < weekRows.length; r++) {
      const row  = weekRows[r];
      const name = row[0];
      if (!name) continue;
      byPlayer[name] = {
        Monday:    (typeof row[2] === 'number' ? row[2] : 0),
        Tuesday:   (typeof row[3] === 'number' ? row[3] : 0),
        Wednesday: (typeof row[4] === 'number' ? row[4] : 0),
        Thursday:  (typeof row[5] === 'number' ? row[5] : 0),
        Friday:    (typeof row[6] === 'number' ? row[6] : 0),
        rank:      (rankCol >= 0 ? String(row[rankCol] || "").trim() || null : null)
      };
    }

    Object.keys(players).forEach(name => {
      const pd = byPlayer[name];
      DAYS.forEach(day => {
        dailyData.players[name][day].push(pd ? (pd[day] || 0) : 0);
      });
      dailyData.players[name].rank.push(pd ? (pd.rank || null) : null);
    });
  }

  // =========================================================
  // RECALCULATE MISSED GOALS
  // =========================================================
  Object.values(players).forEach(player => {
    let missedDaily = 0, missedWeekly = 0;

    player.weeks.forEach((week, i) => {
      if (!week.pushing || week.inProgress) return;
      if (!week.score || week.score <= 0) return;

      if (week.score < WEEKLY_GOAL) missedWeekly++;

      const pd      = dailyData.players[player.name];
      const weekIdx = dailyData.weekOrder.indexOf(week.label);
      if (pd && weekIdx >= 0) {
        DAYS.forEach(day => {
          const dayScore = pd[day][weekIdx];
          if (dayScore > 0 && dayScore < DAILY_GOAL) missedDaily++;
        });
      }
    });

    player.missedDaily  = missedDaily;
    player.missedWeekly = missedWeekly;
  });

  // =========================================================
  // READ ARENA POWER SHEET
  // =========================================================
  // Column layout (0-indexed):
  //   0=Name, 1=Date, 2=Level, 3=Arena Power, 4=HQ Power,
  //   5=Δ Arena session, 6=Δ HQ session,
  //   7=Δ Arena overall, 8=Δ HQ overall,
  //   9=Level note
  //   History groups of 4 starting at col 10:
  //     10=date, 11=level, 12=arena, 13=HQ, 14=date, 15=level, 16=arena, 17=HQ ...
  // =========================================================

  if (workbook.SheetNames.includes(ARENA_POWER_SHEET)) {
    const apRows = XLSX.utils.sheet_to_json(
      workbook.Sheets[ARENA_POWER_SHEET],
      { header: 1, defval: null }
    );

    for (let r = 1; r < apRows.length; r++) {
      const row  = apRows[r];
      const name = row[0];
      if (!name) continue;

      const growth = parseArenaRow(row);

      // Merge into player object if name matches
      if (players[name]) {
        players[name].growth = growth;
      } else {
        // Try case-insensitive match
        const key = Object.keys(players).find(
          k => k.toLowerCase() === name.toLowerCase()
        );
        if (key) players[key].growth = growth;
      }
    }

    console.log("[GT] Arena Power sheet parsed.");
  } else {
    console.warn("[GT] Arena Power sheet not found — growth data unavailable.");
  }

  // =========================================================
  // BUILD ID -> PLAYER NAME LOOKUP FROM ROSTER SHEET
  // =========================================================
  // Roster layout: ID in cols A(0),E(4),I(8),M(12) — Player Name in B(1),F(5),J(9),N(13)
  // Rank group headers are in the first row above each name column (R4/R5, R3, R2, R1).
  // =========================================================
  const idToPlayer = {};
  const rosterRanks = {};
  const ROSTER_SHEET = "Roster";

  if (workbook.SheetNames.includes(ROSTER_SHEET)) {
    const rosterRows = XLSX.utils.sheet_to_json(
      workbook.Sheets[ROSTER_SHEET],
      { header: 1, defval: null }
    );

    const idCols   = [0, 4, 8, 12];
    const nameCols = [1, 5, 9, 13];
    const akaCols  = [3, 7, 11, 15];
    const rankHeaders = rosterRows[0] ? [
      rosterRows[0][1],
      rosterRows[0][5],
      rosterRows[0][9],
      rosterRows[0][13]
    ] : [];

    for (let s = 0; s < idCols.length; s++) {
      const header = String(rankHeaders[s] || "");
      const rank = header.split(" - ")[0].trim() || null;
      for (let r = 1; r < rosterRows.length; r++) {
        const row = rosterRows[r];
        const id   = row[idCols[s]];
        const name = row[nameCols[s]];
        const aka  = row[akaCols[s]];
        // Skip header row and non-numeric IDs
        if (id && name && !isNaN(Number(id))) {
          idToPlayer[Number(id)] = String(name).trim();
          if (rank) {
            const n = String(name).trim();
            rosterRanks[n] = rank;
            if (aka) rosterRanks[String(aka).trim()] = rank;
          }
        }
      }
    }
    console.log("[GT] Roster IDs loaded:", Object.keys(idToPlayer).length);
  } else {
    console.warn("[GT] Roster sheet not found — ID login unavailable.");
  }

  // =========================================================
  // CROSS-TEAM ARENA/HQ POWER MERGE (WPX -> GT)
  // For players who were on BOTH teams (matched by stable roster ID), pull
  // their WPX Arena Power + HQ (personal) power history so their growth card
  // spans both teams. GT's own Arena Power data (once present) stays "current"
  // and the WPX baseline becomes the "starting" values; deltas are recomputed
  // across the full span. Auto-detects IDs present in both rosters unless
  // CROSS_TEAM_PLAYER_IDS is explicitly populated.
  // =========================================================
  if (WPX_DROPBOX_URL) {
    try {
      const wpxResp = await fetchWithRetry(WPX_DROPBOX_URL);
      if (!wpxResp.ok) throw new Error(`WPX Dropbox returned ${wpxResp.status}`);
      const wpxWorkbook = XLSX.read(await wpxResp.arrayBuffer(), { type: "array", cellDates: true });

      // WPX roster: ID -> player name
      const wpxIdToPlayer = {};
      if (wpxWorkbook.SheetNames.includes(ROSTER_SHEET)) {
        const wr = XLSX.utils.sheet_to_json(wpxWorkbook.Sheets[ROSTER_SHEET], { header: 1, defval: null });
        const idCols = [0, 4, 8, 12];
        const nameCols = [1, 5, 9, 13];
        for (let r = 1; r < wr.length; r++) {
          const row = wr[r];
          for (let s = 0; s < idCols.length; s++) {
            const id = row[idCols[s]];
            const nm = row[nameCols[s]];
            if (id && nm && !isNaN(Number(id))) wpxIdToPlayer[Number(id)] = String(nm).trim();
          }
        }
      }

      // WPX Arena Power: lowercased name -> row
      const wpxArenaByName = {};
      if (wpxWorkbook.SheetNames.includes(ARENA_POWER_SHEET)) {
        const ar = XLSX.utils.sheet_to_json(wpxWorkbook.Sheets[ARENA_POWER_SHEET], { header: 1, defval: null });
        for (let r = 1; r < ar.length; r++) {
          const row = ar[r];
          if (row[0]) wpxArenaByName[String(row[0]).trim().toLowerCase()] = row;
        }
      }

      const diff = (a, b) => (a !== null && b !== null) ? Math.round((a - b) * 10) / 10 : null;

      const crossTeamIds = CROSS_TEAM_PLAYER_IDS.length
        ? CROSS_TEAM_PLAYER_IDS
        : Object.keys(wpxIdToPlayer).filter(id => idToPlayer[id]);

      crossTeamIds.forEach(id => {
        const gtName = idToPlayer[id];
        const wpxName = wpxIdToPlayer[id];
        if (!gtName || !wpxName || !players[gtName]) return;
        const wpxRow = wpxArenaByName[wpxName.toLowerCase()];
        if (!wpxRow) return;

        const wpxG = parseArenaRow(wpxRow);
        const gtG = players[gtName].growth || null;
        const hasGtCurrent = !!(gtG && (gtG.currentArena !== null || gtG.currentHQ !== null));

        const currentArena = hasGtCurrent ? gtG.currentArena : wpxG.currentArena;
        const currentHQ    = hasGtCurrent ? gtG.currentHQ    : wpxG.currentHQ;
        const currentLevel = hasGtCurrent ? gtG.currentLevel : wpxG.currentLevel;

        players[gtName].growth = {
          currentLevel,
          currentArena,
          currentHQ,
          firstLevel:   wpxG.firstLevel,
          firstArena:   wpxG.firstArena,
          firstHQ:      wpxG.firstHQ,
          baselineDate: wpxG.baselineDate,
          deltaArenaSession: hasGtCurrent ? gtG.deltaArenaSession : wpxG.deltaArenaSession,
          deltaHQSession:    hasGtCurrent ? gtG.deltaHQSession    : wpxG.deltaHQSession,
          deltaArenaOverall: diff(currentArena, wpxG.firstArena),
          deltaHQOverall:    diff(currentHQ, wpxG.firstHQ),
          levelNote:         hasGtCurrent ? gtG.levelNote : wpxG.levelNote,
          crossTeam:         true
        };
        players[gtName].wpxHistory = wpxG;
        console.log(`[GT] Merged WPX Arena/HQ power for cross-team player ${gtName} (ID ${id}).`);
      });
    } catch (err) {
      console.warn("[GT] WPX cross-team power merge skipped:", err.message);
    }
  }

  // =========================================================
  // FINAL LOGGING
  // =========================================================
  console.log(
    `[GT] Loaded ${Object.keys(players).length} players ` +
    `across ${weekLabels.length} weeks`,
    weekLabels
  );

  // =========================================================
  // RETURN FINAL DATA
  // =========================================================
  return {
    weekLabels,
    playerList: Object.keys(players),
    players,
    dailyData,
    currentWeekLabel,
    notPushingWeeks,
    serverHelpers,
    idToPlayer,
    rosterRanks,
    DAILY_GOAL,
    WEEKLY_GOAL
  };
}
