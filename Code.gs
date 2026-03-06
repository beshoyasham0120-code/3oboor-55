const SHEET_NAMES = {
  QUESTIONS: 'Questions',
  ROUNDS: 'Rounds',
  ANSWERS: 'Answers'
};

const HEADERS = {
  Questions: ['category', 'points', 'question', 'answer', 'isDouble'],
  Rounds: ['roundId', 'sessionId', 'hostName', 'team1Number', 'team2Number', 'score1', 'score2', 'answeredCount', 'doubleSelected', 'mainTimerRemaining', 'mainTimerRunning', 'mainTimerStartedAt', 'eventType', 'timestamp', 'answeredJson', 'activeQuestionJson', 'qTimerRemaining', 'qTimerRunning', 'qTimerStartedAt', 'extraJson', 'updatedAt'],
  Answers: ['roundId', 'sessionId', 'timestamp', 'categoryName', 'questionText', 'basePoints', 'awardedPoints', 'isDouble', 'result', 'winnerTeam', 'winnerTeamNumber', 'score1', 'score2']
};

function doGet(e) {
  return handleRequest_(e, null);
}

function doPost(e) {
  let body = {};
  try {
    const raw = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    body = raw ? JSON.parse(raw) : {};
  } catch (err) {
    return json_({ ok: false, error: 'Invalid JSON body' });
  }
  return handleRequest_(e, body);
}

function handleRequest_(e, body) {
  try {
    ensureAllSheets_();

    const api = getApi_(e, body);
    const payload = getPayload_(body);

    switch (api) {
      case 'getQuestions':
        return json_({ ok: true, data: { questions: getQuestions_() } });

      case 'saveQuestions':
        return json_({ ok: true, data: saveQuestions_(payload) });

      case 'startRound':
        return json_({ ok: true, data: startRound_(payload) });

      case 'saveRoundState':
        return json_({ ok: true, data: saveRoundState_(payload) });

      case 'saveAnswer':
        return json_({ ok: true, data: saveAnswer_(payload) });

      case 'getLiveState':
        return json_({ ok: true, data: getLiveState_(payload) });

      case 'load':
        return json_({ ok: true, data: { questions: getQuestions_() } });

      case 'sync':
        return json_({ ok: true, data: syncLegacy_(payload) });

      default:
        return json_({ ok: false, error: 'Unknown api/action: ' + api });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function getApi_(e, body) {
  const queryApi = e && e.parameter ? (e.parameter.api || e.parameter.action) : '';
  return (body.api || body.action || queryApi || '').toString().trim();
}

function getPayload_(body) {
  if (body && typeof body.payload === 'object' && body.payload !== null) {
    return body.payload;
  }
  return body || {};
}

function ensureAllSheets_() {
  ensureSheetWithHeaders_(SHEET_NAMES.QUESTIONS, HEADERS.Questions);
  ensureSheetWithHeaders_(SHEET_NAMES.ROUNDS, HEADERS.Rounds);
  ensureSheetWithHeaders_(SHEET_NAMES.ANSWERS, HEADERS.Answers);
}

function ensureSheetWithHeaders_(sheetName, expectedHeaders) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
  }

  const currentLastCol = Math.max(sh.getLastColumn(), expectedHeaders.length);
  const firstRow = sh.getRange(1, 1, 1, currentLastCol).getValues()[0];

  const missingOrMismatch = expectedHeaders.some((header, idx) => {
    const cell = (firstRow[idx] || '').toString().trim();
    return cell !== header;
  });

  if (missingOrMismatch) {
    sh.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
  }
}

function getQuestions_() {
  const sh = getSheet_(SHEET_NAMES.QUESTIONS);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const values = sh.getRange(2, 1, lastRow - 1, 5).getValues();
  const groups = {};
  const order = [];

  values.forEach(row => {
    const category = String(row[0] || '').trim();
    const points = Number(row[1] || 0);
    const question = String(row[2] || '').trim();
    const answer = String(row[3] || '').trim();
    const isDouble = toBool_(row[4]);

    if (!category || !question || !points) return;

    if (!groups[category]) {
      groups[category] = [];
      order.push(category);
    }
    groups[category].push({ pts: points, text: question, answer: answer, isDouble: isDouble });
  });

  return order.map(category => {
    const questions = groups[category].sort((a, b) => a.pts - b.pts);
    return { name: category, questions: questions };
  });
}

function saveQuestions_(payload) {
  const questions = (payload && payload.questions) || [];
  if (!Array.isArray(questions)) {
    throw new Error('questions must be an array');
  }

  const sh = getSheet_(SHEET_NAMES.QUESTIONS);

  const rows = [];
  questions.forEach(categoryItem => {
    const categoryName = String((categoryItem && categoryItem.name) || '').trim();
    const categoryQuestions = (categoryItem && categoryItem.questions) || [];
    if (!categoryName || !Array.isArray(categoryQuestions)) return;

    categoryQuestions.forEach(questionItem => {
      const pts = Number((questionItem && questionItem.pts) || 0);
      const text = String((questionItem && questionItem.text) || '').trim();
      const answer = String((questionItem && questionItem.answer) || '').trim();
      const isDouble = toBool_(questionItem && questionItem.isDouble);
      if (!pts || !text) return;
      rows.push([categoryName, pts, text, answer, isDouble]);
    });
  });

  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 5).clearContent();
  }

  if (rows.length > 0) {
    sh.getRange(2, 1, rows.length, 5).setValues(rows);
  }

  return { saved: true, rowsCount: rows.length };
}

function startRound_(payload) {
  const roundId = payload.roundId || generateRoundId_();
  const nowIso = new Date().toISOString();

  const record = {
    roundId: roundId,
    sessionId: valueOr_(payload.sessionId, ''),
    hostName: valueOr_(payload.hostName, ''),
    team1Number: valueOr_(payload.team1Number, ''),
    team2Number: valueOr_(payload.team2Number, ''),
    score1: toNumber_(payload.score1, 0),
    score2: toNumber_(payload.score2, 0),
    answeredCount: toNumber_(payload.answeredCount, 0),
    doubleSelected: toBool_(payload.doubleSelected),
    mainTimerRemaining: toNumber_(payload.mainTimerRemaining, 90),
    mainTimerRunning: toBool_(payload.mainTimerRunning),
    mainTimerStartedAt: valueOr_(payload.mainTimerStartedAt, ''),
    eventType: valueOr_(payload.eventType, 'start-round'),
    timestamp: valueOr_(payload.timestamp, nowIso),
    answeredJson: valueOr_(payload.answeredJson, '[]'),
    activeQuestionJson: valueOr_(payload.activeQuestionJson, ''),
    qTimerRemaining: toNumber_(payload.qTimerRemaining, 90),
    qTimerRunning: toBool_(payload.qTimerRunning),
    qTimerStartedAt: valueOr_(payload.qTimerStartedAt, ''),
    extraJson: valueOr_(payload.extraJson, '{}'),
    updatedAt: nowIso
  };

  appendRowByHeaders_(SHEET_NAMES.ROUNDS, record, HEADERS.Rounds);
  return { roundId: roundId };
}

function saveRoundState_(payload) {
  const roundId = valueOr_(payload.roundId, '');
  if (!roundId) throw new Error('roundId is required for saveRoundState');

  const nowIso = new Date().toISOString();
  const update = {
    roundId: roundId,
    sessionId: valueOr_(payload.sessionId, ''),
    hostName: valueOr_(payload.hostName, ''),
    team1Number: valueOr_(payload.team1Number, ''),
    team2Number: valueOr_(payload.team2Number, ''),
    score1: toNumber_(payload.score1, 0),
    score2: toNumber_(payload.score2, 0),
    answeredCount: toNumber_(payload.answeredCount, 0),
    doubleSelected: toBool_(payload.doubleSelected),
    mainTimerRemaining: toNumber_(payload.mainTimerRemaining, 90),
    mainTimerRunning: toBool_(payload.mainTimerRunning),
    mainTimerStartedAt: valueOr_(payload.mainTimerStartedAt, ''),
    eventType: valueOr_(payload.eventType, 'state-update'),
    timestamp: valueOr_(payload.timestamp, nowIso),
    answeredJson: valueOr_(payload.answeredJson, '[]'),
    activeQuestionJson: valueOr_(payload.activeQuestionJson, ''),
    qTimerRemaining: toNumber_(payload.qTimerRemaining, 90),
    qTimerRunning: toBool_(payload.qTimerRunning),
    qTimerStartedAt: valueOr_(payload.qTimerStartedAt, ''),
    extraJson: valueOr_(payload.extraJson, '{}'),
    updatedAt: nowIso
  };

  const upsertResult = upsertByKey_(SHEET_NAMES.ROUNDS, 'roundId', roundId, update, HEADERS.Rounds);
  return { roundId: roundId, mode: upsertResult };
}

function saveAnswer_(payload) {
  const record = {
    roundId: valueOr_(payload.roundId, ''),
    sessionId: valueOr_(payload.sessionId, ''),
    timestamp: valueOr_(payload.timestamp, new Date().toISOString()),
    categoryName: valueOr_(payload.categoryName, ''),
    questionText: valueOr_(payload.questionText, ''),
    basePoints: toNumber_(payload.basePoints, 0),
    awardedPoints: toNumber_(payload.awardedPoints, 0),
    isDouble: toBool_(payload.isDouble),
    result: valueOr_(payload.result, ''),
    winnerTeam: toNumber_(payload.winnerTeam, 0),
    winnerTeamNumber: valueOr_(payload.winnerTeamNumber, ''),
    score1: toNumber_(payload.score1, 0),
    score2: toNumber_(payload.score2, 0)
  };

  appendRowByHeaders_(SHEET_NAMES.ANSWERS, record, HEADERS.Answers);
  return { saved: true };
}

function syncLegacy_(payload) {
  const stateObj = (payload && payload.state) || {};
  const matchSetup = (payload && payload.matchSetup) || {};

  const roundId = valueOr_(payload.roundId, payload.sessionId || generateRoundId_());
  const update = {
    roundId: roundId,
    sessionId: valueOr_(payload.sessionId, ''),
    hostName: valueOr_(matchSetup.hostName, ''),
    team1Number: valueOr_(matchSetup.team1Number, ''),
    team2Number: valueOr_(matchSetup.team2Number, ''),
    score1: toNumber_(stateObj.scores && stateObj.scores[0], 0),
    score2: toNumber_(stateObj.scores && stateObj.scores[1], 0),
    answeredCount: Array.isArray(stateObj.answered) ? stateObj.answered.length : 0,
    doubleSelected: toBool_(stateObj.doubleSelected),
    mainTimerRemaining: toNumber_(stateObj.mainTimerRemaining, 90),
    mainTimerRunning: false,
    mainTimerStartedAt: '',
    eventType: valueOr_(payload.eventType, 'legacy-sync'),
    timestamp: valueOr_(payload.timestamp, new Date().toISOString()),
    answeredJson: JSON.stringify(Array.isArray(stateObj.answered) ? stateObj.answered : []),
    activeQuestionJson: '',
    qTimerRemaining: 90,
    qTimerRunning: false,
    qTimerStartedAt: '',
    extraJson: JSON.stringify((payload && payload.extra) || {}),
    updatedAt: new Date().toISOString()
  };

  upsertByKey_(SHEET_NAMES.ROUNDS, 'roundId', roundId, update, HEADERS.Rounds);
  return { roundId: roundId, synced: true };
}

function getLiveState_(payload) {
  const allQuestions = getQuestions_();
  const publicQuestions = allQuestions.map(cat => ({
    name: cat.name,
    questions: (cat.questions || []).map(q => ({ pts: toNumber_(q.pts, 0) }))
  }));

  const round = getLatestRoundObject_(payload || {});
  if (!round) {
    return {
      hasRound: false,
      questions: publicQuestions
    };
  }

  const answered = parseJsonSafe_(round.answeredJson, []);
  const activeQuestion = parseJsonSafe_(round.activeQuestionJson, null);

  return {
    hasRound: true,
    serverTime: new Date().toISOString(),
    questions: publicQuestions,
    round: {
      roundId: valueOr_(round.roundId, ''),
      sessionId: valueOr_(round.sessionId, ''),
      hostName: valueOr_(round.hostName, ''),
      team1Number: valueOr_(round.team1Number, ''),
      team2Number: valueOr_(round.team2Number, ''),
      score1: toNumber_(round.score1, 0),
      score2: toNumber_(round.score2, 0),
      answeredCount: toNumber_(round.answeredCount, 0),
      answered: Array.isArray(answered) ? answered : [],
      eventType: valueOr_(round.eventType, ''),
      timestamp: valueOr_(round.timestamp, ''),
      mainTimerRemaining: toNumber_(round.mainTimerRemaining, 0),
      mainTimerRunning: toBool_(round.mainTimerRunning),
      mainTimerStartedAt: valueOr_(round.mainTimerStartedAt, ''),
      qTimerRemaining: toNumber_(round.qTimerRemaining, 0),
      qTimerRunning: toBool_(round.qTimerRunning),
      qTimerStartedAt: valueOr_(round.qTimerStartedAt, ''),
      activeQuestion: activeQuestion && typeof activeQuestion === 'object' ? activeQuestion : null,
      updatedAt: valueOr_(round.updatedAt, '')
    }
  };
}

function getLatestRoundObject_(payload) {
  const sh = getSheet_(SHEET_NAMES.ROUNDS);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;

  const values = sh.getRange(2, 1, lastRow - 1, HEADERS.Rounds.length).getValues();
  const rows = values.map(row => rowToObject_(row, HEADERS.Rounds));

  let filtered = rows;
  const wantedRoundId = String(valueOr_(payload.roundId, '')).trim();
  const wantedSessionId = String(valueOr_(payload.sessionId, '')).trim();

  if (wantedRoundId) {
    filtered = filtered.filter(r => String(valueOr_(r.roundId, '')).trim() === wantedRoundId);
  } else if (wantedSessionId) {
    filtered = filtered.filter(r => String(valueOr_(r.sessionId, '')).trim() === wantedSessionId);
  }

  if (!filtered.length) return null;

  filtered.sort((a, b) => {
    const at = String(valueOr_(a.updatedAt, ''));
    const bt = String(valueOr_(b.updatedAt, ''));
    return at.localeCompare(bt);
  });

  return filtered[filtered.length - 1];
}

function rowToObject_(row, headers) {
  const out = {};
  headers.forEach((h, i) => {
    out[h] = row[i];
  });
  return out;
}

function parseJsonSafe_(raw, fallback) {
  try {
    const text = String(valueOr_(raw, '')).trim();
    if (!text) return fallback;
    return JSON.parse(text);
  } catch (err) {
    return fallback;
  }
}

function upsertByKey_(sheetName, keyHeader, keyValue, record, headers) {
  const sh = getSheet_(sheetName);
  const keyCol = headers.indexOf(keyHeader) + 1;
  if (keyCol <= 0) throw new Error('Key header not found: ' + keyHeader);

  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    const keyValues = sh.getRange(2, keyCol, lastRow - 1, 1).getValues();
    for (let i = 0; i < keyValues.length; i++) {
      const current = String(keyValues[i][0] || '').trim();
      if (current === String(keyValue).trim()) {
        const rowValues = buildRowByHeaders_(record, headers);
        sh.getRange(i + 2, 1, 1, headers.length).setValues([rowValues]);
        return 'updated';
      }
    }
  }

  appendRowByHeaders_(sheetName, record, headers);
  return 'inserted';
}

function appendRowByHeaders_(sheetName, record, headers) {
  const sh = getSheet_(sheetName);
  const row = buildRowByHeaders_(record, headers);
  sh.appendRow(row);
}

function buildRowByHeaders_(record, headers) {
  return headers.map(h => {
    if (!(h in record)) return '';
    return record[h];
  });
}

function getSheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('Sheet not found: ' + name);
  return sh;
}

function generateRoundId_() {
  return 'round_' + new Date().getTime() + '_' + Math.floor(Math.random() * 1000000);
}

function valueOr_(value, fallback) {
  if (value === null || value === undefined) return fallback;
  return value;
}

function toNumber_(value, fallback) {
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
}

function toBool_(value) {
  return value === true || String(value).toLowerCase() === 'true' || Number(value) === 1;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
