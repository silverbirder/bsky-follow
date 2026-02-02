const SHEET_NAMES = {
  CONFIG: 'Config',
  POSTS: 'Posts',
  USERS: 'Users',
};

const CONFIG_KEYS = [
  'IDENTIFIER',
  'APP_PASSWORD',
  'PDS_HOST',
  'SEARCH_QUERY',
  'SEARCH_LIMIT',
  'MAX_PAGES',
  'SEARCH_SORT',
  'SEARCH_LANG',
  'WEBHOOK_TOKEN',
];

const POSTS_HEADERS = [
  'fetched_at',
  'keyword',
  'post_uri',
  'post_text',
  'post_created_at',
  'author_display_name',
];

const USERS_HEADERS = [
  'user_url',
  'author_display_name',
  'action',
  'follow_uri',
  'status',
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Bsky Follow')
    .addItem('Setup sheets', 'setup')
    .addItem('Search & sync', 'searchAndSync')
    .addItem('Apply follow/unfollow', 'applyUserActions')
    .addToUi();
}

function setup() {
  ensureSheet_(SHEET_NAMES.CONFIG, ['key', 'value']);
  ensureSheet_(SHEET_NAMES.POSTS, POSTS_HEADERS);
  ensureSheet_(SHEET_NAMES.USERS, USERS_HEADERS);
  seedConfig_();
}

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  const token = e && e.parameter && e.parameter.token;
  if (!isWebhookAuthorized_(token)) {
    return ContentService.createTextOutput('Unauthorized');
  }

  if (action === 'search') {
    searchAndSync();
    return ContentService.createTextOutput('search: ok');
  }
  if (action === 'apply') {
    applyUserActions();
    return ContentService.createTextOutput('apply: ok');
  }

  return ContentService.createTextOutput(
    'OK. Use ?action=search or ?action=apply'
  );
}

function searchAndSync() {
  const config = getConfig_();
  const session = createSession_(config);

  const query = config.SEARCH_QUERY;
  const limit = clampNumber_(config.SEARCH_LIMIT || 50, 1, 100);
  const maxPages = clampNumber_(config.MAX_PAGES || 1, 1, 20);
  const sort = config.SEARCH_SORT || 'latest';
  const lang = config.SEARCH_LANG || '';

  let cursor = '';
  const allPosts = [];
  for (let i = 0; i < maxPages; i += 1) {
    const result = searchPosts_(session, query, limit, cursor, sort, lang);
    const posts = result && result.posts ? result.posts : [];
    allPosts.push.apply(allPosts, posts);
    cursor = result && result.cursor ? result.cursor : '';
    if (!cursor) {
      break;
    }
  }

  if (!allPosts.length) {
    return;
  }

  writePosts_(allPosts, query);
  writeUsers_(allPosts);
}

function applyUserActions() {
  const config = getConfig_();
  const session = createSession_(config);

  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.USERS);
  if (!sheet) {
    throw new Error('Users sheet not found. Run setup().');
  }

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return;
  }

  const header = values[0];
  const userUrlIndex = header.indexOf('user_url');
  const actionIndex = header.indexOf('action');
  const followUriIndex = header.indexOf('follow_uri');
  const statusIndex = header.indexOf('status');

  const updatedRows = [];

  for (let i = 1; i < values.length; i += 1) {
    const row = values[i];
    const action = (row[actionIndex] || '').toString().trim().toLowerCase();
    if (!action) {
      updatedRows.push(row);
      continue;
    }

    try {
      if (action === 'follow') {
        const did = toDidFromProfileUrl_(row[userUrlIndex]);
        const followUri = followActor_(session, did);
        row[followUriIndex] = followUri;
        row[statusIndex] = 'followed';
      } else if (action === 'unfollow') {
        const followUri = row[followUriIndex];
        unfollowActor_(session, followUri);
        row[statusIndex] = 'unfollowed';
        row[followUriIndex] = '';
      } else {
        row[statusIndex] = 'skipped';
      }
    } catch (err) {
      row[statusIndex] = 'error';
    }

    updatedRows.push(row);
  }

  sheet
    .getRange(2, 1, updatedRows.length, updatedRows[0].length)
    .setValues(updatedRows);
}

function createSession_(config) {
  validateConfig_(config);
  const url = joinUrl_(config.PDS_HOST, '/xrpc/com.atproto.server.createSession');
  const payload = {
    identifier: config.IDENTIFIER,
    password: config.APP_PASSWORD,
  };

  const response = fetchJson_(url, {
    method: 'post',
    payload: JSON.stringify(payload),
  });

  return {
    accessJwt: response.accessJwt,
    did: response.did,
    handle: response.handle,
    pdsHost: config.PDS_HOST,
  };
}

function searchPosts_(session, query, limit, cursor, sort, lang) {
  const params = {
    q: query,
    limit: limit,
    cursor: cursor,
    sort: sort,
    lang: lang,
  };
  const url = joinUrl_(session.pdsHost, '/xrpc/app.bsky.feed.searchPosts');
  const fullUrl = url + '?' + encodeQuery_(params);
  return fetchJson_(fullUrl, {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + session.accessJwt,
    },
  });
}

function followActor_(session, did) {
  const url = joinUrl_(session.pdsHost, '/xrpc/com.atproto.repo.createRecord');
  const payload = {
    repo: session.did,
    collection: 'app.bsky.graph.follow',
    record: {
      $type: 'app.bsky.graph.follow',
      subject: did,
      createdAt: nowIso_(),
    },
  };
  const response = fetchJson_(url, {
    method: 'post',
    headers: {
      Authorization: 'Bearer ' + session.accessJwt,
    },
    payload: JSON.stringify(payload),
  });
  return response.uri;
}

function unfollowActor_(session, followUri) {
  if (!followUri) {
    throw new Error('follow_uri is empty');
  }
  const parsed = parseAtUri_(followUri);
  if (!parsed) {
    throw new Error('Invalid follow_uri');
  }
  if (parsed.collection !== 'app.bsky.graph.follow') {
    throw new Error('follow_uri is not app.bsky.graph.follow');
  }

  const url = joinUrl_(session.pdsHost, '/xrpc/com.atproto.repo.deleteRecord');
  const payload = {
    repo: session.did,
    collection: parsed.collection,
    rkey: parsed.rkey,
  };
  fetchJson_(url, {
    method: 'post',
    headers: {
      Authorization: 'Bearer ' + session.accessJwt,
    },
    payload: JSON.stringify(payload),
  });
}

function writePosts_(posts, keyword) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.POSTS);
  if (!sheet) {
    throw new Error('Posts sheet not found. Run setup().');
  }
  ensureHeaders_(sheet, POSTS_HEADERS);

  const existing = getExistingIndex_(sheet, 'post_uri');
  const rows = [];

  posts.forEach((post) => {
    const postUrl = toBskyPostUrl_(post);
    const row = [
      nowIso_(),
      keyword,
      postUrl,
      (post.record && post.record.text) || '',
      (post.record && post.record.createdAt) || '',
      (post.author && post.author.displayName) || '',
    ];

    if (existing[postUrl]) {
      const rowIndex = existing[postUrl];
      sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
    } else {
      rows.push(row);
    }
  });

  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
}

function writeUsers_(posts) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.USERS);
  if (!sheet) {
    throw new Error('Users sheet not found. Run setup().');
  }
  ensureHeaders_(sheet, USERS_HEADERS);

  const existing = getExistingIndex_(sheet, 'user_url');
  const rows = [];

  posts.forEach((post) => {
    const userUrl = toBskyProfileUrl_(post.author);
    if (!userUrl) {
      return;
    }

    const row = [
      userUrl,
      (post.author && post.author.displayName) || '',
      '',
      '',
      '',
    ];

    if (existing[userUrl]) {
      const rowIndex = existing[userUrl];
      const current = sheet.getRange(rowIndex, 1, 1, USERS_HEADERS.length).getValues()[0];
      current[1] = row[1];
      sheet.getRange(rowIndex, 1, 1, current.length).setValues([current]);
    } else {
      rows.push(row);
    }
  });

  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
}

function ensureSheet_(name, headers) {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaders_(sheet, headers);
  }
  return sheet;
}

function seedConfig_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.CONFIG);
  if (!sheet) {
    return;
  }
  if (sheet.getLastRow() > 1) {
    return;
  }
  const values = CONFIG_KEYS.map((key) => [key, '']);
  sheet.getRange(2, 1, values.length, 2).setValues(values);
}

function getConfig_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.CONFIG);
  if (!sheet) {
    throw new Error('Config sheet not found. Run setup().');
  }
  const values = sheet.getDataRange().getValues();
  const config = {};
  for (let i = 1; i < values.length; i += 1) {
    const key = (values[i][0] || '').toString().trim();
    if (!key) {
      continue;
    }
    config[key] = (values[i][1] || '').toString().trim();
  }
  if (!config.PDS_HOST) {
    config.PDS_HOST = 'https://bsky.social';
  }
  if (!config.SEARCH_LIMIT) {
    config.SEARCH_LIMIT = '50';
  }
  if (!config.MAX_PAGES) {
    config.MAX_PAGES = '1';
  }
  return config;
}

function validateConfig_(config) {
  const required = ['IDENTIFIER', 'APP_PASSWORD', 'SEARCH_QUERY', 'PDS_HOST'];
  required.forEach((key) => {
    if (!config[key]) {
      throw new Error('Config missing: ' + key);
    }
  });
}

function isWebhookAuthorized_(token) {
  const config = getConfig_();
  if (!config.WEBHOOK_TOKEN) {
    return true;
  }
  return token && token === config.WEBHOOK_TOKEN;
}

function getExistingIndex_(sheet, keyHeader) {
  const range = sheet.getDataRange().getValues();
  if (range.length < 2) {
    return {};
  }
  const header = range[0];
  const keyIndex = header.indexOf(keyHeader);
  if (keyIndex === -1) {
    return {};
  }
  const map = {};
  for (let i = 1; i < range.length; i += 1) {
    const key = range[i][keyIndex];
    if (key) {
      map[key] = i + 1;
    }
  }
  return map;
}

function ensureHeaders_(sheet, headers) {
  const existing = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let same = existing.length >= headers.length;
  for (let i = 0; i < headers.length; i += 1) {
    if (existing[i] !== headers[i]) {
      same = false;
      break;
    }
  }
  if (!same) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

function fetchJson_(url, options) {
  const response = UrlFetchApp.fetch(url, {
    method: options.method || 'get',
    contentType: 'application/json',
    headers: options.headers || {},
    payload: options.payload || null,
    muteHttpExceptions: true,
  });
  const status = response.getResponseCode();
  const text = response.getContentText();
  if (status >= 400) {
    throw new Error('HTTP ' + status + ': ' + text);
  }
  return text ? JSON.parse(text) : {};
}

function encodeQuery_(params) {
  const parts = [];
  Object.keys(params).forEach((key) => {
    const value = params[key];
    if (value === undefined || value === null || value === '') {
      return;
    }
    parts.push(encodeURIComponent(key) + '=' + encodeURIComponent(value));
  });
  return parts.join('&');
}

function joinUrl_(base, path) {
  if (base.endsWith('/')) {
    return base.slice(0, -1) + path;
  }
  return base + path;
}

function parseAtUri_(uri) {
  const match = /^at:\/\/([^/]+)\/([^/]+)\/([^/]+)$/.exec(uri);
  if (!match) {
    return null;
  }
  return {
    repo: match[1],
    collection: match[2],
    rkey: match[3],
  };
}

function nowIso_() {
  return new Date().toISOString();
}

function clampNumber_(value, min, max) {
  const num = Number(value);
  if (Number.isNaN(num)) {
    return min;
  }
  return Math.max(min, Math.min(max, num));
}

function toBskyProfileUrl_(author) {
  if (!author) {
    return '';
  }
  const id = author.handle || author.did || '';
  if (!id) {
    return '';
  }
  return 'https://bsky.app/profile/' + encodeURIComponent(id);
}

function toDidFromProfileUrl_(userUrl) {
  if (!userUrl) {
    throw new Error('user_url is empty');
  }
  const match = /^https:\/\/bsky\.app\/profile\/([^/?#]+)(?:[/?#]|$)/.exec(
    String(userUrl)
  );
  if (!match) {
    throw new Error('user_url is invalid');
  }
  const id = decodeURIComponent(match[1]);
  if (id.indexOf('did:') === 0) {
    return id;
  }
  const profile = resolveProfile_(id);
  return profile.did;
}

function resolveProfile_(actor) {
  const config = getConfig_();
  const url = joinUrl_(config.PDS_HOST || 'https://bsky.social', '/xrpc/com.atproto.identity.resolveHandle');
  const fullUrl = url + '?' + encodeQuery_({ handle: actor });
  return fetchJson_(fullUrl, { method: 'get' });
}

function toBskyPostUrl_(post) {
  if (!post) {
    return '';
  }
  const parsed = parseAtUri_(post.uri || '');
  if (!parsed) {
    return '';
  }
  const author = post.author || {};
  const profileId = author.handle || author.did || parsed.repo;
  if (!profileId) {
    return '';
  }
  return (
    'https://bsky.app/profile/' +
    encodeURIComponent(profileId) +
    '/post/' +
    encodeURIComponent(parsed.rkey)
  );
}
