const SHEET_NAMES = {
  CONFIG: 'Config',
  POSTS: 'Posts',
  FOLLOWS: 'Follows',
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
  'IGNORE_AUTHOR_DISPLAY_NAMES',
  'WEBHOOK_TOKEN',
  'AUTO_FOLLOW_MAX_PER_RUN',
  'AUTO_FOLLOW_MAX_PER_DAY',
  'AUTO_FOLLOW_COOLDOWN_DAYS',
  'AUTO_FOLLOW_MIN_DELAY_SECONDS',
  'AUTO_FOLLOW_MAX_DELAY_SECONDS',
  'AUTO_FOLLOW_EXCLUDE_ACTORS',
  'AUTO_UNFOLLOW_AFTER_DAYS',
  'AUTO_UNFOLLOW_MAX_PER_RUN',
  'AUTO_UNFOLLOW_MIN_DELAY_SECONDS',
  'AUTO_UNFOLLOW_MAX_DELAY_SECONDS',
];

const CONFIG_DEFAULTS = {
  PDS_HOST: 'https://bsky.social',
  SEARCH_LIMIT: '50',
  MAX_PAGES: '1',
  SEARCH_SORT: 'latest',
  AUTO_FOLLOW_MAX_PER_RUN: '3',
  AUTO_FOLLOW_MAX_PER_DAY: '20',
  AUTO_FOLLOW_COOLDOWN_DAYS: '30',
  AUTO_FOLLOW_MIN_DELAY_SECONDS: '15',
  AUTO_FOLLOW_MAX_DELAY_SECONDS: '45',
  AUTO_UNFOLLOW_AFTER_DAYS: '7',
  AUTO_UNFOLLOW_MAX_PER_RUN: '10',
  AUTO_UNFOLLOW_MIN_DELAY_SECONDS: '15',
  AUTO_UNFOLLOW_MAX_DELAY_SECONDS: '45',
};

const CONFIG_DESCRIPTIONS = {
  IDENTIFIER: 'BlueskyのログインID。通常はhandle（例: yourname.bsky.social）を入力。',
  APP_PASSWORD: 'BlueskyのApp Password。通常パスワードではなく、必ずApp Passwordを使用。',
  PDS_HOST: 'BlueskyのPDSホスト。通常は https://bsky.social のままでOK。',
  SEARCH_QUERY: '検索キーワード。複数指定する場合はカンマ区切り（例: AI,JavaScript）。',
  SEARCH_LIMIT: '1クエリ1ページあたりの取得件数（1〜100）。大きいほど取得量が増える。',
  MAX_PAGES: '1クエリあたりの取得ページ数（1〜20）。増やすと実行時間も増える。',
  SEARCH_SORT: '検索並び順。latest（新着）推奨。',
  SEARCH_LANG: '言語フィルタ。例: ja / en。空欄なら言語指定なし。',
  IGNORE_AUTHOR_DISPLAY_NAMES: '検索結果から除外する表示名。カンマ区切りで指定。',
  WEBHOOK_TOKEN: 'Webアプリ実行時の簡易認証トークン。設定するとURLに token=... が必須。',
  AUTO_FOLLOW_MAX_PER_RUN: '1回の実行でフォローする最大人数。小さめ（例: 1〜5）推奨。',
  AUTO_FOLLOW_MAX_PER_DAY: '1日あたりのフォロー上限。過剰操作回避のため低めに設定推奨。',
  AUTO_FOLLOW_COOLDOWN_DAYS: '同一アカウントを再試行するまでの日数。重複操作の抑制用。',
  AUTO_FOLLOW_MIN_DELAY_SECONDS: 'フォロー間の最小待機秒数。短すぎる値は避ける。',
  AUTO_FOLLOW_MAX_DELAY_SECONDS: 'フォロー間の最大待機秒数。MIN以上の値を設定。',
  AUTO_FOLLOW_EXCLUDE_ACTORS: '自動フォロー対象から除外するDID/handle。カンマ区切り。',
  AUTO_UNFOLLOW_AFTER_DAYS:
    'フォロー後に自動アンフォロー判定を始めるまでの日数。例: 7（1週間）。',
  AUTO_UNFOLLOW_MAX_PER_RUN: '1回の実行でアンフォローする最大人数。',
  AUTO_UNFOLLOW_MIN_DELAY_SECONDS: 'アンフォロー間の最小待機秒数。短すぎる値は避ける。',
  AUTO_UNFOLLOW_MAX_DELAY_SECONDS: 'アンフォロー間の最大待機秒数。MIN以上の値を設定。',
};

const CONFIG_MENU_INFO_KEYS = [
  'MENU_RUN_SEARCH',
  'MENU_RUN_AUTO_FOLLOW',
  'MENU_RUN_AUTO_UNFOLLOW',
  'MENU_INSTALL_SEARCH_TRIGGER',
  'MENU_INSTALL_AUTO_FOLLOW_TRIGGER',
  'MENU_INSTALL_AUTO_UNFOLLOW_TRIGGER',
];

const CONFIG_MENU_INFO_DESCRIPTIONS = {
  MENU_RUN_SEARCH:
    'メニュー「検索」: 手動で今すぐ1回だけ searchAndSync を実行し、Postsシートを更新する。',
  MENU_RUN_AUTO_FOLLOW:
    'メニュー「自動フォロー」: 手動で今すぐ1回だけ autoFollow を実行する。',
  MENU_RUN_AUTO_UNFOLLOW:
    'メニュー「自動アンフォロー」: 手動で今すぐ1回だけ autoUnfollow を実行する。',
  MENU_INSTALL_SEARCH_TRIGGER:
    'メニュー「定期検索」: searchAndSync を1時間ごとに動かすトリガーを設定。既存の searchAndSync トリガーは一度削除して作り直す。',
  MENU_INSTALL_AUTO_FOLLOW_TRIGGER:
    'メニュー「定期自動フォロー」: autoFollow を1時間ごとに動かすトリガーを設定。既存の autoFollow トリガーは一度削除して作り直す。',
  MENU_INSTALL_AUTO_UNFOLLOW_TRIGGER:
    'メニュー「定期自動アンフォロー」: autoUnfollow を1時間ごとに動かすトリガーを設定。既存の autoUnfollow トリガーは一度削除して作り直す。',
};

const POSTS_HEADERS = [
  'fetched_at',
  'keyword',
  'post_uri',
  'post_at_uri',
  'post_text',
  'post_created_at',
  'author_display_name',
  'author_handle',
  'author_did',
];

const FOLLOWS_HEADERS = [
  'attempted_at',
  'keyword',
  'actor_input',
  'actor_did',
  'source_post_uri',
  'status',
  'reason',
  'follow_uri',
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Bsky Follow')
    .addItem('初期化', 'setup')
    .addItem('検索', 'searchAndSync')
    .addItem('自動フォロー', 'runAutoFollowNow')
    .addItem('自動アンフォロー', 'runAutoUnfollowNow')
    .addItem('定期検索', 'installSearchTrigger')
    .addItem('定期自動フォロー', 'installAutoFollowTrigger')
    .addItem('定期自動アンフォロー', 'installAutoUnfollowTrigger')
    .addItem('リセット', 'resetPosts')
    .addToUi();
}

function setup() {
  ensureSheet_(SHEET_NAMES.CONFIG, ['key', 'value', 'description']);
  ensureSheet_(SHEET_NAMES.POSTS, POSTS_HEADERS);
  ensureSheet_(SHEET_NAMES.FOLLOWS, FOLLOWS_HEADERS);
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
  if (action === 'follow') {
    const result = autoFollowFromPosts_();
    return ContentService.createTextOutput('follow: ' + JSON.stringify(result));
  }
  if (action === 'unfollow') {
    const result = autoUnfollowFromFollows_();
    return ContentService.createTextOutput('unfollow: ' + JSON.stringify(result));
  }
  if (action === 'search_follow') {
    searchAndSync();
    const result = autoFollowFromPosts_();
    return ContentService.createTextOutput(
      'search_follow: ' + JSON.stringify(result)
    );
  }

  return ContentService.createTextOutput(
    'OK. Use ?action=search|follow|unfollow|search_follow'
  );
}

function searchAndSync() {
  const config = getConfig_();
  const session = createSession_(config);

  const limit = clampNumber_(config.SEARCH_LIMIT || 50, 1, 100);
  const maxPages = clampNumber_(config.MAX_PAGES || 1, 1, 20);
  const sort = config.SEARCH_SORT || 'latest';
  const lang = config.SEARCH_LANG || '';

  const queries = parseKeywords_(config.SEARCH_QUERY);
  if (!queries.length) {
    return;
  }

  const ignoredDisplayNames = buildIgnoreSet_(config.IGNORE_AUTHOR_DISPLAY_NAMES);

  queries.forEach((query) => {
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

    const filteredPosts = allPosts.filter((post) =>
      ignoredDisplayNames.size
        ? !isIgnoredDisplayName_(post, ignoredDisplayNames)
        : true
    );

    if (!filteredPosts.length) {
      return;
    }

    writePosts_(filteredPosts, query);
  });
}

function runAutoFollowNow() {
  const result = autoFollowFromPosts_();
  Logger.log('auto follow result: %s', JSON.stringify(result));
}

function autoFollow() {
  const result = autoFollowFromPosts_();
  Logger.log('auto follow result: %s', JSON.stringify(result));
}

function runAutoUnfollowNow() {
  const result = autoUnfollowFromFollows_();
  Logger.log('auto unfollow result: %s', JSON.stringify(result));
}

function autoUnfollow() {
  const result = autoUnfollowFromFollows_();
  Logger.log('auto unfollow result: %s', JSON.stringify(result));
}

function installAutoFollowTrigger() {
  const handler = 'autoFollow';
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((trigger) => {
    if (trigger.getHandlerFunction() === handler) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger(handler).timeBased().everyHours(1).create();
}

function installAutoUnfollowTrigger() {
  const handler = 'autoUnfollow';
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((trigger) => {
    if (trigger.getHandlerFunction() === handler) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger(handler).timeBased().everyHours(1).create();
}

function installSearchTrigger() {
  const handler = 'searchAndSync';
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((trigger) => {
    if (trigger.getHandlerFunction() === handler) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger(handler).timeBased().everyHours(1).create();
}

function autoFollowFromPosts_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const config = getConfig_();

    const session = createSession_(config);
    const postsSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.POSTS);
    if (!postsSheet) {
      throw new Error('Posts sheet not found. Run setup().');
    }
    const followsSheet = ensureSheet_(SHEET_NAMES.FOLLOWS, FOLLOWS_HEADERS);
    ensureHeaders_(postsSheet, POSTS_HEADERS);

    const maxPerRun = clampNumber_(config.AUTO_FOLLOW_MAX_PER_RUN, 1, 20);
    const maxPerDay = clampNumber_(config.AUTO_FOLLOW_MAX_PER_DAY, 1, 200);
    const cooldownDays = clampNumber_(config.AUTO_FOLLOW_COOLDOWN_DAYS, 1, 365);
    const minDelaySec = clampNumber_(config.AUTO_FOLLOW_MIN_DELAY_SECONDS, 0, 300);
    const maxDelaySec = clampNumber_(
      config.AUTO_FOLLOW_MAX_DELAY_SECONDS,
      minDelaySec,
      600
    );
    const excludedActors = buildIgnoreSet_(config.AUTO_FOLLOW_EXCLUDE_ACTORS);

    const followState = buildFollowState_(followsSheet, cooldownDays);
    let dailyCount = countFollowedToday_(followsSheet);
    if (dailyCount >= maxPerDay) {
      return { status: 'daily_limit_reached', followed: 0, dailyCount: dailyCount };
    }

    const rows = postsSheet.getDataRange().getValues();
    if (rows.length < 2) {
      return { status: 'no_posts', followed: 0 };
    }
    const header = rows[0];
    const idxKeyword = header.indexOf('keyword');
    const idxPostUri = header.indexOf('post_uri');
    const idxAuthorDid = header.indexOf('author_did');
    const idxAuthorHandle = header.indexOf('author_handle');

    const didCache = {};
    const logs = [];
    let followed = 0;

    for (let i = rows.length - 1; i >= 1; i -= 1) {
      if (followed >= maxPerRun || dailyCount >= maxPerDay) {
        break;
      }
      const row = rows[i];
      const keyword = idxKeyword >= 0 ? (row[idxKeyword] || '').toString() : '';
      const sourcePostUri = idxPostUri >= 0 ? (row[idxPostUri] || '').toString() : '';
      const actorInput =
        (idxAuthorDid >= 0 ? (row[idxAuthorDid] || '').toString().trim() : '') ||
        (idxAuthorHandle >= 0 ? (row[idxAuthorHandle] || '').toString().trim() : '') ||
        extractActorFromPostUrl_(sourcePostUri);
      if (!actorInput) {
        continue;
      }
      if (excludedActors.has(actorInput.toLowerCase())) {
        logs.push(
          buildFollowLogRow_(
            keyword,
            actorInput,
            '',
            sourcePostUri,
            'skipped',
            'excluded_actor',
            ''
          )
        );
        continue;
      }

      let actorDid;
      try {
        actorDid = resolveDid_(session, actorInput, didCache);
      } catch (err) {
        logs.push(
          buildFollowLogRow_(
            keyword,
            actorInput,
            '',
            sourcePostUri,
            'error',
            String(err),
            ''
          )
        );
        continue;
      }

      if (actorDid === session.did) {
        logs.push(
          buildFollowLogRow_(
            keyword,
            actorInput,
            actorDid,
            sourcePostUri,
            'skipped',
            'self',
            ''
          )
        );
        continue;
      }
      if (followState.followedDids[actorDid]) {
        continue;
      }
      if (
        followState.cooldownUntil[actorDid] &&
        followState.cooldownUntil[actorDid] > Date.now()
      ) {
        continue;
      }

      try {
        const followUri = followActor_(session, actorDid);
        logs.push(
          buildFollowLogRow_(
            keyword,
            actorInput,
            actorDid,
            sourcePostUri,
            'followed',
            '',
            followUri
          )
        );
        followState.followedDids[actorDid] = true;
        followState.cooldownUntil[actorDid] = Date.now() + cooldownDays * 86400000;
        followed += 1;
        dailyCount += 1;
      } catch (err) {
        const message = String(err);
        logs.push(
          buildFollowLogRow_(
            keyword,
            actorInput,
            actorDid,
            sourcePostUri,
            'error',
            message,
            ''
          )
        );
        if (isStopError_(message)) {
          break;
        }
      }

      sleepWithJitter_(minDelaySec, maxDelaySec);
    }

    if (logs.length) {
      followsSheet
        .getRange(followsSheet.getLastRow() + 1, 1, logs.length, logs[0].length)
        .setValues(logs);
    }
    return {
      status: 'ok',
      followed: followed,
      logged: logs.length,
      dailyCount: dailyCount,
    };
  } finally {
    lock.releaseLock();
  }
}

function autoUnfollowFromFollows_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const config = getConfig_();
    const session = createSession_(config);
    const followsSheet = ensureSheet_(SHEET_NAMES.FOLLOWS, FOLLOWS_HEADERS);
    ensureHeaders_(followsSheet, FOLLOWS_HEADERS);

    const unfollowAfterDays = clampNumber_(config.AUTO_UNFOLLOW_AFTER_DAYS, 1, 365);
    const maxPerRun = clampNumber_(config.AUTO_UNFOLLOW_MAX_PER_RUN, 1, 100);
    const minDelaySec = clampNumber_(config.AUTO_UNFOLLOW_MIN_DELAY_SECONDS, 0, 300);
    const maxDelaySec = clampNumber_(
      config.AUTO_UNFOLLOW_MAX_DELAY_SECONDS,
      minDelaySec,
      600
    );
    const candidates = getUnfollowCandidates_(followsSheet, unfollowAfterDays);
    if (!candidates.length) {
      return { status: 'no_candidates', checked: 0, unfollowed: 0, logged: 0 };
    }

    const logs = [];
    let checked = 0;
    let unfollowed = 0;

    for (let i = 0; i < candidates.length; i += 1) {
      if (checked >= maxPerRun) {
        break;
      }
      const candidate = candidates[i];
      checked += 1;

      try {
        const relationship = getRelationship_(session, candidate.actorDid);
        const followingUri = relationship.following || '';
        if (!followingUri) {
          logs.push(
            buildFollowLogRow_(
              candidate.keyword,
              candidate.actorInput,
              candidate.actorDid,
              candidate.sourcePostUri,
              'unfollowed',
              'already_unfollowed',
              ''
            )
          );
        } else if (relationship.followedBy) {
          continue;
        } else {
          unfollowByUri_(session, followingUri);
          logs.push(
            buildFollowLogRow_(
              candidate.keyword,
              candidate.actorInput,
              candidate.actorDid,
              candidate.sourcePostUri,
              'unfollowed',
              'not_following_me',
              followingUri
            )
          );
          unfollowed += 1;
        }
      } catch (err) {
        const message = String(err);
        logs.push(
          buildFollowLogRow_(
            candidate.keyword,
            candidate.actorInput,
            candidate.actorDid,
            candidate.sourcePostUri,
            'error',
            message,
            ''
          )
        );
        if (isStopError_(message)) {
          break;
        }
      }

      sleepWithJitter_(minDelaySec, maxDelaySec);
    }

    if (logs.length) {
      followsSheet
        .getRange(followsSheet.getLastRow() + 1, 1, logs.length, logs[0].length)
        .setValues(logs);
    }

    return {
      status: 'ok',
      checked: checked,
      unfollowed: unfollowed,
      logged: logs.length,
      candidates: candidates.length,
    };
  } finally {
    lock.releaseLock();
  }
}

function resetPosts() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.POSTS);
  if (!sheet) {
    throw new Error('Posts sheet not found. Run setup().');
  }
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1 || lastCol === 0) {
    return;
  }
  sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
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
  return response.uri || '';
}

function getRelationship_(session, actorDid) {
  const url = joinUrl_(session.pdsHost, '/xrpc/app.bsky.graph.getRelationships');
  const fullUrl = url + '?' + encodeQuery_({ actor: session.did, others: actorDid });
  const data = fetchJson_(fullUrl, {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + session.accessJwt,
    },
  });
  const relationships = data && data.relationships ? data.relationships : [];
  for (let i = 0; i < relationships.length; i += 1) {
    const relationship = relationships[i] || {};
    if ((relationship.did || '') === actorDid) {
      return relationship;
    }
  }
  return relationships[0] || {};
}

function unfollowByUri_(session, followUri) {
  const parsed = parseAtUri_(followUri);
  if (!parsed) {
    throw new Error('invalid follow uri: ' + followUri);
  }
  if (parsed.collection !== 'app.bsky.graph.follow') {
    throw new Error('not follow record: ' + followUri);
  }
  if (parsed.repo !== session.did) {
    throw new Error('follow uri repo mismatch: ' + followUri);
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
      post.uri || '',
      (post.record && post.record.text) || '',
      (post.record && post.record.createdAt) || '',
      (post.author && post.author.displayName) || '',
      (post.author && post.author.handle) || '',
      (post.author && post.author.did) || '',
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
  ensureHeaders_(sheet, ['key', 'value', 'description']);
  const data = sheet.getDataRange().getValues();
  const existingKeys = {};
  const rowByKey = {};
  for (let i = 1; i < data.length; i += 1) {
    const key = (data[i][0] || '').toString().trim();
    if (key) {
      existingKeys[key] = true;
      rowByKey[key] = i + 1;
    }
  }
  const values = [];
  CONFIG_KEYS.forEach((key) => {
    if (!existingKeys[key]) {
      values.push([key, CONFIG_DEFAULTS[key] || '', CONFIG_DESCRIPTIONS[key] || '']);
    }
  });
  CONFIG_MENU_INFO_KEYS.forEach((key) => {
    if (!existingKeys[key]) {
      values.push([key, '', CONFIG_MENU_INFO_DESCRIPTIONS[key] || '']);
    }
  });
  if (values.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, values.length, 3).setValues(values);
  }

  CONFIG_KEYS.forEach((key) => {
    const rowIndex = rowByKey[key];
    if (!rowIndex) {
      return;
    }
    const description = CONFIG_DESCRIPTIONS[key] || '';
    sheet.getRange(rowIndex, 3).setValue(description);
    if (!data[rowIndex - 1] || !data[rowIndex - 1][1]) {
      const defaultValue = CONFIG_DEFAULTS[key] || '';
      if (defaultValue !== '') {
        sheet.getRange(rowIndex, 2).setValue(defaultValue);
      }
    }
  });

  CONFIG_MENU_INFO_KEYS.forEach((key) => {
    const rowIndex = rowByKey[key];
    if (!rowIndex) {
      return;
    }
    const description = CONFIG_MENU_INFO_DESCRIPTIONS[key] || '';
    sheet.getRange(rowIndex, 3).setValue(description);
  });

  if (sheet.getFrozenRows() < 1) {
    sheet.setFrozenRows(1);
  }
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
  Object.keys(CONFIG_DEFAULTS).forEach((key) => {
    if (!config[key]) {
      config[key] = CONFIG_DEFAULTS[key];
    }
  });
  return config;
}

function validateConfig_(config) {
  const required = ['IDENTIFIER', 'APP_PASSWORD', 'PDS_HOST'];
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
  Logger.log('API request: %s %s', (options.method || 'get').toUpperCase(), url);
  const response = UrlFetchApp.fetch(url, {
    method: options.method || 'get',
    contentType: 'application/json',
    headers: options.headers || {},
    payload: options.payload || null,
    muteHttpExceptions: true,
  });
  const status = response.getResponseCode();
  const text = response.getContentText();
  Logger.log('API response: %s %s', status, url);
  if (status >= 400) {
    Logger.log('API error body: %s', text);
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

function parseKeywords_(value) {
  if (!value) {
    return [];
  }
  return String(value)
    .split(',')
    .map((item) => item.trim())
    .filter((item) => item);
}

function buildIgnoreSet_(value) {
  const items = parseKeywords_(value);
  const set = new Set();
  items.forEach((item) => {
    if (item) {
      set.add(item.toLowerCase());
    }
  });
  return set;
}

function buildFollowState_(sheet, cooldownDays) {
  const values = sheet.getDataRange().getValues();
  const state = {
    followedDids: {},
    cooldownUntil: {},
  };
  if (values.length < 2) {
    return state;
  }
  const header = values[0];
  const idxAttemptedAt = header.indexOf('attempted_at');
  const idxActorDid = header.indexOf('actor_did');
  const idxStatus = header.indexOf('status');
  const cooldownMs = cooldownDays * 86400000;
  for (let i = 1; i < values.length; i += 1) {
    const actorDid = idxActorDid >= 0 ? (values[i][idxActorDid] || '').toString().trim() : '';
    const status = idxStatus >= 0 ? (values[i][idxStatus] || '').toString().trim() : '';
    if (!actorDid) {
      continue;
    }
    if (status === 'followed' || status === 'unfollowed') {
      state.followedDids[actorDid] = true;
    }
    const attemptedAt = idxAttemptedAt >= 0 ? new Date(values[i][idxAttemptedAt]).getTime() : NaN;
    if (!Number.isNaN(attemptedAt)) {
      const until = attemptedAt + cooldownMs;
      if (!state.cooldownUntil[actorDid] || until > state.cooldownUntil[actorDid]) {
        state.cooldownUntil[actorDid] = until;
      }
    }
  }
  return state;
}

function getUnfollowCandidates_(sheet, unfollowAfterDays) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return [];
  }
  const header = values[0];
  const idxAttemptedAt = header.indexOf('attempted_at');
  const idxKeyword = header.indexOf('keyword');
  const idxActorInput = header.indexOf('actor_input');
  const idxActorDid = header.indexOf('actor_did');
  const idxSourcePostUri = header.indexOf('source_post_uri');
  const idxStatus = header.indexOf('status');
  const thresholdMs = unfollowAfterDays * 86400000;
  const now = Date.now();
  const byActorDid = {};

  for (let i = 1; i < values.length; i += 1) {
    const actorDid = idxActorDid >= 0 ? (values[i][idxActorDid] || '').toString().trim() : '';
    if (!actorDid) {
      continue;
    }
    const status = idxStatus >= 0 ? (values[i][idxStatus] || '').toString().trim() : '';
    if (status !== 'followed' && status !== 'unfollowed') {
      continue;
    }
    const attemptedAt = idxAttemptedAt >= 0 ? new Date(values[i][idxAttemptedAt]).getTime() : NaN;
    if (Number.isNaN(attemptedAt)) {
      continue;
    }

    if (!byActorDid[actorDid]) {
      byActorDid[actorDid] = {
        actorDid: actorDid,
        actorInput: '',
        keyword: '',
        sourcePostUri: '',
        followedAt: 0,
        unfollowedAt: 0,
      };
    }
    const entry = byActorDid[actorDid];
    if (status === 'followed' && attemptedAt >= entry.followedAt) {
      entry.followedAt = attemptedAt;
      entry.actorInput =
        idxActorInput >= 0 ? (values[i][idxActorInput] || '').toString().trim() : '';
      entry.keyword = idxKeyword >= 0 ? (values[i][idxKeyword] || '').toString() : '';
      entry.sourcePostUri =
        idxSourcePostUri >= 0 ? (values[i][idxSourcePostUri] || '').toString() : '';
    }
    if (status === 'unfollowed' && attemptedAt >= entry.unfollowedAt) {
      entry.unfollowedAt = attemptedAt;
    }
  }

  const candidates = [];
  Object.keys(byActorDid).forEach((actorDid) => {
    const entry = byActorDid[actorDid];
    if (!entry.followedAt) {
      return;
    }
    if (entry.unfollowedAt >= entry.followedAt) {
      return;
    }
    if (now - entry.followedAt < thresholdMs) {
      return;
    }
    if (!entry.actorInput) {
      entry.actorInput = actorDid;
    }
    candidates.push(entry);
  });
  candidates.sort((a, b) => a.followedAt - b.followedAt);
  return candidates;
}

function countFollowedToday_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return 0;
  }
  const header = values[0];
  const idxAttemptedAt = header.indexOf('attempted_at');
  const idxStatus = header.indexOf('status');
  const start = new Date();
  start.setHours(0, 0, 0, 0);
  let count = 0;
  for (let i = 1; i < values.length; i += 1) {
    const status = idxStatus >= 0 ? (values[i][idxStatus] || '').toString().trim() : '';
    if (status !== 'followed') {
      continue;
    }
    const ts = idxAttemptedAt >= 0 ? new Date(values[i][idxAttemptedAt]).getTime() : NaN;
    if (!Number.isNaN(ts) && ts >= start.getTime()) {
      count += 1;
    }
  }
  return count;
}

function buildFollowLogRow_(keyword, actorInput, actorDid, sourcePostUri, status, reason, followUri) {
  return [
    nowIso_(),
    keyword || '',
    actorInput || '',
    actorDid || '',
    sourcePostUri || '',
    status || '',
    reason || '',
    followUri || '',
  ];
}

function sleepWithJitter_(minSeconds, maxSeconds) {
  if (maxSeconds <= 0) {
    return;
  }
  const minMs = Math.max(0, minSeconds) * 1000;
  const maxMs = Math.max(minMs, maxSeconds * 1000);
  const waitMs = minMs + Math.floor(Math.random() * (maxMs - minMs + 1));
  Utilities.sleep(waitMs);
}

function isStopError_(message) {
  return /^Error: HTTP (429|5\d\d):/.test(message);
}

function resolveDid_(session, actor, cache) {
  const normalized = String(actor || '').trim();
  if (!normalized) {
    throw new Error('actor is empty');
  }
  if (/^did:/.test(normalized)) {
    return normalized;
  }
  if (cache[normalized]) {
    return cache[normalized];
  }
  const did = resolveHandle_(session, normalized);
  cache[normalized] = did;
  return did;
}

function resolveHandle_(session, handle) {
  const url = joinUrl_(session.pdsHost, '/xrpc/com.atproto.identity.resolveHandle');
  const fullUrl = url + '?' + encodeQuery_({ handle: handle });
  const data = fetchJson_(fullUrl, { method: 'get' });
  if (!data.did) {
    throw new Error('resolve handle failed: ' + handle);
  }
  return data.did;
}

function extractActorFromPostUrl_(postUrl) {
  const match = /^https:\/\/bsky\.app\/profile\/([^/]+)\/post\/[^/]+(?:[/?#]|$)/.exec(
    String(postUrl || '')
  );
  if (!match) {
    return '';
  }
  return decodeURIComponent(match[1]);
}

function isIgnoredDisplayName_(post, ignored) {
  if (!post || !post.author) {
    return false;
  }
  const displayName = (post.author.displayName || '').toString().trim();
  if (!displayName) {
    return false;
  }
  return ignored.has(displayName.toLowerCase());
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
