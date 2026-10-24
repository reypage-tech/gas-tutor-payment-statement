// Code.gs（複数の.gsファイル）が使うGAS組み込みサービス・高度なサービスの最小モック。
// 実行環境はnode（vmコンテキスト）。目的は、実際のApps Script実行環境なしに
// チェック・計算・PDF・下書き・確定・6分制限からの再開・日次上限・冪等性のロジックを
// 確かめること。実際のGoogle Apps Script実行環境では実行していない。
//
// GmailApp・DriveApp・MailApp・UrlFetchAppはこのモックに一切存在しない（意図的）。
// 本体コードがこれらを呼び出せば ReferenceError で即座に失敗するため、
// 「高度なサービス（Gmail.Users.Drafts / Drive.Files）だけを使っている」ことを
// このテストの完走そのもので確認できる。
var fs = require('fs');
var path = require('path');
var vm = require('vm');

function createWorld() {
  return {
    idCounter: 0,
    scriptProperties: {},
    spreadsheet: null, // 1つのスプレッドシートだけを模す（このプロジェクトはcurrentonly運用のため）
    files: {}, // id -> {id,name,mimeType,parents:[],trashed,blob?,docContent?,ssRef?,api?}
    drafts: {}, // id -> {id,raw,deleted}
    triggers: [], // {getHandlerFunction()}
    currentUserEmail: 'staff@example.com'
  };
}

function nextId(world, prefix) {
  world.idCounter += 1;
  return prefix + world.idCounter;
}

// ===== Sheet / Range =====

function makeSheet(world, name) {
  var rows = []; // rows[0] が1行目
  var conditionalRules = [];
  var protections = [];
  var charts = [];

  function ensureRow(r) {
    if (!rows[r]) rows[r] = [];
    return rows[r];
  }

  var sheetApi = {
    __isMockSheet: true,
    getName: function () { return name; },
    setName: function (n) { name = n; return sheetApi; },
    getLastRow: function () {
      for (var i = rows.length - 1; i >= 0; i--) {
        if (rows[i] && rows[i].some(function (v) { return v !== undefined && v !== ''; })) return i + 1;
      }
      return 0;
    },
    getLastColumn: function () {
      var max = 0;
      rows.forEach(function (row) { if (row) max = Math.max(max, row.length); });
      return max;
    },
    getRange: function (r, c, numRows, numCols) {
      numRows = numRows || 1;
      numCols = numCols || 1;
      return {
        setValues: function (values) {
          for (var i = 0; i < values.length; i++) {
            var row = ensureRow(r - 1 + i);
            for (var j = 0; j < values[i].length; j++) row[c - 1 + j] = values[i][j];
          }
          return this;
        },
        getValues: function () {
          var out = [];
          for (var i = 0; i < numRows; i++) {
            var row = rows[r - 1 + i] || [];
            var outRow = [];
            for (var j = 0; j < numCols; j++) outRow.push(row[c - 1 + j] !== undefined ? row[c - 1 + j] : '');
            out.push(outRow);
          }
          return out;
        },
        setValue: function (v) {
          var row = ensureRow(r - 1);
          row[c - 1] = v;
          return this;
        },
        getValue: function () {
          var row = rows[r - 1] || [];
          return row[c - 1] !== undefined ? row[c - 1] : '';
        },
        setFontWeight: function () { return this; },
        setBackground: function () { return this; },
        setNumberFormat: function () { return this; }
      };
    },
    appendRow: function (values) {
      rows.push(values.slice());
    },
    clearContents: function () { rows = []; },
    clearFormat: function () { /* 見た目のみ。ロジックには影響しない */ },
    getConditionalFormatRules: function () { return conditionalRules.slice(); },
    setConditionalFormatRules: function (rules) { conditionalRules = rules.slice(); return sheetApi; },
    protect: function () {
      var editors = [];
      var domainEdit = true;
      var description = '';
      var protection = {
        setDescription: function (d) { description = d; return protection; },
        canDomainEdit: function () { return domainEdit; },
        setDomainEdit: function (v) { domainEdit = v; return protection; },
        getEditors: function () { return editors.slice(); },
        addEditors: function (emails) { editors = editors.concat(emails); return protection; },
        removeEditors: function (emails) { editors = editors.filter(function (e) { return emails.indexOf(e) === -1; }); return protection; },
        remove: function () { protections = protections.filter(function (p) { return p !== protection; }); },
        getDescription: function () { return description; }
      };
      protections.push(protection);
      return protection;
    },
    getProtections: function () { return protections.slice(); },
    getCharts: function () { return charts.slice(); },
    removeChart: function (c) { charts = charts.filter(function (x) { return x !== c; }); },
    newChart: function () {
      var spec = { type: null, ranges: [], position: null, options: {} };
      var builder = {
        setChartType: function (t) { spec.type = t; return builder; },
        addRange: function (range) { spec.ranges.push(range); return builder; },
        setPosition: function () { return builder; },
        setOption: function (k, v) { spec.options[k] = v; return builder; },
        build: function () { return spec; }
      };
      return builder;
    },
    insertChart: function (chart) { charts.push(chart); },
    __dumpRows: function () { return rows.map(function (r) { return (r || []).slice(); }); }
  };
  return sheetApi;
}

function makeSpreadsheet(world, id, name) {
  var sheets = {};
  var ssApi = {
    getId: function () { return id; },
    getName: function () { return name; },
    getUrl: function () { return 'https://mock.sheet.local/' + id; },
    getSheetByName: function (n) { var hit = null; Object.keys(sheets).forEach(function (k) { if (sheets[k].getName() === n) hit = sheets[k]; }); return hit; },
    insertSheet: function (n) {
      var sheet = makeSheet(world, n);
      // 名前を変えたシートと新しいシートが同じキーで上書きされないよう、一意のキーで持つ
      var key = n;
      while (sheets[key]) key = key + '#';
      sheets[key] = sheet;
      return sheet;
    },
    getSheets: function () { return Object.keys(sheets).map(function (n) { return sheets[n]; }); },
    deleteSheet: function (sh) {
      Object.keys(sheets).forEach(function (k) { if (sheets[k] === sh) delete sheets[k]; });
    },
    toast: function () {}
  };
  world.files[id] = { id: id, name: name, mimeType: 'application/vnd.google-apps.spreadsheet', parents: [], trashed: false, ssRef: ssApi };
  return ssApi;
}

function buildSpreadsheetApp(world) {
  var uiAlerts = [];
  var ui = {
    createMenu: function () {
      var menu = {
        addItem: function () { return menu; },
        addSeparator: function () { return menu; },
        addToUi: function () {}
      };
      return menu;
    },
    alert: function (msg) { uiAlerts.push(msg); },
    showSidebar: function () {}
  };
  world.uiAlerts = uiAlerts;
  return {
    getActive: function () { return world.spreadsheet; },
    getUi: function () { return ui; },
    newConditionalFormatRule: function () {
      var spec = { kind: null, value: null, background: null, ranges: [] };
      var builder = {
        whenTextEqualTo: function (v) { spec.kind = 'textEqualTo'; spec.value = v; return builder; },
        setBackground: function (bg) { spec.background = bg; return builder; },
        setRanges: function (ranges) { spec.ranges = ranges; return builder; },
        build: function () { return spec; }
      };
      return builder;
    },
    ProtectionType: { SHEET: 'SHEET' }
  };
}

// ===== Drive 高度なサービス（Drive API v3 の最小モック） =====

var DRIVE_MIME_FOLDER_ = 'application/vnd.google-apps.folder';
var DRIVE_MIME_DOCUMENT_ = 'application/vnd.google-apps.document';

function escapeRegExp_(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

function matchesDriveQuery_(file, q) {
  if (file.trashed) return false;
  var mimeMatch = /mimeType='((?:[^'\\]|\\.)*)'/.exec(q);
  if (mimeMatch && file.mimeType !== unescapeDriveQueryValue_(mimeMatch[1])) return false;
  var nameMatch = /name='((?:[^'\\]|\\.)*)'/.exec(q);
  if (nameMatch && file.name !== unescapeDriveQueryValue_(nameMatch[1])) return false;
  var parentMatch = /'((?:[^'\\]|\\.)*)' in parents/.exec(q);
  if (parentMatch && (file.parents || []).indexOf(unescapeDriveQueryValue_(parentMatch[1])) === -1) return false;
  return true;
}

function unescapeDriveQueryValue_(v) { return v.replace(/\\'/g, "'").replace(/\\\\/g, '\\'); }

function cloneDoc_(doc) {
  return {
    pageWidth: doc.pageWidth,
    pageHeight: doc.pageHeight,
    margins: Object.assign({}, doc.margins),
    elements: doc.elements.map(function (el) {
      if (el.type === 'table') return { type: 'table', rows: el.rows.map(function (row) { return row.slice(); }) };
      return Object.assign({}, el);
    })
  };
}

function renderDocToText_(doc) {
  var lines = [];
  doc.elements.forEach(function (el) {
    if (el.type === 'paragraph') lines.push(el.text);
    else if (el.type === 'table') el.rows.forEach(function (row) { lines.push(row.join(' | ')); });
  });
  return lines.join('\n');
}

function buildDriveAdvancedService(world) {
  function get(fileId, optionalArgs) {
    var file = world.files[fileId];
    if (!file) throw new Error('Drive.Files.get: 見つかりません (id=' + fileId + ')');
    if (optionalArgs && optionalArgs.alt === 'media') {
      if (file.trashed) throw new Error('Drive.Files.get(alt=media): ゴミ箱にあります (id=' + fileId + ')');
      if (!file.blob) throw new Error('Drive.Files.get(alt=media): コンテンツがありません (id=' + fileId + ')');
      return makeBlob(file.blob.bytes, file.blob.mimeType, file.name);
    }
    if (file.trashed) throw new Error('Drive.Files.get: ゴミ箱にあります (id=' + fileId + ')');
    return { id: file.id, name: file.name, mimeType: file.mimeType, trashed: !!file.trashed, parents: (file.parents || []).slice() };
  }

  function create(resource, media) {
    var id = nextId(world, 'file_');
    var mimeType = (resource && resource.mimeType) || (media ? media.getContentType() : 'application/octet-stream');
    var file = { id: id, name: (resource && resource.name) || (media && media.getName()) || id, mimeType: mimeType, parents: (resource && resource.parents) ? resource.parents.slice() : [], trashed: false };
    if (media) file.blob = { bytes: media.getBytes(), mimeType: media.getContentType() };
    world.files[id] = file;
    return { id: id, name: file.name, mimeType: file.mimeType };
  }

  function list(optionalArgs) {
    var q = (optionalArgs && optionalArgs.q) || '';
    var matches = Object.keys(world.files)
      .map(function (id) { return world.files[id]; })
      .filter(function (f) { return matchesDriveQuery_(f, q); });
    return { files: matches.map(function (f) { return { id: f.id, name: f.name, mimeType: f.mimeType }; }) };
  }

  function copy(resource, fileId) {
    var src = world.files[fileId];
    if (!src) throw new Error('Drive.Files.copy: 見つかりません (id=' + fileId + ')');
    var newId = nextId(world, 'file_');
    var copyFile = {
      id: newId,
      name: (resource && resource.name) || src.name,
      mimeType: src.mimeType,
      parents: (resource && resource.parents) ? resource.parents.slice() : [],
      trashed: false
    };
    if (src.docContent) copyFile.docContent = cloneDoc_(src.docContent);
    if (src.blob) copyFile.blob = { bytes: Buffer.from(src.blob.bytes), mimeType: src.blob.mimeType };
    if (src.ssRef) copyFile.ssRef = src.ssRef; // バックアップ用の浅いコピー（テストではシート内容の深い複製は見ない）
    world.files[newId] = copyFile;
    return { id: newId, name: copyFile.name };
  }

  function update(resource, fileId) {
    var file = world.files[fileId];
    if (!file) throw new Error('Drive.Files.update: 見つかりません (id=' + fileId + ')');
    if (resource && resource.trashed !== undefined) file.trashed = !!resource.trashed;
    if (resource && resource.name) file.name = resource.name;
    return { id: file.id, trashed: !!file.trashed };
  }

  function exportFile(fileId, mimeType) {
    if (mimeType !== 'application/pdf') throw new Error('Drive.Files.export: 未対応のmimeTypeです: ' + mimeType);
    var file = world.files[fileId];
    if (!file || file.trashed) throw new Error('Drive.Files.export: 見つかりません (id=' + fileId + ')');
    if (file.mimeType !== DRIVE_MIME_DOCUMENT_) {
      throw new Error('Drive.Files.export: Document以外からのPDF変換はこのモックでは未対応です');
    }
    var text = renderDocToText_(file.docContent);
    return makeBlob(Buffer.from('%PDF-1.4 mock\n' + text), 'application/pdf', file.name + '.pdf');
  }

  return { Files: { get: get, create: create, list: list, copy: copy, update: update, export: exportFile } };
}

// ===== DocumentApp（本文の組み立てのみ。ファイルの実体管理はDrive側のworld.filesに統一） =====

function makeBodyApi(doc) {
  function paragraphWrapper(el) {
    return {
      setHeading: function (h) { el.heading = h; return paragraphWrapper(el); },
      setAlignment: function (a) { el.align = a; return paragraphWrapper(el); },
      setBold: function (b) { el.bold = !!b; return paragraphWrapper(el); }
    };
  }
  function rowWrapper(row) {
    return {
      appendTableCell: function (text) { row.push(String(text)); return { setText: function () {} }; }
    };
  }
  function tableWrapper(el) {
    return {
      appendTableRow: function () {
        var row = [];
        el.rows.push(row);
        return rowWrapper(row);
      },
      getNumRows: function () { return el.rows.length; }
    };
  }
  var body = {
    setPageWidth: function (w) { doc.pageWidth = w; return body; },
    setPageHeight: function (h) { doc.pageHeight = h; return body; },
    setMarginTop: function (v) { doc.margins.top = v; return body; },
    setMarginBottom: function (v) { doc.margins.bottom = v; return body; },
    setMarginLeft: function (v) { doc.margins.left = v; return body; },
    setMarginRight: function (v) { doc.margins.right = v; return body; },
    appendParagraph: function (text) {
      var el = { type: 'paragraph', text: String(text), heading: null };
      doc.elements.push(el);
      return paragraphWrapper(el);
    },
    appendTable: function (initialRows) {
      var el = { type: 'table', rows: initialRows.map(function (r) { return r.slice(); }) };
      doc.elements.push(el);
      return tableWrapper(el);
    },
    getTables: function () {
      return doc.elements.filter(function (el) { return el.type === 'table'; }).map(tableWrapper);
    },
    replaceText: function (pattern, replacement) {
      doc.elements.forEach(function (el) {
        if (el.type === 'paragraph' && el.text.indexOf(pattern) !== -1) {
          el.text = el.text.split(pattern).join(replacement);
        }
      });
    }
  };
  return body;
}

function buildDocApiFor_(world, id) {
  var file = world.files[id];
  if (!file.api) {
    file.api = {
      getId: function () { return id; },
      getBody: function () { return makeBodyApi(file.docContent); },
      saveAndClose: function () {}
    };
  }
  return file.api;
}

function buildDocumentApp(world) {
  return {
    ParagraphHeading: { HEADING1: 'HEADING1', HEADING2: 'HEADING2', HEADING3: 'HEADING3', NORMAL: 'NORMAL' },
    create: function (name) {
      var id = nextId(world, 'file_');
      var docContent = { pageWidth: 0, pageHeight: 0, margins: {}, elements: [] };
      world.files[id] = { id: id, name: name, mimeType: DRIVE_MIME_DOCUMENT_, parents: [], trashed: false, docContent: docContent };
      return buildDocApiFor_(world, id);
    },
    openById: function (id) {
      var file = world.files[id];
      if (!file || file.mimeType !== DRIVE_MIME_DOCUMENT_ || file.trashed) {
        throw new Error('DocumentApp.openById: 見つかりません (id=' + id + ')');
      }
      return buildDocApiFor_(world, id);
    }
  };
}

// ===== Gmail 高度なサービス（Gmail API v1 の最小モック：Drafts.create/get/delete のみ） =====

function buildGmailAdvancedService(world) {
  function create(resource) {
    var id = nextId(world, 'draft_');
    world.drafts[id] = { id: id, raw: resource.message.raw, deleted: false };
    return { id: id, message: { id: 'msg_' + id } };
  }
  function get(userId, id) {
    var d = world.drafts[id];
    if (!d || d.deleted) throw new Error('Gmail.Users.Drafts.get: 見つかりません (id=' + id + ')');
    return { id: d.id };
  }
  function del(userId, id) {
    var d = world.drafts[id];
    if (!d) throw new Error('Gmail.Users.Drafts.delete: 見つかりません (id=' + id + ')');
    d.deleted = true;
  }
  return { Users: { Drafts: { create: create, get: get, delete: del } } };
}

// ===== LockService / PropertiesService / ScriptApp / Utilities / HtmlService / Session / Charts =====

function buildLockService(world) {
  return {
    getScriptLock: function () {
      return {
        waitLock: function () {
          if (world.lockHeld) throw new Error('lock already held');
          world.lockHeld = true;
        },
        releaseLock: function () { world.lockHeld = false; }
      };
    }
  };
}

function buildPropertiesService(world) {
  var scoped = {
    getProperty: function (key) { return world.scriptProperties.hasOwnProperty(key) ? world.scriptProperties[key] : null; },
    setProperty: function (key, value) { world.scriptProperties[key] = value; return scoped; },
    deleteProperty: function (key) { delete world.scriptProperties[key]; return scoped; }
  };
  return { getScriptProperties: function () { return scoped; } };
}

function buildScriptApp(world) {
  return {
    newTrigger: function (handlerFunction) {
      var spec = { handlerFunction: handlerFunction, type: null, afterMs: null };
      var builder = {
        timeBased: function () { spec.type = 'TIME'; return builder; },
        after: function (ms) { spec.afterMs = ms; return builder; },
        create: function () {
          var trigger = {
            getHandlerFunction: function () { return spec.handlerFunction; }
          };
          world.triggers.push(trigger);
          return trigger;
        }
      };
      return builder;
    },
    getProjectTriggers: function () { return world.triggers.slice(); },
    deleteTrigger: function (t) { world.triggers = world.triggers.filter(function (x) { return x !== t; }); }
  };
}

function pad2_(n) { return (n < 10 ? '0' : '') + n; }

function buildUtilities() {
  return {
    formatDate: function (date, tz, format) {
      var y = date.getFullYear(), mo = pad2_(date.getMonth() + 1), d = pad2_(date.getDate());
      var h = pad2_(date.getHours()), mi = pad2_(date.getMinutes()), s = pad2_(date.getSeconds());
      return format
        .replace('yyyy', y)
        .replace('MM', mo)
        .replace('dd', d)
        .replace('HH', h)
        .replace('mm', mi)
        .replace('ss', s);
    },
    newBlob: function (bytes, contentType, name) { return makeBlob(bytes, contentType, name); },
    base64Encode: function (data) {
      var buf = typeof data === 'string' ? Buffer.from(data, 'utf8') : (Buffer.isBuffer(data) ? data : Buffer.from(data));
      return buf.toString('base64');
    },
    base64EncodeWebSafe: function (data) {
      var buf = typeof data === 'string' ? Buffer.from(data, 'utf8') : (Buffer.isBuffer(data) ? data : Buffer.from(data));
      return buf.toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
    },
    getUuid: function () { return 'uuid-' + Math.random().toString(36).slice(2) + '-' + Math.random().toString(36).slice(2); },
    Charset: { UTF_8: 'UTF-8' }
  };
}

function makeBlob(bytes, contentType, name) {
  var buf = Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes);
  return {
    getBytes: function () { return buf; },
    getContentType: function () { return contentType; },
    getName: function () { return name; },
    setName: function (n) { name = n; return this; }
  };
}

function buildHtmlService() {
  return {
    createHtmlOutputFromFile: function () {
      return {
        setTitle: function () { return this; },
        setWidth: function () { return this; },
        addMetaTag: function () { return this; }
      };
    }
  };
}

function buildSession(world) {
  return {
    getEffectiveUser: function () {
      return { getEmail: function () { return world.currentUserEmail; } };
    }
  };
}

function buildCharts() {
  return { ChartType: { COLUMN: 'COLUMN', LINE: 'LINE' } };
}

// ===== エクスポート =====

function buildMockGlobals() {
  var world = createWorld();
  var ssId = nextId(world, 'ss_');
  world.spreadsheet = makeSpreadsheet(world, ssId, '月次の報酬・支払明細（見本）');

  var globals = {
    console: console,
    SpreadsheetApp: buildSpreadsheetApp(world),
    DocumentApp: buildDocumentApp(world),
    Drive: buildDriveAdvancedService(world),
    Gmail: buildGmailAdvancedService(world),
    LockService: buildLockService(world),
    PropertiesService: buildPropertiesService(world),
    ScriptApp: buildScriptApp(world),
    Utilities: buildUtilities(),
    HtmlService: buildHtmlService(),
    Session: buildSession(world),
    Charts: buildCharts()
  };
  return { world: world, globals: globals };
}

module.exports = { buildMockGlobals: buildMockGlobals };
