/**
 * Slide Table Styler（Googleスライド用）
 * 選択中の表に、プリセットのデザインを一括で適用します。
 */

// ===== プリセット定義 =====
// サイドバーのプレビューもこの定義から作られます（getPresets）。
// name: 表示名。日本語（ja）と英語（en）を書きます。
// 色の書き方：'#RRGGBB'（固定の色）／'accent'（アクセントカラー）／'accent/10'（アクセントカラーを白に10%混ぜた色）
//   accent はプリセットの既定のアクセントカラー。サイドバーで選んだ色に置き換わります。
// headerText: 'auto' は見出しの背景色に合わせて、白か濃い灰色の読みやすいほうを選びます。
// borders: 上から順に適用されます。color: null は「罫線なし（透明）」。
//   range: 'header' を付けると、ヘッダー行だけに適用します。
//   dash: 'DOT'（点線）/ 'DASH'（破線）。省略すると実線です。
const PRESETS = {
  cleanGray: {
    name: { ja: 'クリーン・グレー', en: 'Clean Gray' },
    accent: '#7F7F7F',
    headerFill: 'accent/10', headerText: '#333333',
    bodyFill: '#FFFFFF', stripeFill: null, bodyText: '#333333',
    borders: [
      { position: 'ALL', color: null },
      { position: 'INNER_HORIZONTAL', color: '#D9D9D9', weight: 1 },
      { position: 'TOP', color: '#D9D9D9', weight: 1 },
      { position: 'BOTTOM', color: '#D9D9D9', weight: 1 },
    ],
  },
  navyStripe: {
    name: { ja: 'ネイビー・ストライプ', en: 'Navy Stripe' },
    accent: '#1F3A5F',
    headerFill: 'accent', headerText: 'auto',
    bodyFill: '#FFFFFF', stripeFill: 'accent/5', bodyText: '#333333',
    borders: [
      { position: 'ALL', color: null },
      { position: 'INNER_HORIZONTAL', color: 'accent/12', weight: 0.75 },
      { position: 'BOTTOM', color: 'accent', weight: 1.5 },
    ],
  },
  minimal: {
    name: { ja: 'ミニマル（横線のみ）', en: 'Minimal (Rules Only)' },
    accent: '#222222',
    headerFill: '#FFFFFF', headerText: '#222222',
    bodyFill: '#FFFFFF', stripeFill: null, bodyText: '#333333',
    borders: [
      { position: 'ALL', color: null },
      { position: 'INNER_HORIZONTAL', color: '#E0E0E0', weight: 0.75 },
      { position: 'TOP', color: 'accent', weight: 1.5 },
      { position: 'BOTTOM', color: 'accent', weight: 1.5 },
      { position: 'BOTTOM', color: 'accent', weight: 1, range: 'header' },
    ],
  },
  dotted: {
    name: { ja: 'ドット区切り', en: 'Dotted Rows' },
    accent: '#2E6B5E',
    headerFill: 'accent/10', headerText: '#222222',
    bodyFill: '#FFFFFF', stripeFill: null, bodyText: '#333333',
    borders: [
      { position: 'ALL', color: null },
      { position: 'INNER_HORIZONTAL', color: 'accent/45', weight: 1, dash: 'DOT' },
      { position: 'BOTTOM', color: 'accent', weight: 1.5, range: 'header' },
      { position: 'BOTTOM', color: 'accent', weight: 1.5 },
    ],
  },
};

// 列幅の下限（Slides API は 32pt 未満の列幅を受け付けない）
const MIN_COLUMN_WIDTH_PT = 32;

// ===== 表示言語 =====
// 日本語（ja で始まる言語）なら日本語、それ以外は英語で表示する。
// サイドバーは Chrome の表示言語で判定して、opts.lang でサーバーに渡す。
// メニューは Chrome の言語がわからないため、Google アカウントの言語で判定する。
const MESSAGES = {
  ja: {
    openSidebar: 'サイドバーを開く',
    presetNotFound: key => `プリセットが見つかりません: ${key}`,
    appliedOne: (name, rows, cols) => `「${name}」を適用しました（${rows}行 × ${cols}列）`,
    appliedMany: (name, count) => `「${name}」を${count}個の表に適用しました`,
    noTablesInPresentation: 'このプレゼンテーションには表がありません。',
    noSlideSelected: 'スライドが選択されていません。対象のスライドを開いてから、もう一度「適用」を押してください。',
    noTablesOnSlide: 'このスライドには表がありません。',
    noTableSelected: '表が選択されていません。表の中をクリックしてから、もう一度「適用」を押してください。',
    noPermission: 'このスライドを編集する権限がありません。編集権限のあるスライドで実行してください。',
    authExpired: '承認の有効期限が切れています。スライドを再読み込みして、もう一度サイドバーを開いてください。',
    tooBusy: '処理に時間がかかりすぎたか、混み合っています。適用する範囲を狭めるか、しばらくしてからもう一度お試しください。',
    invalidTable: detail => `この表には書式を適用できませんでした。結合したセルなど、表の構造が原因の可能性があります。（詳細：${detail}）`,
    unexpected: detail => `予期しないエラーが発生しました。（詳細：${detail}）`,
  },
  en: {
    openSidebar: 'Open sidebar',
    presetNotFound: key => `Design not found: ${key}`,
    appliedOne: (name, rows, cols) => `Applied "${name}" (${rows} rows × ${cols} columns).`,
    appliedMany: (name, count) => `Applied "${name}" to ${count} tables.`,
    noTablesInPresentation: 'This presentation has no tables.',
    noSlideSelected: 'No slide is selected. Open the slide you want, then click "Apply" again.',
    noTablesOnSlide: 'This slide has no tables.',
    noTableSelected: 'No table is selected. Click inside a table, then click "Apply" again.',
    noPermission: "You don't have permission to edit this presentation. Try a presentation you can edit.",
    authExpired: 'Your authorization has expired. Reload the presentation and open the sidebar again.',
    tooBusy: 'This took too long or the service is busy. Try a smaller scope, or try again in a moment.',
    invalidTable: detail => `Couldn't format this table. Its structure (such as merged cells) may be the cause. (Details: ${detail})`,
    unexpected: detail => `Something went wrong. (Details: ${detail})`,
  },
};

function normalizeLang_(code) {
  return /^ja/i.test(code || '') ? 'ja' : 'en';
}

function messages_(lang) {
  return MESSAGES[normalizeLang_(lang)];
}

// ===== メニューとサイドバー =====
// onOpen は承認前（AuthMode.NONE）にも呼ばれるため、メニュー作成以外のサービスは使わない
function onOpen(e) {
  let locale = '';
  try {
    locale = Session.getActiveUserLocale();
  } catch (err) {
    // 取得できない環境では英語にする
  }
  SlidesApp.getUi()
    .createAddonMenu()
    .addItem(messages_(locale).openSidebar, 'showSidebar')
    .addToUi();
}

// インストール直後は onOpen が呼ばれないので、ここでメニューを出す
function onInstall(e) {
  onOpen(e);
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('Slide Table Styler');
  SlidesApp.getUi().showSidebar(html);
}

// ===== サイドバーから呼ばれる処理 =====
/** サイドバーに表示するプリセット一覧 */
function getPresets() {
  return Object.keys(PRESETS).map(key => Object.assign({ key: key }, PRESETS[key]));
}

/** 開いているスライドのテーマのアクセントカラー（重複を除く）。取れなければ空の配列 */
function getThemeColors() {
  try {
    const presentation = SlidesApp.getActivePresentation();
    const page = presentation.getSelection().getCurrentPage() || presentation.getSlides()[0] || presentation.getMasters()[0];
    const scheme = page.getColorScheme();
    const colors = ['ACCENT1', 'ACCENT2', 'ACCENT3', 'ACCENT4', 'ACCENT5', 'ACCENT6']
      .map(type => scheme.getConcreteColor(SlidesApp.ThemeColorType[type]).asRgbColor().asHexString().toUpperCase());
    return colors.filter((c, i) => colors.indexOf(c) === i);
  } catch (e) {
    console.warn(e);
    return [];
  }
}

/**
 * @param {string} presetKey PRESETSのキー
 * @param {Object} opts {
 *   scope: 'selection' | 'slide' | 'all',
 *   accent: '#RRGGBB'（空ならプリセットの色）,
 *   hAlign: 'KEEP' | 'START' | 'CENTER',
 *   vAlign: 'KEEP' | 'TOP' | 'MIDDLE',
 *   alignNumbers: boolean,
 *   minRowHeight: pt（空なら変更しない）, equalColumns: boolean,
 *   innerDash: 'KEEP' | 'SOLID' | 'DOT' | 'DASH',
 *   fontFamily, headerFontSize, bodyFontSize, boldColumns, headerBold,
 *   lang: 表示言語（'ja' / 'en' など）
 * }
 */
function applyPreset(presetKey, opts) {
  opts = opts || {};
  const lang = normalizeLang_(opts.lang);
  const msg = messages_(lang);
  if (!PRESETS[presetKey]) throw userError_(msg.presetNotFound(presetKey));
  const preset = resolvePreset_(PRESETS[presetKey], opts.accent, opts.innerDash);
  const name = preset.name[lang];

  let tables;
  try {
    tables = getTargetTables_(opts.scope, msg);
    const requests = [];
    tables.forEach(table => requests.push(...tableRequests_(table, preset, opts)));
    Slides.Presentations.batchUpdate({ requests: requests }, SlidesApp.getActivePresentation().getId());
  } catch (e) {
    throw toUserError_(e, msg);
  }

  if (tables.length === 1) {
    return msg.appliedOne(name, tables[0].getNumRows(), tables[0].getNumColumns());
  }
  return msg.appliedMany(name, tables.length);
}

// ===== 表1つ分のリクエストを作る =====
function tableRequests_(table, preset, opts) {
  const tableId = table.getObjectId();
  const rows = table.getNumRows();
  const cols = table.getNumColumns();
  const boldCols = parseColumns_(opts.boldColumns, cols);

  // セルの文字を先に読んでおく（結合で隠れたセルは null）
  const texts = [];
  const spans = [];
  for (let r = 0; r < rows; r++) {
    texts.push([]);
    spans.push([]);
    for (let c = 0; c < cols; c++) {
      const cell = table.getCell(r, c);
      const hidden = cell.getMergeState() === SlidesApp.CellMergeState.MERGED;
      texts[r].push(hidden ? null : cell.getText().asString().trim());
      spans[r].push(hidden ? 0 : cell.getColumnSpan());
    }
  }
  const numericCols = opts.alignNumbers ? findNumericColumns_(texts, spans) : [];

  const requests = [];

  // 1) セルの背景色と縦位置
  const vAlign = opts.vAlign && opts.vAlign !== 'KEEP' ? opts.vAlign : null;
  for (let r = 0; r < rows; r++) {
    let fill = preset.bodyFill;
    if (r === 0) fill = preset.headerFill;
    else if (preset.stripeFill && r % 2 === 0) fill = preset.stripeFill;

    const props = { tableCellBackgroundFill: { solidFill: { color: { rgbColor: hexToRgb_(fill) } } } };
    const fields = ['tableCellBackgroundFill'];
    if (vAlign) {
      props.contentAlignment = vAlign;
      fields.push('contentAlignment');
    }
    requests.push({
      updateTableCellProperties: {
        objectId: tableId,
        tableRange: { location: { rowIndex: r, columnIndex: 0 }, rowSpan: 1, columnSpan: cols },
        tableCellProperties: props,
        fields: fields.join(','),
      },
    });
  }

  // 2) 罫線
  preset.borders.forEach(b => requests.push(borderRequest_(tableId, b, cols)));

  // 3) 行の高さ・列幅（行・列の番号を省略すると、すべての行・列が対象になる）
  const minRowHeight = Number(opts.minRowHeight);
  if (minRowHeight > 0) {
    requests.push({
      updateTableRowProperties: {
        objectId: tableId,
        tableRowProperties: { minRowHeight: { magnitude: minRowHeight, unit: 'PT' } },
        fields: 'minRowHeight',
      },
    });
  }
  if (opts.equalColumns && cols > 1) {
    let total = 0;
    for (let c = 0; c < cols; c++) total += table.getColumn(c).getWidth();
    requests.push({
      updateTableColumnProperties: {
        objectId: tableId,
        tableColumnProperties: { columnWidth: { magnitude: Math.max(total / cols, MIN_COLUMN_WIDTH_PT), unit: 'PT' } },
        fields: 'columnWidth',
      },
    });
  }

  // 4) 文字の書式と横位置（空のセル・結合で隠れたセルはスキップ）
  const hAlign = opts.hAlign && opts.hAlign !== 'KEEP' ? opts.hAlign : null;
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      if (!texts[r][c]) continue;

      const isHeader = r === 0;
      const style = {
        foregroundColor: { opaqueColor: { rgbColor: hexToRgb_(isHeader ? preset.headerText : preset.bodyText) } },
        bold: isHeader ? opts.headerBold !== false : boldCols.includes(c),
      };
      const fields = ['foregroundColor', 'bold'];

      if (opts.fontFamily) {
        style.fontFamily = opts.fontFamily;
        fields.push('fontFamily');
      }
      const size = Number(isHeader ? opts.headerFontSize : opts.bodyFontSize);
      if (size > 0) {
        style.fontSize = { magnitude: size, unit: 'PT' };
        fields.push('fontSize');
      }

      requests.push({
        updateTextStyle: {
          objectId: tableId,
          cellLocation: { rowIndex: r, columnIndex: c },
          textRange: { type: 'ALL' },
          style: style,
          fields: fields.join(','),
        },
      });

      // 数値の列は見出しも含めて右揃え。それ以外は指定があるときだけ変更する
      const alignment = numericCols.includes(c) && spans[r][c] === 1 ? 'END' : hAlign;
      if (alignment) {
        requests.push({
          updateParagraphStyle: {
            objectId: tableId,
            cellLocation: { rowIndex: r, columnIndex: c },
            textRange: { type: 'ALL' },
            style: { alignment: alignment },
            fields: 'alignment',
          },
        });
      }
    }
  }
  return requests;
}

// ===== エラー =====
// 利用者向けのメッセージを持つエラー。サイドバーにはそのまま表示される
function userError_(message) {
  const e = new Error(message);
  e.name = 'UserError';
  return e;
}

// 想定外のエラーを、利用者が次にどうすればよいかわかるメッセージに置き換える
function toUserError_(e, msg) {
  if (e && e.name === 'UserError') return e;
  console.error(e);

  const detail = String((e && e.message) || e);
  if (/permission|PERMISSION_DENIED|forbidden|権限/i.test(detail)) {
    return userError_(msg.noPermission);
  }
  if (/authoriz|UNAUTHENTICATED|承認/i.test(detail)) {
    return userError_(msg.authExpired);
  }
  if (/Exceeded maximum execution time|too many times|RESOURCE_EXHAUSTED|rate limit|時間/i.test(detail)) {
    return userError_(msg.tooBusy);
  }
  if (/Invalid requests|INVALID_ARGUMENT/i.test(detail)) {
    return userError_(msg.invalidTable(detail));
  }
  return userError_(msg.unexpected(detail));
}

// ===== 色と線種の解決 =====
// Sidebar.html の resolvePreset も同じ規則でプレビューを描く。変えるときは両方そろえる
/** プリセットの 'accent' などを実際の色に置き換え、内側の線種の指定を反映したコピーを返す */
function resolvePreset_(preset, accent, innerDash) {
  const base = /^#[0-9A-F]{6}$/i.test(accent || '') ? accent : preset.accent;
  const color = value => resolveColor_(value, base);
  const headerFill = color(preset.headerFill);
  const dash = innerDash && innerDash !== 'KEEP' ? innerDash : null;
  return Object.assign({}, preset, {
    headerFill: headerFill,
    headerText: preset.headerText === 'auto' ? readableTextColor_(headerFill) : color(preset.headerText),
    bodyFill: color(preset.bodyFill),
    stripeFill: color(preset.stripeFill),
    bodyText: color(preset.bodyText),
    borders: preset.borders.map(b => Object.assign({}, b, {
      color: color(b.color),
      dash: dash && b.position.indexOf('INNER') === 0 ? dash : b.dash,
    })),
  });
}

// 'accent' → base、'accent/10' → base を白に10%混ぜた色。それ以外はそのまま
function resolveColor_(value, base) {
  const m = /^accent(?:\/(\d+))?$/.exec(value || '');
  if (!m) return value;
  if (!m[1]) return base;
  const ratio = Number(m[1]) / 100;
  const rgb = hexToRgb_(base);
  return '#' + ['red', 'green', 'blue']
    .map(k => Math.round((rgb[k] * ratio + (1 - ratio)) * 255).toString(16).padStart(2, '0'))
    .join('').toUpperCase();
}

// 背景が明るければ濃い灰色、暗ければ白
function readableTextColor_(background) {
  const c = hexToRgb_(background);
  return 0.299 * c.red + 0.587 * c.green + 0.114 * c.blue > 0.6 ? '#222222' : '#FFFFFF';
}

// ===== ヘルパー =====
function getTargetTables_(scope, msg) {
  const presentation = SlidesApp.getActivePresentation();

  if (scope === 'all') {
    const tables = [];
    presentation.getSlides().forEach(slide => tables.push(...slide.getTables()));
    if (tables.length === 0) throw userError_(msg.noTablesInPresentation);
    return tables;
  }

  if (scope === 'slide') {
    const page = presentation.getSelection().getCurrentPage();
    if (!page) throw userError_(msg.noSlideSelected);
    const tables = page.getTables();
    if (tables.length === 0) throw userError_(msg.noTablesOnSlide);
    return tables;
  }

  return [getSelectedTable_(msg)];
}

function getSelectedTable_(msg) {
  const sel = SlidesApp.getActivePresentation().getSelection();

  // セルやセル内の文字を選択している場合
  const cellRange = sel.getTableCellRange();
  if (cellRange) return cellRange.getTableCells()[0].getParentTable();

  // 表そのもの（枠）を選択している場合
  const range = sel.getPageElementRange();
  if (range) {
    const el = range.getPageElements()
      .find(e => e.getPageElementType() === SlidesApp.PageElementType.TABLE);
    if (el) return el.asTable();
  }
  throw userError_(msg.noTableSelected);
}

// 見出し行を除いて、空でないセルがすべて数値の列を返す（横に結合したセルは判定に使わない）
function findNumericColumns_(texts, spans) {
  const result = [];
  for (let c = 0; c < texts[0].length; c++) {
    let count = 0;
    let allNumeric = true;
    for (let r = 1; r < texts.length; r++) {
      if (!texts[r][c] || spans[r][c] !== 1) continue;
      count++;
      if (!isNumeric_(texts[r][c])) {
        allNumeric = false;
        break;
      }
    }
    if (count > 0 && allNumeric) result.push(c);
  }
  return result;
}

// "1,234" "-12.5%" "¥500" "▲300" "(1,200)" "３件" などを数値とみなす
function isNumeric_(text) {
  const s = text
    .replace(/[０-９．，％＋－]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
    .replace(/\s/g, '');
  return /^[+\-−▲△]?[¥￥$€£]?\(?[+\-−]?\d[\d,]*(\.\d+)?\)?(%|円|万|億|千|倍|件|人|個|pt|x)?$/i.test(s);
}

function borderRequest_(tableId, b, cols) {
  const req = {
    objectId: tableId,
    borderPosition: b.position,
    tableBorderProperties: {
      tableBorderFill: {
        solidFill: { color: { rgbColor: hexToRgb_(b.color || '#000000') }, alpha: b.color ? 1 : 0 },
      },
      weight: { magnitude: b.weight || 1, unit: 'PT' },
      dashStyle: b.dash || 'SOLID',
    },
    fields: 'tableBorderFill,weight,dashStyle',
  };
  if (b.range === 'header') {
    req.tableRange = { location: { rowIndex: 0, columnIndex: 0 }, rowSpan: 1, columnSpan: cols };
  }
  return { updateTableBorderProperties: req };
}

// "1,4" → [0, 3]（列番号は1始まりで入力、内部は0始まり）
function parseColumns_(text, cols) {
  if (!text) return [];
  return String(text).split(/[,、\s]+/)
    .map(s => parseInt(s, 10) - 1)
    .filter(n => !isNaN(n) && n >= 0 && n < cols);
}

function hexToRgb_(hex) {
  const h = hex.replace('#', '');
  return {
    red: parseInt(h.substring(0, 2), 16) / 255,
    green: parseInt(h.substring(2, 4), 16) / 255,
    blue: parseInt(h.substring(4, 6), 16) / 255,
  };
}
