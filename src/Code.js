/**
 * 表デザイン（Googleスライド用）
 * 選択中の表に、プリセットのデザインを一括で適用します。
 */

// ===== プリセット定義 =====
// サイドバーのプレビューもこの定義から作られます（getPresets）。
// borders: 上から順に適用されます。color: null は「罫線なし（透明）」。
// range: 'header' を付けると、ヘッダー行だけに適用します。
const PRESETS = {
  cleanGray: {
    name: 'クリーン・グレー',
    headerFill: '#F2F2F2', headerText: '#333333',
    bodyFill: '#FFFFFF', stripeFill: null, bodyText: '#333333',
    borders: [
      { position: 'ALL', color: '#D9D9D9', weight: 1 },
      { position: 'LEFT', color: null },
      { position: 'RIGHT', color: null },
    ],
  },
  navyStripe: {
    name: 'ネイビー・ストライプ',
    headerFill: '#1F3A5F', headerText: '#FFFFFF',
    bodyFill: '#FFFFFF', stripeFill: '#F3F6FA', bodyText: '#333333',
    borders: [
      { position: 'ALL', color: null },
      { position: 'INNER_HORIZONTAL', color: '#E3E8EF', weight: 0.75 },
      { position: 'BOTTOM', color: '#1F3A5F', weight: 1.5 },
    ],
  },
  minimal: {
    name: 'ミニマル（横線のみ）',
    headerFill: '#FFFFFF', headerText: '#222222',
    bodyFill: '#FFFFFF', stripeFill: null, bodyText: '#333333',
    borders: [
      { position: 'ALL', color: null },
      { position: 'INNER_HORIZONTAL', color: '#E0E0E0', weight: 0.75 },
      { position: 'TOP', color: '#222222', weight: 1.5 },
      { position: 'BOTTOM', color: '#222222', weight: 1.5 },
      { position: 'BOTTOM', color: '#222222', weight: 1, range: 'header' },
    ],
  },
};

// ===== メニューとサイドバー =====
function onOpen() {
  SlidesApp.getUi()
    .createMenu('表デザイン')
    .addItem('サイドバーを開く', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('表デザイン');
  SlidesApp.getUi().showSidebar(html);
}

// ===== サイドバーから呼ばれる処理 =====
/** サイドバーに表示するプリセット一覧 */
function getPresets() {
  return Object.keys(PRESETS).map(key => Object.assign({ key: key }, PRESETS[key]));
}

/**
 * @param {string} presetKey PRESETSのキー
 * @param {Object} opts {
 *   scope: 'selection' | 'slide' | 'all',
 *   hAlign: 'KEEP' | 'START' | 'CENTER',
 *   vAlign: 'KEEP' | 'TOP' | 'MIDDLE',
 *   alignNumbers: boolean,
 *   fontFamily, headerFontSize, bodyFontSize, boldColumns, headerBold
 * }
 */
function applyPreset(presetKey, opts) {
  const preset = PRESETS[presetKey];
  if (!preset) throw new Error('プリセットが見つかりません: ' + presetKey);

  const tables = getTargetTables_(opts.scope);
  const requests = [];
  tables.forEach(table => requests.push(...tableRequests_(table, preset, opts)));

  Slides.Presentations.batchUpdate({ requests: requests }, SlidesApp.getActivePresentation().getId());

  if (tables.length === 1) {
    return `「${preset.name}」を適用しました（${tables[0].getNumRows()}行 × ${tables[0].getNumColumns()}列）`;
  }
  return `「${preset.name}」を${tables.length}個の表に適用しました`;
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

  // 3) 文字の書式と横位置（空のセル・結合で隠れたセルはスキップ）
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

// ===== ヘルパー =====
function getTargetTables_(scope) {
  const presentation = SlidesApp.getActivePresentation();

  if (scope === 'all') {
    const tables = [];
    presentation.getSlides().forEach(slide => tables.push(...slide.getTables()));
    if (tables.length === 0) throw new Error('このプレゼンテーションには表がありません。');
    return tables;
  }

  if (scope === 'slide') {
    const page = presentation.getSelection().getCurrentPage();
    if (!page) throw new Error('スライドが選択されていません。対象のスライドを開いてから、もう一度「適用」を押してください。');
    const tables = page.getTables();
    if (tables.length === 0) throw new Error('このスライドには表がありません。');
    return tables;
  }

  return [getSelectedTable_()];
}

function getSelectedTable_() {
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
  throw new Error('表が選択されていません。表の中をクリックしてから、もう一度「適用」を押してください。');
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
      dashStyle: 'SOLID',
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
