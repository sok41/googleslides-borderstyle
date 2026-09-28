/**
 * 表デザイン（Googleスライド用）
 * 選択中の表に、プリセットのデザインを一括で適用します。
 */

// ===== プリセット定義 =====
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

// ===== サイドバーから呼ばれるメイン処理 =====
/**
 * @param {string} presetKey PRESETSのキー
 * @param {Object} opts { fontFamily, headerFontSize, bodyFontSize, boldColumns, headerBold }
 */
function applyPreset(presetKey, opts) {
  const preset = PRESETS[presetKey];
  if (!preset) throw new Error('プリセットが見つかりません: ' + presetKey);

  const table = getSelectedTable_();
  const tableId = table.getObjectId();
  const rows = table.getNumRows();
  const cols = table.getNumColumns();
  const boldCols = parseColumns_(opts.boldColumns, cols);

  const requests = [];

  // 1) セルの背景色と縦位置（上揃え）
  for (let r = 0; r < rows; r++) {
    let fill = preset.bodyFill;
    if (r === 0) fill = preset.headerFill;
    else if (preset.stripeFill && r % 2 === 0) fill = preset.stripeFill;

    requests.push({
      updateTableCellProperties: {
        objectId: tableId,
        tableRange: { location: { rowIndex: r, columnIndex: 0 }, rowSpan: 1, columnSpan: cols },
        tableCellProperties: {
          tableCellBackgroundFill: { solidFill: { color: { rgbColor: hexToRgb_(fill) } } },
          contentAlignment: 'TOP',
        },
        fields: 'tableCellBackgroundFill,contentAlignment',
      },
    });
  }

  // 2) 罫線
  preset.borders.forEach(b => requests.push(borderRequest_(tableId, b, cols)));

  // 3) 文字の書式（空のセル・結合で隠れたセルはスキップ）
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const cell = table.getCell(r, c);
      if (cell.getMergeState() === SlidesApp.CellMergeState.MERGED) continue;
      if (cell.getText().asString().trim() === '') continue;

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
      requests.push({
        updateParagraphStyle: {
          objectId: tableId,
          cellLocation: { rowIndex: r, columnIndex: c },
          textRange: { type: 'ALL' },
          style: { alignment: 'START' },
          fields: 'alignment',
        },
      });
    }
  }

  Slides.Presentations.batchUpdate({ requests: requests }, SlidesApp.getActivePresentation().getId());
  return `「${preset.name}」を適用しました（${rows}行 × ${cols}列）`;
}

// ===== ヘルパー =====
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