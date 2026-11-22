/**
 * 底稿填充助手 · 核心逻辑
 *
 * 用途：把序时账按客户汇总后，填进事务所底稿的明细表空行，
 *      不动任何公式格，输出「一键填充宏」或直接生成的 xlsx。
 *
 * 设计原则（沿用前四个工具）：
 *   1. 只填空格和空行，公式格一律不碰
 *   2. 金额整数分
 *   3. 拿不准就拒绝，并说明第几行
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.WpCore = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var VERSION = '0.1.0';

  /* ---------------- 基础 ---------------- */

  function str(v) {
    if (v === null || v === undefined) return '';
    return String(v);
  }
  function norm(s) {
    return str(s).replace(/[\s\u00a0\u3000]/g, '').replace(/[（(][^）)]*[）)]/g, '').replace(/[：:]/g, '').toLowerCase();
  }
  function toCents(v) {
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'number') return isFinite(v) ? Math.round(v * 100) : null;
    var s = str(v).trim().replace(/[¥￥$,，\s]/g, '');
    if (!s || !/^-?\d+(\.\d+)?$/.test(s)) return null;
    return Math.round(parseFloat(s) * 100);
  }
  function moneyText(c) {
    if (c === null || c === undefined) return '';
    var n = c < 0, a = Math.abs(c);
    var y = Math.floor(a / 100), f = a % 100;
    var s = String(y).replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '.' + (f < 10 ? '0' + f : f);
    return (n ? '-' : '') + s;
  }

  /* ---------------- 序时账 ---------------- */

  var J_DATE = /日期|date/;
  var J_CODE = /科目代码|科目编码|科目编号|account.*code/;
  var J_NAME = /科目名称|科目|account.*name/;
  var J_DEBIT = /借方|debit/;
  var J_CREDIT = /贷方|credit/;
  var J_PARTY = /客户|供应商|往来|单位名称|客商/;
  var J_PARTY_CODE = /代码|编码|编号/;
  var J_SUMMARY = /摘要|说明|事由/;

  /** 找序时账表头（表头可能在第 2-5 行，前面有标题行） */
  function detectJournal(rows) {
    var best = null;
    var limit = Math.min(rows.length, 12);
    for (var r = 0; r < limit; r++) {
      var row = rows[r] || [];
      var cols = { date: -1, code: -1, name: -1, debit: -1, credit: -1, party: -1, partyCode: -1, summary: -1 };
      for (var c = 0; c < row.length; c++) {
        var h = str(row[c]).trim();
        if (!h || h.length > 16) continue;
        var n = norm(h);
        if (cols.date < 0 && J_DATE.test(n)) cols.date = c;
        else if (cols.debit < 0 && J_DEBIT.test(n)) cols.debit = c;
        else if (cols.credit < 0 && J_CREDIT.test(n)) cols.credit = c;
        else if (cols.code < 0 && J_CODE.test(n)) cols.code = c;
        else if (cols.name < 0 && J_NAME.test(n)) cols.name = c;
        else if (J_PARTY.test(n)) {
          // 「客户档案代码」和「客户档案名称」都命中客户，取名称那一列
          if (J_PARTY_CODE.test(n)) { if (cols.partyCode < 0) cols.partyCode = c; }
          else if (cols.party < 0) cols.party = c;
        }
        else if (cols.summary < 0 && J_SUMMARY.test(n)) cols.summary = c;
      }
      var score = (cols.debit >= 0 ? 2 : 0) + (cols.credit >= 0 ? 2 : 0) + (cols.code >= 0 ? 2 : 0) +
        (cols.name >= 0 ? 1 : 0) + (cols.party >= 0 ? 1 : 0) + (cols.date >= 0 ? 1 : 0);
      if (cols.debit >= 0 && cols.credit >= 0 && cols.code >= 0 && (!best || score > best.score)) {
        best = { rowIndex: r, cols: cols, score: score, heads: row };
      }
    }
    return best;
  }

  function readJournal(rows, header) {
    var out = [], skipped = 0;
    if (!header) return { records: out, skipped: 0 };
    var partyCol = header.cols.party >= 0 ? header.cols.party : header.cols.partyCode;
    for (var r = header.rowIndex + 1; r < rows.length; r++) {
      var row = rows[r] || [];
      var code = str(row[header.cols.code]).trim();
      var name = header.cols.name >= 0 ? str(row[header.cols.name]).trim() : '';
      var d = header.cols.debit >= 0 ? toCents(row[header.cols.debit]) : null;
      var c = header.cols.credit >= 0 ? toCents(row[header.cols.credit]) : null;
      if (!code && !name && d === null && c === null) continue;
      if (!code) { skipped++; continue; }
      out.push({
        code: code,
        name: name,
        debit: d || 0,
        credit: c || 0,
        party: partyCol >= 0 ? str(row[partyCol]).trim() : '',
        summary: header.cols.summary >= 0 ? str(row[header.cols.summary]).trim() : ''
      });
    }
    return { records: out, skipped: skipped };
  }

  /**
   * 按科目筛选并汇总到客户。
   * subject: 科目代码前缀（如 '2203'）或科目名称片段
   */
  /** ERP 导出的「客户档案－XXX」前缀去掉，底稿里只留单位名 */
  function cleanParty(name) {
    return str(name).replace(/^客户档案\s*[－\-—–]\s*/, '').replace(/^客商\s*[－\-—–]\s*/, '').trim();
  }

  function summarizeByParty(records, subject) {
    var key = norm(subject);
    var map = Object.create(null), order = [];
    for (var i = 0; i < records.length; i++) {
      var rec = records[i];
      var code = norm(rec.code), name = norm(rec.name);
      if (key && !(code.indexOf(key) === 0 || name.indexOf(key) === 0 || name.indexOf(key) >= 0)) continue;
      var party = cleanParty(rec.party) || '（未填客户）';
      if (!map[party]) { map[party] = { party: party, debit: 0, credit: 0, count: 0 }; order.push(party); }
      map[party].debit += rec.debit;
      map[party].credit += rec.credit;
      map[party].count++;
    }
    return order.map(function (p) {
      var o = map[p];
      o.balance = o.debit - o.credit;
      return o;
    });
  }

  /* ---------------- 底稿明细表 ---------------- */

  var D_NO = /^序号$|^编号$/;
  var D_NAME = /项目名称|客户名称|单位名称|明细|名称/;
  var D_OPEN = /期初金额|期初数|期初余额/;
  var D_DEBIT = /本期借方|借方发生/;
  var D_CREDIT = /本期贷方|贷方发生/;
  var D_CLOSE = /^期末金额$|期末余额|期末数/;

  /**
   * 找明细表的两级表头：上行可能是「序号/项目名称/本期未审」，下行是「期初金额/本期借方…」
   * 返回列映射 + 数据起始行
   */
  function detectDetail(rows) {
    for (var r = 0; r < Math.min(rows.length, 20); r++) {
      var row = rows[r] || [];
      var map = { no: -1, name: -1, open: -1, debit: -1, credit: -1, close: -1 };
      for (var c = 0; c < row.length; c++) {
        var h = norm(str(row[c]));
        if (!h || h.length > 14) continue;
        if (map.no < 0 && D_NO.test(h)) map.no = c;
        if (map.open < 0 && D_OPEN.test(h)) map.open = c;
        if (map.debit < 0 && D_DEBIT.test(h)) map.debit = c;
        if (map.credit < 0 && D_CREDIT.test(h)) map.credit = c;
        if (map.close < 0 && D_CLOSE.test(h)) map.close = c;
        if (map.name < 0 && D_NAME.test(h)) map.name = c;
      }
      // 「序号 + 项目名称」这一行是分组行；金额列在下一行
      if (map.name >= 0 && (map.no >= 0 || map.open >= 0)) {
        var next = rows[r + 1] || [];
        var m2 = { open: -1, debit: -1, credit: -1, close: -1 };
        for (var c2 = 0; c2 < next.length; c2++) {
          var h2 = norm(str(next[c2]));
          if (!h2 || h2.length > 14) continue;
          if (m2.open < 0 && D_OPEN.test(h2)) m2.open = c2;
          if (m2.debit < 0 && D_DEBIT.test(h2)) m2.debit = c2;
          if (m2.credit < 0 && D_CREDIT.test(h2)) m2.credit = c2;
          if (m2.close < 0 && D_CLOSE.test(h2)) m2.close = c2;
        }
        if (map.open < 0) map.open = m2.open;
        if (map.debit < 0) map.debit = m2.debit;
        if (map.credit < 0) map.credit = m2.credit;
        if (map.close < 0) map.close = m2.close;
        var headerRow = (m2.open >= 0 || m2.debit >= 0) ? r + 1 : r;
        return {
          usable: map.name >= 0 && map.debit >= 0 && map.credit >= 0,
          headerRow: headerRow,
          groupRow: r,
          cols: map,
          heads: row,
          heads2: next
        };
      }
    }
    return { usable: false, headerRow: -1, groupRow: -1, cols: {} };
  }

  /**
   * 找可填的空行：从表头下一行开始，跳过「一/二/三」这类分组标题行，
   * 只看「项目名称」为空、且金额列没有公式的行。
   */
  function findBlankRows(rows, detail, isFormula) {
    var out = [];
    var start = detail.headerRow + 1;
    for (var r = start; r < rows.length; r++) {
      var row = rows[r] || [];
      var nameCell = str(row[detail.cols.name]).trim();
      var noCell = detail.cols.no >= 0 ? str(row[detail.cols.no]).trim() : '';
      // 分组行：「一 / 二 / 三」或「合计」
      if (/^[一二三四五六七八九十]+$/.test(nameCell) || /合计|小计|总计/.test(nameCell)) continue;
      if (nameCell) continue;                       // 已经有内容，不覆盖
      var hasFormula = false;
      [detail.cols.open, detail.cols.debit, detail.cols.credit, detail.cols.close].forEach(function (c) {
        if (c >= 0 && isFormula(r, c)) hasFormula = true;
      });
      if (hasFormula) continue;                     // 公式格所在行不当空白行
      out.push(r);
    }
    return out;
  }

  /* ---------------- 生成写入计划 ---------------- */

  /**
   * parties: summarizeByParty 的结果
   * blanks: findBlankRows 的结果（Excel 行号，0 基）
   * detail: detectDetail 的结果
   * opts.openCents: 期初（客户级拿不到时统一给 0，让公式算）
   */
  function planWrites(parties, blanks, detail, opts) {
    opts = opts || {};
    var plan = [], problems = [];
    if (parties.length > blanks.length) {
      problems.push('客户有 ' + parties.length + ' 户，明细表只剩 ' + blanks.length + ' 个空行。' +
        '请先在底稿里插入足够的空行，或缩小科目范围。');
      return { plan: [], problems: problems };
    }
    for (var i = 0; i < parties.length; i++) {
      var p = parties[i], row = blanks[i];
      var open = opts.openCents ? (opts.openCents[p.party] || 0) : 0;
      var close = open + p.debit - p.credit;
      var cells = [];
      if (detail.cols.name >= 0) cells.push({ col: detail.cols.name, value: p.party, text: true });
      if (detail.cols.open >= 0) cells.push({ col: detail.cols.open, value: open / 100 });
      if (detail.cols.debit >= 0) cells.push({ col: detail.cols.debit, value: p.debit / 100 });
      if (detail.cols.credit >= 0) cells.push({ col: detail.cols.credit, value: p.credit / 100 });
      if (detail.cols.close >= 0) cells.push({ col: detail.cols.close, value: close / 100 });
      plan.push({ row: row, party: p.party, cells: cells, debit: p.debit, credit: p.credit, close: close });
    }
    return { plan: plan, problems: problems };
  }

  /* ---------------- 输出：宏 ---------------- */

  function colLetter(c) {
    var s = '';
    c = c + 1;
    while (c > 0) {
      var m = (c - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      c = Math.floor((c - 1) / 26);
    }
    return s;
  }

  /** 生成 VBA：只写计划里的格，绝不碰公式 */
  function buildVBA(plan, sheetName, opts) {
    opts = opts || {};
    var lines = [];
    lines.push("' 底稿填充助手 生成");
    lines.push("' 用法：Excel 里 Alt+F11 → 插入 → 模块 → 粘贴本段 → F5");
    lines.push("' 只写下面这些格，公式和其他内容一律不动");
    lines.push('Sub 填充底稿明细()');
    lines.push('    Dim ws As Worksheet');
    lines.push('    Set ws = ThisWorkbook.Worksheets("' + String(sheetName).replace(/"/g, '""') + '")');
    if (opts.backup !== false) {
      lines.push("    ' 先备份一份，填错了还能退回来");
      lines.push('    ThisWorkbook.SaveCopyAs ThisWorkbook.Path & "\\" & ThisWorkbook.Name & ".bak.xlsx"');
    }
    for (var i = 0; i < plan.length; i++) {
      var item = plan[i];
      for (var j = 0; j < item.cells.length; j++) {
        var cell = item.cells[j];
        var addr = colLetter(cell.col) + (item.row + 1);
        if (cell.text) {
          lines.push('    ws.Range("' + addr + '").Value = "' + String(cell.value).replace(/"/g, '""') + '"');
        } else {
          lines.push('    ws.Range("' + addr + '").Value = ' + Number(cell.value).toFixed(2));
        }
      }
    }
    lines.push('    MsgBox "已填入 ' + plan.length + ' 户，公式未改动。"');
    lines.push('End Sub');
    return lines.join('\r\n');
  }

  /** 生成可粘贴的对照表（CSV 文本） */
  function buildCsv(plan, detail) {
    var cols = ['客户/项目名称', '期初', '本期借方', '本期贷方', '期末'];
    var lines = [cols.join(',')];
    for (var i = 0; i < plan.length; i++) {
      var p = plan[i];
      var vals = { name: p.party, open: 0, debit: p.debit, credit: p.credit, close: p.close };
      var get = function (c) {
        for (var k = 0; k < p.cells.length; k++) if (p.cells[k].col === c) return p.cells[k].value;
        return '';
      };
      lines.push([
        '"' + String(p.party).replace(/"/g, '""') + '"',
        fmt(get(detail.cols.open)),
        fmt(get(detail.cols.debit)),
        fmt(get(detail.cols.credit)),
        fmt(get(detail.cols.close))
      ].join(','));
    }
    function fmt(v) { return typeof v === 'number' ? v.toFixed(2) : ''; }
    return lines.join('\r\n');
  }

  return {
    VERSION: VERSION,
    str: str,
    norm: norm,
    toCents: toCents,
    moneyText: moneyText,
    detectJournal: detectJournal,
    readJournal: readJournal,
    summarizeByParty: summarizeByParty,
    detectDetail: detectDetail,
    findBlankRows: findBlankRows,
    planWrites: planWrites,
    buildVBA: buildVBA,
    buildCsv: buildCsv,
    colLetter: colLetter,
    cleanParty: cleanParty
  };
});
