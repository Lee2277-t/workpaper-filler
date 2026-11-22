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
  function detectJournal(rows, firstRow) {
    var best = null;
    var start = firstRow || 0;
    var limit = Math.min(rows.length, start + 12);
    for (var r = start; r < limit; r++) {
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
      var record = {
        code: code,
        name: name,
        debit: d || 0,
        credit: c || 0,
        party: partyCol >= 0 ? str(row[partyCol]).trim() : '',
        summary: header.cols.summary >= 0 ? str(row[header.cols.summary]).trim() : ''
      };
      if (d === null && str(row[header.cols.debit]).trim()) {
        record.amountError = '第 ' + (r + 1) + ' 行借方金额读不出：「' + row[header.cols.debit] + '」';
      } else if (c === null && str(row[header.cols.credit]).trim()) {
        record.amountError = '第 ' + (r + 1) + ' 行贷方金额读不出：「' + row[header.cols.credit] + '」';
      }
      out.push(record);
    }
    return { records: out, skipped: skipped };
  }

  /** ERP 导出的「客户档案－XXX」前缀去掉，底稿里只留单位名 */
  function cleanParty(name) {
    return str(name).replace(/^客户档案\s*[－\-—–]\s*/, '').replace(/^客商\s*[－\-—–]\s*/, '').trim();
  }

  function matchesSubject(record, key) {
    if (!key) return true;
    var code = norm(record.code);
    if (code === key) return true;
    // 科目名直接包含
    if (norm(record.name).indexOf(key) >= 0) return true;
    if (code.indexOf(key) !== 0) return false;
    var rest = code.slice(key.length);
    if (!rest) return true;
    // 1221.01 / 1221-01 这种带分隔符的
    if (/^[.\-_/]/.test(rest)) return true;
    // 纯数字尾：4-2-2 编码（1221 → 122101）算下级；
    // 但 1221 → 12210 这种只多一位的是别的科目，不算
    if (/^\d+$/.test(rest)) return rest.length === 2;
    return false;
  }

  function findJournalAmountIssue(records, subject) {
    var key = norm(subject);
    return records.find(function (record) { return record.amountError && matchesSubject(record, key); });
  }

  /** 按科目代码前缀或科目名称片段筛选并汇总到客户。 */
  function summarizeByParty(records, subject) {
    var key = norm(subject);
    var map = Object.create(null), order = [];
    for (var i = 0; i < records.length; i++) {
      var rec = records[i];
      if (!matchesSubject(rec, key)) continue;
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
  var D_AMOUNTS = {
    open: { pattern: /期初金额|期初数|期初余额/, label: '期初金额' },
    debit: { pattern: /本期借方|借方发生/, label: '本期借方' },
    credit: { pattern: /本期贷方|贷方发生/, label: '本期贷方' },
    close: { pattern: /^期末金额$|期末余额|期末数/, label: '期末金额' }
  };

  function detectDetailAmounts(row) {
    var cols = {}, duplicates = [];
    Object.keys(D_AMOUNTS).forEach(function (key) {
      var field = D_AMOUNTS[key];
      cols[key] = -1;
      for (var c = 0; c < row.length; c++) {
        var h = norm(row[c]);
        if (!h || h.length > 14 || !field.pattern.test(h)) continue;
        if (cols[key] >= 0) {
          duplicates.push(field.label);
          break;
        }
        cols[key] = c;
      }
    });
    return { cols: cols, duplicates: duplicates };
  }

  /**
   * 找明细表的两级表头：上行可能是「序号/项目名称/本期未审」，下行是「期初金额/本期借方…」
   * 返回列映射 + 数据起始行
   */
  function detectDetail(rows, firstRow) {
    var start = firstRow || 0;
    for (var r = start; r < Math.min(rows.length, start + 20); r++) {
      var row = rows[r] || [];
      var amounts = detectDetailAmounts(row);
      var map = amounts.cols;
      map.no = map.name = -1;
      for (var c = 0; c < row.length; c++) {
        var h = norm(str(row[c]));
        if (!h || h.length > 14) continue;
        if (map.no < 0 && D_NO.test(h)) map.no = c;
        if (map.name < 0 && D_NAME.test(h)) map.name = c;
      }
      // 「序号 + 项目名称」这一行是分组行；金额列在下一行
      if (map.name >= 0 && (map.no >= 0 || map.open >= 0)) {
        var next = rows[r + 1] || [];
        var second = detectDetailAmounts(next);
        var m2 = second.cols;
        var duplicates = amounts.duplicates.concat(second.duplicates);
        Object.keys(m2).forEach(function (key) {
          if (map[key] < 0) map[key] = m2[key];
        });
        var headerRow = (m2.open >= 0 || m2.debit >= 0) ? r + 1 : r;
        return {
          usable: !duplicates.length && map.name >= 0 && map.debit >= 0 && map.credit >= 0,
          error: duplicates.length ? '「' + duplicates[0] + '」列重复，请保留一个明确的写入列。' : '',
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
   * 找可填的空行。
   *
   * 明细表不是一直到底都空着：客户明细区下面还有账龄表、组合计提表，
   * 那些地方也有空格，但填进去就错了。所以到第一个「合计」行就停。
   * 分组（一、关联方 / 二、非关联方）单独列出来，让用户选往哪组填。
   */
  /**
   * 这一行是不是合计/小计行。
   * 不要求整格恰好等于「合计」——实际底稿里有「合计：」「合  计」「本项目小计」
   * 「2025年12月合计」各种写法，认漏了就会把客户填到合计行下面的账龄表里去。
   */
  function isTotalLabel(text) {
    var v = norm(str(text)).replace(/[：:，,。.、\s]/g, '');
    if (!v) return false;
    return /(合计|小计|总计|共计|累计)/.test(v);
  }
  /** 明细区的总合计（而不是组内小计） */
  function isGrandTotal(text) {
    var v = norm(str(text)).replace(/[：:，,。.、\s]/g, '');
    return /^(合计|总计|共计|累计|合计数|合计金额)$/.test(v) || /(^|[^本页本表])(合计|总计)$/.test(v);
  }

  function findBlankRows(rows, detail, isFormula, opts) {
    opts = opts || {};
    var out = [];
    var start = detail.headerRow + 1;
    var endRow = opts.endRow === undefined ? rows.length : opts.endRow;
    for (var r = start; r < endRow; r++) {
      var row = rows[r] || [];
      var nameCell = str(row[detail.cols.name]).trim();
      // 合计/小计行：这段明细到此结束，后面的空格是别的表
      if (isTotalLabel(nameCell)) break;
      if ((opts.isHidden && opts.isHidden(r))) continue;   // 隐藏行不写，用户看不见
      if (/^[一二三四五六七八九十]+$/.test(nameCell)) continue;
      if (nameCell) continue;                       // 已经有内容，不覆盖
      // 名称格本身是公式 → 不是给人填的行
      if (detail.cols.name >= 0 && isFormula(r, detail.cols.name)) continue;
      // 这一行已经有金额了（哪怕没有名字）→ 多半是小计/说明行，不能往里塞客户
      // 合并区的从属格：写进去 Excel 会静默丢掉，不能当空行
      var isMergedSub = opts.isMergedSubordinate || function () { return false; };
      if (detail.cols.name >= 0 && isMergedSub(r, detail.cols.name)) continue;
      var isOccupied = opts.isOccupied || function () { return false; };
      var occupied = false;
      [detail.cols.open, detail.cols.debit, detail.cols.credit, detail.cols.close].forEach(function (c) {
        if (c >= 0 && isOccupied(r, c)) occupied = true;
      });
      if (occupied) continue;
      // 至少得有一列是空着能写的；全被公式占满的行没意义
      // （很多模板「期末金额」整列是公式，那是它自己算的，不该拦住这行）
      var writable = 0;
      [detail.cols.open, detail.cols.debit, detail.cols.credit, detail.cols.close].forEach(function (c) {
        if (c >= 0 && !isFormula(r, c)) writable++;
      });
      if (!writable) continue;
      out.push(r);
    }
    return out;
  }

  /**
   * 列出客户明细区里的分组（「一、关联方」「二、非关联方」），
   * 每组的可填行。找不到分组就整块算一组。
   */
  function listGroups(rows, detail, isFormula, opts) {
    opts = opts || {};
    var start = detail.headerRow + 1;
    var marks = [];
    var grandRow = -1;
    for (var r = start; r < rows.length; r++) {
      var row = rows[r] || [];
      var nameCell = str(row[detail.cols.name]).trim();
      // 总合计：明细区到此为止
      if (isGrandTotal(nameCell) && grandRow < 0) { grandRow = r; break; }
      if (isTotalLabel(nameCell)) continue;          // 组内小计：是边界，不是终点
      var isGroup = nameCell.length <= 24 && !/合计|小计|总计/.test(nameCell) &&
        (/^[一二三四五六七八九十]+[、.．]?/.test(nameCell) || /关联方|组合计提|按款项性质/.test(nameCell)) &&
        !/[为是]$|占比|金额为/.test(nameCell);
      if (isGroup) marks.push({ row: r, title: nameCell });
    }
    var areaEnd = grandRow >= 0 ? grandRow : rows.length;
    // 没有分组标题：整块当一组（表头下方到合计行之前）
    if (!marks.length) {
      return [{
        title: '（未分组）',
        startRow: start,
        endRow: areaEnd,
        blanks: findBlankRows(rows, detail, isFormula, opts)
      }];
    }
    var out = [];
    for (var i = 0; i < marks.length; i++) {
      var from = marks[i].row + 1;
      var to = i + 1 < marks.length ? marks[i + 1].row : areaEnd;
      out.push({
        title: marks[i].title,
        startRow: from,
        endRow: to,
        blanks: findBlankRows(rows, detail, isFormula, {
          endRow: to,
          isOccupied: opts.isOccupied,
          isMergedSubordinate: opts.isMergedSubordinate,
          isHidden: opts.isHidden
        }).filter(function (rr) { return rr >= from; })
      });
    }
    // 表头下有行、但第一个分组标题之前也有空行：单独成组，排在最后，不抢默认
    if (marks.length && marks[0].row > start) {
      var headBlanks = findBlankRows(rows, detail, isFormula, {
        endRow: marks[0].row,
        isOccupied: opts.isOccupied,
        isMergedSubordinate: opts.isMergedSubordinate
      }).filter(function (rr) { return rr >= start && rr < marks[0].row; });
      if (headBlanks.length) {
        out.push({ title: '（分组标题以上）', startRow: start, endRow: marks[0].row, blanks: headBlanks });
      }
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
    var canWrite = opts.canWrite || function () { return true; };
    var isMergedSub = opts.isMergedSubordinate || function () { return false; };
    var plan = [], problems = [];
    // 空行不够不再拒绝：Excel 本来就该插行，插进去公式会跟着往下走。
    // 这里算出缺几行、从哪一行续，宏里负责插。
    var missing = Math.max(0, parties.length - blanks.length);
    var lastBlank = blanks.length ? blanks[blanks.length - 1] : -1;
    var rows = blanks.slice();
    for (var k = 0; k < missing; k++) rows.push(lastBlank + 1 + k);

    for (var i = 0; i < parties.length; i++) {
      var p = parties[i], row = rows[i];
      var open = opts.openCents ? (opts.openCents[p.party] || 0) : 0;
      var close = open + p.debit - p.credit;
      var cells = [];
      // 逐格判断：公式格不写，让它自己算。
      // 新插出来的行还没有公式（宏会先复制上一行再写），这里照样写。
      var add = function (col, value, text) {
        if (col < 0) return;
        if (isMergedSub(row, col)) return;          // 合并区从属格写进去会丢
        if (row <= lastBlank && !canWrite(row, col)) return;
        cells.push({ col: col, value: value, text: !!text });
      };
      add(detail.cols.name, p.party, true);
      add(detail.cols.open, open / 100);
      add(detail.cols.debit, p.debit / 100);
      add(detail.cols.credit, p.credit / 100);
      add(detail.cols.close, close / 100);
      plan.push({ row: row, party: p.party, cells: cells, debit: p.debit, credit: p.credit, close: close, inserted: row > lastBlank });
    }

    if (missing > 0) {
      problems.push('客户有 ' + parties.length + ' 户，明细表只有 ' + blanks.length + ' 个空行，' +
        '要往下补 ' + missing + ' 行。用「下载填充宏」会自动插：照第 ' + (lastBlank + 1) +
        ' 行复制，公式跟着往下走；「直接生成 xlsx」不支持插行，这种情况请用宏。');
    }
    return { plan: plan, problems: problems, missing: missing, insertAfter: lastBlank };
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

  /** VBA 字符串字面量：双引号成对转义，换行拆成 Chr(10) 拼接（否则 .bas 编译不过） */
  function vbaString(value) {
    var parts = String(value).split(/\r\n|\r|\n/);
    var escaped = parts.map(function (seg) { return '"' + seg.replace(/"/g, '""') + '"'; });
    return escaped.join(' & Chr(10) & ');
  }

  /** 生成 VBA：只写计划里的格，绝不碰公式；不够的行自动插，公式跟着复制 */
  function buildVBA(plan, sheetName, opts) {
    opts = opts || {};
    var lines = [];
    var missing = opts.missing || 0;
    var insertAfter = opts.insertAfter;      // 0 基；插在它下面
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
    if (missing > 0 && insertAfter >= 0) {
      // 把最后一行明细整行复制下来插入：格式和公式都跟着，相对引用自动顺延
      var src = insertAfter + 1;               // Excel 行号
      lines.push("    ' 空行不够：照第 " + src + " 行复制 " + missing + ' 行，公式会跟着往下走');
      lines.push('    Application.ScreenUpdating = False');
      lines.push('    ws.Rows(' + src + ').Copy');
      lines.push('    ws.Rows(' + (src + 1) + ':' + (src + missing) + ').Insert Shift:=xlDown');
      lines.push('    Application.CutCopyMode = False');
      lines.push('    Application.ScreenUpdating = True');
    }
    var guards = [], writes = [];
    for (var i = 0; i < plan.length; i++) {
      var item = plan[i];
      for (var j = 0; j < item.cells.length; j++) {
        var cell = item.cells[j];
        var addr = colLetter(cell.col) + (item.row + 1);
        var target = 'ws.Range("' + addr + '")';
        guards.push('    If ' + target + '.HasFormula Or ' + target + '.HasArray Then GoTo FormulaFound');
        guards.push('    If ' + target + '.MergeCells And ' + target + '.Address <> ' + target + '.MergeArea.Cells(1).Address Then GoTo MergeFound');
        writes.push('    ' + target + '.Value = ' + (cell.text
          ? vbaString(cell.value)
          : Number(cell.value).toFixed(2)));
      }
    }
    lines = lines.concat(guards);
    lines = lines.concat(writes);
    lines.push('    MsgBox "已填入 ' + plan.length + ' 户' + (missing > 0 ? '，并补了 ' + missing + ' 行' : '') +
      '。公式未改动，请核对合计行是否已把新行算进去。"');
    lines.push('    Exit Sub');
    lines.push('FormulaFound:');
    lines.push('    MsgBox "目标格包含公式，已停止，没有写入任何单元格。"');
    lines.push('    Exit Sub');
    lines.push('MergeFound:');
    lines.push('    MsgBox "目标格是合并单元格的一部分，写进去会被 Excel 丢掉，已停止。"');
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
    findJournalAmountIssue: findJournalAmountIssue,
    summarizeByParty: summarizeByParty,
    detectDetail: detectDetail,
    findBlankRows: findBlankRows,
    listGroups: listGroups,
    planWrites: planWrites,
    buildVBA: buildVBA,
    buildCsv: buildCsv,
    colLetter: colLetter,
    cleanParty: cleanParty
  };
});
