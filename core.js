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

  var VERSION = '0.4.3';

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
    for (var r = start; r < Math.min(rows.length, start + 60); r++) {
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

  /** 这一行是不是明细区的边界（合计/小计）。
   *  合计不一定写在项目名称列，有时在 A 列，所以名称列为空时再看整行。 */
  function isRowBoundary(row, detail, nameCell) {
    if (isTotalLabel(nameCell)) return true;
    if (nameCell) return false;
    for (var c = 0; c < row.length; c++) {
      if (c === detail.cols.name) continue;
      if (isTotalLabel(str(row[c]).trim())) return true;
    }
    return false;
  }

  function findBlankRows(rows, detail, isFormula, opts) {
    opts = opts || {};
    var out = [];
    // 从哪一行开始扫：第二个分组要跳过第一个分组的小计，
    // 否则从表头扫起一碰见「小计」就 break，二组以后永远 0 个空行
    var start = opts.startRow === undefined ? detail.headerRow + 1 : opts.startRow;
    var endRow = opts.endRow === undefined ? rows.length : opts.endRow;
    var isOccupied = opts.isOccupied || function () { return false; };
    var isMerged = opts.isMergedSubordinate || function () { return false; };
    var seenRow = false;   // 已经进过明细行（空行或有金额的行）
    for (var r = start; r < endRow; r++) {
      var row = rows[r] || [];
      var nameCell = str(row[detail.cols.name]).trim();
      // 合计/小计行：这段明细到此结束。
      // 「小计在上、明细在下」的块标题由 listGroups 当成独立分组，扫不到这里。
      if (isRowBoundary(row, detail, nameCell)) break;
      var hasAmount = false;
      [detail.cols.open, detail.cols.debit, detail.cols.credit, detail.cols.close].forEach(function (c) {
        if (c < 0) return;
        var v = row[c];
        if (v !== null && v !== undefined && String(v).trim() !== '') hasAmount = true;
      });
      // 明细区里冒出一行「只有文字、没有任何金额」的标题（账龄分析 / 说明 / 备注），
      // 而且它自己不在任何合计范围内 → 那是下一张表，到此为止。
      // 在合计范围内的说明行只是跳过，不截断（否则同一段里后面还有空行就填不到了）
      if (nameCell && !hasAmount && seenRow && opts.inTotalRange && !opts.inTotalRange(r)) break;
      if ((opts.isHidden && opts.isHidden(r))) continue;   // 隐藏行不写，用户看不见
      if (/^[一二三四五六七八九十]+$/.test(nameCell)) continue;
      if (nameCell) { if (hasAmount) seenRow = true; continue; }   // 已经有内容，不覆盖
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
      // 硬判据放在最后：先排除「不是空行」的，剩下的空行才谈「在不在合计范围内」
      if (opts.inTotalRange && !opts.inTotalRange(r)) {
        if (opts.uncovered) opts.uncovered.push(r);
        continue;
      }
      seenRow = true;
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
      // 块标题小计（SUM 指向下方）自成一段：1、合并范围内小计 / 2、非合并范围内小计
      if (opts.isBlockHeader && opts.isBlockHeader(r)) {
        var label = nameCell;
        if (!label) {
          for (var bc = 0; bc < row.length; bc++) {
            var bv = str(row[bc]).trim();
            if (bv && isTotalLabel(bv)) { label = bv; break; }
          }
        }
        marks.push({ row: r, title: label || ('第 ' + (r + 1) + ' 行小计') });
        continue;
      }
      if (isTotalLabel(nameCell)) continue;          // 组内小计：是边界，不是终点
      // 标题可能合并进 A 列（A:B / A:F），此时名称列是空从属格，要看整行
      var titleText = nameCell;
      if (!titleText) {
        for (var tc = 0; tc < row.length; tc++) {
          var tv = str(row[tc]).trim();
          if (tv && /^[一二三四五六七八九十]+[、.．]?/.test(tv) || /关联方|组合计提|按款项性质/.test(tv)) {
            if (tv && tv.length <= 24) { titleText = tv; break; }
          }
        }
      }
      var isGroup = titleText.length <= 24 && !/合计|小计|总计/.test(titleText) &&
        (/^[一二三四五六七八九十]+[、.．]?/.test(titleText) || /关联方|组合计提|按款项性质/.test(titleText)) &&
        !/[为是]$|占比|金额为/.test(titleText);
      if (isGroup) marks.push({ row: r, title: titleText });
    }
    var areaEnd = grandRow >= 0 ? grandRow : rows.length;
    // 没有分组标题：整块当一组（表头下方到合计行之前）
    if (!marks.length) {
      var headUncovered = [];
      return [{
        title: '（未分组）',
        startRow: start,
        endRow: areaEnd,
        uncovered: headUncovered,
        blanks: findBlankRows(rows, detail, isFormula, {
          uncovered: headUncovered,
          startRow: start,
          endRow: areaEnd,
          isOccupied: opts.isOccupied,
          isMergedSubordinate: opts.isMergedSubordinate,
          isHidden: opts.isHidden,
          inTotalRange: opts.inTotalRange
        })
      }];
    }
    var out = [];
    for (var i = 0; i < marks.length; i++) {
      var from = marks[i].row + 1;
      var to = i + 1 < marks.length ? marks[i + 1].row : areaEnd;
      // 空行但不在合计公式覆盖范围里：填了合计也不会算进去，单独记下来提示用户
      var uncovered = [];
      var groupBlanks = findBlankRows(rows, detail, isFormula, {
        uncovered: uncovered,
        startRow: from,
        endRow: to,
        isOccupied: opts.isOccupied,
        isMergedSubordinate: opts.isMergedSubordinate,
        isHidden: opts.isHidden,
        inTotalRange: opts.inTotalRange
      }).filter(function (rr) { return rr >= from; });
      out.push({
        title: marks[i].title,
        startRow: from,
        endRow: to,
        uncovered: uncovered,
        blanks: groupBlanks
      });
    }
    // 表头下有行、但第一个分组标题之前也有空行：单独成组，排在最后，不抢默认
    if (marks.length && marks[0].row > start) {
      var headBlanks = findBlankRows(rows, detail, isFormula, {
        startRow: start,
        endRow: marks[0].row,
        isOccupied: opts.isOccupied,
        isMergedSubordinate: opts.isMergedSubordinate,
        isHidden: opts.isHidden,
          inTotalRange: opts.inTotalRange
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
  /**
   * 解析用户粘进来的期初（从 Excel 复制两列：客户名 + 期初金额）。
   * 制表符、逗号、多个空格都认；表头行自动跳过。
   */
  function parseOpening(text) {
    var map = Object.create(null);
    var bad = [];
    var lines = str(text).split(/\r?\n/);
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i];
      if (!line.trim()) continue;
      var parts = line.indexOf('\t') >= 0 ? line.split('\t') : line.split(/[,，]/);
      if (parts.length < 2) { bad.push({ line: i + 1, text: line }); continue; }
      var name = str(parts[0]).trim();
      var amount = toCents(parts[1]);
      if (!name) continue;
      // 表头行
      if (i === 0 && /客户|名称|单位|期初/.test(name) && amount === null) continue;
      if (amount === null) { bad.push({ line: i + 1, text: line }); continue; }
      map[cleanParty(name)] = amount;
    }
    return { map: map, bad: bad };
  }

  function planWrites(parties, blanks, detail, opts) {
    opts = opts || {};
    var canWrite = opts.canWrite || function () { return true; };
    var isMergedSub = opts.isMergedSubordinate || function () { return false; };
    var plan = [], problems = [], missingOpen = [], mergedOut = [];
    // 空行数为 0 时无法判断从哪一行开始续，硬算会把客户写进标题行/表头
    if (!blanks.length) {
      problems.push('这一组没有可填的空行。请先在底稿里给这一组插入几行空行，再重新生成。');
      return { plan: [], problems: problems, missing: 0, insertAfter: -1 };
    }
    // 空行不够不再拒绝：Excel 本来就该插行，插进去公式会跟着往下走。
    // 这里算出缺几行、从哪一行续，宏里负责插。
    // 但用户手填了范围时不许越界插行——那会写到他没指定的地方。
    var missing = Math.max(0, parties.length - blanks.length);
    if (opts.noInsert && missing > 0) {
      var dropped = parties.slice(blanks.length).map(function (x) { return x.party; });
      problems.push('你指定的范围里只有 ' + blanks.length + ' 个能填的空行，客户有 ' +
        parties.length + ' 户，多出来的 ' + missing + ' 户没有填：' +
        dropped.slice(0, 5).join('、') + (dropped.length > 5 ? ' 等' : '') + '。把范围改大一些再来一次。');
      parties = parties.slice(0, blanks.length);
      missing = 0;
    }
    var lastBlank = blanks.length ? blanks[blanks.length - 1] : -1;
    var rows = blanks.slice();
    for (var k = 0; k < missing; k++) rows.push(lastBlank + 1 + k);

    for (var i = 0; i < parties.length; i++) {
      var p = parties[i], row = rows[i];
      var open = opts.openCents ? (opts.openCents[p.party] || 0) : 0;
      if (opts.openCents && !(p.party in opts.openCents)) missingOpen.push(p.party);
      // 期末方向：资产类 期初+借−贷；负债/权益类 期初−借+贷
      var sign = opts.closeSign === -1 ? -1 : 1;
      var close = open + sign * p.debit - sign * p.credit;
      var cells = [];
      // 新插的行是照最后一行复制的，公式和合并都一起下来了：一律按模板行判断
      var refRow = row <= lastBlank ? row : lastBlank;
      // 逐格判断：公式格不写，让它自己算。
      var add = function (col, value, text) {
        if (col < 0) return;
        if (isMergedSub(refRow, col)) return;       // 合并区从属格写进去会丢
        if (!canWrite(refRow, col)) return;
        cells.push({ col: col, value: value, text: !!text });
      };
      // 整行检查：有任何一个要写的格被合并盖住 → 这一户整行不填，
      // 否则会出现「期初没写进去、期末却按期初算了」的勾稽不平
      var wanted = [detail.cols.name, detail.cols.open, detail.cols.debit, detail.cols.credit, detail.cols.close];
      var blocked = null;
      for (var w = 0; w < wanted.length; w++) {
        var wc = wanted[w];
        if (wc >= 0 && isMergedSub(refRow, wc)) { blocked = wc; break; }
      }
      if (blocked !== null) {
        mergedOut.push({ party: p.party, row: row, col: blocked });
        continue;
      }
      add(detail.cols.name, p.party, true);
      add(detail.cols.open, open / 100);
      add(detail.cols.debit, p.debit / 100);
      add(detail.cols.credit, p.credit / 100);
      add(detail.cols.close, close / 100);
      plan.push({ row: row, party: p.party, cells: cells, debit: p.debit, credit: p.credit, close: close, inserted: row > lastBlank });
    }

    if (mergedOut.length) {
      problems.push('有 ' + mergedOut.length + ' 户在明细表里被合并单元格挡住了（' +
        mergedOut.slice(0, 3).map(function (m) { return m.party + ' 第 ' + (m.row + 1) + ' 行'; }).join('、') +
        (mergedOut.length > 3 ? ' 等' : '') + '），这些户没有填，请先取消合并再生成。');
    }
    if (missingOpen.length) {
      problems.push('有 ' + missingOpen.length + ' 户在期初里没给数（' + missingOpen.slice(0, 3).join('、') +
        (missingOpen.length > 3 ? ' 等' : '') + '），这些户的期初按 0 填，期末会偏小。');
    }

    if (missing > 0) {
      problems.push('客户有 ' + parties.length + ' 户，明细表只有 ' + blanks.length + ' 个空行，' +
        '要往下补 ' + missing + ' 行。用「下载填充宏」会自动插：照第 ' + (lastBlank + 1) +
        ' 行复制，发现合计范围漏掉新行时会询问；「直接生成 xlsx」不支持插行，这种情况请用宏。');
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

  /** 生成 VBA：填充计划格，插行时经用户确认才扩展合计公式 */
  function buildVBA(plan, sheetName, opts) {
    opts = opts || {};
    var lines = [];
    var missing = opts.missing || 0;
    var insertAfter = opts.insertAfter;      // 0 基；插在它下面
    var willInsert = missing > 0 && insertAfter >= 0;
    lines.push("' 底稿填充助手 生成");
    lines.push("' 用法：Excel 里 Alt+F11 → 插入 → 模块 → 粘贴本段 → F5");
    lines.push(willInsert ? "' 填写计划格，扩展合计公式前先询问" : "' 只写下面这些格，公式和其他内容一律不动");
    lines.push("' ===== 开关 =====");
    lines.push("' 不想看到弹窗就把它改成 True：照样填数，只是不提示");
    lines.push('Private Const QUIET As Boolean = False');
    lines.push('');
    lines.push("' 逐格相加的合计个数（=D3+D4 这种），插行后不会自动含新行");
    lines.push("' VBA 要求模块级变量都写在过程之前，不能放在 End Sub 后面");
    lines.push('Private plainTotals As Long');
    lines.push('');
    lines.push("' 所有提示都走这里，方便你一处关掉");
    lines.push('Private Sub Say(msg As String, Optional icon As Long = 64)');
    lines.push('    If Not QUIET Then MsgBox msg, icon, "底稿填充助手"');
    lines.push('End Sub');
    lines.push('');
    lines.push('Sub 填充底稿明细()');
    lines.push('    plainTotals = 0');
    lines.push('    Dim ws As Worksheet');
    lines.push('    On Error GoTo Fail');
    // 名字写错、工作表被删时 1004 会先炸在这里，给一句人话
    lines.push('    On Error Resume Next');
    lines.push('    Set ws = ThisWorkbook.Worksheets("' + String(sheetName).replace(/"/g, '""') + '")');
    lines.push('    On Error GoTo Fail');
    lines.push('    If ws Is Nothing Then');
    lines.push('        Say "找不到工作表「' + String(sheetName).replace(/"/g, '""') + '」。' +
      '可能改了表名，或者宏跑在另一个工作簿上（要在底稿那个工作簿里跑）。", vbExclamation');
    lines.push('        Exit Sub');
    lines.push('    End If');
    lines.push('    If ws.ProtectContents Then');
    lines.push('        Say "工作表「' + String(sheetName).replace(/"/g, '""') + '」有保护，写不进去。" & vbCrLf & _');
    lines.push('            "请先在「审阅」→「撤销工作表保护」里解开，再跑一次。", vbExclamation');
    lines.push('        Exit Sub');
    lines.push('    End If');
    if (opts.backup !== false) {
      lines.push("    ' 先备份一份，填错了还能退回来（没保存过的新文件没有路径，跳过）");
      lines.push('    If ThisWorkbook.Path <> "" Then');
      lines.push('        On Error Resume Next');
      lines.push('        ThisWorkbook.SaveCopyAs ThisWorkbook.Path & "\\" & _');
      lines.push('            Left$(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1) & "_填充前备份.xlsx"');
      lines.push('        On Error GoTo Fail');
      lines.push('    End If');
    }
    if (willInsert) {
      // 把最后一行明细整行复制下来插入：格式和公式都跟着，相对引用自动顺延
      var src = insertAfter + 1;               // Excel 行号
      lines.push('    Dim totalChanges As New Collection, change As Variant, expandTotals As Boolean');
      lines.push('    expandTotals = ConfirmTotalExpansion(ws, ' + src + ', ' + missing + ', totalChanges)');
      lines.push("    ' 空行不够：照第 " + src + " 行复制 " + missing + ' 行，公式会跟着往下走');
      lines.push('    On Error GoTo InsertFail');
      lines.push('    Application.ScreenUpdating = False');
      lines.push('    ws.Rows(' + src + ').Copy');
      // 插在最后一行空行的位置上（不是它下面）：
      // 插入点落在合计公式的引用范围之内，Excel 会自动把范围撑开，公式不用改
      lines.push('    ws.Rows("' + src + ':' + (src + missing - 1) + '").Insert Shift:=xlDown');
      lines.push('    Application.CutCopyMode = False');
      lines.push('    Application.ScreenUpdating = True');
      lines.push('    On Error GoTo Fail');
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
    if (willInsert) {
      lines.push('    Dim result As String');
      lines.push('    result = "已填入 ' + plan.length + ' 户，插入了 ' + missing + ' 行。"');
      lines.push('    If expandTotals Then');
      lines.push('        For Each change In totalChanges');
      lines.push('            ws.Cells(change(0), change(1)).Formula = change(2)');
      lines.push('        Next change');
      lines.push('        result = result & vbCrLf & "扩了 " & totalChanges.Count & ' +
        '" 个合计公式的范围，合计已含新增行。"');
      lines.push('    ElseIf totalChanges.Count > 0 Then');
      lines.push('        result = result & vbCrLf & totalChanges.Count & ' +
        '" 个合计公式的范围没动，它们会漏掉新增行，需要你自己核一下。"');
      lines.push('    ElseIf plainTotals > 0 Then');
      lines.push('        result = result & vbCrLf & plainTotals & " 个合计是逐格相加写的（比如 =D3+D4），" & _');
      lines.push('            "插行后 Excel 不会把它们扩到新行，合计会漏，请手工改。"');
      lines.push('    Else');
      lines.push('        result = result & vbCrLf & "新行插在合计公式的范围内，Excel 已自动把它算进去，公式没动。"');
      lines.push('    End If');
      lines.push('    Say result');
    } else {
      var doneMsg = '已填入 ' + plan.length + ' 户。公式没动。';
      if (opts.outsideNote) doneMsg += ' ' + opts.outsideNote;
      lines.push('    Say "' + doneMsg + '"');
    }
    lines.push('    Exit Sub');
    lines.push('FormulaFound:');
    lines.push('    Say "目标格包含公式，已停止，没有写入任何单元格。"');
    lines.push('    Exit Sub');
    lines.push('MergeFound:');
    lines.push('    Say "目标格是合并单元格的一部分，写进去会被 Excel 丢掉，已停止。"');
    lines.push('    Exit Sub');
    // 任何没预料到的错误（多半是 1004）都给一句人话，不要弹「应用程序定义或对象定义错误」
    if (willInsert) {
      lines.push('InsertFail:');
      lines.push('    On Error Resume Next');
      lines.push('    Application.CutCopyMode = False');
      lines.push('    Application.ScreenUpdating = True');
      lines.push('    Say "插第 ' + (insertAfter + 2) + ' 行时失败了，没有写入任何单元格。" & vbCrLf & vbCrLf & _');
      lines.push('        "最常见的原因是这一行的合并单元格跨到了下面一行——" & vbCrLf & _');
      lines.push('        "先手工在明细区插入 ' + missing + ' 行，再跑一次这个宏（行数够了就不会再插）。", vbCritical');
      lines.push('    Exit Sub');
    }
    lines.push('Fail:');
    lines.push('    Application.ScreenUpdating = True');
    lines.push('    Application.CutCopyMode = False');
    lines.push('    Say "宏跑到一半出错了：" & Err.Number & " " & Err.Description & vbCrLf & vbCrLf & _');
    lines.push('        "常见原因：" & vbCrLf & _');
    lines.push('        "· 工作表被保护——先撤销保护" & vbCrLf & _');
    lines.push('        "· 宏跑在了别的工作簿上——要在底稿那个工作簿里跑" & vbCrLf & _');
    lines.push('        "· 目标行被隐藏或合并且公式要求整行插入" & vbCrLf & vbCrLf & _');
    lines.push('        "已写入的单元格不会自动撤回，可用同目录的「_填充前备份.xlsx」恢复。", vbCritical');
    lines.push('End Sub');
    if (willInsert) {
      lines.push('',
        'Private Function ConfirmTotalExpansion(ws As Worksheet, ByVal lastDetailRow As Long, ByVal insertedRows As Long, changes As Collection) As Boolean',
        '    Dim formulas As Range, cell As Range, sumRange As Range',
        '    Dim formula As String, refText As String, address As String, proposed As String, examples As String',
        '    Dim startRef As String, endRef As String, parts() As String',
        '    On Error Resume Next',
        '    Set formulas = ws.UsedRange.SpecialCells(xlCellTypeFormulas)',
        '    On Error GoTo 0',
        '    If formulas Is Nothing Then Exit Function',
        '    For Each cell In formulas.Cells',
        '        If cell.Row <= lastDetailRow Or cell.HasArray Then GoTo NextFormula',
        '        formula = cell.Formula',
        '        If Left$(UCase$(formula), 5) <> "=SUM(" Then',
        '            If InStr(formula, "+") > 0 Then plainTotals = plainTotals + 1',
        '            GoTo NextFormula',
        '        End If',
        '        If Right$(formula, 1) <> ")" Then GoTo NextFormula',
        '        refText = Replace(Mid$(formula, 6, Len(formula) - 6), " ", "")',
        '        Set sumRange = Nothing',
        '        On Error Resume Next',
        '        Set sumRange = ws.Range(refText)',
        '        On Error GoTo 0',
        '        If sumRange Is Nothing Then GoTo NextFormula',
        '        address = sumRange.Address(False, False)',
        '        If address <> Replace(refText, "$", "") And address & ":" & address <> Replace(refText, "$", "") Then GoTo NextFormula',
        '        If sumRange.Columns.Count <> 1 Then GoTo NextFormula',
        '        If sumRange.Row + sumRange.Rows.Count - 1 >= lastDetailRow Then GoTo NextFormula',
        '        If InStr(refText, ":") = 0 Then refText = refText & ":" & refText',
        '        parts = Split(refText, ":")',
        '        startRef = parts(0)',
        '        endRef = parts(1)',
        '        proposed = "=SUM(" & startRef & ":" & _',
        '            ws.Cells(lastDetailRow + insertedRows, sumRange.Column).Address(False, False) & ")"',
        '        changes.Add Array(cell.Row + insertedRows, cell.Column, proposed)',
        '        If changes.Count <= 3 Then examples = examples & vbCrLf & cell.Address(False, False) & "（插行前） " & formula & " → " & proposed',
        'NextFormula:',
        '    Next cell',
        '    If plainTotals > 0 Then Say plainTotals & " 个合计是逐格相加写的（比如 =D3+D4），" & _',
        '        "插行后不会自动包含新行，请手工核一下。", vbExclamation',
        '    If changes.Count = 0 Then Exit Function',
        '    If changes.Count > 3 Then examples = examples & vbCrLf & "其余 " & (changes.Count - 3) & " 个合计公式同样扩至新增行末尾。"',
        '    If QUIET Then ConfirmTotalExpansion = True: Exit Function',
        '    ConfirmTotalExpansion = (MsgBox(changes.Count & " 个合计公式将漏掉新增的 " & insertedRows & " 行。" & examples & vbCrLf & _',
        '        "是否扩展这些合计范围？选否将照常插行并填数，保留这些公式的原范围。", vbYesNo + vbDefaultButton2 + vbExclamation, "合计范围确认") = vbYes)',
        'End Function');
    }
    return lines.join('\r\n');
  }

  /**
   * 用户手动指定范围时用：只按「这一行是不是能填的空行」判断，
   * 不再看合计范围（范围是人定的，机器只负责别写坏）。
   */
  function rowsInRange(rows, detail, isFormula, from, to, opts) {
    opts = opts || {};
    var b = findBlankRows(rows, detail, isFormula, {
      startRow: from,
      endRow: Math.min(to + 1, rows.length),
      isOccupied: opts.isOccupied,
      isMergedSubordinate: opts.isMergedSubordinate,
      isHidden: opts.isHidden,
      // 不看合计范围：用户指定了就按他指定的
      inTotalRange: null
    });
    return b.filter(function (r) { return r >= from && r <= to; });
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
    isRowBoundary: isRowBoundary,
    rowsInRange: rowsInRange,
    listGroups: listGroups,
    planWrites: planWrites,
    parseOpening: parseOpening,
    buildVBA: buildVBA,
    buildCsv: buildCsv,
    colLetter: colLetter,
    cleanParty: cleanParty
  };
});
