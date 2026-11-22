/* 底稿填充助手 · 界面层 */
(function () {
  'use strict';

  var Core = window.WpCore;
  var XLSX = window.XLSX;

  var state = {
    tplName: '', tplWb: null,
    dataName: '', dataWb: null,
    sheetName: '', detail: null, rows: null,
    parties: [], blanks: [], plan: [],
    problems: [], error: ''
  };

  function $(id) { return document.getElementById(id); }
  function esc(s) {
    return String(s === null || s === undefined ? '' : s).replace(/[&<>"]/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
    });
  }
  function money(c) { return Core.moneyText(c); }

  function sheetRows(wb, name) {
    var ws = wb.Sheets[name];
    if (!ws || !ws['!ref']) return { rows: [], firstRow: 0 };
    var range = XLSX.utils.decode_range(ws['!ref']);
    // 输出使用实际工作表坐标，表头搜索仍从表的有效范围开始。
    range.s.c = 0;
    var rows = XLSX.utils.sheet_to_json(ws, { header: 1, range: range, raw: true, defval: null, blankrows: true });
    return { rows: new Array(range.s.r).concat(rows), firstRow: range.s.r };
  }

  /**
   * 明细区的硬判据：某一行只要被某个金额列的 SUM/SUBTOTAL 覆盖，就属于明细区。
   * 账龄区、说明区不在任何明细合计范围内，自然被排除。
   * 支持 $F$4:$F$8、SUBTOTAL(9,F4:F8)、跨表引用前缀。
   */
  function totalRanges(ws) {
    var out = [];
    Object.keys(ws).forEach(function (k) {
      if (k.charAt(0) === '!') return;
      var cell = ws[k];
      if (!cell || !cell.f) return;
      var f = String(cell.f);
      var re = /(?:SUM|SUBTOTAL\s*\(\s*9\s*,)\s*\(?\s*(\$?[A-Z]{1,3}\$?\d+)\s*:\s*(\$?[A-Z]{1,3}\$?\d+)/gi;
      var m;
      while ((m = re.exec(f))) {
        var a = XLSX.utils.decode_cell(m[1].replace(/\$/g, ''));
        var b = XLSX.utils.decode_cell(m[2].replace(/\$/g, ''));
        if (a.c !== b.c) continue;
        out.push({ c: a.c, top: Math.min(a.r, b.r), bottom: Math.max(a.r, b.r) });
      }
    });
    return out;
  }

  function mergeRanges(ws) {
    return (ws["!merges"] || []).map(function (m) { return { s: m.s, e: m.e }; });
  }
  function isMergedAnchor(ranges, r, c) {
    return ranges.some(function (rg) { return rg.s.r === r && rg.s.c === c; });
  }
  /** 合并区里但不是我这一格（即从属格）→ 写进去 Excel 会丢 */
  function isMergedSubordinate(ranges, r, c) {
    return ranges.some(function (rg) {
      var inside = r >= rg.s.r && r <= rg.e.r && c >= rg.s.c && c <= rg.e.c;
      var anchor = rg.s.r === r && rg.s.c === c;
      return inside && !anchor;
    });
  }

  function arrayRanges(ws) {
    return Object.keys(ws).filter(function (addr) {
      return ws[addr].f && ws[addr].F;
    }).map(function (addr) { return XLSX.utils.decode_range(ws[addr].F); });
  }
  function inArray(ranges, r, c) {
    return ranges.some(function (rg) {
      return r >= rg.s.r && r <= rg.e.r && c >= rg.s.c && c <= rg.e.c;
    });
  }

  function formulaChecker(ws) {
    var arrays = Object.keys(ws).filter(function (addr) {
      return ws[addr].f && ws[addr].F;
    }).map(function (addr) {
      return XLSX.utils.decode_range(ws[addr].F);
    });
    return function (r, c) {
      var cell = ws[XLSX.utils.encode_cell({ r: r, c: c })];
      if (cell && (cell.f || cell.F)) return true;
      // Some members of an array formula have no stored cell of their own.
      return arrays.some(function (range) {
        return r >= range.s.r && r <= range.e.r && c >= range.s.c && c <= range.e.c;
      });
    };
  }
  function decodeCsv(buf) {
    var u8 = new Uint8Array(buf);
    if (u8.length > 3 && u8[0] === 0xEF && u8[1] === 0xBB && u8[2] === 0xBF) {
      return new TextDecoder('utf-8').decode(u8.subarray(3));
    }
    try { return new TextDecoder('utf-8', { fatal: true }).decode(u8); }
    catch (e) { try { return new TextDecoder('gbk').decode(u8); } catch (e2) { return new TextDecoder('utf-8').decode(u8); } }
  }
  function readWorkbook(file) {
    return new Promise(function (resolve) {
      var reader = new FileReader();
      var isCsv = /\.csv$/i.test(file.name || '');
      reader.onload = function () {
        try {
          var wb = isCsv
            ? XLSX.read(decodeCsv(reader.result), { type: 'string', cellFormula: true })
            : XLSX.read(new Uint8Array(reader.result), { type: 'array', cellFormula: true, cellStyles: true });
          resolve({ ok: true, wb: wb, fileName: file.name });
        } catch (e) { resolve({ ok: false, error: e.message }); }
      };
      reader.onerror = function () { resolve({ ok: false, error: '文件读不了' }); };
      reader.readAsArrayBuffer(file);
    });
  }

  /* ---------------- 底稿 ---------------- */

  /** 这一格已经有金额了（不是公式、也不是空） */
  function cellHasValue(ws, r, c) {
    var cell = ws[XLSX.utils.encode_cell({ r: r, c: c })];
    if (!cell || cell.f || cell.F) return false;
    return cell.v !== undefined && cell.v !== null && cell.v !== '';
  }
  function rowOccupied(ws, ranges, r, detail) {
    return [detail.cols.open, detail.cols.debit, detail.cols.credit, detail.cols.close].some(function (c) {
      return c >= 0 && (cellHasValue(ws, r, c) || inArray(ranges, r, c));
    });
  }
  function useRowHasValue_REMOVE(ws, r, detail) {
    return [detail.cols.open, detail.cols.debit, detail.cols.credit, detail.cols.close].some(function (c) {
      return c >= 0 && cellHasValue(ws, r, c);
    });
  }

  function listDetailSheets(wb, problems) {
    var out = [];
    wb.SheetNames.forEach(function (name) {
      var grid = sheetRows(wb, name);
      var det = Core.detectDetail(grid.rows, grid.firstRow);
      if (det.usable) out.push({ name: name, det: det, rows: grid.rows });
      if (det.error && problems) problems.push('「' + name + '」' + det.error);
    });
    return out;
  }

  /* ---------------- 数据源 ---------------- */

  function pickJournalSheet(wb) {
    var best = null;
    wb.SheetNames.forEach(function (name) {
      var grid = sheetRows(wb, name);
      var jh = Core.detectJournal(grid.rows, grid.firstRow);
      if (!jh) return;
      var parsed = Core.readJournal(grid.rows, jh);
      var withParty = parsed.records.filter(function (r) { return r.party; }).length;
      if (!best || withParty > best.withParty) {
        best = { name: name, header: jh, parsed: parsed, withParty: withParty };
      }
    });
    return best;
  }

  /* ---------------- 计算 ---------------- */

  /**
   * 期末方向：模板自己的公式最可信。
   * 预收款项 H7 = ROUND(E7-F7+G7,2)   → 期初 − 借 + 贷（负债，sign = -1）
   * 应收账款 K8 = ROUND(H8+I8-J8,2)   → 期初 + 借 − 贷（资产，sign = +1）
   * 认不出公式就看科目代码：2 负债 / 3 权益 → -1
   */
  function inferCloseSign(ws, detail, subject) {
    var col = detail.cols.close, dbt = detail.cols.debit, crd = detail.cols.credit, opn = detail.cols.open;
    if (ws && col >= 0 && dbt >= 0) {
      // 表头下面几行里，跳过合计/SUMIF 之类的汇总公式，找到第一个方向公式
      for (var k = 1; k <= 10; k++) {
        var cell = ws[XLSX.utils.encode_cell({ r: detail.headerRow + k, c: col })];
        if (!cell || !cell.f) continue;
        var f = String(cell.f).replace(/^=?\s*ROUND\s*\(/i, '').replace(/,\s*[-\d.]+\s*\)\s*$/i, '');
        if (/^SUM(IF|IFS)?\s*\(/i.test(f)) continue;          // 汇总公式不是方向公式
        // IF 包着的方向公式：=IF(B3="",0,ROUND(D3-E3+F3,2))
        // 条件里的引用（B3）不是金额，不能算进三列，否则整条被判掉、退回默认方向
        var body = f;
        var mRound = body.match(/ROUND\s*\(([^()]*)\)/i);
        if (mRound) body = mRound[1];
        else if (/^IF\s*\(/i.test(body)) body = body.replace(/^IF\s*\([^,]*,\s*[^,]*,\s*/i, '').replace(/\)\s*$/, '');
        var re = /([+-]?)\s*\$?([A-Z]{1,3})\$?(\d+)/g, m, cbt = 0, seen = [], ok = true;
        while ((m = re.exec(body))) {
          var ci = XLSX.utils.decode_col(m[2]);
          seen.push(ci);
          if (ci === dbt) cbt += m[1] === '-' ? -1 : 1;
        }
        // 方向公式只该由 期初/借方/贷方 三格组成，各出现一次；
        // =D4+D10、=D3+D4 这类跨行加总是合计，不是方向
        var want = [opn, dbt, crd].filter(function (x) { return x >= 0; }).sort().join(',');
        var got = seen.slice().sort().join(',');
        if (seen.length !== 3 || got !== want) ok = false;
        if (ok && cbt !== 0) return { sign: cbt > 0 ? 1 : -1, from: 'template' };
      }
    }
    // 认不出公式：按科目代码 1/5 借、2/4 贷；6 开头看名字；3 共同类不猜
    var s = String(subject || '').trim();
    var code = (s.match(/^\d+/) || [''])[0].charAt(0);
    if (code === '1' || code === '5') return { sign: 1, from: 'code' };
    if (code === '2' || code === '4') return { sign: -1, from: 'code' };
    if (code === '6') return { sign: /收入|收益/.test(s) ? -1 : 1, from: 'code' };
    if (code === '3') return { sign: 1, from: 'unknown' };     // 共同类，借贷都可能
    if (/预收|应付|应交|合同负债|递延收益|预计负债|短期借款|长期借款|实收资本|盈余公积/.test(s)) {
      return { sign: -1, from: 'code' };
    }
    return { sign: 1, from: 'default' };
  }

  function recompute() {
    state.error = '';
    state.problems = [];
    state.plan = [];
    state.parties = [];
    state.blanks = [];
    if (!state.tplWb || !state.detail) {
      renderAll();
      return;
    }
    var subject = String($('subject').value || '').trim();
    if (!subject) {
      state.error = '请填科目代码或名称（比如 1221 或 其他应收款）。工具只按这个科目汇总，不会全表乱填。';
      renderAll();
      return;
    }
    if (!state.dataWb) {
      state.error = '还没传数据源（科目余额表、序时账那个文件）。';
      renderAll();
      return;
    }
    var js = pickJournalSheet(state.dataWb);
    if (!js) {
      state.error = '数据源里找不到序时账（需要「日期 + 科目代码 + 借方 + 贷方」几列）。';
      renderAll();
      return;
    }
    state.journal = js;
    var issue = Core.findJournalAmountIssue(js.parsed.records, subject);
    if (issue) {
      state.error = '序时账「' + js.name + '」' + issue.amountError + '，请核对后重新上传。';
      renderAll();
      return;
    }
    var parties = Core.summarizeByParty(js.parsed.records, subject);
    if (!parties.length) {
      state.error = '科目「' + subject + '」在序时账里没有带客户档案的记录。' +
        '如果是这个科目本来就没有往来明细，那是正常的；否则请确认科目代码填对了。';
      renderAll();
      return;
    }
    state.parties = parties;
    var ws = state.tplWb.Sheets[state.sheetName];
    var isF = formulaChecker(ws);
    var ranges = arrayRanges(ws);
    var merges = mergeRanges(ws);
    var moneyCols = [state.detail.cols.open, state.detail.cols.debit,
      state.detail.cols.credit, state.detail.cols.close].filter(function (c) { return c >= 0; });
    var tRanges = totalRanges(ws).filter(function (x) {
      return moneyCols.indexOf(x.c) >= 0 && x.bottom > x.top;
    });
    var inTotalRange = tRanges.length ? function (r) {
      for (var i = 0; i < tRanges.length; i++) {
        if (r >= tRanges[i].top && r <= tRanges[i].bottom) return true;
      }
      return false;
    } : null;
    // 块标题：带合计/小计字样，且某个金额列的 SUM **正好从下一行开始**
    // （1、合并范围内小计 =SUM(H8:H11)）。用已解析出的范围判断，
    // 这样 $H$8、SUBTOTAL(9,…) 也能认。
    var blockHeaderCache = {};
    var isBlockHeader = function (r) {
      if (blockHeaderCache[r] !== undefined) return blockHeaderCache[r];
      blockHeaderCache[r] = false;
      var row = state.rows[r] || [];
      var hasLabel = false;
      for (var c1 = 0; c1 < row.length; c1++) {
        var v = row[c1];
        if (v === null || v === undefined) continue;
        var t = String(v).trim();
        if (t && /合计|小计|总计|共计|累计/.test(t.replace(/[\s（）()：:]/g, ''))) { hasLabel = true; break; }
      }
      if (!hasLabel) return false;
      var below = tRanges.some(function (x) { return x.top === r + 1; });
      blockHeaderCache[r] = below;
      return below;
    };
    state.groups = Core.listGroups(state.rows, state.detail, isF, {
      inTotalRange: inTotalRange,
      isBlockHeader: isBlockHeader,
      isOccupied: function (r, c) { return cellHasValue(ws, r, c) || inArray(ranges, r, c); },
      isMergedSubordinate: function (r, c) { return isMergedSubordinate(merges, r, c); },
      isHidden: function (r) { var info = (ws["!rows"] || [])[r]; return !!(info && info.hidden); }
    });

    var gi = Number($('groupSel').value || 0);
    if (!state.groups[gi] || !state.groups[gi].blanks.length) {
      // 默认选第一个真的有空行的组：账龄区那种 0 空行的组不该抢默认
      var firstWithBlanks = -1;
      state.groups.forEach(function (g, ix) {
        if (firstWithBlanks < 0 && g.blanks.length) firstWithBlanks = ix;
      });
      gi = firstWithBlanks >= 0 ? firstWithBlanks : 0;
    }
    state.groupIndex = gi;
    var group = state.groups[gi];
    state.blanks = (group ? group.blanks : []).filter(function (r) {
      return !rowOccupied(ws, ranges, r, state.detail);
    });
    // 只有用户**真的改过**行号，才走手填范围这条路。
    // 框里的数字是程序回填的建议值时，仍按自动判断走。
    var manualFrom = Number($('rowFrom').value || 0);
    var manualTo = Number($('rowTo').value || 0);
    var useManual = false;
    state.manualRange = null;
    if (state.rangeDirty) {
      // 校验不过就只提示、**不动用户填的数字**：改回去会让「没授权却被当授权」
      if (!(manualFrom >= 1) || !(manualTo >= 1)) {
        state.problems.push('填写范围还没填完整（起止行都要有数），这一轮先按自动判断走。');
      } else if (manualFrom !== Math.floor(manualFrom) || manualTo !== Math.floor(manualTo)) {
        state.problems.push('行号要填整数（现在是 ' + manualFrom + '–' + manualTo + '），这一轮先按自动判断走。');
      } else if (manualTo < manualFrom) {
        state.problems.push('范围填反了（第 ' + manualFrom + ' 行 > 第 ' + manualTo + ' 行），这一轮先按自动判断走。');
      } else if (manualFrom <= state.detail.headerRow + 1) {
        state.problems.push('第 ' + manualFrom + ' 行还在表头里（表头在第 ' + (state.detail.headerRow + 1) +
          ' 行），这一轮先按自动判断走。');
      } else {
        useManual = true;
      }
    }
    if (useManual) {
      state.manualRange = { from: manualFrom - 1, to: manualTo - 1 };
      var beforeAuto = state.blanks;
      state.blanks = Core.rowsInRange(state.rows, state.detail, isF, manualFrom - 1, manualTo - 1, {
        isOccupied: function (r, c) { return cellHasValue(ws, r, c) || inArray(ranges, r, c); },
        isMergedSubordinate: function (r, c) { return isMergedSubordinate(merges, r, c); },
        isHidden: function (r) { var info = (ws['!rows'] || [])[r]; return !!(info && info.hidden); }
      });
      if (!state.blanks.length) {
        var why = '';
        var hitRow = -1;
        for (var q = manualFrom - 1; q < manualTo && q < state.rows.length; q++) {
          var qr = state.rows[q] || [];
          if (Core.isRowBoundary(qr, state.detail, String(qr[state.detail.cols.name] || '').trim())) { hitRow = q; break; }
        }
        if (hitRow >= 0) {
          why = '第 ' + (hitRow + 1) + ' 行是合计/小计行，扫到那就停了。';
        }
        state.problems.push('你指定的范围（第 ' + manualFrom + '–' + manualTo + ' 行）里没有能填的空行：' +
          (why || '要么已经填了客户，要么是公式格/合并格/隐藏行。') + '换个范围试试。');
      } else if (state.blanks[0] > manualFrom - 1) {
        state.problems.push('第 ' + manualFrom + '–' + state.blanks[0] + ' 行不能填（合计/小计/已填/公式/合并），' +
          '实际从第 ' + (state.blanks[0] + 1) + ' 行开始。');
      }
      // 手填的范围如果跑到合计公式之外，钱写进去合计不算——要说清楚，但照填
      state.outsideNote = '';
      if (inTotalRange && state.blanks.length) {
        var outside = state.blanks.filter(function (r) { return !inTotalRange(r); });
        if (outside.length) {
          var oFirst = outside[0] + 1, oLast = outside[outside.length - 1] + 1;
          var oText = '其中第 ' + oFirst + (oLast > oFirst ? '–' + oLast : '') + ' 行（共 ' +
            outside.length + ' 行）不在合计公式范围内，合计不会算上它们。';
          state.problems.push('注意：你指定的范围里，第 ' + oFirst + (oLast > oFirst ? '–' + oLast : '') +
            ' 行（共 ' + outside.length + ' 行）不在任何合计公式范围内。填进去合计不会算上这几行，请自己确认这样对不对。');
          // 把这句话带进宏和导出，别只留在浏览器里
          state.outsideNote = oText;
        }
      }
      // 手填时提一句「自动判断本来会给哪几行」，方便对照
      if (beforeAuto && beforeAuto.length) {
        state.problems.push('（自动判断会给第 ' + beforeAuto.map(function (r) { return r + 1; }).join('、') + ' 行）');
      }
    } else if (!state.rangeDirty && $('rowFrom') && $('rowTo')) {
      // 自动判断：把结果回填到框里当建议。用户改过就不覆盖，免得打字被吞
      $('rowFrom').value = state.blanks.length ? state.blanks[0] + 1 : '';
      $('rowTo').value = state.blanks.length ? state.blanks[state.blanks.length - 1] + 1 : '';
    }
    if (!useManual && group && group.title && group.title !== '（未分组）') {
      state.problems.push('往「' + group.title + '」这一组填（第 ' + (group.startRow + 1) + '–' + group.endRow +
        ' 行），一组填完不会串到下一组。要填别的组请在上面切换。');
    }
    if (group && group.uncovered && group.uncovered.length) {
      var uu = group.uncovered;
      state.problems.push('第 ' + (uu[0] + 1) + '–' + (uu[uu.length - 1] + 1) + ' 行（共 ' + uu.length +
        ' 行）是空的，但不在这个科目的合计公式范围内——填进去合计也不会算上。' +
        '要填这几行，请先在底稿里把合计公式的范围扩到它们。');
    }
    var opening = Core.parseOpening($('openBox') ? $('openBox').value : '');
    state.opening = opening.map;
    if (opening.bad.length) {
      state.problems.push('期初里有 ' + opening.bad.length + ' 行读不出（第 ' +
        opening.bad.map(function (b) { return b.line; }).join('、') +
        ' 行），应该是「客户名 + 金额」两列。');
    }
    var wsNow = state.tplWb.Sheets[state.sheetName];
    var signInfo = inferCloseSign(wsNow, state.detail, $('subject').value);
    state.closeSign = signInfo.sign;
    state.closeSignFrom = signInfo.from;
    var res = Core.planWrites(parties, state.blanks, state.detail, {
      noInsert: !!state.manualRange,
      closeSign: signInfo.sign,
      openCents: Object.keys(opening.map).length ? opening.map : null,
      canWrite: function (r, c) { return !isF(r, c); },
      isMergedSubordinate: function (r, c) { return isMergedSubordinate(merges, r, c); }
    });
    state.plan = res.plan;
    state.missing = res.missing || 0;
    state.insertAfter = res.insertAfter;
    state.problems = state.problems.concat(res.problems);
    // 两行一户的模板（明细区里有上下合并）：工具按一户一行填，填完得自己看第二行
    var tallMerge = merges.filter(function (m) {
      return m.e.r > m.s.r && m.e.r >= (state.blanks[0] === undefined ? 0 : state.blanks[0]);
    })[0];
    if (tallMerge && res.plan.length) {
      state.problems.push('明细区里有上下合并的单元格（第 ' + (tallMerge.s.r + 1) + '–' + (tallMerge.e.r + 1) +
        ' 行那种，通常是「一户占两行」的版式）。工具按一户一行填，填完请核对第二行是不是空的。');
    }
    renderAll();
  }

  /* ---------------- 渲染 ---------------- */

  function renderGroups() {
    var sel = $('groupSel');
    if (!sel) return;
    var groups = state.groups || [];
    sel.innerHTML = '';
    groups.forEach(function (g, i) {
      var opt = document.createElement('option');
      opt.value = String(i);
      opt.textContent = g.title + '（空 ' + g.blanks.length + ' 行）';
      sel.appendChild(opt);
    });
    sel.value = String(state.groupIndex || 0);
    sel.disabled = groups.length <= 1;
  }

  function renderStats() {
    if (!state.parties.length) { $('statBox').innerHTML = ''; return; }
    var debit = state.parties.reduce(function (s, p) { return s + p.debit; }, 0);
    var credit = state.parties.reduce(function (s, p) { return s + p.credit; }, 0);
    var html = '';
    html += '<div class="stat"><span class="k">科目</span><span class="v">' + esc($('subject').value) + '</span></div>';
    html += '<div class="stat"><span class="k">客户户数</span><span class="v">' + state.parties.length + ' 户</span></div>';
    html += '<div class="stat"><span class="k">借方合计</span><span class="v">' + money(debit) + '</span></div>';
    html += '<div class="stat"><span class="k">贷方合计</span><span class="v">' + money(credit) + '</span></div>';
    var dir = state.closeSign === -1 ? '期初 − 借 + 贷' : '期初 + 借 − 贷';
    var from = state.closeSignFrom === 'template' ? '照模板公式'
      : (state.closeSignFrom === 'code' ? '按科目性质'
        : (state.closeSignFrom === 'unknown' ? '共同类，请自己核' : '默认'));
    html += '<div class="stat"><span class="k">期末方向</span><span class="v">' + dir +
      '（' + from + '）</span></div>';
    if (state.closeSignFrom === 'unknown') {
      html += '<div class="note warn" style="margin-top:8px">科目 ' + esc($('subject').value) +
        ' 是共同类，借贷方向不固定，工具认不出方向，期末按「期初 + 借 − 贷」算。' +
        '期末列是公式时不受影响；不是公式的话请自己核一下这个数。</div>';
    }
    html += '<div class="stat"><span class="k">明细表空行</span><span class="v">' + state.blanks.length + ' 行</span></div>';
    $('statBox').innerHTML = html;
  }

  function renderPlan() {
    var box = $('planBody');
    if (!state.plan.length) {
      box.innerHTML = '';
      $('planPanel').classList.add('hidden');
      return;
    }
    var html = '';
    state.plan.forEach(function (p, i) {
      var open = 0, debit = 0, credit = 0, close = null;
      p.cells.forEach(function (c) {
        if (c.col === state.detail.cols.open) open = c.value;
        if (c.col === state.detail.cols.debit) debit = c.value;
        if (c.col === state.detail.cols.credit) credit = c.value;
        if (c.col === state.detail.cols.close) close = c.value;
      });
      // 期末列是公式时我们不写它，也就不该在预览里替它算一个数——
      // 往来科目有借有贷，算错了反而误导
      var closeCell = close === null
        ? '<span class="byFormula">公式算</span>'
        : num(close);
      html += '<tr>' +
        '<td>' + (i + 1) + '</td>' +
        '<td>' + (p.row + 1) + '</td>' +
        '<td>' + esc(p.party) + '</td>' +
        '<td class="num">' + num(open) + '</td>' +
        '<td class="num">' + num(debit) + '</td>' +
        '<td class="num">' + num(credit) + '</td>' +
        '<td class="num">' + closeCell + '</td>' +
        '</tr>';
    });
    box.innerHTML = html;
    $('planPanel').classList.remove('hidden');
  }
  function num(v) {
    if (typeof v !== 'number') return '';
    return v.toLocaleString('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function renderAll() {
    var info = $('infoBox');
    var lines = [];
    if (state.error) lines.push('<div class="note bad">' + esc(state.error) + '</div>');
    state.problems.forEach(function (p) { lines.push('<div class="note warn">' + esc(p) + '</div>'); });
    renderMacroPreview();
    if (!state.error && state.detail && state.tplWb) {
      var where = state.blanks.length
        ? '可填空行 ' + state.blanks.length + ' 行（第 ' + (state.blanks[0] + 1) + ' 行起）'
        : '这个范围里没有能填的空行';
      lines.push('<div class="note">明细表「' + esc(state.sheetName) + '」：表头第 ' + (state.detail.headerRow + 1) +
        ' 行，' + where + '。填数时跳过公式格。</div>');
    }
    if (!state.tplWb) lines.push('<div class="note">先选底稿模板（带公式的那个 xlsx）。</div>');
    else if (!state.sheetName) lines.push('<div class="note bad">这个底稿里没找到明细表（表头要在前 60 行里，' +
      '并且有「项目名称 / 本期借方 / 本期贷方」这几列）。</div>');
    info.innerHTML = lines.join('');
    renderGroups();
    renderStats();
    renderPlan();
    refreshButtons();
  }

  function refreshButtons() {
    var ok = state.plan.length > 0;
    var needInsert = (state.missing || 0) > 0;
    $('genVba').disabled = !ok;
    $('genCsv').disabled = !ok;
    // 直接写 xlsx 没法插行：行数不够时只给宏和 CSV
    // 宏底稿（.xlsm）直接用 SheetJS 重写会把 VBA 弄丢，只能走宏
    $('genXlsx').disabled = !ok || needInsert || !!state.isMacroTemplate;
    $('genXlsx').title = needInsert
      ? '行数不够，需要插行，请用填充宏'
      : (state.isMacroTemplate ? '这是宏底稿，直接生成会丢掉 VBA，请用填充宏' : '');
  }

  /* ---------------- 导出 ---------------- */

  function buildMacroText() {
    if (!state.plan.length || !state.sheetName) return '';
    return Core.buildVBA(state.plan, state.sheetName, {
      missing: state.missing || 0,
      insertAfter: state.insertAfter,
      outsideNote: state.outsideNote || ''
    });
  }

  /**
   * 把宏的开头几行显示出来。粘贴进 VBE 的如果和这里不一样，
   * 说明模块里混了别的东西——这是用户唯一能自查的线索。
   */
  function renderMacroPreview() {
    var box = $('macroDetails');
    if (!box) return;
    var text = buildMacroText();
    if (!text) { box.classList.add('hidden'); return; }
    var lines = text.split(/\r?\n/).slice(0, 20);
    $('macroPreview').textContent = lines.join('\n') + '\n' +
      '…（共 ' + text.split(/\r?\n/).length + ' 行）';
    box.classList.remove('hidden');
  }

  function exportVba() {
    var vba = buildMacroText();
    download(new Blob([vba], { type: 'text/plain;charset=utf-8' }),
      '填充宏_' + state.sheetName + '.bas');
    $('genInfo').textContent = (state.missing || 0) > 0
      ? '已生成宏：会照最后一行复制插入 ' + state.missing + ' 行并填数。' +
        '发现可扩展的合计范围时，会先询问是否把新行计入。Alt+F11 → 插入模块 → 粘贴 → F5。'
      : '已生成宏：在 Excel 里 Alt+F11 → 插入模块 → 粘贴 → F5。公式不会被改动。';
  }

  function exportCsv() {
    var csv = '\ufeff' + Core.buildCsv(state.plan, state.detail);
    download(new Blob([csv], { type: 'text/csv;charset=utf-8' }),
      '明细行_' + state.sheetName + '.csv');
    $('genInfo').textContent = '已生成明细行 CSV，可以直接粘贴进底稿。';
  }

  function exportXlsx() {
    var sheets = Object.assign({}, state.tplWb.Sheets);
    var wb = Object.assign({}, state.tplWb, { Sheets: sheets });
    // 每次从原模板生成，避免上一次导出留下的客户进入新文件。
    var ws = sheets[state.sheetName] = Object.assign({}, sheets[state.sheetName]);
    var isFormula = formulaChecker(ws);
    // 只写计划里的格；公式格已经在 planWrites 阶段排除
    state.plan.forEach(function (item) {
      item.cells.forEach(function (c) {
        var addr = XLSX.utils.encode_cell({ r: item.row, c: c.col });
        if (isFormula(item.row, c.col)) return;        // 双保险：公式格绝不覆盖
        ws[addr] = c.text
          ? { t: 's', v: String(c.value) }
          : { t: 'n', v: Number(c.value) };
      });
    });
    XLSX.writeFile(wb, '已填充_' + (state.tplName || '底稿'));
    $('genInfo').textContent = '已生成填好的底稿。注意：这个方式可能丢掉部分格式，' +
      '要完全保真请用宏。' + (state.outsideNote ? ' ' + state.outsideNote : '');
  }

  function download(blob, name) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(function () { document.body.removeChild(a); URL.revokeObjectURL(url); }, 0);
  }

  /* ---------------- 事件 ---------------- */

  function bind() {
    $('tplFile').addEventListener('change', function () {
      var f = this.files && this.files[0];
      if (!f) return;
      state.error = '';
      $('tplInfo').textContent = '正在读 ' + f.name + ' …';
      readWorkbook(f).then(function (res) {
        if (!res.ok) { state.error = '底稿读不了：' + res.error; renderAll(); return; }
        state.tplWb = res.wb;
        state.tplName = res.fileName;
        // 换了一份底稿 → 上一份的行号一律作废，否则会套到新表上
        state.rangeDirty = false;
        if ($('rowFrom')) $('rowFrom').value = '';
        if ($('rowTo')) $('rowTo').value = '';
        // 宏底稿重写会丢 VBA，只能走填充宏
        state.isMacroTemplate = /\.xlsm$/i.test(f.name || res.fileName || '');
        var problems = [];
        var list = listDetailSheets(res.wb, problems);
        $('tplInfo').textContent = f.name + '（' + res.wb.SheetNames.length + ' 张表，认出 ' + list.length + ' 张明细表）' +
          (problems.length ? '。' + problems.join('；') : '');
        var sel = $('sheetSel');
        sel.innerHTML = '';
        list.forEach(function (s) {
          var opt = document.createElement('option');
          opt.value = s.name;
          opt.textContent = s.name;
          sel.appendChild(opt);
        });
        if (list.length) {
          sel.value = list[0].name;
          state.sheetName = list[0].name;
          state.detail = list[0].det;
          state.rows = list[0].rows;
        } else {
          state.sheetName = ''; state.detail = null; state.rows = null;
        }
        recompute();
      });
    });

    $('dataFile').addEventListener('change', function () {
      var f = this.files && this.files[0];
      if (!f) return;
      $('dataInfo').textContent = '正在读 ' + f.name + ' …';
      readWorkbook(f).then(function (res) {
        if (!res.ok) { state.error = '数据源读不了：' + res.error; renderAll(); return; }
        state.dataWb = res.wb;
        state.dataName = res.fileName;
        var js = pickJournalSheet(res.wb);
        $('dataInfo').textContent = f.name + (js
          ? '（序时账：' + js.name + '，' + js.parsed.records.length + ' 行，其中 ' + js.withParty + ' 行带客户档案）'
          : '（没找到序时账）');
        recompute();
      });
    });

    $('sheetSel').addEventListener('change', function () {
      var list = listDetailSheets(state.tplWb);
      var hit = list.filter(function (s) { return s.name === $('sheetSel').value; })[0];
      state.sheetName = hit ? hit.name : '';
      state.rangeDirty = false;                 // 换表 → 人工范围作废
      state.detail = hit ? hit.det : null;
      state.rows = hit ? hit.rows : null;
      recompute();
    });

    $('groupSel').addEventListener('change', function () {
      state.rangeDirty = false;                     // 换组 → 重新用自动判断
      if (state.tplWb) recompute();
    });
    // 手改行号 → 严格按用户给的范围填，机器不再自己划
    ['rowFrom', 'rowTo'].forEach(function (id) {
      if (!$(id)) return;
      $(id).addEventListener('input', function () {
        state.rangeDirty = true;
        if (state.tplWb) recompute();
      });
    });
    if ($('useAuto')) {
      $('useAuto').addEventListener('click', function () {
        state.rangeDirty = false;
        if ($('rowFrom')) $('rowFrom').value = '';
        if ($('rowTo')) $('rowTo').value = '';
        if (state.tplWb) recompute();
      });
    }
    $('subject').addEventListener('input', function () {
      state.rangeDirty = false;                 // 换科目 → 人工范围作废
      if (state.tplWb) recompute();
    });
    if ($('openBox')) $('openBox').addEventListener('input', function () { if (state.tplWb) recompute(); });
    $('genVba').addEventListener('click', exportVba);
    $('genCsv').addEventListener('click', exportCsv);
    $('genXlsx').addEventListener('click', exportXlsx);
  }

  function boot() {
    bind();
    $('genVba').disabled = true;
    $('genXlsx').disabled = true;
    $('genCsv').disabled = true;
    renderAll();
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();

  window.WpApp = {
    recompute: recompute, exportVba: exportVba, exportCsv: exportCsv, exportXlsx: exportXlsx,
    state: state
  };
})();
