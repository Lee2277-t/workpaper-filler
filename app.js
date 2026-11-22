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
    state.groups = Core.listGroups(state.rows, state.detail, isF, {
      isOccupied: function (r, c) { return cellHasValue(ws, r, c) || inArray(ranges, r, c); },
      isMergedSubordinate: function (r, c) { return isMergedSubordinate(merges, r, c); },
      isHidden: function (r) { var info = (ws["!rows"] || [])[r]; return !!(info && info.hidden); }
    });

    var gi = Number($('groupSel').value || 0);
    if (!state.groups[gi]) gi = 0;
    state.groupIndex = gi;
    var group = state.groups[gi];
    state.blanks = (group ? group.blanks : []).filter(function (r) {
      return !rowOccupied(ws, ranges, r, state.detail);
    });
    if (group && group.title && group.title !== '（未分组）') {
      state.problems.push('往「' + group.title + '」这一组填（第 ' + (group.startRow + 1) + '–' + group.endRow +
        ' 行），一组填完不会串到下一组。要填别的组请在上面切换。');
    }
    var opening = Core.parseOpening($('openBox') ? $('openBox').value : '');
    state.opening = opening.map;
    if (opening.bad.length) {
      state.problems.push('期初里有 ' + opening.bad.length + ' 行读不出（第 ' +
        opening.bad.map(function (b) { return b.line; }).join('、') +
        ' 行），应该是「客户名 + 金额」两列。');
    }
    var res = Core.planWrites(parties, state.blanks, state.detail, {
      openCents: Object.keys(opening.map).length ? opening.map : null,
      canWrite: function (r, c) { return !isF(r, c); },
      isMergedSubordinate: function (r, c) { return isMergedSubordinate(merges, r, c); }
    });
    state.plan = res.plan;
    state.missing = res.missing || 0;
    state.insertAfter = res.insertAfter;
    state.problems = state.problems.concat(res.problems);
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
      var open = 0, debit = 0, credit = 0, close = 0;
      p.cells.forEach(function (c) {
        if (c.col === state.detail.cols.open) open = c.value;
        if (c.col === state.detail.cols.debit) debit = c.value;
        if (c.col === state.detail.cols.credit) credit = c.value;
        if (c.col === state.detail.cols.close) close = c.value;
      });
      html += '<tr>' +
        '<td>' + (i + 1) + '</td>' +
        '<td>' + (p.row + 1) + '</td>' +
        '<td>' + esc(p.party) + '</td>' +
        '<td class="num">' + num(open) + '</td>' +
        '<td class="num">' + num(debit) + '</td>' +
        '<td class="num">' + num(credit) + '</td>' +
        '<td class="num">' + num(close) + '</td>' +
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
    if (!state.error && state.detail && state.tplWb) {
      lines.push('<div class="note">明细表「' + esc(state.sheetName) + '」：表头第 ' + (state.detail.headerRow + 1) +
        ' 行，可填空行 ' + state.blanks.length + ' 行（从第 ' + ((state.blanks[0] || 0) + 1) + ' 行起）。' +
        '填数时跳过公式格。</div>');
    }
    if (!state.tplWb) lines.push('<div class="note">先选底稿模板（带公式的那个 xlsx）。</div>');
    else if (!state.sheetName) lines.push('<div class="note bad">这个底稿里没找到明细表（表头要有「项目名称 / 本期借方 / 本期贷方」）。</div>');
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
    $('genXlsx').disabled = !ok || needInsert;
    $('genXlsx').title = needInsert ? '行数不够，需要插行，请用填充宏' : '';
  }

  /* ---------------- 导出 ---------------- */

  function exportVba() {
    var vba = Core.buildVBA(state.plan, state.sheetName, {
      missing: state.missing || 0,
      insertAfter: state.insertAfter
    });
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
      '要完全保真请用宏。';
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
      state.detail = hit ? hit.det : null;
      state.rows = hit ? hit.rows : null;
      recompute();
    });

    $('groupSel').addEventListener('change', function () { if (state.tplWb) recompute(); });
    $('subject').addEventListener('input', function () { if (state.tplWb) recompute(); });
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
