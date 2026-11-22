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
    if (!ws) return [];
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
  }
  function isFormula(wb, name, r, c) {
    var ws = wb.Sheets[name];
    if (!ws) return false;
    var cell = ws[XLSX.utils.encode_cell({ r: r, c: c })];
    return !!(cell && cell.f);
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

  function listDetailSheets(wb) {
    var out = [];
    wb.SheetNames.forEach(function (name) {
      var rows = sheetRows(wb, name);
      var det = Core.detectDetail(rows);
      if (det.usable) out.push({ name: name, det: det, rows: rows });
    });
    return out;
  }

  /* ---------------- 数据源 ---------------- */

  function pickJournalSheet(wb) {
    var best = null;
    wb.SheetNames.forEach(function (name) {
      var rows = sheetRows(wb, name);
      var jh = Core.detectJournal(rows);
      if (!jh) return;
      var parsed = Core.readJournal(rows, jh);
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
    var parties = Core.summarizeByParty(js.parsed.records, subject);
    if (!parties.length) {
      state.error = '科目「' + subject + '」在序时账里没有带客户档案的记录。' +
        '如果是这个科目本来就没有往来明细，那是正常的；否则请确认科目代码填对了。';
      state.parties = [];
      renderAll();
      return;
    }
    state.parties = parties;
    state.blanks = Core.findBlankRows(state.rows, state.detail, function (r, c) {
      return isFormula(state.tplWb, state.sheetName, r, c);
    });
    var res = Core.planWrites(parties, state.blanks, state.detail, {});
    state.plan = res.plan;
    state.problems = res.problems;
    renderAll();
  }

  /* ---------------- 渲染 ---------------- */

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
    if (!state.plan.length) { box.innerHTML = ''; return; }
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
        '公式格一律不碰。</div>');
    }
    if (!state.tplWb) lines.push('<div class="note">先选底稿模板（带公式的那个 xlsx）。</div>');
    else if (!state.sheetName) lines.push('<div class="note bad">这个底稿里没找到明细表（表头要有「项目名称 / 本期借方 / 本期贷方」）。</div>');
    info.innerHTML = lines.join('');
    renderStats();
    renderPlan();
    refreshButtons();
  }

  function refreshButtons() {
    var ok = state.plan.length > 0;
    $('genVba').disabled = !ok;
    $('genXlsx').disabled = !ok;
    $('genCsv').disabled = !ok;
  }

  /* ---------------- 导出 ---------------- */

  function exportVba() {
    var vba = Core.buildVBA(state.plan, state.sheetName, {});
    download(new Blob([vba], { type: 'text/plain;charset=utf-8' }),
      '填充宏_' + state.sheetName + '.bas');
    $('genInfo').textContent = '已生成宏：在 Excel 里 Alt+F11 → 插入模块 → 粘贴 → F5。公式不会被改动。';
  }

  function exportCsv() {
    var csv = '\ufeff' + Core.buildCsv(state.plan, state.detail);
    download(new Blob([csv], { type: 'text/csv;charset=utf-8' }),
      '明细行_' + state.sheetName + '.csv');
    $('genInfo').textContent = '已生成明细行 CSV，可以直接粘贴进底稿。';
  }

  function exportXlsx() {
    var wb = state.tplWb;
    var ws = wb.Sheets[state.sheetName];
    // 只写计划里的格；公式格已经在 planWrites 阶段排除
    state.plan.forEach(function (item) {
      item.cells.forEach(function (c) {
        var addr = XLSX.utils.encode_cell({ r: item.row, c: c.col });
        var old = ws[addr];
        if (old && old.f) return;                      // 双保险：公式格绝不覆盖
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
        var list = listDetailSheets(res.wb);
        $('tplInfo').textContent = f.name + '（' + res.wb.SheetNames.length + ' 张表，认出 ' + list.length + ' 张明细表）';
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

    $('subject').addEventListener('input', function () { if (state.tplWb) recompute(); });
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
