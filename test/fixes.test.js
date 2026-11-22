const { createApp, workbook, detailRows, journal, runSuite, XLSX } = require('./app-fixture');

const data = journal([['2026-01-05', '1221', '其他应收款', '合成甲', 100, 5]]);

async function prepareDetail(template, source = data) {
  const app = createApp();
  await app.loadTemplate(template);
  await app.loadData(source);
  app.setSubject('1221');
  return app;
}

async function exportDetail(app) {
  app.fire('genXlsx', 'click');
  const bytes = (await app.readDownload('xlsx')).bytes;
  return XLSX.read(bytes, { type: 'buffer', cellStyles: true }).Sheets.明细;
}

async function exportMacro(app) {
  app.fire('genVba', 'click');
  return (await app.readDownload('bas')).bytes.toString('utf8');
}

runSuite('缺陷回归', {
  async '1 合计变体截断明细，账龄分析的空行不纳入'(check) {
    for (const total of ['本项目小计', '合计：', '合  计', '小计：']) {
      const rows = detailRows(2);
      rows[4][1] = total;
      rows.push([], ['', '账龄分析'], ['', '账龄区间甲'], [], ['', '账龄区间乙'], ['', '合计']);
      const template = workbook({ 明细: rows });
      for (const col of ['C', 'D', 'E', 'F']) {
        template.Sheets.明细[col + '5'] = { t: 'n', f: `SUM(${col}3:${col}4)`, v: 0 };
      }
      const app = await prepareDetail(template, journal([
        ['2026-01-05', '1221', '其他应收款', '合成甲', 100, 0],
        ['2026-01-06', '1221', '其他应收款', '合成乙', 200, 0]
      ]));
      check.match(app.el('infoBox').innerHTML, /可填空行 2 行（从第 3 行起）/, total);
      const out = await exportDetail(app);
      check.deepEqual([out.B3?.v, out.B4?.v, out.B6, out.B9], ['合成甲', '合成乙', undefined, undefined]);
      check.equal(out.F5?.f, 'SUM(F3:F4)');
    }
  },
  async '2 合并标题行不写，仍逐格保护合并从属格'(check) {
    const rows = detailRows(4);
    rows[3] = ['一、合并范围内关联方'];
    rows[6][1] = '本项目小计';
    const template = workbook({ 明细: rows });
    template.Sheets.明细['!merges'] = ['A4:F4', 'C5:D5'].map(XLSX.utils.decode_range);
    const app = await prepareDetail(template, journal([
      ['2026-01-05', '1221', '其他应收款', '合成甲', 100, 0],
      ['2026-01-06', '1221', '其他应收款', '合成乙', 200, 0],
      ['2026-01-07', '1221', '其他应收款', '合成丙', 300, 0]
    ]));
    const out = await exportDetail(app);
    check.deepEqual([out.B3?.v, out.B4, out.B5?.v, out.B6?.v], ['合成甲', undefined, '合成乙', '合成丙']);
    check.equal(out.A4?.v, '一、合并范围内关联方');
    check.deepEqual(out['!merges'].map(XLSX.utils.encode_range), ['A4:F4', 'C5:D5']);
    check.equal(out.D5, undefined);
    const vba = await exportMacro(app);
    check.doesNotMatch(vba, /ws\.Range\("(?:[B-F]4|D5)"\)\.Value =/);
    check.match(vba, /If ws\.Range\("B5"\)\.MergeCells .*MergeArea\.Cells\(1\)\.Address Then GoTo MergeFound/);
  },
  async '3 默认选明细分组，分组标题以前的空行排在最后'(check) {
    const rows = detailRows(8);
    rows[3] = ['', '一、合并范围内关联方'];
    rows[6] = ['', '小计'];
    rows[7] = ['', '二、合并范围外关联方'];
    rows[9] = ['', '小计'];
    const app = await prepareDetail(workbook({ 明细: rows }));
    const selector = app.el('groupSel');
    check.match(selector.children[Number(selector.value)].textContent, /^一、合并范围内关联方/);
    check.match(selector.children[1].textContent, /^二、合并范围外关联方/);
    check.match(selector.children.at(-1).textContent, /^（分组标题以上）/);
    const out = await exportDetail(app);
    check.equal(out.B3, undefined);
    check.equal(out.B5?.v, '合成甲');
  },
  async '4 隐藏的空行原样保留，客户写到下一可见行'(check) {
    const template = workbook({ 明细: detailRows(2) });
    template.Sheets.明细['!rows'] = [];
    template.Sheets.明细['!rows'][2] = { hidden: true };
    const app = await prepareDetail(template);
    const out = await exportDetail(app);
    check.equal(out.B3, undefined);
    check.equal(out.B4?.v, '合成甲');
    check.equal(out['!rows'][2].hidden, true);
  },
  async '5 客户名称的换行生成单行 VBA 赋值'(check) {
    const app = await prepareDetail(workbook({ 明细: detailRows(1) }), journal([
      ['2026-01-05', '1221', '其他应收款', '北京\n华远公司', 100, 0]
    ]));
    check.match(await exportMacro(app), /^    ws\.Range\("B3"\)\.Value = "北京" & Chr\(10\) & "华远公司"$/m);
    check.equal((await exportDetail(app)).B3?.v, '北京\n华远公司');
  },
  async '6 科目 1221 汇总 122101，但排除 12210'(check) {
    const app = await prepareDetail(workbook({ 明细: detailRows(1) }), journal([
      ['2026-01-05', '1221', '其他应收款', '合成甲', 1000, 0],
      ['2026-01-06', '12210', '其他应收款', '合成甲', 500, 0],
      ['2026-01-07', '122101', '其他应收款', '合成甲', 200, 0]
    ]));
    check.equal((await exportDetail(app)).D3?.v, 1200);
    check.match(await exportMacro(app), /^    ws\.Range\("D3"\)\.Value = 1200\.00$/m);
  },
  async '7 插在合计范围内时 Excel 自己扩，问都不用问'(check) {
    const rows = detailRows(2);
    rows[2] = ['', '已有客户', 0, 50, 0, 50];
    const template = workbook({ 明细: rows });
    template.Sheets.明细.F5 = { t: 'n', f: 'SUM(F3:F4)', v: 50 };
    const app = await prepareDetail(template, journal([100, 200, 300, 400].map((amount, i) =>
      ['2026-01-05', '1221', '其他应收款', '合成客户' + (i + 1), amount, 0]
    )));
    const vba = await exportMacro(app);
    // 插在最后一个空行的位置上，落在 SUM(F3:F4) 之内 → Excel 自动撑开
    check.match(vba, /ws\.Rows\("4:6"\)\.Insert Shift:=xlDown/);
    check.match(vba, /ConfirmTotalExpansion\(ws, 4, 3, totalChanges\)/);
    // 范围内的情况不进提示分支
    check.match(vba, /If sumRange\.Row \+ sumRange\.Rows\.Count - 1 >= lastDetailRow Then GoTo NextFormula/);
    for (const [row, amount] of [[4, 100], [5, 200], [6, 300], [7, 400]]) {
      check.match(vba, new RegExp(`^    ws\\.Range\\("F${row}"\\)\\.Value = ${amount}\\.00$`, 'm'));
    }
    check.match(vba, /未扩展.*漏计新增行，需要核对/);
    check.match(vba, /已扩展.*合计已包含新增行/);
  },
  async '7b 合计范围止于插入点之前时才提示，默认不扩'(check) {
    const rows = detailRows(2);
    rows[2] = ['', '已有客户', 0, 50, 0, 50];
    const template = workbook({ 明细: rows });
    // 合计只到 F3，插在 F4 之下 → 落在范围之外，必须问
    template.Sheets.明细.F5 = { t: 'n', f: 'SUM(F3:F3)', v: 50 };
    const app = await prepareDetail(template, journal([100, 200, 300, 400].map((amount, i) =>
      ['2026-01-05', '1221', '其他应收款', '合成客户' + (i + 1), amount, 0]
    )));
    const vba = await exportMacro(app);
    check.match(vba, /vbYesNo \+ vbDefaultButton2/);
  },
  async '起始范围在 C3 的模板仍使用实际单元格坐标'(check) {
    const template = workbook({ 明细: [[], [], [null, null, '合成底稿'],
      [null, null, '序号', '项目名称', '期初金额', '本期借方', '本期贷方', '期末金额'],
      [], [null, null, null, '合计']] });
    template.Sheets.明细.H6 = { t: 'n', f: 'H5', v: 0 };
    template.Sheets.明细['!ref'] = 'C3:H6';
    const bytes = XLSX.write(template, { type: 'buffer', bookType: 'xlsx' });
    check.equal(XLSX.read(bytes, { type: 'buffer' }).Sheets.明细['!ref'], 'C3:H6', '夹具从 C3 起');
    const app = createApp();
    await app.loadTemplate(template);
    await app.loadData(data);
    app.setSubject('1221');
    check.match(app.el('infoBox').innerHTML, /表头第 4 行.*从第 5 行起/);
    app.fire('genVba', 'click');
    check.match((await app.readDownload('bas')).bytes.toString('utf8'), /ws\.Range\("D5"\)\.Value = "合成甲"/);
    app.fire('genXlsx', 'click');
    const out = XLSX.read((await app.readDownload('xlsx')).bytes, { type: 'buffer' }).Sheets.明细;
    check.equal(out.D5?.v, '合成甲');
    check.equal(out.F5?.v, 100);
    check.equal(out.H5?.v, 95);
    check.equal(out.H6?.f, 'H5');
    check.equal(out.B3, undefined);
  },
  async '名称列的空字符串公式和数组公式从属格都保留'(check) {
    const template = workbook({ 明细: detailRows(5) });
    const sheet = template.Sheets.明细;
    sheet.B3 = { t: 's', f: 'IF(TRUE,"","已有客户")', v: '' };
    XLSX.utils.sheet_set_array_formula(sheet, 'D4:D6', 'ROW(D4:D6)');
    sheet.D4.v = 4;
    sheet.D5.v = 5;
    delete sheet.D6; // 数组范围引用仍包括 D6，文件里没有单独的 D6 节点
    const app = createApp();
    await app.loadTemplate(template);
    await app.loadData(data);
    app.setSubject('1221');
    app.fire('genXlsx', 'click');
    const out = XLSX.read((await app.readDownload('xlsx')).bytes, { type: 'buffer' }).Sheets.明细;
    check.equal(out.B3?.f, 'IF(TRUE,"","已有客户")');
    check.equal(out.B3?.v, '');
    check.equal(out.D4?.f, 'ROW(D4:D6)');
    check.equal(out.D4?.F, 'D4:D6');
    check.equal(out.D5?.F, 'D4:D6');
    check.equal(out.D5?.v, 5);
    check.equal(out.B5, undefined);
    check.equal(out.B6, undefined);
    check.equal(out.D6, undefined);
    check.equal(out.B7?.v, '合成甲');
    check.equal(out.D7?.v, 100);
  },
  async '直接导出后改科目，下一份只填新科目'(check) {
    const app = createApp();
    await app.loadTemplate(workbook({ 明细: detailRows(2) }));
    await app.loadData(journal([
      ['2026-01-05', '1221', '其他应收款', '合成甲', 100, 5],
      ['2026-01-06', '1221', '其他应收款', '合成乙', 200, 4],
      ['2026-01-07', '2203', '预收款项', '合成丙', 50, 1]
    ]));
    app.setSubject('1221');
    app.fire('genXlsx', 'click');
    const first = XLSX.read((await app.readDownload('xlsx')).bytes, { type: 'buffer' }).Sheets.明细;
    check.equal(first.B3?.v, '合成甲');
    check.equal(first.B4?.v, '合成乙');

    app.setSubject('2203');
    app.fire('genXlsx', 'click');
    const second = XLSX.read((await app.readDownload('xlsx')).bytes, { type: 'buffer' }).Sheets.明细;
    check.equal(second.B3?.v, '合成丙');
    check.equal(second.D3?.v, 50);
    check.equal(second.B4, undefined, '上次导出的客户不能残留');
    check.equal(second.D4, undefined, '上次导出的金额不能残留');
  },
  async '宏运行前检查全部目标公式，发现任意公式就停止'(check) {
    const app = createApp();
    await app.loadTemplate(workbook({ 明细: detailRows(1) }));
    await app.loadData(data);
    app.setSubject('1221');
    app.fire('genVba', 'click');
    const lines = (await app.readDownload('bas')).bytes.toString('utf8').split('\r\n');
    // 守卫语句现在排在备份之后，取到第一句写入之前就行
    const firstWrite = lines.findIndex(line => /^    ws\\.Range\\('[A-Z]+\\d+'\\)\\.Value = /.test(line));
    const beforeWrite = lines.slice(0, firstWrite);
    const guards = beforeWrite.flatMap(line => {
      const match = line.match(/^    If ws\.Range\("([A-Z]+\d+)"\)\.HasFormula Or ws\.Range\("\1"\)\.HasArray Then GoTo FormulaFound$/);
      return match ? [match[1]] : [];
    });
    check.deepEqual(guards, ['B3', 'C3', 'D3', 'E3', 'F3']);
    check.match(lines.join('\n'), /Exit Sub\nFormulaFound:\n    MsgBox ".*公式.*没有写入/);
  }
});
