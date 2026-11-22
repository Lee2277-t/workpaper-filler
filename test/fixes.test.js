const { createApp, workbook, detailRows, journal, runSuite, XLSX } = require('./app-fixture');

const data = journal([['2026-01-05', '1221', '其他应收款', '合成甲', 100, 5]]);

runSuite('缺陷回归', {
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
