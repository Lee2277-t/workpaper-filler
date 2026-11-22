const { createApp, workbook, detailRows, journal, exportButtons, runSuite } = require('./app-fixture');

const normal = ['2026-01-05', '1221', '其他应收款', '合成甲', 100, 5];

async function prepare(data, rows = detailRows(5)) {
  const app = createApp();
  await app.loadTemplate(workbook({ 合成明细: rows }));
  await app.loadData(data);
  app.setSubject('1221');
  return app;
}

async function checkBlocked(check, app) {
  for (const [id, kind] of exportButtons.map((id, i) => [id, ['bas', 'csv', 'xlsx'][i]])) {
    check.equal(app.el(id).disabled, true, id + ' 已禁用');
    app.fire(id, 'click');
    check.equal(await app.readDownload(kind), null, kind + ' 没有生成文件');
  }
  check.equal(app.el('planBody').innerHTML, '');
}

runSuite('边界输入', {
  async '序时账只有表头'(check) {
    const app = await prepare(journal([]));
    check.match(app.el('dataInfo').textContent, /序时账：序时账，0 行/);
    check.match(app.el('infoBox').innerHTML, /没有带客户档案的记录/);
    await checkBlocked(check, app);
  },
  async '明细表没有金额列'(check) {
    const app = await prepare(journal([normal]), [['序号', '项目名称'], [], ['', '合计']]);
    check.deepEqual(app.el('sheetSel').children, []);
    check.match(app.el('infoBox').innerHTML, /没找到明细表/);
    await checkBlocked(check, app);
  },
  async '同一表头有两列本期借方金额'(check) {
    const rows = detailRows(5);
    rows[1].push('本期借方（调整后）');
    const app = await prepare(journal([normal]), rows);
    check.match(app.el('tplInfo').textContent, /合成明细.*本期借方.*重复/);
    check.deepEqual(app.el('sheetSel').children, []);
    await checkBlocked(check, app);
  },
  async '100 户只有 5 个空行，排满并说明要补几行'(check) {
    const customers = Array.from({ length: 100 }, (_, i) => [
      '2026-01-05', '1221', '其他应收款', '合成客户' + (i + 1), 10, 1
    ]);
    const app = await prepare(journal(customers));
    check.match(app.el('infoBox').innerHTML, /客户有 100 户.*只有 5 个空行.*要往下补 95 行/);
    check.match(app.el('statBox').innerHTML, />100 户<.*>1,000\.00<.*>100\.00<.*>5 行</);
    // 行数不够也让做：宏会插行，直接写 xlsx 不行
    check.equal(app.el('genXlsx').disabled, true, '直接生成 xlsx 要禁用');
    app.fire('genVba', 'click');
    const vba = (await app.readDownload('bas')).bytes.toString('utf8');
    check.match(vba, /ws\.Rows\(\d+\)\.Copy/);
    check.match(vba, /Insert Shift:=xlDown/);
  },
  async '客户名里的双引号在 VBA 中成对转义'(check) {
    const app = await prepare(journal([
      ['2026-01-05', '1221', '其他应收款', '合成"引号"客户', 100, 5]
    ]));
    check.match(app.el('planBody').innerHTML, /合成&quot;引号&quot;客户/);
    app.fire('genVba', 'click');
    const macro = await app.readDownload('bas');
    check.ok(macro, '宏下载存在');
    check.match(macro.bytes.toString('utf8'), /^    ws\.Range\("B3"\)\.Value = "合成""引号""客户"$/m);
  },
  async '文字金额指出源文件行和金额列，改选有效科目仍可生成'(check) {
    const app = createApp();
    await app.loadTemplate(workbook({ 合成明细: detailRows(5) }));
    await app.loadCsv('\ufeff日期,科目代码,科目名称,客户档案名称,借方金额,贷方金额\r\n' +
      '2026-01-05,1221,其他应收款,合成甲,25,5\r\n' +
      '2026-01-06,1221,其他应收款,合成乙,待核实,0\r\n' +
      '2026-01-07,2203,预收款项,合成丙,,17.50');
    app.setSubject('1221');
    check.match(app.el('infoBox').innerHTML, /第 3 行.*借方金额.*待核实/);
    check.equal(app.el('statBox').innerHTML, '');
    await checkBlocked(check, app);

    app.setSubject('2203');
    check.equal(app.el('genXlsx').disabled, false);
    check.match(app.el('statBox').innerHTML, />1 户<.*>0\.00<.*>17\.50</);
    check.match(app.el('planBody').innerHTML, /合成丙/);
  }
});
