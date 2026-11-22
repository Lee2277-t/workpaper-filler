const { createApp, workbook, detailRows, journal, exportButtons, runSuite, Core } = require('./app-fixture');

const rows = detailRows(4);
rows[3] = ['', '已有客户', 0, 42, 3, 39];
const data = journal([
  ['2026-01-05', '1221.01', '其他应收款', '合成甲', 1234.56, 0],
  ['2026-01-06', '1221.01', '其他应收款', '合成甲', 0, 234.56],
  ['2026-01-07', '1221.02', '其他应收款', '合成乙', 500, 100],
  ['2026-01-08', '1002', '银行存款', '合成丙', 999, 999]
]);

runSuite('界面冒烟', {

  async '手填填写范围：只填指定的行，改科目就作废'(check) {
    const rows = detailRows(6);
    const app = createApp();
    await app.loadTemplate(workbook({ 明细: rows }));
    await app.loadData(journal([
      ['2026-01-05', '1221', '其他应收款', '甲', 100, 0],
      ['2026-01-06', '1221', '其他应收款', '乙', 200, 0]
    ]));
    app.setSubject('1221');

    // 没碰输入框：按自动判断
    const autoFrom = app.el('rowFrom').value;
    const autoTo = app.el('rowTo').value;
    check.equal(autoFrom !== '', true, '自动建议回填到框里');
    const planRows = () => [...app.el('planBody').innerHTML.matchAll(/<tr><td>\d+<\/td><td>(\d+)<\/td>/g)]
      .map(m => Number(m[1]));

    // 手填 5-6：只写这两行
    app.el('rowFrom').value = '5';
    app.el('rowTo').value = '6';
    app.fire('rowFrom', 'input');
    check.deepEqual(planRows(), [5, 6], '手填范围严格生效');

    // 换科目：手填作废，回到自动判断
    app.fire('subject', 'input');
    check.deepEqual(planRows(), [3, 4], '换科目后回到自动判断');

    // 范围填反：拒绝并回退
    app.el('rowFrom').value = '9';
    app.el('rowTo').value = '7';
    app.fire('rowFrom', 'input');
    const info = app.el('infoBox').innerHTML;
    check.match(info, /范围填反了/, '填反要提示');
    check.deepEqual(planRows(), [3, 4], '填反不写入');

    // 小数行号：拒绝并回退
    app.el('rowFrom').value = '3.5';
    app.fire('rowFrom', 'input');
    check.match(app.el('infoBox').innerHTML, /行号要填整数/, '小数要提示');

    // 范围比户数小：不插行、不越界
    app.el('rowFrom').value = '5';
    app.el('rowTo').value = '5';
    app.fire('rowFrom', 'input');
    check.deepEqual(planRows(), [5], '范围只给一行就只填一行');
    check.match(app.el('infoBox').innerHTML, /多出来的 1 户没有填/, '超出范围要说清楚');

    // 填反之后：框里的数字不动，但重算不能再把它当授权（历史上踩过）
    app.el('rowFrom').value = '9';
    app.el('rowTo').value = '7';
    app.fire('rowFrom', 'input');
    app.el('openBox').value = '甲 10';
    app.fire('openBox', 'input');
    check.deepEqual(planRows(), [3, 4], '填反之后重算仍走自动判断');
    check.equal(/多出来的/.test(app.el('infoBox').innerHTML), false, '没授权就不该说户数不够');

    // 用户改过之后，键入不能被回填吞掉
    app.el('useAuto').click ? app.fire('useAuto', 'click') : null;
    app.el('rowFrom').value = '';
    app.fire('rowFrom', 'input');
    check.equal(app.el('rowFrom').value, '', '清空后不被回填');
    app.el('rowFrom').value = '12';
    app.fire('rowFrom', 'input');
    check.equal(app.el('rowFrom').value, '12', '键入 12 不被吞成 2');

    // 「用自动判断」清掉人工范围
    app.fire('useAuto', 'click');
    check.equal(String(app.el('rowFrom').value), '3', '用自动判断后框里回到自动值');
  },

  async '选择工作表、核对汇总和明细行，再下载宏'(check) {
    const app = createApp();
    for (const id of exportButtons) check.equal(app.el(id).disabled, true, id + ' 初始禁用');
    await app.loadTemplate(workbook({ 封面: [['合成底稿']], 应收明细: rows, 应付明细: detailRows(2) }));
    check.deepEqual(app.el('sheetSel').children.map(option => [option.value, option.textContent]), [
      ['应收明细', '应收明细'], ['应付明细', '应付明细']
    ]);
    check.equal(app.el('sheetSel').value, '应收明细');
    check.match(app.el('tplInfo').textContent, /3 张表，认出 2 张明细表/);

    await app.loadData(data);
    app.setSubject('1221');
    for (const id of exportButtons) check.equal(app.el(id).disabled, false, id + ' 可以下载');
    const stats = app.el('statBox').innerHTML;
    check.match(stats, /客户户数<\/span><span class="v">2 户<\/span>/);
    check.match(stats, /借方合计<\/span><span class="v">1,734\.56<\/span>/);
    check.match(stats, /贷方合计<\/span><span class="v">334\.56<\/span>/);
    check.match(stats, /明细表空行<\/span><span class="v">3 行<\/span>/);
    check.equal(app.el('planPanel').classList.contains('hidden'), false);
    const plannedRows = [...app.el('planBody').innerHTML.matchAll(/<tr><td>\d+<\/td><td>(\d+)<\/td>/g)]
      .map(match => Number(match[1]));
    check.deepEqual(plannedRows, [3, 5]);
    check.deepEqual(plannedRows, Core.findBlankRows(rows, Core.detectDetail(rows), () => false)
      .slice(0, 2).map(row => row + 1));

    app.fire('genVba', 'click');
    const download = await app.readDownload('bas');
    check.equal(download?.name, '填充宏_应收明细.bas');
    const vba = download.bytes.toString('utf8');
    check.match(vba, /Worksheets\("应收明细"\)/);
    check.match(vba, /ws\.Range\("B3"\)\.Value = "合成甲"/);
    check.match(vba, /ws\.Range\("B5"\)\.Value = "合成乙"/);
    check.match(app.el('genInfo').textContent, /已生成宏/);

    app.el('sheetSel').value = '应付明细';
    app.fire('sheetSel', 'change');
    check.match(app.el('infoBox').innerHTML, /明细表「应付明细」/);
    app.fire('genVba', 'click');
    check.match((await app.readDownload('bas')).bytes.toString('utf8'), /Worksheets\("应付明细"\)/);
  },

  async '没有匹配客户、空行不够和切换到无明细底稿'(check) {
    const app = createApp();
    await app.loadTemplate(workbook({ 应收明细: rows, 容量不足: detailRows(1) }));
    await app.loadData(data);
    app.setSubject('1221');
    check.equal(app.el('genVba').disabled, false);
    app.setSubject('9999');
    for (const id of exportButtons) check.equal(app.el(id).disabled, true);
    check.match(app.el('infoBox').innerHTML, /科目「9999」.*没有带客户档案的记录/);
    check.equal(app.el('planBody').innerHTML, '');
    check.equal(app.el('statBox').innerHTML, '');

    app.setSubject('1221');
    app.el('sheetSel').value = '容量不足';
    app.fire('sheetSel', 'change');
    // 行数不够也能做：宏负责插行，直接写 xlsx 不行
    check.equal(app.el('genVba').disabled, false, '宏可用');
    check.equal(app.el('genXlsx').disabled, true, 'xlsx 禁用');
    check.match(app.el('infoBox').innerHTML, /客户有 2 户.*只有 1 个空行.*要往下补 1 行/);
    app.fire('genVba', 'click');
    const vba = (await app.readDownload('bas')).bytes.toString('utf8');
    check.match(vba, /Insert Shift:=xlDown/);

    await app.loadTemplate(workbook({ 封面: [['无明细表的合成底稿']] }));
    check.deepEqual(app.el('sheetSel').children, []);
    check.match(app.el('infoBox').innerHTML, /没找到明细表/);
    check.equal(app.el('statBox').innerHTML, '', '换掉底稿后清除旧汇总');
    check.equal(app.el('planPanel').classList.contains('hidden'), true, '没有计划时隐藏预览');
    for (const id of exportButtons) check.equal(app.el(id).disabled, true);
  }
});
