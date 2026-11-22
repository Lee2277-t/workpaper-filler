const fs = require('fs');
/**
 * 底稿填充助手 · 核心逻辑回归
 */
const path = require('path');
const Core = require(path.join(__dirname, '..', 'core.js'));

let pass = 0, fail = [];
function ok(cond, msg) { if (cond) pass++; else fail.push(msg); }
function eq(a, b, msg) {
  ok(a === b, msg + ' — 期望 ' + JSON.stringify(b) + '，实际 ' + JSON.stringify(a));
}

/* ---------------- 金额 ---------------- */

console.log('[金额]');
eq(Core.toCents('1,234.56'), 123456, '千分位');
eq(Core.toCents('-500'), -50000, '负数');
eq(Core.toCents('待核实'), null, '文字读不出');
eq(Core.moneyText(123456), '1,234.56', '显示');

/* ---------------- 认序时账 ---------------- */

console.log('\n[认序时账]');
const journal = [
  ['记账凭证'],
  ['2026年满足条件查询的全部'],
  ['日期', '字', '号', '摘要', '科目代码', '科目名称', '借方金额', '贷方金额', '客户档案代码', '客户档案名称'],
  ['2026-01-05', '记', '1', '收货款', '1221.01', '其他应收款', 1000, 0, 'KH001', '华远'],
  ['2026-01-06', '记', '2', '付款', '1221.01', '其他应收款', 0, 300, 'KH001', '华远'],
  ['2026-01-07', '记', '3', '借款', '1221.02', '其他应收款', 500, 0, 'KH002', '东方'],
  ['2026-01-08', '记', '4', '转账', '1002', '银行存款', 0, 200, '', '']
];
const jh = Core.detectJournal(journal);
eq(jh.rowIndex, 2, '表头在第 3 行');
eq(jh.cols.code, 4, '科目代码列');
eq(jh.cols.debit, 6, '借方列');
eq(jh.cols.credit, 7, '贷方列');
eq(jh.cols.party, 9, '客户档案列');

const jr = Core.readJournal(journal, jh);
eq(jr.records.length, 4, '四条记录');
eq(jr.records[0].debit, 100000, '借方金额');
eq(jr.records[1].credit, 30000, '贷方金额');
eq(jr.records[0].party, '华远', '客户名');

/* ---------------- 按客户汇总 ---------------- */

console.log('\n[按客户汇总]');
eq(Core.cleanParty('客户档案－中国金谷'), '中国金谷', '去掉 ERP 的客户档案前缀');
eq(Core.cleanParty('华远建设'), '华远建设', '没有前缀的不动');

const p1 = Core.summarizeByParty(jr.records, '1221');
eq(p1.length, 2, '科目 1221 有两户');
eq(p1[0].party, '华远', '第一户');
eq(p1[0].debit, 100000, '华远借方');
eq(p1[0].credit, 30000, '华远贷方');
eq(p1[0].balance, 70000, '华远净额');
eq(p1[1].party, '东方', '第二户');

const p2 = Core.summarizeByParty(jr.records, '1221.02');
eq(p2.length, 1, '按明细科目筛只有一户');

const p3 = Core.summarizeByParty(jr.records, '不存在的科目');
eq(p3.length, 0, '筛不到就是空');

const p4 = Core.summarizeByParty(jr.records, '其他应收款');
eq(p4.length, 2, '按科目名也能筛');

/* ---------------- 认明细表 ---------------- */

console.log('\n[认明细表]');
const detail = [
  ['', '预收款项明细表'],
  ['', '', '', '', '', '', '', ''],
  ['', '序号', '项目名称', '款项性质', '本期未审'],
  ['科目编号', '', '', '', '期初金额', '本期借方', '本期贷方', '期末金额'],
  ['', '一', '合并范围内关联方', '——', '', '', '', ''],
  ['', '', '', '', '', '', '', '']
];
const d = Core.detectDetail(detail);
eq(d.usable, true, '能认出来');
eq(d.headerRow, 3, '金额表头在第 4 行');
eq(d.cols.no, 1, '序号列');
eq(d.cols.name, 2, '项目名称列');
eq(d.cols.open, 4, '期初列');
eq(d.cols.debit, 5, '借方列');
eq(d.cols.credit, 6, '贷方列');
eq(d.cols.close, 7, '期末列');

const notDetail = Core.detectDetail([['日期', '客户', '金额'], ['2025-01-01', '华远', 100]]);
eq(notDetail.usable, false, '普通表不当明细表');

/* ---------------- 找空行 ---------------- */

console.log('\n[找空行]');
const blankRows = Core.findBlankRows(detail, d, function (r, c) {
  // 第 5 行（索引 4）的分组合计格是公式
  return r === 4 && c >= 4;
});
eq(blankRows.length, 1, '只有最后一行是可填的空行');
eq(blankRows[0], 5, '空行是第 6 行（0 基索引 5）');

const filled = detail.slice();
filled[5] = ['', '', '已有客户', '', 1, 2, 3, 4];
eq(Core.findBlankRows(filled, d, function () { return false; }).length, 0, '已有内容的行不覆盖');

/* ---------------- 生成计划 ---------------- */

console.log('\n[生成计划]');
const parties = [
  { party: '华远', debit: 100000, credit: 30000, count: 2, balance: 70000 },
  { party: '东方', debit: 50000, credit: 0, count: 1, balance: 50000 }
];
const blanks = [5, 6];
const res = Core.planWrites(parties, blanks, d, {});
eq(res.plan.length, 2, '两户两行');
eq(res.plan[0].row, 5, '第一户写第 6 行');
eq(res.plan[0].party, '华远', '客户名');
const cell = (p, col) => p.cells.filter(c => c.col === col)[0];
eq(cell(res.plan[0], 2).value, '华远', '名称写成文本');
eq(cell(res.plan[0], 5).value, 1000, '借方 1000 元');
eq(cell(res.plan[0], 6).value, 300, '贷方 300 元');
eq(cell(res.plan[0], 7).value, 700, '期末 = 0 + 1000 - 300');

const tooMany = Core.planWrites(parties, [5], d, {});
eq(tooMany.plan.length, 2, '空行不够也照排，缺的行后面补');
eq(tooMany.missing, 1, '算出缺 1 行');
ok(tooMany.problems.join('').indexOf('补 1 行') >= 0, '说明要补几行');


/* ---------------- 宏 ---------------- */

console.log('\n[期初]');
{
  const o = Core.parseOpening('客户\t期初\n华远\t1,000.00\n东方\t500\n坏行');
  eq(Object.keys(o.map).length, 2, '两户期初');
  eq(o.map['华远'], 100000, '华远期初 1000 元');
  eq(o.bad.length, 1, '一行读不出');
  const withOpen = Core.planWrites(parties, [5, 6], d, { openCents: o.map });
  eq(cell(withOpen.plan[0], 4).value, 1000, '期初写进第 4 列');
  eq(withOpen.plan[0].close, 170000, '期末 = 1000 + 1000 - 300');
  ok(withOpen.problems.length === 0, '都给了期初就不提示');
  const partial = Core.planWrites(parties.concat([{ party: '南方', debit: 100, credit: 0 }]), [5, 6, 7], d, { openCents: o.map });
  ok(partial.problems.join('').indexOf('没给数') >= 0, '缺期初的户会提示');
}


console.log('\n[分组边界 / 空行]');
{
  const det2 = { cols: { no: 0, name: 1, open: 2, debit: 3, credit: 4, close: 5 }, headerRow: 1, groupRow: 1 };
  const grid = [
    ['序号', '项目名称', '期初金额', '本期借方', '本期贷方', '期末金额'],
    [1, '', null, null, null, null],
    ['', '本项目小计', 1, 2, 3, 4],
    ['', '二、非关联方', null, null, null, null],
    [2, '', null, null, null, null],
    [3, '', null, null, null, null],
    ['', '合计', null, null, null, null]
  ];
  const gs = Core.listGroups(grid, det2, () => false);
  const g2 = gs.filter(g => g.title === '二、非关联方')[0];
  ok(!!g2, '认出第二个分组');
  eq(g2.blanks.map(r => r + 1).join(','), '5,6', '二组前面有「小计」也照样能扫到空行');

  // 合计写在 A 列（不在项目名称列）也要截断
  const grid2 = [
    ['序号', '项目名称', '期初金额', '本期借方', '本期贷方', '期末金额'],
    [1, '', null, null, null, null],
    ['合计', '', null, null, null, null],
    ['', '', null, null, null, null]
  ];
  const det2b = Object.assign({}, det2, { headerRow: 0, groupRow: 0 });
  const b2 = Core.findBlankRows(grid2, det2b, () => false);
  eq(b2.map(r => r + 1).join(','), '2', '合计在 A 列也能截断，不会把账龄区当空行');

  // 0 个空行不推算行号
  const zero = Core.planWrites([{ party: '甲', debit: 100, credit: 0 }], [], det2, {});
  eq(zero.plan.length, 0, '0 空行时不出计划');
  ok(zero.problems.join('').indexOf('没有可填的空行') >= 0, '0 空行时说清楚要先插行');
}


console.log('\n[期末方向]');
{
  const d2 = { cols: { name: 1, open: 4, debit: 5, credit: 6, close: 7 }, headerRow: 0 };
  const p = [{ party: '甲', debit: 20000, credit: 30000 }];
  const asAsset = Core.planWrites(p, [4], d2, { openCents: { 甲: 100000 } });
  eq(asAsset.plan[0].close, 90000, '资产口径：1000 + 200 - 300 = 900');
  const asLiability = Core.planWrites(p, [4], d2, { openCents: { 甲: 100000 }, closeSign: -1 });
  eq(asLiability.plan[0].close, 110000, '负债口径：1000 - 200 + 300 = 1100');
  const cell7 = cell(asLiability.plan[0], 7);
  eq(cell7 && cell7.value, 1100, '负债口径的期末写进第 7 列');
}

console.log('\n[未来行的合并判断]');
{
  const d2 = { cols: { name: 1, open: 4, debit: 5, credit: 6, close: 7 }, headerRow: 0 };
  // 第 5 行（0基 4）是最后一个空行，第 6 行是合计行且被合并
  const isMerged = (r, cc) => r === 5 && cc >= 0;
  const res = Core.planWrites(
    [{ party: '甲', debit: 100, credit: 0 }, { party: '乙', debit: 200, credit: 0 }],
    [4], d2, { isMergedSubordinate: isMerged, canWrite: () => true });
  eq(res.plan.length, 2, '合并的合计行不该挡住要插进来的客户');
  eq(res.plan[1].row, 5, '第二户排在模板行之后，插行后落到它下面');
}

console.log('\n[版本号一致]');
{
  // 版本号漂过一次：core.js 改了、index.html 忘了改，界面上显示的就不是发出去的版本
  const html = fs.readFileSync(__dirname + '/../index.html', 'utf8');
  ok(/^\d+\.\d+\.\d+$/.test(Core.VERSION), 'core.js 版本号是 x.y.z');
  ok(html.indexOf('v' + Core.VERSION) >= 0, 'index.html 里的版本号与 core.js 一致');
}

console.log('\n[生成宏]');
const vba = Core.buildVBA(res.plan, '预收款项明细表', { backup: false });
ok(vba.indexOf('Sub 填充底稿明细()') >= 0, '有 Sub 入口');
ok(vba.indexOf('Worksheets("预收款项明细表")') >= 0, '指向正确的工作表');
ok(vba.indexOf('ws.Range("C6").Value = "华远"') >= 0, '名称写 C6');
ok(vba.indexOf('ws.Range("F6").Value = 1000.00') >= 0, '借方写 F6');
ok(vba.indexOf('ws.Range("H6").Value = 700.00') >= 0, '期末写 H6');
ok(vba.indexOf('ws.Range("G6").Value = 300.00') >= 0, '贷方写 G6');
ok(vba.indexOf('End Sub') >= 0, '有结束');
ok(vba.indexOf('MsgBox') >= 0, '有完成提示');

const vbaQuoted = Core.buildVBA([{ row: 0, party: 'x', cells: [{ col: 0, value: '带"引号"的名字', text: true }] }], 'A', { backup: false });
ok(vbaQuoted.indexOf('""引号""') >= 0, 'VBA 里的引号要转义');

eq(Core.colLetter(0), 'A', '第 0 列是 A');
eq(Core.colLetter(25), 'Z', '第 25 列是 Z');
eq(Core.colLetter(26), 'AA', '第 26 列是 AA');
eq(Core.colLetter(27), 'AB', '第 27 列是 AB');

/* ---------------- CSV ---------------- */

console.log('\n[明细行 CSV]');
const csv = Core.buildCsv(res.plan, d);
const lines = csv.split('\r\n');
eq(lines.length, 3, '表头 + 两行');
ok(lines[0].indexOf('客户') >= 0, '有表头');
ok(lines[1].indexOf('华远') >= 0, '有华远');
ok(lines[1].indexOf('1000.00') >= 0, '有借方金额');

console.log('\n============================================================');
if (fail.length) {
  console.log('失败 ' + fail.length + ' / ' + (pass + fail.length) + ' 项：');
  fail.forEach((m, i) => console.log('  ' + (i + 1) + '. ' + m));
  process.exit(1);
}
console.log('核心逻辑全部通过：' + pass + ' 项断言');
