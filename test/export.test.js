const { createApp, workbook, journal, runSuite, XLSX } = require('./app-fixture');

const sheetName = '其他应收款明细';
const amounts = ['序号', '项目名称', '期初金额', '本期借方', '本期贷方'];
amounts[26] = '期末金额'; // AA，验证 VBA 的多字母列地址
const template = workbook({
  [sheetName]: [['合成底稿'], [], amounts, ['', '一'], [], [null, null, null, 9, 4], [], ['', '合计']],
  审定表: [['合成审定数据']]
});
const detail = template.Sheets[sheetName];
detail.AA6 = { t: 'n', f: 'D6+E6', v: 13 };
detail.AA8 = { t: 'n', f: 'SUM(AA5:AA7)', v: 13 };
template.Sheets.审定表.A1 = { t: 'n', f: "'" + sheetName + "'!AA8", v: 13 };

const writes = {
  B5: '合成甲', C5: 0, D5: 1234.56, E5: 234.56, AA5: 1000,
  B7: '合成乙', C7: 0, D7: 0, E7: 25.75, AA7: -25.75
};

runSuite('导出回读', {
  async '从按钮下载的 XLSX 读回公式、缓存值和写入数据'(check) {
    const app = createApp();
    await app.loadTemplate(template);
    await app.loadData(journal([
      ['2026-01-05', '1221', '其他应收款', '合成甲', 1234.56, 234.56],
      ['2026-01-06', '1221', '其他应收款', '合成乙', null, 25.75]
    ]));
    app.setSubject('其他应收款');

    app.fire('genXlsx', 'click');
    const xlsx = await app.readDownload('xlsx');
    check.equal(xlsx?.name, '已填充_合成底稿.xlsx');
    check.equal(xlsx.bytes.subarray(0, 4).toString('hex'), '504b0304', '下载内容是 XLSX 字节');
    const exported = XLSX.read(xlsx.bytes, { type: 'buffer', cellFormula: true });
    check.deepEqual(exported.SheetNames, [sheetName, '审定表']);
    const output = exported.Sheets[sheetName];
    for (const [address, formula, value] of [
      ['AA6', 'D6+E6', 13], ['AA8', 'SUM(AA5:AA7)', 13]
    ]) {
      check.equal(output[address]?.f, formula, address + ' 公式原文');
      check.equal(output[address]?.v, value, address + ' 公式缓存值');
    }
    check.equal(exported.Sheets.审定表.A1?.f, "'" + sheetName + "'!AA8");
    check.equal(exported.Sheets.审定表.A1?.v, 13);
    check.equal(output.B6, undefined, '公式行没有写入客户');
    check.equal(output.D6?.v, 9, '公式行的借方未改');
    check.equal(output.E6?.v, 4, '公式行的贷方未改');
    for (const [address, value] of Object.entries(writes)) {
      check.equal(output[address]?.v, value, address + ' 写入值');
      check.equal(output[address]?.t, typeof value === 'string' ? 's' : 'n', address + ' 单元格类型');
    }

    app.fire('genVba', 'click');
    const macro = await app.readDownload('bas');
    check.equal(macro?.name, '填充宏_' + sheetName + '.bas');
    const vba = macro.bytes.toString('utf8');
    check.match(vba, /Worksheets\("其他应收款明细"\)/);
    const assignments = [...vba.matchAll(/^    ws\.Range\("([A-Z]+\d+)"\)\.Value = (.+)$/gm)]
      .map(([, address, value]) => [address, value]);
    check.deepEqual(assignments, Object.entries(writes).map(([address, value]) =>
      [address, typeof value === 'string' ? '"' + value + '"' : value.toFixed(2)]
    ), '逐一核对宏地址和值，含 AA 列、跳过第 6 行和合计行');

    app.fire('genCsv', 'click');
    const csv = await app.readDownload('csv');
    check.equal(csv?.name, '明细行_' + sheetName + '.csv');
    check.equal(csv.bytes.subarray(0, 3).toString('hex'), 'efbbbf');
    check.equal(csv.bytes.toString('utf8').slice(1), [
      '客户/项目名称,期初,本期借方,本期贷方,期末',
      '"合成甲",0.00,1234.56,234.56,1000.00',
      '"合成乙",0.00,0.00,25.75,-25.75'
    ].join('\r\n'));
  }
});
