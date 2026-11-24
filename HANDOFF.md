# 底稿填充助手 · 交接稿

> 给接手的 AI / 开发者看的。假设你没有任何上下文，读完这份就能改代码、跑测试、发版。
> 仓库：https://github.com/Lee2277-t/workpaper-filler　当前版本：**v0.4.3**

---

## 1. 这个工具是干什么的

审计实习里最耗时的活：把序时账按客户汇总，填进底稿明细表的空白行，**同时不能覆盖表里自带的公式**。

用户流程：

1. 选底稿模板（带公式的 `.xlsx`，或带 VBA 的 `.xlsm`）
2. 选序时账（`.xlsx` / `.csv`）
3. 填科目代码，核对界面上的汇总
4. 生成 Excel 宏 → 在 Excel 里跑；或直接下载填好的 xlsx / CSV

**纯离线单页 HTML**：解压双击 `index.html`，不联网、不安装、数据不出本机。这是硬约束，不要引入任何后端、CDN、构建步骤。

---

## 2. 技术栈与目录

零依赖（除 SheetJS），ES5 语法（要能在旧版 Excel 附带的 IE 内核里跑，别用 `let`/箭头函数）。

```
core.js      779 行  纯逻辑，UMD，可在 node 里 require
app.js       696 行  DOM 绑定、读文件、渲染、导出
index.html    94 行  三个步骤的表单
app.css       36 行
vendor/xlsx.full.min.js   SheetJS（别升级，升级要重跑全部导出测试）

test/
  app-fixture.js    假 DOM + 假 FileReader，163 行，所有界面测试的地基
  core.test.js      88 项断言   核心算法
  fixes.test.js     84 项断言   历次复核缺陷的回归
  limits.test.js    47 项断言   边界输入
  export.test.js    38 项断言   导出字节回读
  ui-smoke.test.js  56 项断言   真点按钮、真读产物
  run-all.ps1       全量回归入口
  pack.ps1          打 zip
  local-paths.ps1   （gitignored）本机 node 路径
```

**总计 313 项断言。改任何逻辑前先跑一遍，改完再跑一遍。**

---

## 3. 数据流

```
读底稿 wb ──► detectDetail(rows)        认出表头行 + 六个列（序号/项目名称/期初/借/贷/期末）
          └─► totalRanges(ws)          扫全表公式，解析 SUM/SUBTOTAL 引用范围
          └─► listGroups(rows, ...)    划出「分组」(合并范围内关联方 / 非合并范围内小计 …)
                  └─► findBlankRows()  每组里哪些行能填
读序时账 ──► detectJournal(rows)        认列
          └─► readJournal()            逐行读成记录
          └─► summarizeByParty()       按客户汇总成 {party, debit, credit, balance}
                └─► planWrites(parties, blanks, detail, opts)
                        └─► plan[]     [{row, party, open, debit, credit, close}, ...]
                              ├─► buildVBA()  生成 .bas 宏
                              ├─► buildCsv()  生成可粘贴文本
                              └─► 直接写 xlsx（SheetJS）
```

`planWrites` 返回 `{plan, problems, missing, insertAfter}`：
- `missing > 0` 表示行数不够，宏会插行；直接写 xlsx 时禁用该入口
- `insertAfter` 是插行的基准行

---

## 4. core.js 的 20 个导出

```js
VERSION              '0.4.3'
str(v)               任何值 → 去空白字符串
norm(v)              全角转半角、去空白，用于比对科目名
toCents(v)           '1,234.56' / 1234.56 → 123456（**金额一律整数分**）
moneyText(cents)     123456 → '1,234.56'
cleanParty(s)        客户名清洗（去掉尾部编号、括号备注）

detectJournal(rows, firstRow)   扫前 12 行找表头，返回 {rowIndex, cols:{date,code,name,debit,credit,party,partyCode,summary}, score}
readJournal(rows, header)       返回 {records, skipped}；records 里金额已是分
findJournalAmountIssue(records) 找出「有摘要没金额」「借贷都为空」的可疑行

summarizeByParty(records, code) 按客户汇总；**没填客户的会归到「（未填客户）」哨兵**
detectDetail(rows)              返回 {headerRow, cols:{no,name,open,debit,credit,close}}
isRowBoundary(row, detail, nameCell)  这一行是不是合计/小计（名称列命中，或名称列空但整行有合计字样）
findBlankRows(rows, detail, isFormula, opts)  核心扫描，见 §5
rowsInRange(rows, detail, isFormula, from, to, opts)  用户手填范围时用，**不看合计范围**
listGroups(rows, detail, isFormula, opts)    返回 [{title, startRow, endRow, blanks, uncovered}]
planWrites(parties, blanks, detail, opts)    生成写入计划
parseOpening(text)              '客户 1000' 每行一条 → {map, bad}
buildVBA(plan, sheetName, opts) 生成完整 .bas
buildCsv(plan, detail)          生成 CSV 文本
colLetter(n)                    0 → 'A'
```

---

## 5. 最关键的一段：`findBlankRows` 的判定顺序

顺序不能改，每一条都是踩过坑加上的：

```js
for (r = start; r < endRow; r++) {
  if (isRowBoundary(...)) break;                    // ① 合计/小计 → 本段结束
  hasAmount = 四个金额列有没有值
  if (nameCell && !hasAmount && seenRow
      && inTotalRange && !inTotalRange(r)) break;   // ② 明细区里冒出无金额的文字标题
                                                     //    （账龄分析/说明），且它自己不在合计范围 → 下一张表
  if (isHidden(r)) continue;                        // ③ 隐藏行不写
  if (nameCell) continue;                           // ④ 已有内容不覆盖
  if (isFormula(r, nameCol)) continue;              // ⑤ 名称格是公式
  if (isMergedSubordinate(r, nameCol)) continue;    // ⑥ 合并从属格
  if (任一金额列有值) continue;                       // ⑦ 有小计/说明行混在明细区
  if (四个金额列全被公式占满) continue;                // ⑧
  if (inTotalRange && !inTotalRange(r)) {           // ⑨ 硬判据：必须在某个金额列的
      uncovered.push(r); continue;                  //    SUM/SUBTOTAL 范围内
  }
  out.push(r);
}
```

**⑨ 是必要条件，不是充分条件。** 账龄表自己也带 `SUM()`，所以它挡不住所有情况——这正是为什么界面上要有「手填范围」。别把它当万能钥匙，也别因为发现它挡不住就删掉；它挡住了「明显不属于明细区」的一大类。

`uncovered` 要一路传到界面（`group.uncovered`），用户才知道「这几行是空的但合计不含它们」。

---

## 6. 手填范围（v0.4.x 的核心交互）

界面上两个输入框 `rowFrom` / `rowTo`，默认填自动判断的结果。

```js
var useManual = state.rangeDirty && manualFrom > 0 && manualTo > 0;
```

- **`rangeDirty` 只有用户真的改过才为 true**。程序回填的建议值不算授权。
- 校验不过（小数、填反、落在表头）→ **只提示，不修改用户填的数字**，这一轮按自动判断走。
  > 曾经写成「校验不过就把框改回建议值」，结果 `rangeDirty` 还是 true，下一次重算（比如粘期初）就把建议值当成了用户指定的范围——静默只填一半客户。
- 换表 / 换组 / 换科目 / 重新上传底稿 → `rangeDirty = false`，并清空两个框。
- 手填时 `planWrites` 传 `noInsert: true`：**行数不够也不插行**，直接告诉用户哪几户没填。
- 手填的行不在合计范围内 → 警告，但仍照填；这句话要跟进宏完成语和导出提示（`state.outsideNote`）。

---

## 7. 历次复核踩过的坑（照着这张表别再踩）

| 症状 | 根因 | 现在的做法 |
|---|---|---|
| 客户被写进账龄分析区 | 账龄表自带 `SUM()`，被当成明细区 | 硬判据 + 手填范围 + `break` 规则 |
| 「小计在上、明细在下」的版式 0 空行 | 扫到小计就 break | `isBlockHeader`：SUM 指向下方（`top === r+1`）→ 块标题自成一组 |
| 明细区已满时把账龄行当空行 | 用「后面 3 行有没有空行」猜块标题 | 删掉这个启发式，只看公式方向 |
| 手填范围写进表头以上 | 没校验 | `manualFrom <= headerRow+1` → 拒绝并回退 |
| 范围小于户数却静默插行 | 没区分手动/自动 | `noInsert` |
| 期末方向写反（负债类） | 方向符号取错 | `inferCloseSign`：三引用规则 + IF/ROUND 内层提取 + 没把握就警告 |
| `.xlsm` 重写后 VBA 丢失 | SheetJS 不保留 `vbaProject.bin` | `.xlsm` 禁用「直接下载 xlsx」，只给宏 |
| 宏编译报「End Sub 后面只能出现注释」 | 模块级变量写在过程之后 | 所有 `Private` 声明放在第一个 `Sub` 之前 |
| 合计没包含新插入的行 | 插行点在 SUM 范围之外 | 宏里 `ConfirmTotalExpansion` 检查 `sumRange` 末端 |
| 合计是逐格相加（`=D3+D4`）时不扩 | `SUM` 判断失效 | 数 `plainTotals` 并提示用户手工改 |
| 二次运行把上次的数再加一遍 | `plainTotals` 是模块级变量 | `Sub` 第一句 `plainTotals = 0` |
| 合并单元格里写数被静默丢掉 | 写到从属格 | `isMergedSubordinate` 跳过 |
| 两行一户版式错位 | 纵向合并 | 只提示，不拦（Excel 插行会自己挪） |

---

## 8. 测试怎么写

`test/app-fixture.js` 提供假 DOM 和假 `FileReader`，直接跑真的 `app.js`：

```js
const { createApp, workbook, journal, detailRows, XLSX } = require('./app-fixture.js');

const app = createApp();
await app.loadTemplate(workbook({ 明细: detailRows(4) }));   // 造一张假底稿
await app.loadData(journal([['2026-01-05', '1221', '其他应收款', '甲', 100, 0]]));
app.setSubject('1221');

app.el('genVba');                       // 取元素
app.fire('genVba', 'click');            // 真点
const out = await app.readDownload('bas');   // 真读产物字节
```

断言用 `check.equal / deepEqual / match`，每个用例写清「这条断言在守什么」。

**新修一个 bug，就在 `fixes.test.js` 里加一条用例，用例名写上症状。** 这是这个项目唯一有效的防腐手段。

跑：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File test\run-all.ps1
```

需要 node 在 PATH，或在 `test/local-paths.ps1` 里写：

```powershell
$env:DSH_NODE = "C:\path\to\node.exe"
```

---

## 9. 打包与发布

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File test\pack.ps1
# → 仓库上一级目录的 底稿填充助手-v<版本>.zip（7 个文件，平铺，无外层目录）
```

发版流程（每次都要走完）：

1. 改 `core.js` 的 `VERSION` 和 `index.html` 的 `<span class="ver">v…</span>`，**两处必须一致**（`core.test.js` 里有断言）
2. 跑全量回归
3. `pack.ps1`
4. `git add -A && git commit && git tag -a v<版本>`
5. `git push origin master --tags`
6. 建 release、传 zip（资产名用 ASCII）
7. **把 zip 下载回来重算 SHA256，和本地比**

发布前自查：`git status` 干净、版本号两处一致、断言全绿。

---

## 10. 已知限制（不是 bug，是有意不做）

- 期初余额要用户粘（序时账里没有客户级期初）
- 关联方分组靠用户在下拉里选，工具不猜
- 一次一个科目
- 序号、款项性质列不填
- `.xlsm` 只能走宏
- 两行一户版式只提示不拦
- 逐格相加的合计只警告

---

## 11. 待办（按建议优先级）

1. **多段 SUM 只解析第一段**：`SUM(F3:F4,F6:F8)` 的第二段被忽略，导致本该能填的行被判成「范围外」。
2. **跨列 SUM 被跳过**：`SUM(C3:F8)` 里 `a.c !== b.c` 直接丢弃。
3. **`SUMIF` / `INDIRECT` / 跨表引用** 认不出 → 退化成「无判据」。至少要提示「这份底稿的合计认不出来，请手填范围」。
4. **`useRowHasValue_REMOVE`** 是死代码，删掉。
5. **表头识别窗口** 现在是 60 行，超出就整个模板被拒。
6. **批量科目**：现在一次一个科目，实习时一个科目跑一遍很烦。
7. **序号自动填**、**款项性质自动填**（可从科目余额表明细取）。

---

## 12. 工作纪律

- **绝不接触真实客户底稿**：测试只用合成夹具，仓库里不得出现任何真实客户名、金额、文件路径。原开发者机器上的客户数据目录不在本仓库内，也不要去找。
- 改 `core.js` 必补断言；只改 `app.js` 也要补界面用例（历史上「新功能零覆盖」直接导致两个版本带病发布）。
- 金额只用整数分，不要用浮点。
- **不要在 PowerShell 里用 `node -e @"..."@` 写含中文和引号的 JS**：`$` 和 `\r` 会被吞，历史上因此把 `$('rowFrom')` 改成 `if (owFrom)` 差点发出去。用编辑工具，或把脚本写成 `.js` 文件再 `node 文件`。
- ES5 语法，别用 `const`/`let`/箭头函数/模板字符串（`app.js` 和 `core.js` 都要能在旧浏览器跑；测试文件可以用新语法）。

---

## 13. 接手第一小时

```powershell
cd <仓库目录>
powershell -NoProfile -ExecutionPolicy Bypass -File test\run-all.ps1   # 应显示 313 项全过
git log --oneline | Select-Object -First 20                            # 看改动脉络
```

然后读这个顺序：`README.md` → `core.js` 的 `findBlankRows` / `listGroups` / `planWrites` → `app.js` 的 `recompute` → `test/fixes.test.js`（历史坑都在这）。

改完第一件事：跑回归，看是不是还是 313 项。
