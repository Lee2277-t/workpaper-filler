const assert = require('node:assert/strict');
const { Blob, File } = require('node:buffer');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.join(__dirname, '..');
const XLSX = require('../vendor/xlsx.full.min.js');
const Core = require('../core');
const source = ['vendor/xlsx.full.min.js', 'core.js', 'app.js'].map(file =>
  new vm.Script(fs.readFileSync(path.join(root, file), 'utf8'), { filename: file })
);
const html = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const exportButtons = ['genVba', 'genCsv', 'genXlsx'];

function createApp() {
  const elements = new Map();
  const downloads = new Map();
  const objectUrls = new Map();
  let urlCount = 0;

  function makeElement(tag) {
    const handlers = new Map();
    const classes = new Set();
    let markup = '';
    return {
      value: '', textContent: '', disabled: false, files: null, style: {}, download: '',
      children: [],
      get innerHTML() { return markup; },
      set innerHTML(value) { markup = value; this.children.length = 0; },
      classList: {
        add: name => classes.add(name),
        remove: name => classes.delete(name),
        contains: name => classes.has(name)
      },
      addEventListener(type, handler) {
        if (!handlers.has(type)) handlers.set(type, []);
        handlers.get(type).push(handler);
      },
      dispatchEvent(event) {
        event.target = this;
        for (const handler of handlers.get(event.type) || []) handler.call(this, event);
      },
      appendChild(child) { this.children.push(child); },
      removeChild(child) { this.children.splice(this.children.indexOf(child), 1); },
      click() {
        if (this.disabled) return;
        this.dispatchEvent({ type: 'click' });
        if (tag === 'a' && this.download) {
          const extension = path.extname(this.download).slice(1);
          downloads.set(extension, { name: this.download, blob: objectUrls.get(this.href) });
        }
      }
    };
  }

  for (const [, tag, before, id, after] of html.matchAll(/<(\w+)([^>]*?)\bid="([^"]+)"([^>]*)>/g)) {
    const attrs = before + after;
    const element = makeElement(tag);
    element.disabled = /\bdisabled\b/.test(attrs);
    for (const name of (attrs.match(/\bclass="([^"]+)"/)?.[1] || '').split(/\s+/).filter(Boolean)) {
      element.classList.add(name);
    }
    elements.set(id, element);
  }

  function el(id) {
    if (!elements.has(id)) throw new Error('找不到页面元素 ' + id);
    return elements.get(id);
  }
  function fire(id, type) {
    if (type === 'click') el(id).click();
    else el(id).dispatchEvent({ type });
  }
  class FileReader {
    readAsArrayBuffer(file) {
      file.arrayBuffer().then(result => {
        this.result = result;
        this.onload();
      }, error => this.onerror(error));
    }
  }

  const browser = {
    document: { readyState: 'complete', getElementById: el, createElement: makeElement, body: makeElement('body') },
    Blob, FileReader, TextDecoder,
    URL: {
      createObjectURL(blob) {
        const url = 'blob:fixture/' + ++urlCount;
        objectUrls.set(url, blob);
        return url;
      },
      revokeObjectURL(url) { objectUrls.delete(url); }
    },
    // SheetJS schedules URL cleanup after 60 seconds. That timer need not keep Node alive.
    setTimeout(callback, delay) { return setTimeout(callback, delay).unref(); }
  };
  browser.window = browser.self = browser;
  const context = vm.createContext(browser);
  for (const script of source) script.runInContext(context);

  async function loadFile(id, content, name) {
    el(id).files = [new File([content], name)];
    fire(id, 'change');
    await new Promise(setImmediate);
  }
  function loadWorkbook(id, wb, name) {
    return loadFile(id, XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }), name);
  }
  return {
    el, fire,
    loadTemplate(wb) { return loadWorkbook('tplFile', wb, '合成底稿.xlsx'); },
    loadData(wb) { return loadWorkbook('dataFile', wb, '合成序时账.xlsx'); },
    loadCsv(csv) { return loadFile('dataFile', csv, '合成序时账.csv'); },
    setSubject(value) {
      el('subject').value = value;
      fire('subject', 'input');
    },
    async readDownload(kind) {
      const download = downloads.get(kind);
      if (!download) return null;
      return { name: download.name, bytes: Buffer.from(await download.blob.arrayBuffer()) };
    }
  };
}

function workbook(sheets) {
  const wb = XLSX.utils.book_new();
  for (const [name, rows] of Object.entries(sheets)) {
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), name);
  }
  return wb;
}

function detailRows(blankCount) {
  return [
    ['合成明细'],
    ['序号', '项目名称', '期初金额', '本期借方', '本期贷方', '期末金额'],
    ...Array.from({ length: blankCount }, () => []),
    ['', '合计']
  ];
}

const journalHeader = ['日期', '科目代码', '科目名称', '客户档案名称', '借方金额', '贷方金额'];
function journal(rows) {
  return workbook({ 序时账: [journalHeader, ...rows] });
}

async function runSuite(label, cases) {
  let count = 0;
  const check = {};
  for (const method of ['equal', 'deepEqual', 'match', 'doesNotMatch', 'ok']) {
    check[method] = (...args) => {
      assert[method](...args);
      count++;
    };
  }
  const failures = [];
  for (const [name, test] of Object.entries(cases)) {
    try {
      await test(check);
    } catch (error) {
      failures.push(name + '\n' + error.stack);
    }
  }
  if (failures.length) {
    console.error(label + '失败 ' + failures.length + ' 项用例');
    failures.forEach((failure, i) => console.error('  ' + (i + 1) + '. ' + failure));
    process.exitCode = 1;
  } else {
    console.log(label + '全部通过：' + count + ' 项断言');
  }
}

module.exports = { createApp, workbook, detailRows, journal, exportButtons, runSuite, XLSX, Core };
