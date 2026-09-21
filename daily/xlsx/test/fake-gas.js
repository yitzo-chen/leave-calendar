// 測試用假 Apps Script 環境：同步 zip / Drive / 試算表（不依賴網路）
const zlib = require("zlib");
const fs = require("fs");
const vm = require("vm");

// ---- 同步 zip ----
const CRC_TABLE = (() => { const t = []; for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
function crc32(buf) { let c = 0xffffffff; for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; }
function readZip(buf) {
  let eocd = buf.length - 22;
  while (eocd >= 0 && buf.readUInt32LE(eocd) !== 0x06054b50) eocd--;
  const count = buf.readUInt16LE(eocd + 10);
  let p = buf.readUInt32LE(eocd + 16);
  const out = [];
  for (let i = 0; i < count; i++) {
    const method = buf.readUInt16LE(p + 10), csize = buf.readUInt32LE(p + 20);
    const nlen = buf.readUInt16LE(p + 28), elen = buf.readUInt16LE(p + 30), clen = buf.readUInt16LE(p + 32), off = buf.readUInt32LE(p + 42);
    const name = buf.slice(p + 46, p + 46 + nlen).toString("utf8");
    const lnlen = buf.readUInt16LE(off + 26), lelen = buf.readUInt16LE(off + 28);
    const data = buf.slice(off + 30 + lnlen + lelen, off + 30 + lnlen + lelen + csize);
    out.push({ name, data: method === 8 ? zlib.inflateRawSync(data) : data });
    p += 46 + nlen + elen + clen;
  }
  return out;
}
function writeZip(entries) {
  const locals = [], centrals = []; let offset = 0;
  for (const e of entries) {
    const name = Buffer.from(e.name, "utf8"), raw = e.data, comp = zlib.deflateRawSync(raw), crc = crc32(raw);
    const lh = Buffer.alloc(30); lh.writeUInt32LE(0x04034b50, 0); lh.writeUInt16LE(20, 4); lh.writeUInt16LE(0x0800, 6); lh.writeUInt16LE(8, 8);
    lh.writeUInt32LE(crc, 14); lh.writeUInt32LE(comp.length, 18); lh.writeUInt32LE(raw.length, 22); lh.writeUInt16LE(name.length, 26);
    const ch = Buffer.alloc(46); ch.writeUInt32LE(0x02014b50, 0); ch.writeUInt16LE(20, 4); ch.writeUInt16LE(20, 6); ch.writeUInt16LE(0x0800, 8); ch.writeUInt16LE(8, 10);
    ch.writeUInt32LE(crc, 16); ch.writeUInt32LE(comp.length, 20); ch.writeUInt32LE(raw.length, 24); ch.writeUInt16LE(name.length, 28); ch.writeUInt32LE(offset, 42);
    locals.push(lh, name, comp); centrals.push(ch, name);
    offset += 30 + name.length + comp.length;
  }
  const cd = Buffer.concat(centrals), end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(entries.length, 8); end.writeUInt16LE(entries.length, 10); end.writeUInt32LE(cd.length, 12); end.writeUInt32LE(offset, 16);
  return Buffer.concat([...locals, cd, end]);
}

class Blob {
  constructor(bytes, type, name) { this.bytes = Buffer.from(bytes); this.type = type; this.name = name; }
  getName() { return this.name; }
  setName(n) { this.name = n; return this; }
  getBytes() { return [...this.bytes]; }
  getDataAsString() { return this.bytes.toString("utf8"); }
  setContentType(t) { this.type = t; return this; }
}

/** 建立一個假環境：sheetsData = {分頁名: 二維陣列(含表頭)}，drive = {資料夾檔案 {名稱: Buffer}} */
function makeGlueEnv(opts) {
  const s = { sheetsData: opts.sheetsData, files: {}, nextId: 1, props: {}, patches: 0, calls: [] };
  function mkFile(name, bytes) {
    const id = "F" + s.nextId++;
    const f = { id, name, bytes: Buffer.from(bytes), trashed: false };
    s.files[id] = f;
    return {
      getId: () => id, getName: () => f.name, isTrashed: () => f.trashed, setTrashed: (t) => { f.trashed = t; },
      getBlob: () => new Blob(f.bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", f.name), getSize: () => f.bytes.length,
      getUrl: () => "https://drive.google.com/file/d/" + id, _f: f,
      getParents: () => ({ hasNext: () => false }),
    };
  }
  const folder = {
    getFilesByName(name) {
      const list = Object.values(s.files).filter((f) => f.name === name && !f.trashed);
      let i = 0; return { hasNext: () => i < list.length, next: () => mkFile.wrap(list[i++]) };
    },
    createFile(blob) { const w = mkFile(blob.name, blob.bytes); return w; },
  };
  mkFile.wrap = (f) => ({
    getId: () => f.id, getName: () => f.name, isTrashed: () => f.trashed, setTrashed: (t) => { f.trashed = t; },
    getBlob: () => new Blob(f.bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", f.name), getSize: () => f.bytes.length, getUrl: () => "https://drive.google.com/file/d/" + f.id,
  });
  for (const [name, bytes] of Object.entries(opts.driveFiles || {})) mkFile(name, bytes);

  function sheet(name) {
    return { getDataRange: () => ({ getValues: () => JSON.parse(JSON.stringify(s.sheetsData[name])) }) };
  }
  const ctx = {
    console, Logger: { log: () => {} }, JSON, Date, Object, Array, String, Number, Math, RegExp, Error, parseInt, isNaN, isFinite,
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getId: () => "SS1", getSheetByName: (n) => (s.sheetsData[n] ? sheet(n) : null) }) },
    DriveApp: {
      getFileById(id) { if (id === "SS1") return { getParents: () => ({ hasNext: () => true, next: () => folder }) }; const f = s.files[id]; if (!f) throw new Error("no file " + id); return mkFile.wrap(f); },
      getRootFolder: () => folder,
    },
    Utilities: {
      unzip: (blob) => { if (blob.type !== "application/zip") throw new Error("Exception: Invalid argument: ContentType. Should be of type: application/zip"); return readZip(blob.bytes).map((e) => new Blob(e.data, "text/xml", e.name)); },
      zip: (blobs, name) => new Blob(writeZip(blobs.map((b) => ({ name: b.name, data: b.bytes }))), "application/zip", name),
      newBlob: (data, type, name) => new Blob(typeof data === "string" ? Buffer.from(data, "utf8") : Buffer.from(data), type, name),
      formatDate: (d) => d.toISOString().slice(0, 10),
    },
    Session: { getScriptTimeZone: () => "Asia/Taipei" },
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => (k in s.props ? s.props[k] : null), setProperty: (k, v) => { s.props[k] = v; } }) },
    ScriptApp: { getOAuthToken: () => "fake-token" },
    UrlFetchApp: {
      fetch(url, o) {
        const id = /files\/([^?]+)\?uploadType=media/.exec(url)[1];
        s.patches++; s.calls.push({ url, method: o.method });
        if (!s.files[id]) return { getResponseCode: () => 404, getContentText: () => "nf" };
        s.files[id].bytes = Buffer.from(o.payload);
        return { getResponseCode: () => 200, getContentText: () => "{}" };
      },
    },
  };
  vm.createContext(ctx);
  const run = (path) => vm.runInContext(fs.readFileSync(path, "utf8"), ctx, { filename: path });
  const root = require("path").join(__dirname, "..", ".."); // daily/（此檔位於 daily/xlsx/test/）
  run(root + "/google-apps-script.gs");     // SHEET_* 常數與 normalizeDate
  run(root + "/xlsx/xlsx-builder.js");
  run(root + "/xlsx/xlsx-drive.gs");
  return { state: s, ctx, call: (code) => vm.runInContext(code, ctx), readZip, writeZip, Blob };
}
module.exports = { makeGlueEnv, readZip, writeZip };
