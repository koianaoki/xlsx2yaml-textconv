const xlsx = require("xlsx");
const jsyaml = require("js-yaml");

// sheet_to_all_json内で呼ばれるため元コードをそのまま引用
function safe_decode_range(range) {
        var o = {s:{c:0,r:0},e:{c:0,r:0}};
        var idx = 0, i = 0, cc = 0;
        var len = range.length;
        for(idx = 0; i < len; ++i) {
                if((cc=range.charCodeAt(i)-64) < 1 || cc > 26) break;
                idx = 26*idx + cc;
        }
        o.s.c = --idx;

        for(idx = 0; i < len; ++i) {
                if((cc=range.charCodeAt(i)-48) < 0 || cc > 9) break;
                idx = 10*idx + cc;
        }
        o.s.r = --idx;

        if(i === len || range.charCodeAt(++i) === 58) { o.e.c=o.s.c; o.e.r=o.s.r; return o; }

        for(idx = 0; i != len; ++i) {
                if((cc=range.charCodeAt(i)-64) < 1 || cc > 26) break;
                idx = 26*idx + cc;
        }
        o.e.c = --idx;

        for(idx = 0; i != len; ++i) {
                if((cc=range.charCodeAt(i)-48) < 0 || cc > 9) break;
                idx = 10*idx + cc;
        }
        o.e.r = --idx;
        return o;
}

// セルの中身がない場合もnullを出すように必要分だけ書き換えている
// see https://npmdoc.github.io/node-npmdoc-xlsx/build..beta..travis-ci.org/apidoc.html#apidoc.element.xlsx.utils.sheet_to_json
xlsx.utils.sheet_to_all_json = function (sheet, opts){
  let val, row, range, header = 0, offset = 1, r, hdr = [], isempty, R, C, v, vv;
  let o = opts != null ? opts : {};
  let raw = o.raw;
  if(sheet == null || sheet["!ref"] == null) return [];
  range = o.range != null ? o.range : sheet["!ref"];
  if(o.header === 1) header = 1;
  else if(o.header === "A") header = 2;
  else if(Array.isArray(o.header)) header = 3;
  switch(typeof range) {
    case 'string': r = safe_decode_range(range); break;
    case 'number': r = safe_decode_range(sheet["!ref"]); r.s.r = range; break;
    default: r = range;
  }
  if(header > 0) offset = 0;
  let rr = this.encode_row(r.s.r);
  let cols = new Array(r.e.c-r.s.c+1);
  let out = new Array(r.e.r-r.s.r-offset+1);
  let outi = 0;
  for(C = r.s.c; C <= r.e.c; ++C) {
    cols[C] = this.encode_col(C);
    val = sheet[cols[C] + rr];
    switch(header) {
      case 1: hdr[C] = C; break;
      case 2: hdr[C] = cols[C]; break;
      case 3: hdr[C] = o.header[C - r.s.c]; break;
      default:
        if(val == null) continue;
        // 書き換え開始
        vv = v = this.convert_value(this.format_cell(val));
        // 書き換え終了
        let counter = 0;
        for(let CC = 0; CC < hdr.length; ++CC) if(hdr[CC] == vv) vv = v + "_" + (++counter);
        hdr[C] = vv;
    }
  }

  for (R = r.s.r + offset; R <= r.e.r; ++R) {
    rr = this.encode_row(R);
    isempty = true;
    if(header === 1) row = [];
    else {
      row = {};
      if(Object.defineProperty) try { Object.defineProperty(row, '__rowNum__', {value:R, enumerable:false}); } catch(e) { row.__rowNum__
 = R; }
      else row.__rowNum__ = R;
    }
    for (C = r.s.c; C <= r.e.c; ++C) {
      val = sheet[cols[C] + rr];
      // 書き換え開始
      // 見つからない場合はnullを入れてからskip
      if(val === undefined || val.t === undefined) {
        row[hdr[C]] = null;
        continue;
      }
      // 書き換え終了
      v = val.v;
      switch(val.t){
        case 'z': continue;
        case 'e': continue;
        case 's': case 'd': case 'b': case 'n': break;
        default: throw new Error('unrecognized type ' + val.t);
      }
      if(v !== undefined) {
        // 書き換え開始
        _value = this.convert_value(raw ? v : this.format_cell(val,v));
        row[hdr[C]] = _value;
        // 書き換え終了
        isempty = false;
      }
    }
    if(isempty === false || header === 1) out[outi++] = row;
  }
  out.length = outi;
  return out;
}

// yaml変換
xlsx.utils.sheet_to_yaml = function (sheet, opts) {
  let sheet_json = this.sheet_to_all_json(sheet, opts);
  return jsyaml.safeDump(sheet_json);
}

// 値変換
xlsx.utils.convert_value = function (value) {
  if (isNaN(value)) {
    return value.replace(/\\n/g, '\n').replace(/\r/g, '');
  }
  return value;
}

// 行取得(左から見ていき,値がない場合は探索を打ち切る)
xlsx.utils.fetch_rows = function (sheet, row_index) {
  let row_num = this.encode_row(row_index)
  let cell_object = this.decode_cell("A" + row_num);
  let rows = [];
  while(true) {
    let cell = sheet[this.encode_cell(cell_object)];
    let value = this.format_cell(cell);
    if (!value.length) break;
    rows.push(value);
    cell_object["c"] += 1;
  };
  return rows;
}

// 読み取り範囲をきめる
xlsx.utils.table_range = function (sheet, field_row) {
  let range = this.decode_range(sheet["!ref"]);
  if (field_row) {
    field_cell = this.fetch_rows(sheet, field_row).length
    if (field_cell > 0) {
      range.s.r = field_row;
      range.e.c = field_cell;
    }
  }
  return range;
}

module.exports = xlsx;
