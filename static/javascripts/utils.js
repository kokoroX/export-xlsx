/* eslint-disable no-undef*/
function datenum (v, date1904) {
  if (date1904) v += 1462
  var epoch = Date.parse(v)
  return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000)
}

function sheetFromArrayOfArrays (data, opts) {
  var ws = {}
  var range = {s: {c: 10000000, r: 10000000}, e: { c: 0, r: 0 }}
  for (var R = 0; R !== data.length; ++R) {
    for (var C = 0; C !== data[R].length; ++C) {
      if (range.s.r > R) range.s.r = R
      if (range.s.c > C) range.s.c = C
      if (range.e.r < R) range.e.r = R
      if (range.e.c < C) range.e.c = C
      var cell = { v: data[R][C] }
      if (cell.v == null) continue
      var cellRef = XLSX.utils.encode_cell({ c: C, r: R })

      if (typeof cell.v === 'number') cell.t = 'n'
      else if (typeof cell.v === 'boolean') cell.t = 'b'
      else if (cell.v instanceof Date) {
        cell.t = 'n'
        cell.z = XLSX.SSF._table[14]
        cell.v = datenum(cell.v)
      } else {
        cell.t = 's'
      }
      ws[cellRef] = cell
    }
  }
  if (range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range)
  return ws
}

function Workbook () {
  if (!(this instanceof Workbook)) return new Workbook()
  this.SheetNames = []
  this.Sheets = {}
}

function s2ab (s) {
  var buf = new ArrayBuffer(s.length)
  var view = new Uint8Array(buf)
  for (var i = 0; i !== s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF
  return buf
}

/**
 * generate xlsx file, and download it.
 * @method generateAndDownloadXLSX
 * @param  {[type]}                data     =             []     [用于导出的数据列表]
 * @param  {[type]}                columns  =             []     [需要导出的数据key { name: 'colName' }格式]
 * @param  {[type]}                xlsxName =             '失败数据' [表名]
 * @return {[type]}                         [description]
 * $author kokoro
 * date 2016-07-25
 */
function generateAndDownloadXLSX (data = [], columns = [], xlsxName = '失败数据') {
  const array = []
  let firstRow = []
  firstRow = columns.map(grid => grid.name)
  let dataList = data.map(line => firstRow.map(colName => line[colName]))
  array.push(firstRow)
  array.push.apply(array, dataList)

  var ws = sheetFromArrayOfArrays(array)

  var workbook = new Workbook()
  workbook.SheetNames.push(xlsxName)
  workbook.Sheets[xlsxName] = ws

  var wbout = XLSX.write(workbook, {
    bookType: 'xlsx',
    bookSST: true,
    type: 'binary'
  })

  saveAs(new Blob([s2ab(wbout)], { type: '' }), `${xlsxName}.xlsx`)
}
