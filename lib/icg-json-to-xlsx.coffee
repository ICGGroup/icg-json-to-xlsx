fs = require("fs")
path = require("path")
_ = require("lodash")
XLSX = require('xlsx')
isodate = require("isodate")

datenum = (v, date1904) ->
  v += 1462  if date1904
  epoch = Date.parse(v)
  (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000)

findHeaders = (data, lookInFirst) ->
  limit = lookInFirst or 30
  headers = []
  _.each data, (r, i) ->
    for p of r
      headers.push p  unless _.contains(headers, p)
    false  if i is (limit - 1)

  final = _.map headers, (h)->
    # titleCase the headers
    h.replace(/^([a-z])/, (r)-> r.toUpperCase())
     .replace(/\w\_\w/g,(r)-> r[0] + " " + r[2].toUpperCase() )
     .replace(/\w([A-Z])/g, (r)-> r[0] + " " + r[1] )

  final


getAttrs = (data, lookInFirst) ->
  limit = lookInFirst or 30
  attrs = []
  _.each data, (r, i) ->
    for p of r
      attrs.push p  unless _.contains(attrs, p)
    false  if i is (limit - 1)
  attrs

sheetFromJson = (data, opts) ->
  ws = {}
  range =
    s:
      c: 10000000
      r: 10000000

    e:
      c: 0
      r: 0

  opts = opts or {}
  if opts.headers
    headers = opts.headers
  else
    headers = findHeaders(data)
  # derive the headers, based on property
  headerObj = {}
  _.each headers, (h, idx) ->
    attr = opts.attributes[idx]
    if attr
      headerObj[attr] = h
  data.unshift headerObj
  _.each data, (row, R) ->
    C = 0
    console.log(opts)
    _.each opts.attributes, (prop) ->
      val = row[prop] || ""
      range.s.r = R  if range.s.r > R
      range.s.c = C  if range.s.c > C
      range.e.r = R  if range.e.r < R
      range.e.c = C  if range.e.c < C
      cell = v: val
      return true  unless cell.v?
      cell_ref = XLSX.utils.encode_cell(
        c: C
        r: R
      )
      if typeof cell.v is "number" or  typeof cell.v is "'number'"
        cell.t = "n"
      else if typeof cell.v is "boolean"
        cell.t = "b"
      else if cell.v instanceof Date
        cell.t = "n"
        cell.z = XLSX.SSF._table[14]
        cell.v = datenum(cell.v)
      else
        r = /\d{4}\-[01]\d\-[0-3]\dT[0-2]\d\:[0-5]\d\:[0-5]\d([+-][0-2]\d:[0-5]\d|Z)/
        if cell.v.match(r)
          cell.t = "n"
          cell.z = XLSX.SSF._table[14]
          cell.v = datenum(isodate(cell.v))
        else
          cell.t = "s"
      ws[cell_ref] = cell
      C++
      return

    return

  ws["!ref"] = XLSX.utils.encode_range(range)  if range.s.c < 10000000
  ws

# sheetFromArrayOfArrays = (data, opts) ->
#   ws = {}
#   range =
#     s:
#       c: 10000000
#       r: 10000000
#
#     e:
#       c: 0
#       r: 0
#
#   R = 0
#
#   while R isnt data.length
#     C = 0
#
#     while C isnt data[R].length
#       range.s.r = R  if range.s.r > R
#       range.s.c = C  if range.s.c > C
#       range.e.r = R  if range.e.r < R
#       range.e.c = C  if range.e.c < C
#       cell = v: data[R][C]
#       continue  unless cell.v?
#       cell_ref = XLSX.utils.encode_cell(
#         c: C
#         r: R
#       )
#       if typeof cell.v is "number"
#         cell.t = "n"
#       else if typeof cell.v is "boolean"
#         cell.t = "b"
#       else if cell.v instanceof Date
#         cell.t = "n"
#         cell.z = XLSX.SSF._table[14]
#         cell.v = datenum(cell.v)
#       else
#         cell.t = "s"
#       ws[cell_ref] = cell
#       ++C
#     ++R
#   ws["!ref"] = XLSX.utils.encode_range(range)  if range.s.c < 10000000
#   ws
Workbook = ->
  return new Workbook()  unless this instanceof Workbook
  @SheetNames = []
  @Sheets = {}
  return


buildWorkbook = (data, options)->
  options ||= {}
  if not options.headers
    options.headers = findHeaders(data)

  if not options.attributes
    attrs = getAttrs(data)
    console.log(attrs)
    options.attributes = attrs

  if not options.sheetName
    options.sheetName = "Sheet 1"


  opts = {headers:options.headers, attributes:options.attributes}
  wb = new Workbook()
  ws = sheetFromJson(data, opts);

  wb.SheetNames.push(options.sheetName)
  wb.Sheets[options.sheetName] = ws

  wb


module.exports =
  writeFile: (filename, data, options)->

    if not filename or not data
      throw new Error("filename and data parameters are required.")
    else
      wb = buildWorkbook(data, options)
      XLSX.writeFile(wb, filename)

    return filename

  writeBuffer: (data, options)->
    buffer = null
    if not data
      throw new Error("data parameter is required.")
    else
      wb = buildWorkbook(data, options)
      buffer = XLSX.write(wb, {type:"buffer"})

    return buffer
