fs = require('fs')
path = require('path')
jsonXlsx = require('../lib/icg-json-to-xlsx')

jsonData = [
  {"IsMember" : true, "Name" : "John", "Age" : 24},
  {"IsMember" : false, "Name" : "Paul", "Age" : 44},
  {"IsMember" : true, "Name" : "George", "Age" : 12}
]


filename = path.join(__dirname, "basic-sheet-via-buffer-output.xlsx")

buffer = jsonXlsx.writeBuffer(jsonData)
if buffer
  fs.writeFile filename, buffer, (err)->
    console.log filename
