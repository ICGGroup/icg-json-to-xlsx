path = require('path')
jsonXlsx = require('../lib/icg-json-to-xlsx')

jsonData = [
  {"IsMember" : true, "Name" : "John", "Age" : 24},
  {"IsMember" : false, "Name" : "Paul", "Age" : 44},
  {"IsMember" : true, "Name" : "George", "Age" : 12}
]


filename = path.join(__dirname, "basic-sheet-output.xlsx")

outputFile = jsonXlsx.writeFile(filename, jsonData)

console.log outputFile
