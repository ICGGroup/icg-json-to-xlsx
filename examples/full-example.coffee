path = require('path')
jsonXlsx = require('../lib/icg-json-to-xlsx')

jsonData = [
  {"IsMember" : true, "Name" : "John", "Age" : 24},
  {"IsMember" : false, "Name" : "Paul", "Age" : 44},
  {"IsMember" : true, "Name" : "George", "Age" : 12}
]


filename = path.join(__dirname, "full-example.xlsx")
headers = ["Is User Member?", "First Name", "Age"]

outputFile = jsonXlsx.writeFile(filename, jsonData, {headers:headers, sheetName:"Roster"})

console.log outputFile
