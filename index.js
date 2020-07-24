const XLSX = require('xlsx')
let template = require('./template.js')

// read the feed file and output file using xlsx 
const feedFile = XLSX.readFile(`${__dirname}/resources/FeedFile.xlsx`)
// const outputPath = XLSX.readFile(`${__dirname}\\resources\\OutputFile.xlsx`)

let outputSheetColumnHeaders = []
let feedSheetColumnHeaders = []

// read specific sheets
// const outputSheet = outputFile.Sheets['Output']
const feedSheet = feedFile.Sheets['Feedfile']

// convert feed file to json
const feedData = XLSX.utils.sheet_to_json(feedSheet)

// output template consumed as json from template.js file
let outputTemp = template

// retrieve all the column names of output by fetching json keys of outputTemp
const outputSheetValues = Object.keys(outputTemp)
outputSheetValues.forEach(key => {
    outputSheetColumnHeaders.push(key)
})

// retrieve all the column names of feed file by fetching json keys of feedData

const feedSheetValues = Object.keys(feedData[0])
feedSheetValues.forEach(key => {
    feedSheetColumnHeaders.push(key)
})

// logic to look for matching column names between sheets and extract the column headers
let matchingOutputHeaders = []
let matchingFeedHeaders = []

for (let i = 0; i < outputSheetColumnHeaders.length; i++) {

    feedSheetColumnHeaders.forEach((header, index) => {

        let searchString = outputSheetColumnHeaders[i]

        // regex to replace underscore and merge words
        regExSearchString = searchString.replace(/_/g, "")

        if (header.includes(regExSearchString)) {

            // console.log(`Feedsheet -> ${header} : OutputSheet -> ${searchString}`)

            // store the column headers in array
            matchingOutputHeaders.push(searchString)
            matchingFeedHeaders.push(header)

        }

    })
}

// construct the output json data

let newOutputData = []

feedData.forEach((feed, index) => {
    matchingOutputHeaders.forEach((header, headindex) => {
        outputTemp.SR_NO = index + 1
        outputTemp[header] = feed[matchingFeedHeaders[headindex]]

    })
    const temp = JSON.stringify(outputTemp)
    newOutputData.push(JSON.parse(temp))

})

console.log(newOutputData)

// write the json data to excel

let outputWB = XLSX.utils.book_new()
let outputWS = XLSX.utils.json_to_sheet(newOutputData)
XLSX.utils.book_append_sheet(outputWB, outputWS, 'Output')
XLSX.writeFile(outputWB, `${__dirname}/output/OutputFile.xlsx`)