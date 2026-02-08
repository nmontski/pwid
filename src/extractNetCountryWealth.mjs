/*

mnninci999  (=) net national income
mndproi999      (+) net domestic product
mgdproi999          (+) gross domestic product
mconfci999          (-) consumption of fixed capital 
mnnfini999      (+) foreign income

---------------------------------------------------

code	      Description	                                                 SNA code / sector
mnninci999	(=) net national income	                                            B5n	S1
mprigoi999	    (+) net primary income of the general government	              B5n	S13
mprihni999	    (+) net primary income of households and non-profits	          B5n	S14 + S15
mprihoi999	        (+) net primary income of households	                      B5n	S14
mprinpi999	        (+) net primary income of non-profits	                      B5n	S15
mpricoi999	    (+) net primary income of corporations	                        B5n	S11 + S12
mprinfi999	        (+) net primary income of non-financial corporations	      B5n	S11
mprifci999	        (+) net primary income of financial corporations	          B5n	S12

----------------------------------------------------

code	          Description	                                              SNA code	SNA sector
mnninci999	(=) net national income	B5n	S1
msecgoi999	    (+) net secondary income of the general government	            B6n + D7	S13
msechni999	    (+) net secondary income of households and non-profits	        B6n + D7	S14 + S15
msechoi999	        (+) net secondary income of households	                    B6n + D7	S14
msecnpi999	        (+) net secondary income of non-profits	                    B6n + D7	S15
mseccoi999	    (+) net secondary income of corporations	                      B6n + D7	S11 + S12
msecnfi999	        (+) net secondary income of non-financial corporations	    B6n + D7	S11
msecfci999	        (+) net secondary income of financial corporations	        B6n + D7	S12

----------------------------------------------------

code	      Description	                                              SNA code	SNA sector
nninc	(=) net national income	                                              B5n	S1
comhn	    (+) compensation of employees	                                D1	S14 + S15
fkpin	    (+) net capital income		
prphn	        (+) property income distributed to households and NPISH	  D4	S14 + S15
prpgo	        (+) property income distributed to the government	          D4	S13
nsrhn	        (+) net operating surplus of the households and NPISH	      B2n	S14 + S15
prico	        (+) net primary income of corporations	                  B5n	S11 + S12
nmxho	    (+) net mixed income of households	                          B3n	S14
ptxgo	    (+) taxes on products and production	                        D2 – D3	S13

*/

import { MongoClient } from "mongodb"
// import { formatDecimal } from "basikon-common-utils"
import fs from "fs";
import xlsx from "xlsx";
import { parse, transform, stringify } from "csv";

const euroCountries = [
  { name: "Austria", code: "AT" },
  { name: "Belgium", code: "BE" },
  { name: "Bulgaria", code: "BG" },
  { name: "Croatia", code: "HR" },
  { name: "Cyprus", code: "CY" },
  { name: "Estonia", code: "EE" },
  { name: "Finland", code: "FI" },
  { name: "France", code: "FR" },
  { name: "Germany", code: "DE" },
  { name: "Greece", code: "GR" },
  { name: "Ireland", code: "IE" },
  { name: "Italy", code: "IT" },
  { name: "Latvia", code: "LV" },
  // { name: "Lithuania", code: "LR" },
  { name: "Luxembourg", code: "LU" },
  { name: "Malta", code: "MT" },
  { name: "Netherlands", code: "NL" },
  { name: "Portugal", code: "PT" },
  { name: "Slovenia", code: "SI" },
  { name: "Slovakia", code: "SK" },
  { name: "Spain", code: "ES" },
]

const countries = [
  ...euroCountries, 
  { name: "Eurozone", code: "EU" },
  { name: "USA", code: "US" },
  { name: "Canada", code: "CA" },
]

const client = new MongoClient("mongodb://localhost:27017");
const database = client.db('pudd');
const dataCollection = database.collection('data');

/*
code explanation: https://wid.world/codes-dictionary/#one-letter-code 
[Net public wealth] = [Government non-financial assets] + [Government financial assets] - [Government debt]
[Net private wealth] = [Private non-financial assets] + [Private financial assets] - [Private debt]

*/

async function convertCsvSemicolumns(fileNoExtension) {
    fs
      .createReadStream(`data/wid_all_data/${fileNoExtension}.csv`)
      .pipe(parse({ delimiter: ";", quote: '"' }))
      // .pipe(transform(record => {
        // if (first <= 0) {
          // // do something with header
          // record = migrateHeader(mappings, record)
          // // append (or insert) fields here
          // // ...
        // } else {
          // // do something with record
          // record = migrateRecord(mappings, record)
          // // append (or insert) fields here
          // // ...
        // }
        // first++
        // return record
      // }))
      .pipe(stringify())
      .pipe(fs.createWriteStream(`data/${fileNoExtension}_with_commas.csv`))
}

/*
public: owned by the general government sector (central government, state government, local government, and social security funds)
private: held by private owners (households and foundations)
personal: held by households
non-profit: held by non-profit institutions
=> "private" is probably "personal + non-profit"
*/

const dataDefs = [
  // { field: "year", code: null },
  { field: "gdp", code: "mgdproi999" },
  { field: "netTotalWealth", code: null, formula: "netPublicWealth + netPrivateWealth" },
  { field: "netPublicWealth", code: "mgweali999" },
  { field: "netPrivateWealth", code: "mpweali999" },
  { field: "netResidualWealth", code: "mcwresi999" },
  { field: "netPersonalWealth", code: "mhweali999" },
  { field: "netNonProfitWealth", code: "miweali999" },
  { field: "netMarketValueWealth", code: "mnweali999" },
  { field: "netBookValueWealth", code: "mnwbooi999" },
  { field: "netNationalIncome", code: "mnninci999" },
  { field: "netGovernmentPrimaryIncome", code: "mprigoi999" },
  { field: "netLaborPrimaryIncome", code: "mprihni999" },
  { field: "netCorporatePrimaryIncome", code: "mpricoi999" },
  { field: "netNonFinancialPrimaryIncome", code: "mprinfi999" },
  { field: "netFinancialPrimaryIncome", code: "mprifci999" },
  { field: "netCorporateSecondaryIncome", code: "mseccoi999" },
  { field: "netDomesticProduct", code: "mndproi999" },
  { field: "netForeignIncome", code: "mnnfini999" },
  { field: "netSalaries", code: "mcomhni999" },
  { field: "netCapitalIncome", code: "mfkpini999" },
  { field: "netMixedIncome", code: "mnmxhoi999" },
  { field: "netTaxesOnProd", code: "mptxgoi999" },
  { field: "priceIndex", code: "inyixxi999" },
  { field: "inflation", formula: "priceIndex / priceIndex-1 - 1", format: "0.00%" },
  { field: "population", code: "npopuli999", format: "#,###" },
  { field: "populationOver20", code: "npopuli992", format: "#,###" },
  { field: "populationUpTo20", code: null, formula: "population - populationOver20", format: "#,###" },
  { field: "capitalRevenue", code: null, formula: "0.3 * netMixedIncome + netCapitalIncome" },
  { field: "capitalRevenuePercentage", code: null, formula: "capitalRevenue / netMarketValueWealth", format: "0.00%" },
  { field: "gdpGrowth", code: null, formula: "gdp / gdp-1 - 1", format: "0.00%" },
  { field: "nationalIncomeGrowth", code: null, formula: "netNationalIncome / netNationalIncome-1 - 1", format: "0.00%" },
  { field: "overduePercentage", code: null, formula: "capitalRevenuePercentage - gdpGrowth", format: "0.00%" },
  { field: "ddEnvelope", code: null, formula: "netMarketValueWealth * overduePercentage" },
  { field: "monthlyDd", code: null, formula: "ddEnvelope * 1000000000 / population / 12" },
]
  
async function makeNetCountryWealthCsv(country) {
  
  const dataDefsObj = {}
  dataDefs.forEach((dataDef, index) => dataDefsObj[dataDef.field] = dataDef)
  
  dataDefs.forEach((dataDef, index) => {
    dataDef.lineNumber = index + 2
    if (dataDef.formula) {
      const tokens = dataDef.formula.split(" ") 
      
      // this converts formula of a line to Excel formula based on cell ref
      dataDef.getExcelFormula = columnName => {
        const excelFormulaItems = []
        for (let token of tokens) {
          if (token.endsWith("-1")) {
            if (columnName === "C") return undefined
            token = token.substring(0, token.length - 2)
            const columnNumber = xlsx.utils.decode_col(columnName)
            const minusOneColumn = xlsx.utils.encode_col(columnNumber - 1)
            // if (columnName === "AA") console.log(columnName, columnNumber, minusOneColumn)
            excelFormulaItems.push(minusOneColumn + dataDefsObj[token].lineNumber)
          } else if (dataDefsObj[token]) {
            excelFormulaItems.push(columnName + dataDefsObj[token].lineNumber)
          } else {
            excelFormulaItems.push(token)
          }
        }
        return "=" + excelFormulaItems.join("")
      }
      
      // this executes a formula for one cell
      dataDef.executeFormula = (yearsObj, columnIndex) => {
        const stack = []
        for (let token of tokens) {
          let isMinusOne = false
          if (token.endsWith("-1")) {
            if (columnIndex === 0) return undefined
            token = token.substring(0, token.length - 2)
            isMinusOne = true
          }
          let value
          let isValue = false
          if (dataDefsObj[token]) {
            // console.log(columnIndex, yearsObj)
            value = yearsObj[isMinusOne ? columnIndex - 1 : columnIndex][token]?.v
            isValue = true
          } else {
            value = parseFloat(token)
            isValue = !isNaN(value)
          }
          if (isValue) {
            if (typeof stack.at(-1) === "function") {
              // console.log(stack, value)
              const func = stack.pop()
              const a = stack.pop()
              let value2 = func(a, value)
              // console.log("=>", value2)
              stack.push(value2)
            } else {
              stack.push(value)
            }
          } else {
            let func = {
              "+": (a, b) => a + b,
              "-": (a, b) => { /*console.log("---", a, b);*/ return a - b },
              "*": (a, b) => a * b,
              "/": (a, b) => a / b,
            }[token]
            if (!func) throw Error("Invalid operation " + token) 
            stack.push(func)
          }
        }
        // console.log(dataDef.field, columnIndex, stack[0])
        return stack[0]
      }
      
    }
  })
  
  // console.log(dataDefsObj["MonthlyDd"].getExcelFormula("E"))
  // console.log(dataDefsObj["gdpGrowth"].getExcelFormula("ZZ"))
  // for (let i = 0; i < 28; i++) {
    // console.log("----", i, "=>", getExcelcolumnName(i))
  // }
  // return
  
  const startYear = 1995
  
  // get data from Mongo
  const collectionData = {}
  for (let dataDef of dataDefs.filter(it => it.code)) {
    if (country === "EU") {
      const getPipeline = variable => [
        { 
          $match: { 
            variable, 
            country: { $in: euroCountries.map(it => it.code) }, 
            year: { $gte: startYear } 
          },
        },
        { 
          $group: {
            _id: "$year",
            value: { $sum: "$value" }, 
          },
        },
        {
          $project: {
            country: "EU",
            variable,
            year: "$_id",
            value: 1,
          }
        },
        { $sort: { year: 1 } },
      ]
      collectionData[dataDef.field] = await dataCollection.aggregate(getPipeline(dataDef.code)).toArray()
      // console.log("hhh0", collectionData[dataDef.field][0])

      // priceIndex does not sum, we need to compute the barycenter according to gdp => let's load gdp data also
      if (dataDef.field === "priceIndex") {
        const priceIndexPerCountryObj = {}
        const gdpPerCountryObj = {}
        for (let { code: country } of euroCountries) {
          priceIndexPerCountryObj[country] = await dataCollection.find({ variable: dataDefsObj["priceIndex"].code, country, year: { $gte: startYear } }).sort({ year: 1 }).toArray()
          gdpPerCountryObj[country] = await dataCollection.find({ variable: dataDefsObj["gdp"].code, country, year: { $gte: startYear } }).sort({ year: 1 }).toArray()
        }
        for (let i = 0; i < priceIndexPerCountryObj["FR"].length; i++) {
          let barycenter = 0
          let sum = 0
          for (let { code: country } of euroCountries) {
            const coeff = gdpPerCountryObj[country][i]?.value
            // console.log("hhh", country)
            const priceIndex = priceIndexPerCountryObj[country][i]?.value
            // if (i === 29) console.log("hhh2", country, priceIndex)
            sum += coeff
            barycenter += coeff * priceIndex
          }
          // if (i === 29) console.log("hhh3", barycenter, sum)
          barycenter /= sum
          if (!isNaN(barycenter)) collectionData[dataDef.field][i].value = barycenter
        }
      }      

    } else {
      collectionData[dataDef.field] = await dataCollection.find({ variable: dataDef.code, country, year: { $gte: startYear } }).sort({ year: 1 }).toArray()
      // console.log("hhh1", collectionData[dataDef.field][0])
    }
  }
  
  // console.log(collectionData)
  
  const yearsObj = {}
  for (let key of Object.keys(collectionData)) {
    const isBillions = key.startsWith("net") || key === "gdp"
    for (let it of collectionData[key]) {
      yearsObj[it.year] ||= {}
      yearsObj[it.year][key] = { v: isBillions ? it.value / 1000000000 : it.value }
    }
  }
  
  // for (let y in yearsObj) {
    // const year = yearsObj[y]
    // year.netTotalWealth = year.netPublicWealth + year.netPrivateWealth
  // }
  // for (let y in yearsObj) {
    // const year = yearsObj[y]
    // year.populationUpTo20 = year.population - year.populationOver20
  // }
  
  const years = []
  const yearStrings = Object.keys(yearsObj)
  let prevY
  for (let [index, y] of yearStrings.entries()) {
    const year = yearsObj[y]
    years.push(year)
    year.year = { v: y.toString() }
    // // year.netTotalWealthB = year.netTotalWealth
    // if (year.netCapitalIncome && year.netMixedIncome) {
      // year.capitalRevenue = year.netCapitalIncome + 0.3 * year.netMixedIncome
      // if (year.netMarketValueWealth) {
        // year.capitalRevenuePercentage = year.capitalRevenue / year.netMarketValueWealth
      // }
      // if (prevY) {
        // const prevYear = yearsObj[prevY]
        // if (year.netNationalIncome && prevYear.netNationalIncome) {
        // }
      // }
    // }
    for (let { field, code, getExcelFormula, executeFormula, format } of dataDefs) {
        // console.log("field", field, getExcelFormula)
      if (getExcelFormula) {
        const columnIndex = parseInt(y) - startYear
        const columnName = xlsx.utils.encode_col(2 + columnIndex)
        // console.log("columnName", y, 2 + columnIndex, columnName)
        const f = getExcelFormula(columnName)
        const v = executeFormula(years, columnIndex)
        year[field] = { f, v, t: "n", z: format || "0.00" }
      } else {
        year[field] = { v: year[field]?.v, t: "n", z: format || "0.00" }
      }
      
    }
    prevY = y
  }
  
  // check average
  let sum = 0
  let n = 0
  for (let y = 1996; y < 2023; y++) {
    const columnIndex = y - startYear
    sum += yearsObj[yearStrings[columnIndex]].monthlyDd?.v || 0
  // console.log("hhh", yearStrings[columnIndex], yearsObj[yearStrings[columnIndex]].monthlyDd?.v) 
    n++
  }
  const average = sum / n
  console.log(country, average) 
  
  return yearsObj

  // const csvFileStream = fs.createWriteStream(`./DD-${country}.csv`);
  // for (let { field } of dataDefs) {
    // const code = dataDefs.find(it => it.field === field)?.code || ""
    // await csvFileStream.write(field + ",")
    // await csvFileStream.write(code + ",")
    // for (let i = 0; i < years.length; i++) {
      // const year = years[i]
      // // console.log("hhh", i, field, year[field])
      // await csvFileStream.write((year[field] || "").toString())
      // if (i < years.length) {
        // await csvFileStream.write(",")
      // }
    // }
    // await csvFileStream.write("\n")
  // }
  // csvFileStream.close()
}

function yearsObjToXlsxData(yearsObj) {
  const yearStrings = Object.keys(yearsObj)
  const xlsxData = []
  xlsxData.push(["name", "code", ...yearStrings])
  for (let { field, code } of dataDefs) {
    xlsxData.push([field, code, ...yearStrings.map(y => {
      const cell = yearsObj[y][field]
      if (cell.f) {
        const cell2 = { ...cell }
        delete cell2.v
        return cell2
      } else {
        return cell
      }
    })])
  }
  return xlsxData
}

async function saveExcel(countryDataObj) {
  
  const workbook = xlsx.utils.book_new()
  for (let country in countryDataObj) {
    const xlsxData = yearsObjToXlsxData(countryDataObj[country])
  
    const worksheet = xlsx.utils.aoa_to_sheet(xlsxData, { cellDates: true })

    worksheet['!views'] = [
      {
        state: 'frozen',
        xSplit: 2, // Freeze the first column (1-based index)
        ySplit: 1, // Freeze the first row (1-based index)
      }
    ];
    
    worksheet['!cols'] = [
      { wch: 28 }, // Width for column A
      { wch: 12 }, // Width for column B
      // { wch: 25 }  // Width for column C
    ];
  
    xlsx.utils.book_append_sheet(workbook, worksheet, country)
  }
  await xlsx.writeFile(workbook, `DD.xlsx`)
}  

async function buildExcelFile() {
  const countryDataObj = {}
  for (let { code } of countries) {
    countryDataObj[code] = await makeNetCountryWealthCsv(code)
  }
  await saveExcel(countryDataObj)
}

async function convertAllCsvSemicolumns() {
  for (let { code } of countries) {
    await convertCsvSemicolumns(`wid_data_${code}`)
  // await convertCsvSemicolumns("wid_metadata_FR")
  }
}

await buildExcelFile()
// await convertAllCsvSemicolumns()

await client.close()
