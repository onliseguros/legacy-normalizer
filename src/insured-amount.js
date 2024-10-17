import xlsx from 'xlsx';

// Replace with your file name
const fileName = 'AGRUPADO SUBIDA ONLI'

const file = `spreadsheets/${fileName}.xlsx`;
const newFile = `spreadsheets/${fileName}-is-normalizada.xlsx`;

function normalizeInsuredAmount() {
  const workbook = xlsx.readFile(file);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

  // remove pointers
  const normalizedData = JSON.parse(JSON.stringify(data))

  if (data.length < 1) {
    console.log('Not data found.')
    return
  }

  data.forEach((row, index) => {
    // ignore table headers
    if (index === 0) {
      return
    }
    // ignore empty row
    if (row.length < 1) {
      return
    }

    console.log('row:', index + 1)

    const emissionType = row[16].trim().toLowerCase() // Q
    const externalKey = row[17]  // R
    const hubConnectionId = row[18]  // S

    if (emissionType !== 'endosso') {
      return
    }

    if (!hubConnectionId) {
      console.log('Column S is empty')
      return
    }

    if (!externalKey) {
      console.log('Column R is empty')
      return
    }

    const associatedEmissions = []

    // find policy
    for (const row of data) {
      // ignore table headers
      if (index === 0) {
        continue
      }
      // ignore empty row
      if (row.length < 1) {
        continue
      }

      const emissionType = row[16].trim().toLowerCase() // Q
      const externalKey = row[17]  // R

      if (emissionType !== 'apólice' && emissionType !== 'apolice') {
        continue
      }

      if (externalKey !== hubConnectionId) {
        continue
      }

      associatedEmissions.push(row)
      break
    }

    // find associated endorsements
    for (const row of data) {
      // ignore table headers
      if (index === 0) {
        continue
      }
      // ignore empty row
      if (row.length < 1) {
        continue
      }

      const emissionType = row[16].trim().toLowerCase() // Q
      const hubConnectionId2 = row[18]  // S

      if (emissionType !== 'endosso') {
        continue
      }

      if (hubConnectionId2 !== hubConnectionId) {
        continue
      }

      associatedEmissions.push(row)
    }

    // sort by date
    associatedEmissions.sort((a, b) => {
      const dateA = a[21]
      const dateB = b[21]

      if (dateA > dateB) {
        return 1
      } else if (dateA < dateB) {
        return -1
      } else {
        return 0
      }
    })

    let insuredAmount = 0
    for (let i = 0; i < associatedEmissions.length; i++) {
      const emission = associatedEmissions[i]

      const externalKey2 = emission[17] // R
      if (externalKey2 !== externalKey) {
        continue
      }

      // now you have the index of the current endorsement
      for (let j = 0; j <= i; j++) {
        const amount = associatedEmissions[j][2]
        if (typeof amount === 'number') {
          insuredAmount += amount
        }
      }

      break
    }

    normalizedData[index][2] = insuredAmount
  });

  // convert excel dates to JS dates
  normalizedData.forEach((row, index) => {
    // ignore table headers
    if (index === 0) {
      return
    }
    // ignore empty row
    if (row.length < 1) {
      return
    }

    const lifetimeStart = row[5]
    const lifetimeEnd = row[6]
    const date = row[21]

    // it is necessary to convert excel date to JS date
    if (typeof lifetimeStart === 'number') {
      normalizedData[index][5] = convertDateFromExcelToJS(lifetimeStart)
    } else {
      console.log('Lifetime start is not a number')
    }
    if (typeof lifetimeEnd === 'number') {
      normalizedData[index][6] = convertDateFromExcelToJS(lifetimeEnd)
    } else {
      console.log('Lifetime end is not a number')
    }
    if (typeof date === 'number') {
      normalizedData[index][21] = convertDateFromExcelToJS(date)
    } else {
      console.log('Date is not a number')
    }
  })

  // create a new .xlsx file
  const newWorkbook = xlsx.utils.book_new()
  const newWorksheet = xlsx.utils.aoa_to_sheet(normalizedData);
  xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, 'IS normalizada');
  xlsx.writeFile(newWorkbook, newFile);
}

function convertDateFromExcelToJS(excelDate) {
  return new Date((excelDate - 25568) * 86400 * 1000);
}

normalizeInsuredAmount()