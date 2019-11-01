const readXlsxFile = require('read-excel-file/node');
const FILE = './sample.xlsx';
let staffResults = new Map();
let rows = [];
let sheetsArr;
let date;

async function processExcel() {
    let promises = [];

    await new Promise((resolve, reject) => {
        readXlsxFile(FILE, { getSheets: true }).then((sheets) => {
            // sheets === [{ name: 'Sheet1' }, { name: 'Sheet2' }]
            sheetsArr = sheets;
            resolve();
        });
    });


    sheetsArr.map(sheet => {
        promises.push(new Promise(async (resolve, reject) => {
            await readXlsxFile(FILE, { sheet: sheet.name }).then((data) => {
                data.shift();
                resolve(data)
            })
        }));
    });

    await Promise.all(promises).then(results => {
        results.forEach(result => {
            rows = rows.concat(result);
        })

    })

    rows.forEach((row, index) => {

        if (index % 2 == 0) {
            date = row[1] ? row[1] : date;
        }
        else {
            row.forEach(col => {
                if (col) {
                    staffsArr = `${col}`.split('\r\n');
                    staffsArr.forEach(staff => {

                        let staffPreviousObj = staffResults.get(staff);
                        let a = staffObj(staffPreviousObj, staff, date);

                        staffResults.set(staff, a);

                    })
                }

            })
        }
    })

    console.log('Final Result::', staffResults)


}


function staffObj(currentStaffObj, staffName, leaveDate) {
    let name = staffName;
    let count = 1;
    let leave = [leaveDate]
    if (currentStaffObj) {
        count = currentStaffObj.count + 1;
        leave = leave.concat(currentStaffObj.leave);
    }

    return {
        name,
        count,
        leave
    }
}

processExcel();