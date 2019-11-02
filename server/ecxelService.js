const readXlsxFile = require('read-excel-file/node');
const FILE = './sample.xlsx';


class ExcelService {

    async processExcel() {
        let staffResults = new Map();
        let rows = [];
        let sheetsArr;
        let date;
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
                        let staffsArr = `${col}`.split('\r\n');
                        staffsArr.forEach(staff => {

                            let staffPreviousObj = staffResults.get(staff);
                            let a = this._staffObj(staffPreviousObj, staff, date);

                            staffResults.set(staff, a);

                        })
                    }

                })
            }
        })
        return this.strMapToObj(staffResults)
    }

    strMapToObj(strMap) {
        let obj = Object.create(null);
        for (let [k, v] of strMap) {
            // We don’t escape the key '__proto__'
            // which can cause problems on older engines
            obj[k] = v;
        }
        return obj;
    }

    _staffObj(currentStaffObj, staffName, leaveDate) {
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
}


module.exports = ExcelService;