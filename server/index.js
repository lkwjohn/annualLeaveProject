const express = require('express')
const app = express()
const port = 2000
const ExcelService = require('./ecxelService');

app.get('/excel', async (req, res) => {
    try {
        let ecxelService = new ExcelService();
        let result = await ecxelService.processExcel();
        res.header("Access-Control-Allow-Origin", "*");
        res.send(result);
    }
    catch (err) {
        res.header("Access-Control-Allow-Origin", "*");
        res.status(500);
        res.send(err);
    }

})

app.listen(port, () => console.log(`Example app listening on port ${port}!`))