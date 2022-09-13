const express = require('express')
const fs = require('fs')
const xlsx = require('xlsx')
const app = express()
const exel = xlsx.readFile('conso_mix_RTE_2022.ods')
let sheet = []
/* api qui parse l'exel on json  */
app.get('/getJson', (req, res) => {
    for (const sheetName of exel.SheetNames) {
        sheet = xlsx.utils.sheet_to_json(exel.Sheets[sheetName], { header: ['date', 'PrevisionJ-1', 'PrevisionJ'] })
        let date = ''
        let sheet2 = []
        sheet.map((el) => {
            /* pour avoir le date */
            if (el.date?.includes(' du ')) {
                date = el.date.substring(el.date.length - 10)
            }
            /* création d'un nouvel array */
            if (date.length > 0 && isNaN(el.PrevisionJ) === false) {
                sheet2 = [...sheet2, { date: date + ' ' + el.date, 'PrevisionJ-1': el['PrevisionJ-1'], PrevisionJ: el.PrevisionJ }]
            }
        })
        /* le code de création de nouvel exel se fait ici */
        res.status(200).send(sheet2)
    }
})

app.listen(3000, () => {
    console.log('app is runnign in 3000')
})