const xlsx = require('xlsx')
const math = require('mathjs')

const workbook = xlsx.readFile('./Book2.xlsx')

const worksheet = workbook.Sheets['New Data']



const loads = xlsx.utils.sheet_to_json(worksheet)
console.log(loads)
let voltage_i = math.complex(1000,0);

for (let i = 0; i < loads.length; i++) {
    const z_i = math.complex(loads[i].r, loads[i].x); 
    const current_i = math.complex(loads[i].ir,loads[i].ii); 
    const voltage_i_1 = math.subtract(voltage_i, math.multiply(z_i, current_i)); 
   
    voltage_i = voltage_i_1; 
    loads[i].v_1 = `${voltage_i_1.toPolar().r}∠${voltage_i_1.toPolar().phi*(3.14/180)}`
    console.log(`${voltage_i_1.toPolar().r}∠${voltage_i_1.toPolar().phi}`)
}



const newWorkbook = xlsx.utils.book_new()
const newWorksheet = xlsx.utils.json_to_sheet(loads)
xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, 'New Data')

xlsx.writeFile(newWorkbook, 'new Data File .xlsx')
