const xlsx = require('xlsx')
const math = require('mathjs')

const workbook = xlsx.readFile('./Book1.xlsx')

const worksheet = workbook.Sheets['Sheet1']



const loads = xlsx.utils.sheet_to_json(worksheet)


const kva_i = math.complex(loads[loads.length - 1].p, loads[loads.length - 1].q);

const voltage_i = math.complex(loads[loads.length - 1].rv, loads[loads.length - 1].iv);

let current_i = math.divide(kva_i, voltage_i);


loads[loads.length - 1].ir = math.re(current_i)
loads[loads.length - 1].ii = math.im(current_i) * -1



for (let i = loads.length - 1; i >= 2; i--) {

  const kva_i_1 = math.complex(loads[i - 1].p, loads[i - 1].q);
  const voltage_i_1 = math.complex(loads[i - 1].rv, loads[i - 1].iv);
  const current_i_1 = math.divide(kva_i_1, voltage_i_1);
  if (i == loads.length - 1) {
      loads[i-1].ir = math.re(current_i_1)
    loads[i - 1].ii = math.im(current_i_1) * -1
     
  }


  const current_i_2 = math.add(current_i_1, current_i)
  loads[i-2].ir = math.re(current_i_2)
  loads[i-2].ii = math.im(current_i_2)*-1
  current_i = current_i_2
  

}



const newWorkbook = xlsx.utils.book_new()
const newWorksheet = xlsx.utils.json_to_sheet(loads)
xlsx.utils.book_append_sheet(newWorkbook,newWorksheet,'New Data')

xlsx.writeFile(newWorkbook, 'new Data File .xlsx')
