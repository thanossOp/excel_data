const e = require('express')
const express = require('express')
const app = express()
const xlsx = require('xlsx')

// create excel file
const excel = xlsx.utils.book_new();

const excelsheet = xlsx.utils.json_to_sheet([],{header : ['Name' , 'Age' ,'Gender']})

xlsx.utils.book_append_sheet(excel,excelsheet,'Sheet2');

const data = [
    {Name : 'Nisarg',Age : 22,Gender : 'M'},
    {Name : 'pankaj',Age : 26,Gender : 'M'},
    {Name : 'shrutayy',Age : 21,Gender : 'F'},
    {Name : 'srushti',Age : 24,Gender : 'F'},
]

xlsx.utils.sheet_add_json(excelsheet,data,{origin : 1 ,skipHeader : true})


//read excel file
// const excel = xlsx.readFile('./excel_file/user.xlsx')
// const excelsheet = excel.Sheets['Sheet2']

// const data = xlsx.utils.sheet_to_json(excelsheet)

// console.log(data)


// //add new data

const datatoadd = [
    ['Avneet', 22, 'F'],
    ['John', 30, 'M'],
    ['Sarah', 25, 'F'],
    ['Michael', 40, 'M']
  ];

xlsx.utils.sheet_add_aoa(excelsheet,datatoadd,{origin : -1})

const newData = xlsx.utils.sheet_to_json(excelsheet,{header : 1})

const updateddata = xlsx.utils.aoa_to_sheet(newData)

excel.Sheets['Sheet2']=updateddata

// const newData = xlsx.
// update any cell value

// excelsheet['A8'].v = "Roshni"

//delete row

// const deleterow = 7;
// const range = xlsx.utils.decode_range(excelsheet['!ref'])


// if(deleterow <= range.e.r){

//     const rows = range.e.r - range.s.r +1;

//     const sheetdata = xlsx.utils.sheet_to_json(excelsheet,{header : 1})
//     // console.log(sheetdata)

//     if(deleterow >= -1 && deleterow<= rows){
//         const deletedrow = sheetdata.splice(deleterow,1)
//         console.log(deletedrow)

//     }

//     const newData = xlsx.utils.aoa_to_sheet(sheetdata)

//     excel.Sheets['Sheet2'] = newData
// }else{
//     console.log('Row is not found')
// }



//delete cell

// const range = xlsx.utils.decode_range(excelsheet['!ref'])
// const celladdress = xlsx.utils.encode_cell({r:4,c:1})

// if(excelsheet[celladdress]){
//     console.log('deleted cell :' ,excelsheet[celladdress].v)
//     delete excelsheet[celladdress];
// }else{
//     console.log("cell is not exist")
// }

//delete column

// const range = xlsx.utils.decode_range(excelsheet['!ref'])

// const deletecolumn = 2

// if (deletecolumn <= range.e.c) {

//     const column = range.e.c - range.s.c + 1

//     const sheetdata = xlsx.utils.sheet_to_json(excelsheet, { header: 1 })

//     // console.log(sheetdata)

//     if (deletecolumn >= -1 && deletecolumn <= column) {
//         sheetdata.forEach(row => row.splice(deletecolumn, 1))
//     }

//     const newData = xlsx.utils.aoa_to_sheet(sheetdata)

//     excel.Sheets['Sheet2'] = newData
// }
// else{
//     console.log('column not found')
// }

xlsx.writeFile(excel, './excel_file/user.xlsx')
