// Requiring the module
const reader = require('xlsx')

// Reading our test file
const file = reader.readFile('./FullRank.xlsx')

let data = []

const sheets = file.SheetNames

for(let i = 0; i < sheets.length; i++)
{
    const temp = reader.utils.sheet_to_json(
            file.Sheets[file.SheetNames[i]])
    temp.forEach((res) => {
        data.push(res)
    })
}



// Reading our test file
const file2 = reader.readFile('./Rank.xlsx')

let data2 = []

const sheets2 = file2.SheetNames

for(let i = 0; i < sheets2.length; i++)
{
    const temp2 = reader.utils.sheet_to_json(
            file2.Sheets[file2.SheetNames[i]])
    temp2.forEach((res) => {
        data2.push(res)
    })
}

// console.log(data2);
// // Printing data
let retData = [];
for (let index = 0; index < data2.length; index++) {
    // const element = data2[index];
    for (let v = 0; v < data.length; v++) {
        if (data[v].RANK == data2[index].RANK) {
            retData.push({"RANK":data2[index].RANK,"ID":data[v].ID})
        }
    }
    
}

// console.log(retData);

const convertToExcel = ()=>{
    
    let ws = reader.utils.json_to_sheet(retData)
    let workbook = reader.utils.book_new()
    reader.utils.book_append_sheet(workbook,ws,"Sheet3")
    //generate Buffer
    reader.write(workbook,{bookType:'xlsx',type:'buffer'});
    //generate binary
    reader.write(workbook,{bookType:'xlsx',type:'binary'});
    reader.writeFile(workbook,"resultData.xlsx")

}
convertToExcel();