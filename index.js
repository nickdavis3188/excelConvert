// Requiring the module
const reader = require('xlsx')

// Reading our test file
const file = reader.readFile('./Book5.xlsx')

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

console.log(data);
// Reading our test file
const file2 = reader.readFile('./Service.xlsx')

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


// Printing data
let retData = [];
for (let index = 0; index < data2.length; index++) {
    // const element = data2[index];
    for (let v = 0; v < data.length; v++) {
        if (data[v].ServiceName == data2[index].Service) {
            retData.push({"serviceName":data2[index].Service,"serviceId":data[v].ServiceId})
        }
    }
    
}



const convertToExcel = ()=>{
    
    let ws = reader.utils.json_to_sheet(retData)
    let workbook = reader.utils.book_new()
    reader.utils.book_append_sheet(workbook,ws,"Sheet3")
    //generate Buffer
    reader.write(workbook,{bookType:'xlsx',type:'buffer'});
    //generate binary
    reader.write(workbook,{bookType:'xlsx',type:'binary'});
    reader.writeFile(workbook,"mainData.xlsx")

}
convertToExcel();