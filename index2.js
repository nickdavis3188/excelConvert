// Requiring the module
const reader = require('xlsx')

// // Reading our test file
const file = reader.readFile('./FullRank.xlsx')

let rank = []

const sheets = file.SheetNames

for(let i = 0; i < sheets.length; i++)
{
    const temp = reader.utils.sheet_to_json(
            file.Sheets[file.SheetNames[i]])
    temp.forEach((res) => {
        rank.push(res)
    })
}

// console.log(data);
// Reading our test file
const file3 = reader.readFile('./rankData.xlsx')

let ordination = []

const sheets3 = file3.SheetNames

for(let i = 0; i < sheets3.length; i++)
{
    const temp = reader.utils.sheet_to_json(
            file3.Sheets[file3.SheetNames[i]])
    temp.forEach((res) => {
        ordination.push(res)
    })
}

// console.log(ordination);



let ordinationResult = ordination.map(obj => {
  // Extract properties that start with 'Rank_'
  let ranks = Object.keys(obj)
    .filter(key => key.startsWith('RANK'))
    .map(key => obj[key]);

  // Extract properties that start with 'YEAR_'
  let years = Object.keys(obj)
    .filter(key => key.startsWith('YEAR'))
    .map(key => obj[key]);

  // Filter out undefined or null values in the 'years' array
  years = years.filter(year => year !== undefined && year !== null);

  // Create a new object with the desired structure
  return {
    RANKS: ranks,
    YEARS: years,
    SURNAME: obj['SURNAME'],
    OTHERNAME: obj['OTHERNAME'],
    GENDER: obj['GENDER'],
    BRANCHID: obj['BRANCHID']
  };
});
// console.log("res",ordinationResult);
let numberGet = ()=>{
    return Math.floor(Math.random()*(7-5+1)+5)
  }
  let resultProcess = ordinationResult.map(obj=>{
    let ranks = obj.RANKS
    let years = obj.YEARS
    let rankD = []
    
    if(obj.GENDER === "Male"){
        let mrf = rank.filter((v=> v.GENDER !== "F" )).sort((a,b)=> a + b)
        console.log(mrf)
        if(ranks.length > 0){
          ranks.forEach((a,i)=>{
            rank.forEach(e=>{
                if(e.RANK ===a){
                  if(e.GENDER ==="M"){
                    let nextRM = mrf.find(r=>r.RANKORDER === e.RANKORDER+1)
                    rankD.push({RANK:a,ID:e.ID,NEXTRANK:e.RANKORDER === 1?29:e.RANKORDER ===6?18:e.RANKORDER===3?25:nextRM === undefined?"":nextRM.ID,YEAR:years[i] !== undefined?years[i]:0})
                  }
                }
            })
          })
        }
      }else{
        let frf = rank.filter((v=> v.GENDER !== "M")).sort((a,b)=> a + b)
        if(ranks.length > 0){
          ranks.forEach((a,i)=>{
            rank.forEach(e=>{
                if(e.RANK ===a){
                  if(e.GENDER ==="F"){
                    let nextRF = frf.find(r=>r.RANKORDER === e.RANKORDER+1)
                    rankD.push({RANK:a,ID:e.ID,NEXTRANK:e.RANKORDER === 1?29:e.RANKORDER ===11?numberGet():nextRF === undefined?"":nextRF.ID,YEAR:years[i] !== undefined?years[i]:0})
                  }
                }
            })
          })
        }
      }
    let re2 = rankD.map(d=>{
        return{
            
            OtherName:obj.OTHERNAME,
            SureName: obj.SURNAME,
            RankId:d.ID,
            Year:d.YEAR,
            BranchId:obj.BRANCHID,
            NextRankId:d.NEXTRANK
        }
    }) 
    return re2
  })

let retData = [];
resultProcess.forEach(e=>{
    e.forEach(a=> retData.push(a))
})
// console.log("res",retData);


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