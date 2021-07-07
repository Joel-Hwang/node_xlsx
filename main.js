const XLSX = require('xlsx');
const fs = require('fs');
const workbook = XLSX.readFile(`C:\\workspace\\Docs\\PLM.Phase3\\BPFC\\bonding정보\\bonding 기준 정보 20210625(IT).xlsx`);
let sheet = workbook.Sheets['Sheet1'];

//convert pbfc chemical to JSON
function main(){
    let cnt = 100;
    let res = [];
    let endCell = sheet["!ref"].split(":")[1];
    let rows = Number(endCell.substr(1, endCell.length - 1));
    
    for(let i = 2; i<=rows; i++){
        let obj = {};
        obj._sheet = getCell(`A${i}`);
        obj._vendor = getCell(`B${i}`);
        obj._proc_name = getCell(`C${i}`);
        obj._chemical = getCell(`D${i}`);
        obj._drawio_proc_name = getCell(`E${i}`);
        obj._hrd = getCell(`F${i}`) == ""? []:getCell(`F${i}`).split('|');
        obj._condition = getCell(`G${i}`) == ""?[]:getCell(`G${i}`).split('|');
        res.push(obj);
        //console.log(obj);
    }
    fs.writeFile("test.txt",JSON.stringify(res),'utf8',(err,data)=>{
        console.log(data);
    });
    console.log(JSON.stringify(res));
   // XLSX.writeFile(workbook, 'test2.xls');
}

function getCell(cell){
    let result = sheet[cell];
    if (!result) return "";
    else return result.w.toString();
}

function pad(n, width) {
    n = n + '';
    return n.length >= width ? n : new Array(width - n.length + 1).join('0') + n;
}
main();


