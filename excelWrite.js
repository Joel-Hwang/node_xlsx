var XLSX = require('xlsx');
var workbook = XLSX.readFile(`C:\\workspace\\Docs\\PLM\.Phase2\\IE\\proporder.xlsx`);
var sheet = workbook.Sheets[workbook.SheetNames[0]];

function main(){
    let cnt = 19000;
    let endCell = sheet["!ref"].split(":")[1];
    let rows = Number(endCell.substr(2, endCell.length - 1));
    let cols = ['B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R']
    for(let i = 76; i<=126; i+=5){
        try{
            for(let col of cols){
                if(sheet[`${col}${i}`]){
                    sheet[`${col}${i+3}`].v = 'xp-st_'+pad(cnt,5) + getCell(`${col}${i}`).substr(8);
                    cnt+=100;
                }
            }
        }catch(e){
            console.log(i + e.toString());
        }
        
    }
    XLSX.writeFile(workbook, 'test2.xls');
    //XLSX.write(workbook, {bookType:'xlsx', bookSST:true, type: 'base64'}) 
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