var XLSX = require('xlsx');
var workbook = XLSX.readFile("C:\\workspace\\Docs\\Digital Engineering\\CE\\Test Request.xls");
var sheet = workbook.Sheets[workbook.SheetNames[0]];

function main(){
    let endCell = sheet["!ref"].split(":")[1];
    let rows = Number(endCell.substr(2, endCell.length - 1));

    for(let i = 7; i<=rows; i++){
        console.log('/*===='+i+'=====*/');
        let date = getCell(`A${i}`);
        let time = getCell(`B${i}`);
        let ce = getCell(`C${i}`);
        let labDataYn = getCell(`D${i}`);
        let season = getCell(`E${i}`);
        let model = getCell(`F${i}`);
        let category = getCell(`G${i}`);
        let round = getCell(`I${i}`);
        let dpa = getCell(`K${i}`);
        let colorway = getCell(`L${i}`);
        let tempColorway =getCell(`L${i}`).split('/');
        colorway = tempColorway[0];

        let bom = getCell(`M${i}`);
        let style = getCell(`M${i}`);
        let tempBomStyle = getCell(`M${i}`).split('/');
        bom = tempBomStyle[0];
        style = tempBomStyle.length == 2?tempBomStyle[1]:"";

        let dev = getCell(`O${i}`);
        let prod = getCell(`O${i}`);
        let tempDevProd = getCell(`O${i}`).split('/');
        dev = tempDevProd[0];
        prod = tempDevProd.length == 2?tempDevProd[1]:"";


        let td = getCell(`P${i}`);
        let modelId = getCell(`Q${i}`);
        let pm = getCell(`R${i}`);
        let recQty = getCell(`AA${i}`);
        recQty = recQty===""?"0":recQty;
        let remQty = getCell(`AB${i}`);
        remQty = remQty===""?"0":remQty;
        let tempQty = getCell(`AC${i}`);
        tempQty = tempQty===""?"0":tempQty;
        let surfQty = getCell(`AD${i}`);
        surfQty = surfQty===""?"0":surfQty;
        let tPull = getCell(`AE${i}`);
        let tAssem = getCell(`AF${i}`);
        let tStock = getCell(`AG${i}`);
        let tFace = getCell(`AH${i}`);
        let tlini = getCell(`AI${i}`);
        let tHfwe = getCell(`AJ${i}`);
        let tFlex = getCell(`AK${i}`);
        let tWash = getCell(`AL${i}`);
        let tAge = getCell(`AM${i}`);
        let tQuv = getCell(`AN${i}`);
        let tG97 = getCell(`AO${i}`);
        let tChina = getCell(`AP${i}`);
        let tPh = getCell(`AQ${i}`);
        let etc = getCell(`AR${i}`);
        let destroyDate = getCell(`AS${i}`);
        let labeling = getCell(`AT${i}`);

        console.log(`INSERT INTO TEMP_LABTESTREQ VALUES ( '${date}','${time}','${ce}','${labDataYn}','${season}','${model}','${category}','${round}','${dpa}','${colorway}','${bom}','${style}','${dev}','${prod}','${td}','${modelId}','${pm}',${recQty},${remQty},${tempQty},${surfQty},'${tPull}','${tAssem}','${tStock}','${tFace}','${tlini}','${tHfwe}','${tFlex}','${tWash}','${tAge}','${tQuv}','${tG97}','${tChina}','${tPh}','${etc}','${destroyDate}','${labeling}' );`);

        /*console.log(date},${time},${ce},${labDataYn},${season},${model},${category},${round},${dpa
            },${colorway},${bom},${style},${dev},${prod},${td},${modelId},${pm},${recQty},${remQty},${tempQty},${surfQty
            },${tPull},${tAssem},${tStock},${tFace},${tlini},${tHfwe},${tFlex},${tWash},${tAge},${tQuv},${tG97},${tChina},${tPh
            },${etc},${destroyDate},${labeling);*/
    }
}

function getCell(cell){
    let result = sheet[cell];
    if (!result) return "";
    else return result.w.toString();
}
main();