export class CreateXLS {
    constructor() {
        XLSX = require('xlsx');
        const path = 'tmp/partdir_NOVOv2.xls';
        const finalPath = 'consolidado.xls';
        const wbNewData = XLSX.readFile(path, {cellStyles:true});
        const wbSavedData = XLSX.readFile(finalPath, {cellStyles:true});
        
        /* novos dados recuperados do servidor*/
        let wsNewData = wbNewData.Sheets[wbNewData.SheetNames[0]];
        let jsonNewData = XLSX.utils.sheet_to_json(wsNewData,{header:'A', skipHeader:true});
        
        /* dados já armazenados */
        let wsSavedData = wbSavedData.Sheets[wbSavedData.SheetNames[0]];
        let jsonSavedData = XLSX.utils.sheet_to_json(wsSavedData,{header:'A', skipHeader:true});
        
        let jsonRetorno = jsonNewData;
        jsonRetorno.push('');
        jsonSavedData.forEach(row => {
            jsonRetorno.push(row);
        });
        var workbook = XLSX.utils.book_new();
        let updatedWS = XLSX.utils.json_to_sheet(jsonRetorno)
        XLSX.utils.book_append_sheet(workbook, updatedWS, 'resultado');
        XLSX.writeFile(workbook, finalPath);
        
    }
}
