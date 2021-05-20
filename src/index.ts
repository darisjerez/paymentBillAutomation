import * as excel from 'exceljs';
import { ModifyExcelDto } from './dto/modifyExcel.dto';

export class PaymentAutomation {
    public workbook = new excel.Workbook();

    modifyExcel(modifyExcelParams: ModifyExcelDto):void{
        this.workbook.xlsx.readFile(modifyExcelParams.templateExcel)
            .then(function () {
                let worksheet = this.workbook.getWorksheet(1);
                let rows = [worksheet.getRow(modifyExcelParams.rowForBillNumber), worksheet.getRow(modifyExcelParams.rowForDate), worksheet.getRow(modifyExcelParams.rowForConcept)] 
                rows[0].getCell(modifyExcelParams.cellForDate).value = 5; 
                rows[1].getCell(modifyExcelParams.cellForBillNumber).value = 5; 
                rows[2].getCell(modifyExcelParams.cellForConcept).value = 5; 
                rows[0].commit();
                rows[1].commit();
                rows[2].commit();
                
                return this.workbook.xlsx.writeFile(modifyExcelParams.newExcel);
            })
    }


}