export class ModifyExcelDto {
    templateExcel: string;
    newExcel: string;
    rowForBillNumber: number;
    cellForBillNumber: number;
    rowForDate: number;
    cellForDate: number;
    rowForConcept: number;
    cellForConcept: number;
    concept: string;
}