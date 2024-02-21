const ExcelJs = require('exceljs');
async function writeExcelTest(searchText,replaceText,change,filepath)
{
    const workbook = new ExcelJs.Workbook();
    await workbook.xlsx.readFile(filepath)
    const worksheet = workbook.getWorksheet("FirstData");
    
    const output = await readExcel(worksheet,searchText);

    cell = worksheet.getCell(output.row,output.coloumn+change.colChange)
    cell.value = replaceText;
    console.log("Writing at row = " + cell.row + " and col = " + cell.col);
    await workbook.xlsx.writeFile(filepath);
    console.log("cell updated succcessfully !!!");
}

async function readExcel(worksheet,searchText)
{
    let output = {row:-1, coloumn :-1}
    worksheet.eachRow((row,rowNumber) =>
    {
        row.eachCell((cell,colNumber) =>
        {
            if (cell.value === searchText)
            {
                output.row = cell.row; 
                output.coloumn = cell.col;
            }
        })
    })
    return output;

}


writeExcelTest("Avocado",400,{rowChange:0,colChange:1},"/Users/meenakshipal/Documents/test.xlsx");