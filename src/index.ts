
import * as path from 'path';
import Excel from 'exceljs';
const filePath = path.resolve(__dirname, '../'+process.argv[2]);//iw-tech-test-retailer-data.xlsx');

const getCellValue = (row:  Excel.Row, cellIndex: number) => {
    const cell = row.getCell(cellIndex);
    return cell.value ? cell.value.toString() : '';
};

const getObject = (row:  Excel.Row) => {
    let category = getCellValue(row,1);
    return {category:category.split(";"),
        email:getCellValue(row,3),
        fax:getCellValue(row,4),
        mobile:getCellValue(row,5),
        phone:getCellValue(row,6),
        website:getCellValue(row,7),
        id:getCellValue(row,8),
        slug:getCellValue(row,9),
        body:getCellValue(row,10),
        street:getCellValue(row,11),
        city:getCellValue(row,12),
        country:getCellValue(row,13),
        latitude:getCellValue(row,15),
        longitude:getCellValue(row,16),
        zip:getCellValue(row,17),
        state:getCellValue(row,18),
        facebook:getCellValue(row,19),
        googleplus:getCellValue(row,20),
        twitter:getCellValue(row,21),
        status:getCellValue(row,22),
        title:getCellValue(row,23)
    };
};

const main = async () => {
    const workbook = new Excel.Workbook();
    const content = await workbook.xlsx.readFile(filePath);
    const worksheet: Excel.Worksheet | undefined = content.getWorksheet(1)
    const rowStartIndex = 2;
    if(worksheet==undefined){
        return;
    }
    const numberOfRows = worksheet.rowCount - 1;
    const rows = worksheet.getRows(rowStartIndex, numberOfRows) ?? [];
    let companies : any=[]
    if(rows.length>0){
        companies = rows.map((row): any => {
            let keyindex = getCellValue(row,9)
            return {
                [keyindex]:getObject(row)
            }
        });
        console.log(JSON.stringify(companies))
    }
};

main()