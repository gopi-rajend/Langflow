import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class ExcelReader {
    public static void main(String[] args) throws IOException {
        FileInputStream file = new FileInputStream("data.xlsx");
        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);

        for (Row row : sheet) {
            for (Cell cell : row) {
                switch (cell.getCellType()) {
                    case STRING -> System.out.print(cell.getStringCellValue() + "\t");
                    case NUMERIC -> System.out.print(cell.getNumericCellValue() + "\t");
                    case BOOLEAN -> System.out.print(cell.getBooleanCellValue() + "\t");
                    default -> System.out.print("UNKNOWN\t");
                }
            }
            System.out.println();
        }

        workbook.close();
        file.close();
    }
}

typescript=================================================================================================

import * as XLSX from 'xlsx';
import * as fs from 'fs';

const fileBuffer = fs.readFileSync('data.xlsx');
const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];

const jsonData = XLSX.utils.sheet_to_json(sheet);
console.log(jsonData);

=========================python-------------------------------------------------------------

import pandas as pd

# Load the Excel file
df = pd.read_excel("data.xlsx", sheet_name="Sheet1")  # You can also use sheet_name=0 for the first sheet

# Display the entire DataFrame
print(df)

# Access specific columns
print("Names:", df["Name"].tolist())  # Replace "Name" with your actual column header

# Access specific rows
print("First row:", df.iloc[0])