package com.ames;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class TestReadFormula {
    private static FormulaEvaluator evaluator;
    public static void main(String[] args) throws IOException {
        InputStream is=new FileInputStream("ReadFormula.xls");
        HSSFWorkbook wb=new HSSFWorkbook(is);
        Sheet sheet=wb.getSheetAt(0);

        evaluator=wb.getCreationHelper().createFormulaEvaluator();

        for (int i = 1; i <4; i++) {
            Row  row=sheet.getRow(i);
            for (Cell cell : row) {
                System.out.println(getCellValue(cell));
            }
        }
        wb.close();


    }

    private static String getCellValue(Cell cell) {
        if (cell==null) {
            return "isNull";
        }
        System.out.println("rowIdx:"+cell.getRowIndex()+",colIdx:"+cell.getColumnIndex());
        String cellValue = null;
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                System.out.print("STRING :");
                cellValue=cell.getStringCellValue();
                break;

            case Cell.CELL_TYPE_NUMERIC:
                System.out.print("NUMERIC:");
                cellValue=String.valueOf(cell.getNumericCellValue());
                break;

            case Cell.CELL_TYPE_FORMULA:
                System.out.print("FORMULA:");
                cellValue=getCellValue(evaluator.evaluate(cell));
                break;
            default:
                System.out.println("Has Default.");
                break;
        }

        return cellValue;
    }

    private static String getCellValue(CellValue cell) {
        String cellValue = null;
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                System.out.print("String :");
                cellValue=cell.getStringValue();
                break;

            case Cell.CELL_TYPE_NUMERIC:
                System.out.print("NUMERIC:");
                cellValue=String.valueOf(cell.getNumberValue());
                break;
            case Cell.CELL_TYPE_FORMULA:
                System.out.print("FORMULA:");
                break;
            default:
                break;
        }

        return cellValue;
    }

}
