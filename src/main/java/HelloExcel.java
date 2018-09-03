import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class HelloExcel {
    @Test
    public void Test() throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("表一");
        HSSFRow row = sheet.createRow(0);
        HSSFCell cell = row.createCell(2);
        cell.setCellValue("你好EXCEL");
        workbook.write(new File("d:\\hello.xls"));
        System.out.println("success");
    }

    @Test
    public void Test2() throws IOException {
        FileInputStream inputStream = new FileInputStream("d:\\1.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        int sheetnum = workbook.getNumberOfSheets();
        for(int i=0;i<sheetnum;i++) {
            //获取到单个工作表
            Sheet sheet = workbook.getSheetAt(i);
            //获取工作表下的所有行数
            int rownum = sheet.getPhysicalNumberOfRows();
            //获取第一行的单元格个数

            for (int j = 0; j < rownum; j++) {
                //获取到每一行
                Row row = sheet.getRow(j);
                int cellnum = row.getPhysicalNumberOfCells();
                //获取每一行下的全部单元格
                for (int x = 0; x < cellnum; x++) {
                    Cell cell = row.getCell(x);
                    //获取每个单元格下的内容
                    if (cell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
                        System.out.print(cell.getStringCellValue() + "\t" + "\t" + "\t" + "\t");
                    } else if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
                        System.out.print(cell.getNumericCellValue() + "\t" + "\t" + "\t" + "\t");
                    } else if (cell.getCellType() == HSSFCell.CELL_TYPE_BOOLEAN) {
                        System.out.println(cell.getBooleanCellValue() + "\t" + "\t" + "\t" + "\t");
                    } else if (cell.getCellType() == HSSFCell.CELL_TYPE_BLANK) {
                        System.out.print("-" + "\t" + "\t" + "\t" + "\t" + "\t");
                    } else {
                        System.out.print(cell.getDateCellValue() + "\t" + "\t" + "\t");
                    }
                }
                System.out.println("");
            }

        }
        //遍历excel得到每一行再得到这一行第索引列的内容
//        for(int i=0;i<sheet.getLastRowNum();i++){
//        HSSFRow row = sheet.getRow(i);
//            System.out.println(row.getCell(4).getStringCellValue());
//        }
//        for(int i=0;i<sheet.getLastRowNum();i++){
//            Row row = sheet.getRow(i);
//            System.out.println(row.getCell(5).getStringCellValue());
//        }
//        HSSFRow row = sheet.getRow(2);
//        //获取当前第二行第四列的值
//        System.out.println(row.getCell(4).getStringCellValue());
//        //设置第二行第四列的值为0520
//        row.getCell(4).setCellValue("0520");
//        //获取修改后的第二行第四列的值
//        System.out.println(row.getCell(4).getStringCellValue());
    }
}
