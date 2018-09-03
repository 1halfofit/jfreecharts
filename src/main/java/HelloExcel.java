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
        HSSFSheet sheet = workbook.createSheet("��һ");
        HSSFRow row = sheet.createRow(0);
        HSSFCell cell = row.createCell(2);
        cell.setCellValue("���EXCEL");
        workbook.write(new File("d:\\hello.xls"));
        System.out.println("success");
    }

    @Test
    public void Test2() throws IOException {
        FileInputStream inputStream = new FileInputStream("d:\\1.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        int sheetnum = workbook.getNumberOfSheets();
        for(int i=0;i<sheetnum;i++) {
            //��ȡ������������
            Sheet sheet = workbook.getSheetAt(i);
            //��ȡ�������µ���������
            int rownum = sheet.getPhysicalNumberOfRows();
            //��ȡ��һ�еĵ�Ԫ�����

            for (int j = 0; j < rownum; j++) {
                //��ȡ��ÿһ��
                Row row = sheet.getRow(j);
                int cellnum = row.getPhysicalNumberOfCells();
                //��ȡÿһ���µ�ȫ����Ԫ��
                for (int x = 0; x < cellnum; x++) {
                    Cell cell = row.getCell(x);
                    //��ȡÿ����Ԫ���µ�����
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
        //����excel�õ�ÿһ���ٵõ���һ�е������е�����
//        for(int i=0;i<sheet.getLastRowNum();i++){
//        HSSFRow row = sheet.getRow(i);
//            System.out.println(row.getCell(4).getStringCellValue());
//        }
//        for(int i=0;i<sheet.getLastRowNum();i++){
//            Row row = sheet.getRow(i);
//            System.out.println(row.getCell(5).getStringCellValue());
//        }
//        HSSFRow row = sheet.getRow(2);
//        //��ȡ��ǰ�ڶ��е����е�ֵ
//        System.out.println(row.getCell(4).getStringCellValue());
//        //���õڶ��е����е�ֵΪ0520
//        row.getCell(4).setCellValue("0520");
//        //��ȡ�޸ĺ�ĵڶ��е����е�ֵ
//        System.out.println(row.getCell(4).getStringCellValue());
    }
}
