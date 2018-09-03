package com.offcn.control;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

@Controller
public class ExcelExport {
    @RequestMapping("/exceldownload")
    public void downLoadExcelModel(HttpServletRequest request, HttpServletResponse response) throws Exception {
        // ��ȡ����·��
        String path = request.getSession().getServletContext().getRealPath("/upload/");
        //���ɵ�PDF�ĵ�����
        String fileName = "demo.xlsx";
        //����Excel
        createExcel(path + fileName);
        File file = new File(path + fileName); // �����ļ�·�����File�ļ�

        response.setContentType("application/x-xls;charset=GBK");
        // �ļ���
        response.setHeader("Content-Disposition",
                "attachment;filename=\"" + new String(fileName.getBytes(), "ISO8859-1") + "\"");
        response.setContentLength((int) file.length());
        byte[] buffer = new byte[4096];// ������
        BufferedOutputStream output = null;
        BufferedInputStream input = null;
        try {
            output = new BufferedOutputStream(response.getOutputStream());
            input = new BufferedInputStream(new FileInputStream(file));
            int n = -1;
            // ��������ʼ����
            while ((n = input.read(buffer, 0, 4096)) > -1) {
                output.write(buffer, 0, n);
            }
            output.flush();
            response.flushBuffer();
        } catch (Exception e) {
            // �쳣�Լ���׽
        } finally {
            // �ر�����������
            if (input != null)
                input.close();
            if (output != null)
                output.close();
        }
    }

    public void createExcel(String dest) throws IOException {
        //1������������
        XSSFWorkbook workbook=new XSSFWorkbook();
        //2���½�������
        XSSFSheet sheet=workbook.createSheet("�½�������1");
        //3��������
        for(int i=0;i<10;i++){
            XSSFRow row=sheet.createRow(i);
            //4���½���Ԫ��
            for(int j=0;j<10;j++){
                XSSFCell cell=row.createCell(j);
                //5������Ԫ���������
                cell.setCellValue((i+1)*(j+1));
            }
        }
        //6�����湤����
        FileOutputStream fileout=new FileOutputStream(dest);
        workbook.write(fileout);
        System.out.println("����2007��ʽ��Excel�ɹ�");
    }
}
