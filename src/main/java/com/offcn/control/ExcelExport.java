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
        // 获取下载路径
        String path = request.getSession().getServletContext().getRealPath("/upload/");
        //生成的PDF文档名称
        String fileName = "demo.xlsx";
        //创建Excel
        createExcel(path + fileName);
        File file = new File(path + fileName); // 根据文件路径获得File文件

        response.setContentType("application/x-xls;charset=GBK");
        // 文件名
        response.setHeader("Content-Disposition",
                "attachment;filename=\"" + new String(fileName.getBytes(), "ISO8859-1") + "\"");
        response.setContentLength((int) file.length());
        byte[] buffer = new byte[4096];// 缓冲区
        BufferedOutputStream output = null;
        BufferedInputStream input = null;
        try {
            output = new BufferedOutputStream(response.getOutputStream());
            input = new BufferedInputStream(new FileInputStream(file));
            int n = -1;
            // 遍历，开始下载
            while ((n = input.read(buffer, 0, 4096)) > -1) {
                output.write(buffer, 0, n);
            }
            output.flush();
            response.flushBuffer();
        } catch (Exception e) {
            // 异常自己捕捉
        } finally {
            // 关闭流，不可少
            if (input != null)
                input.close();
            if (output != null)
                output.close();
        }
    }

    public void createExcel(String dest) throws IOException {
        //1、创建工作簿
        XSSFWorkbook workbook=new XSSFWorkbook();
        //2、新建工作表
        XSSFSheet sheet=workbook.createSheet("新建工作表1");
        //3、新增行
        for(int i=0;i<10;i++){
            XSSFRow row=sheet.createRow(i);
            //4、新建单元格
            for(int j=0;j<10;j++){
                XSSFCell cell=row.createCell(j);
                //5、给单元格填充数据
                cell.setCellValue((i+1)*(j+1));
            }
        }
        //6、保存工作簿
        FileOutputStream fileout=new FileOutputStream(dest);
        workbook.write(fileout);
        System.out.println("创建2007格式的Excel成功");
    }
}
