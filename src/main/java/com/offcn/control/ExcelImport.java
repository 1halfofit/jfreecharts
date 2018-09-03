package com.offcn.control;

import java.io.File;
import java.io.FileInputStream;
import java.text.DecimalFormat;
import javax.servlet.http.HttpServletRequest;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.stereotype.Controller;
import org.springframework.ui.ModelMap;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

@Controller
public class ExcelImport {
    @RequestMapping(value = "importexcel", method = RequestMethod.POST)
    public String importCustomer(@RequestParam("file") MultipartFile file, HttpServletRequest request, ModelMap model) {
        System.out.println("开始");
        //获得Servlet上下文引用，从上下文引用获得上传目录，真实目录地址
        String path = request.getSession().getServletContext().getRealPath("upload");
        //获得上传文件名字
        String fileName = file.getOriginalFilename();
        //根据上传文件目录和上传文件名创建上传文件File对象
        File targetFile = new File(path, fileName);
        //判断上传文件对象是否为空，如果为空，创建上传文件目录包含父目录
        if (!targetFile.exists()) {
            targetFile.mkdirs();
        }
        //准备存储EXCEL表格数据的缓存对象BUF
        StringBuffer buf = new StringBuffer();
        try {
            //把上传的文件保存到指定的目录
            file.transferTo(targetFile);
            //从制定目录读取文件
            FileInputStream filein = new FileInputStream(targetFile);
            //读取文件到WORKBOOK
            Workbook workbook = WorkbookFactory.create(filein);
            //读取第一张工作表
            Sheet sheet = workbook.getSheetAt(0);
            //判断工作表里面有几行内容
            int rowNum = sheet.getPhysicalNumberOfRows();
            //判断工作表里面第一行有几个单元格（因为表格模板所有行的单元格数量一致）
            int cellNum = sheet.getRow(0).getPhysicalNumberOfCells();
            //循环遍历工作表里面的所有行
            for (int j = 0; j < rowNum; j++) {
                buf.append("<tr>");
                //获得具体每一行对象
                Row row = sheet.getRow(j);
                //循环遍历每一行内容
                for (short k = 0; k < cellNum; k++) {
                    //获得一行里面的每一个单元格
                    Cell cell = row.getCell(k);
                    //根据单元格类型，来读取数据
                    if (cell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
                        //把读取到的数据内容，写入缓冲区对象
                        buf.append("<td>" + cell.getStringCellValue() + "</td>");
                        System.out.print(cell.getStringCellValue() + "\t");
                    } else if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
                        //创建数字格式化工具类
                        DecimalFormat df = new DecimalFormat("####");
                        //针对数字类型的数据，进行格式化处理，防止自动转换成科学计数法格式
                        buf.append("<td>" + String.valueOf(df.format(cell.getNumericCellValue())) + "</td>");
                        System.out.print(cell.getNumericCellValue() + "\t");
                    } else if (cell.getCellType() == HSSFCell.CELL_TYPE_BOOLEAN) {
                        buf.append("<td>" + cell.getBooleanCellValue() + "</td>");
                        System.out.print(cell.getBooleanCellValue() + "\t");
                    } else if (cell.getCellType() == HSSFCell.CELL_TYPE_BLANK) {
                        buf.append("<td>" + "NULL" + "</td>");
                        System.out.print("NULL" + "\t");
                    } else {
                        buf.append("<td>" + cell.getDateCellValue() + "</td>");
                        System.out.print(cell.getDateCellValue() + "\t");
                    }
                }
                buf.append("</tr>");
                System.out.println("");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        //把EXCEL读取到的内容、文件上传路径等信息传递给前端页面
        model.addAttribute("text", buf.toString());
        model.addAttribute("fileUrl", request.getContextPath() + "/upload/" + fileName);
        return "resultupload";
    }
}