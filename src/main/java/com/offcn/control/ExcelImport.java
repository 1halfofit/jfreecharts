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
        System.out.println("��ʼ");
        //���Servlet���������ã������������û���ϴ�Ŀ¼����ʵĿ¼��ַ
        String path = request.getSession().getServletContext().getRealPath("upload");
        //����ϴ��ļ�����
        String fileName = file.getOriginalFilename();
        //�����ϴ��ļ�Ŀ¼���ϴ��ļ��������ϴ��ļ�File����
        File targetFile = new File(path, fileName);
        //�ж��ϴ��ļ������Ƿ�Ϊ�գ����Ϊ�գ������ϴ��ļ�Ŀ¼������Ŀ¼
        if (!targetFile.exists()) {
            targetFile.mkdirs();
        }
        //׼���洢EXCEL������ݵĻ������BUF
        StringBuffer buf = new StringBuffer();
        try {
            //���ϴ����ļ����浽ָ����Ŀ¼
            file.transferTo(targetFile);
            //���ƶ�Ŀ¼��ȡ�ļ�
            FileInputStream filein = new FileInputStream(targetFile);
            //��ȡ�ļ���WORKBOOK
            Workbook workbook = WorkbookFactory.create(filein);
            //��ȡ��һ�Ź�����
            Sheet sheet = workbook.getSheetAt(0);
            //�жϹ����������м�������
            int rowNum = sheet.getPhysicalNumberOfRows();
            //�жϹ����������һ���м�����Ԫ����Ϊ���ģ�������еĵ�Ԫ������һ�£�
            int cellNum = sheet.getRow(0).getPhysicalNumberOfCells();
            //ѭ�����������������������
            for (int j = 0; j < rowNum; j++) {
                buf.append("<tr>");
                //��þ���ÿһ�ж���
                Row row = sheet.getRow(j);
                //ѭ������ÿһ������
                for (short k = 0; k < cellNum; k++) {
                    //���һ�������ÿһ����Ԫ��
                    Cell cell = row.getCell(k);
                    //���ݵ�Ԫ�����ͣ�����ȡ����
                    if (cell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
                        //�Ѷ�ȡ�����������ݣ�д�뻺��������
                        buf.append("<td>" + cell.getStringCellValue() + "</td>");
                        System.out.print(cell.getStringCellValue() + "\t");
                    } else if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
                        //�������ָ�ʽ��������
                        DecimalFormat df = new DecimalFormat("####");
                        //����������͵����ݣ����и�ʽ��������ֹ�Զ�ת���ɿ�ѧ��������ʽ
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
        //��EXCEL��ȡ�������ݡ��ļ��ϴ�·������Ϣ���ݸ�ǰ��ҳ��
        model.addAttribute("text", buf.toString());
        model.addAttribute("fileUrl", request.getContextPath() + "/upload/" + fileName);
        return "resultupload";
    }
}