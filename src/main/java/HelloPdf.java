import com.itextpdf.io.font.FontConstants;
import com.itextpdf.kernel.color.Color;
import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfDocumentInfo;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Cell;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.property.TextAlignment;
import com.itextpdf.layout.property.VerticalAlignment;
import org.junit.Test;

import javax.swing.text.StyleConstants;
import java.io.FileNotFoundException;
import java.io.IOException;

public class HelloPdf {
    public static void main(String[] args) throws IOException {
        PdfWriter writer = new PdfWriter("d:\\javaPDF.pdf");
        PdfDocument pdfDocument = new PdfDocument(writer);
        PdfDocumentInfo info = pdfDocument.getDocumentInfo();
        info.setAuthor("博");
        info.setCreator("博");
        info.setKeywords("java,pdf,test");
        info.setSubject("学习java");
        info.setTitle("pdf");
        Document document = new Document(pdfDocument, PageSize.A4);
        Paragraph paragraph = new Paragraph("学习");
        paragraph.setTextAlignment(TextAlignment.CENTER);
        paragraph.setFontColor(Color.BLUE);
        PdfFont pdfFont = PdfFontFactory.createFont("STSongStd-Light", "UniGB-UCS2-H", false);
        document.setFont(pdfFont);
        paragraph.setFontSize(30);
        //往文档中添加（一个段落（内容））
        document.add(paragraph);
        document.close();
    }

    //生成带有表格的PDF文件
    @Test
    public void Test1() throws IOException {
        PdfWriter writer = new PdfWriter("d:\\table.pdf");
        PdfDocument pdfDocument = new PdfDocument(writer);
        //A4纸横向
        Document document = new Document(pdfDocument, PageSize.A4.rotate());
        //设置页边距（上右下左）
        document.setMargins(20, 20, 20, 20);
        PdfFont font = PdfFontFactory.createFont(FontConstants.HELVETICA);
        //创建表格对象（设置表格的大小）
        Table table = new Table(new float[]{40, 20, 20, 40});
        table.setWidthPercent(100);
        Cell cell = new Cell(1, 4);
        cell.setBackgroundColor(Color.BLUE);
        table.addHeaderCell(cell.add(new Paragraph("Job tile").setTextAlignment(TextAlignment.CENTER).setFont(font).setBold()));
        table.addCell(new Cell().add(new Paragraph("1").setFont(font)));
        table.addCell(new Cell().add(new Paragraph("2").setFont(font)));
        table.addCell(new Cell().add(new Paragraph("3").setFont(font)));
        table.addCell(new Cell().add(new Paragraph("4").setFont(font)));
        for (int i = 1; i < 17; i++) {
            table.addCell(new Cell().add(new Paragraph(i + "").setFont(font)));
        }
        Cell cell2 = new Cell(2, 1);
        //设定垂直居中
        cell2.setVerticalAlignment(VerticalAlignment.MIDDLE);
        cell2.setBackgroundColor(Color.GREEN);
        table.addCell(cell2);
        for (int i = 1; i < 7; i++) {
            table.addCell(new Cell().setBackgroundColor(Color.CYAN).add(new Paragraph(i + "").setFont(font)));
        }
        document.add(table);
        document.close();

    }
}
