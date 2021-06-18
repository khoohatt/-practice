import java.io.File;
import java.io.FileOutputStream;
import java.io.FileInputStream;

import org.apache.poi.xwpf.usermodel.Borders;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.poi.xwpf.usermodel.VerticalAlign;

import java.lang.reflect.Constructor;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;

import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;

import org.apache.poi.util.Units;

import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import java.math.BigInteger;

public class CreateDocument {
    
    public static void main(String[] args) {
        
        try {
            // создание пустого документа
            XWPFDocument document = new XWPFDocument();

            // запись документа в файловую систему
            FileOutputStream out = new FileOutputStream(new File("Word Testing.docx"));
            
            SetBorders(document);
            
            CreateHeader(document);
            
            CreateFooter(document);
            
            WriteText(document);

            CreateTable(document);
            
            CreateBorders(document);
            
            TestFonts(document);

            SetTextToRight(document);
            
            SetTextToCenter(document);
            
            PutImg(document);

            document.write(out);
            out.close();
            System.out.println("файл Word Testing.docx успешно записан.");
        } catch (Exception e) {
            e.printStackTrace();
        } 
    }
    
    private static void SetBorders(XWPFDocument document) {
                    
        CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
        CTPageMar pageMar = sectPr.addNewPgMar();
        pageMar.setLeft(BigInteger.valueOf(1000L));
        pageMar.setTop(BigInteger.valueOf(1500L));
        pageMar.setRight(BigInteger.valueOf(1000L));
        pageMar.setBottom(BigInteger.valueOf(1500L));
    }
    
    private static void CreateHeader(XWPFDocument document) {
        
        CTP ctpHead = CTP.Factory.newInstance();
        CTText textHead = ctpHead.addNewR().addNewT();
        textHead.setStringValue("here goes the header.");
        XWPFParagraph parsHead[] = new XWPFParagraph[1];
        parsHead[0] = new XWPFParagraph(ctpHead, document);
        XWPFHeaderFooterPolicy hfp = document.createHeaderFooterPolicy();
        hfp.createHeader(XWPFHeaderFooterPolicy.DEFAULT, parsHead); 
    }
    
    private static void CreateFooter(XWPFDocument document) {
        
        CTP ctpFoot = CTP.Factory.newInstance();
        CTText textFoot = ctpFoot.addNewR().addNewT();
        textFoot.setStringValue("aaand here goes the footer!");
        XWPFParagraph parsFoot[] = new XWPFParagraph[1];
        parsFoot[0] = new XWPFParagraph(ctpFoot, document);
        XWPFHeaderFooterPolicy hfp = document.createHeaderFooterPolicy();
        hfp.createFooter(XWPFHeaderFooterPolicy.DEFAULT, parsFoot); 
    }
    
    private static void CreateTable(XWPFDocument document) {
                
        XWPFTable table = document.createTable();

        // создание первой строки
        XWPFTableRow tableRowOne = table.getRow(0);
        tableRowOne.getCell(0).setText("col one, row one");
        tableRowOne.addNewTableCell().setText("col two, row one");
        tableRowOne.addNewTableCell().setText("col three, row one");

        // создание второй строки
        XWPFTableRow tableRowTwo = table.createRow();
        tableRowTwo.getCell(0).setText("col one, row two");
        tableRowTwo.getCell(1).setText("col two, row two");
        tableRowTwo.getCell(2).setText("col three, row two");

        // создание третьей строки
        XWPFTableRow tableRowThree = table.createRow();
        tableRowThree.getCell(0).setText("col one, row three");
        tableRowThree.getCell(1).setText("col two, row three");
        tableRowThree.getCell(2).setText("col three, row three");
        
        // создание нового параграфа - запись пойдет с новой строки
        XWPFParagraph paragraph = document.createParagraph();
    }
    
    private static void CreateBorders(XWPFDocument document) {
        
        XWPFParagraph paragraph = document.createParagraph();

        // создание нижней границы таблицы
        paragraph.setBorderBottom(Borders.BASIC_BLACK_DASHES);

        // создание левой границы таблицы
        paragraph.setBorderLeft(Borders.BASIC_BLACK_DASHES);

        // создание правой границы таблицы
        paragraph.setBorderRight(Borders.BASIC_BLACK_DASHES);

        // создание верхней границы таблицы
        paragraph.setBorderTop(Borders.BASIC_BLACK_DASHES);

        XWPFRun run1 = paragraph.createRun();
        run1.setText("aaand some more text for another example.");
    }
    
    private static void TestFonts(XWPFDocument document) {
                
        XWPFParagraph paragraph = document.createParagraph();
        
        // запись текста жирным шрифтом
        XWPFRun paragraphOneRunOne = paragraph.createRun();
        paragraphOneRunOne.setBold(true);
        paragraphOneRunOne.setText("font style number one!");
        paragraphOneRunOne.addBreak();
        
        // запись текста курсивом
        XWPFRun paragraphOneRunTwo = paragraph.createRun();
        paragraphOneRunTwo.setItalic(true);
        paragraphOneRunTwo.setText("font style number two!!");
        
        // определение позиции текста - добавление разделения абзацев
        paragraphOneRunTwo.setTextPosition(100);

        // изменения вертикального выравнивания
        XWPFRun paragraphOneRunThree = paragraph.createRun();
        // изменение размера шрифта
        paragraphOneRunThree.setFontSize(20);
        // изменение цвета (сейчас здесь - темно-синий)
        paragraphOneRunThree.setColor("06357a");
                
        CTR ctr = paragraphOneRunThree.getCTR();
        ctr.addNewTab();
        ctr.addNewTab();
        paragraphOneRunThree.setText("font style number three!!! with tabulation...");
    }
    
    private static void SetTextToRight(XWPFDocument document) {
        
        XWPFParagraph paragraph = document.createParagraph();
        
        // установка вертикального выравнивания по правому краю
        paragraph.setAlignment(ParagraphAlignment.RIGHT);
        XWPFRun run = paragraph.createRun();
        run.setText("there is some text on the right side of the list.");
    }
    
    private static void SetTextToCenter(XWPFDocument document) {
        
        XWPFParagraph paragraph = document.createParagraph();
        
        // установка вертикального выравнивания по центру
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun run = paragraph.createRun();
        run.setText("and there is some text on the center of the list!");
    }
    
    private static void WriteText(XWPFDocument document) {
        
        XWPFParagraph paragraph = document.createParagraph();
        
        XWPFRun run = paragraph.createRun();
        run.setText("some text for an example.");

        paragraph = document.createParagraph();
    }
    
    private static void PutImg(XWPFDocument document) {
                    
        try {
            XWPFParagraph paragraph = document.createParagraph();
            
            XWPFRun run = paragraph.createRun();
            String imgFile = "humanity restored.jpg";
            FileInputStream is = new FileInputStream(imgFile);
            run.addBreak();
            run.addPicture(is, XWPFDocument.PICTURE_TYPE_JPEG, imgFile, Units.toEMU(200), Units.toEMU(200));   // 200x200 пикселей
            is.close();
        } catch (Exception e) {
            e.printStackTrace();
        } 
    }
}
