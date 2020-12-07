import org.apache.poi.xwpf.usermodel.*;

import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;

import java.io.File;
import java.io.FileOutputStream;
import java.math.BigInteger;

public class TestPOIWordOut {
    public static void main(String[] args) {

        //Blank Document
        XWPFDocument document = new XWPFDocument();

        //Write the Document in file system
        FileOutputStream out;
        try {
            out = new FileOutputStream(new File("poi_word.docx"));

            //添加标题
            XWPFParagraph titleParagraph = document.createParagraph();
            //设置段落居中
            titleParagraph.setAlignment(ParagraphAlignment.CENTER);

            XWPFRun titleParagraphRun = titleParagraph.createRun();
            titleParagraphRun.setText("自学指导（标题）");
            titleParagraphRun.setColor("000000");
            titleParagraphRun.setFontSize(20);

            //段落
            XWPFParagraph firstParagraph = document.createParagraph();
            XWPFRun run = firstParagraph.createRun();
            run.setText("   自学遇到问题怎么办？");
            run.setColor("698869");
            run.setFontSize(16);

            //设置段落背景颜色
//            CTShd cTShd = run.getCTR().addNewRPr().addNewShd();
            CTShd ctShd = run.getCTR().addNewRPr().addNewShd();
            ctShd.setVal(STShd.CLEAR);
            ctShd.setFill("97FF45");

            //换行
            XWPFParagraph paragraph1 = document.createParagraph();
            XWPFRun paragraphRun1 = paragraph1.createRun();
            paragraphRun1.setText("\r");

            //基本信息表格
            XWPFTable infoTable = document.createTable();
            //去表格边框
            infoTable.getCTTbl().getTblPr().unsetTblBorders();

            //列宽自动分割
            CTTblWidth infoTableWidth = infoTable.getCTTbl().addNewTblPr().addNewTblW();
            infoTableWidth.setType(STTblWidth.DXA);
            infoTableWidth.setW(BigInteger.valueOf(9072));

            //表格第一行
            XWPFTableRow infoTableRowOne = infoTable.getRow(0);
            infoTableRowOne.getCell(0).setText("职位");
            infoTableRowOne.addNewTableCell().setText(": xxxx");

            //表格第二行
            XWPFTableRow infoTableRowTwo = infoTable.createRow();
            infoTableRowTwo.getCell(0).setText("姓名");
            infoTableRowTwo.getCell(1).setText(": xxxx");

            //表格第三行
            XWPFTableRow infoTableRowThree = infoTable.createRow();
            infoTableRowThree.getCell(0).setText("生日");
            infoTableRowThree.getCell(1).setText(": xxx-xx-xx");

            //表格第四行
            XWPFTableRow infoTableRowFour = infoTable.createRow();
            infoTableRowFour.getCell(0).setText("性别");
            infoTableRowFour.getCell(1).setText(": x");

            //表格第五行
            XWPFTableRow infoTableRowFive = infoTable.createRow();
            infoTableRowFive.getCell(0).setText("现居地");
            infoTableRowFive.getCell(1).setText(": xx");

            //两个表格之间加个换行
            XWPFParagraph paragraph = document.createParagraph();
            XWPFRun paragraphRun = paragraph.createRun();
            paragraphRun.setText("\r");

            //工作经历表格
            XWPFTable ComTable = document.createTable();

            //列宽自动分割
            CTTblWidth comTableWidth = ComTable.getCTTbl().addNewTblPr().addNewTblW();
            comTableWidth.setType(STTblWidth.DXA);
            comTableWidth.setW(BigInteger.valueOf(9072));

            //表格第一行
            XWPFTableRow comTableRowOne = ComTable.getRow(0);
            comTableRowOne.getCell(0).setText("开始时间");
            comTableRowOne.addNewTableCell().setText("结束时间");
            comTableRowOne.addNewTableCell().setText("公司名称");
            comTableRowOne.addNewTableCell().setText("职位");

            //表格第二行
            XWPFTableRow comTableRowTwo = ComTable.createRow();
            comTableRowTwo.getCell(0).setText("20xx-xx-xx");
            comTableRowTwo.getCell(1).setText("至今");
            comTableRowTwo.getCell(2).setText("自学指导网站");
            comTableRowTwo.getCell(3).setText("Java开发工程师");

            //表格第三行
            XWPFTableRow comTableRowThree = ComTable.createRow();
            comTableRowThree.getCell(0).setText("20xx-xx-xx");
            comTableRowThree.getCell(1).setText("至今");
            comTableRowThree.getCell(2).setText("s.jf3q.com");
            comTableRowThree.getCell(3).setText("Java开发工程师");

            document.write(out);
            out.close();
        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        System.out.println("create_table document written success.");
    }
}
