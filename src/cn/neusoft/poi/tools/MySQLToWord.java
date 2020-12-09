package cn.neusoft.poi.tools;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;
import cn.neusoft.poi.util.DBHelper;

import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.sql.ResultSet;
import java.sql.SQLException;

public class MySQLToWord {

    private final DBHelper db = new DBHelper();
    private final XWPFDocument document = new XWPFDocument();
    private final XWPFTable comTable = document.createTable();
    private XWPFTableRow wordRow;
    private ResultSet resultSet;
    private int countColumnNum;

    /***
     * 创建word表格
     * 列宽自动分割
     * 通过循环方式将数据放入剩余表格中
     */
    public void createWordForm()
    {
        db.connectToMySql();
        String sql = "select * from Student";
        countColumnNum = db.countColumn(sql);

        resultSet = db.search(sql);
        int countColumnNum = db.countColumn(sql);

        try {

            String[] names = new String[countColumnNum];

            for (int i = 0; i<countColumnNum; i++)
            {
                names[i] = resultSet.getMetaData().getColumnName(++i);
                //System.out.println(resultSet.getMetaData().getColumnName(i+1));
            }
            CTTblWidth comTableWidth = comTable.getCTTbl().addNewTblPr().addNewTblW();
            comTableWidth.setType(STTblWidth.DXA);
            comTableWidth.setW(BigInteger.valueOf(9072));

            wordRow = comTable.getRow(0);
            wordRow.getCell(0).setText(names[0]);//直接将0行0列放置于word表格中

            for (int j = 1 ; j < countColumnNum; j++)
            {
                wordRow.addNewTableCell().setText(names[j]);
            }

        } catch (SQLException se) {
            se.printStackTrace();
        }
    }

    /***
     * 添加数据至表格中
     */
    public void addAllDateToWordForm()
    {
        try {
            while (resultSet.next())
            {
                wordRow = comTable.createRow();
                for (int j = 0; j < countColumnNum; j++)
                {
                    XWPFTableCell tableCell = wordRow.getCell(j);
                    tableCell.setText(resultSet.getString(++j));
                }
            }
        } catch (SQLException se) {
            se.printStackTrace();
        }
    }

    /***
     * 将变更写入文件，导出docx文件
     * @param filePath
     */
    public void writeFormToDocx(String filePath)
    {
        try {
            FileOutputStream studentWord;
            studentWord = new FileOutputStream(filePath);
            document.write(studentWord);

            if (resultSet != null){
                resultSet.close();
            }
            System.out.println("\ndocx文件输出成功");
        } catch (IOException | SQLException e) {
            e.printStackTrace();
        }
    }
}
