import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;
import util.DBHelper;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.sql.ResultSet;
import java.sql.SQLException;

public class MySQLToWord {

    private final DBHelper db = new DBHelper();
    private final XWPFDocument document = new XWPFDocument();
    private final XWPFTable comTable = document.createTable();
    private ResultSet resultSet;
    private XWPFTableRow wordRow;
    private int countColumnNum;

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
                names[i] = resultSet.getMetaData().getColumnName(i+1);
                System.out.println(resultSet.getMetaData().getColumnName(i+1));
            }
            //列宽自动分割
            CTTblWidth comTableWidth = comTable.getCTTbl().addNewTblPr().addNewTblW();
            comTableWidth.setType(STTblWidth.DXA);
            comTableWidth.setW(BigInteger.valueOf(9072));

            wordRow = comTable.getRow(0);
            wordRow.getCell(0).setText(names[0]);//直接将0行0列放置于word表格中

            //再通过循环方式将数据放入剩余表格中
            for (int j = 1 ; j < countColumnNum; j++)
            {
                wordRow.addNewTableCell().setText(names[j]);
            }

        } catch (SQLException se) {
            se.printStackTrace();
        }
    }

    public void addAllDateToWordForm()
    {
        try {
            while (resultSet.next())
            {
                wordRow = comTable.createRow();
                for (int j = 0; j < countColumnNum; j++)
                {
                    XWPFTableCell tableCell = wordRow.getCell(j);//空，无法找到行，需要创建行
                    tableCell.setText(resultSet.getString(j+1));
                }
            }
        } catch (SQLException se) {
            se.printStackTrace();
        }
    }

    public void writeFormToDocx(String filePath)
    {
        try {
            FileOutputStream studentWord;
            studentWord = new FileOutputStream(filePath);
            document.write(studentWord);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

