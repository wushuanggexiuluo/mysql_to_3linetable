package com.lancer;

import com.alibaba.druid.pool.DruidDataSource;
import com.alibaba.druid.util.JdbcUtils;
import com.lowagie.text.Font;
import com.lowagie.text.*;
import com.lowagie.text.pdf.BaseFont;
import com.lowagie.text.rtf.RtfWriter2;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.commons.lang3.StringUtils;
import org.assertj.core.util.Lists;

import javax.sql.DataSource;
import java.awt.*;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.sql.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;


public class Utils {
    public static String out_path = "X:\\";
    public static String dbHost = "127.0.0.1";
    public static int dbPort = 3306;
    public static String dbName = "snowy-cloud";
    public static String userName = "root";
    public static String password = "scm13503905942";
    // 单元格5号字体
    public static final float tbTitleFontSize = 10.5F;
    // 单元格头字体大小
    public static final float cellFontSize = 10.5F;
    // 单元格头字体大小
    public static final float cellHeaderFontSize = 10.5F;
    // 需要排除的字段
    final static List<String> excludeWord = Arrays.asList("create_time", "update_time", "create_user", "update_user", "delete_flag");
    static BaseFont chinaTxtFont;

    public static void mysqlToWordTb(DataSource ds) throws SQLException {
        String fileName=dbName + ".doc";
        List<TableInfo> tables = getTableInfos(ds, dbName);
        Document document = new Document(PageSize.A4);
        try {
            File dir = new File(out_path);
            if (!dir.exists()) {
                dir.mkdirs();
            }
            fileName = out_path + File.separator + fileName;
            File file = new File(fileName);
            if (file.exists() && file.isFile()) {
                file.delete();
            }
            try{
                file.createNewFile();
            }catch (Exception e){
                System.out.println("当前路径没有权限访问，请使用其他路径！");
            }
            // 写入文件信息
            RtfWriter2.getInstance(document, Files.newOutputStream(Paths.get(fileName)));
            document.open();
            genTableStructDesc(document, tables, ds);
            document.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        printMsg("所有表【共%d个】已经处理完成", tables.size());
    }


    private static void genTableStructDesc(Document document, List<TableInfo> tables, DataSource ds) throws DocumentException, SQLException, IOException {

        Paragraph p = new Paragraph("表结构描述", new Font(chinaTxtFont, tbTitleFontSize));
        p.setAlignment(Element.ALIGN_CENTER);
        document.add(p);
        printMsg("共需要处理%d个表", tables.size());
        int colNum = 7;
        //循环处理每一张表
        for (int i = 0; i < tables.size(); i++) {
            TableInfo tableInfo = tables.get(i);
            String tblName = tableInfo.getTblName();
            String tblComment = tableInfo.getTblComment();

            printMsg("处理%s表开始", tableInfo);
            List<TableFiled> fileds = new ArrayList<>();
            List<TableFiled> fileds2 = getTableFields(ds, tables.get(i).getTblName());
            // 去除常见的
            for (TableFiled filed : fileds2) {
                String lowerCase = filed.getField().toLowerCase();
                if (!excludeWord.contains(lowerCase)) {
                    fileds.add(filed);
                }
            }
            Table table = new Table(colNum);
            int[] widths = new int[]{160, 250, 350, 160, 80, 120, 80};
            table.setWidths(widths);
            table.setPadding(0);
            table.setSpacing(0);

            //添加表名行
            String tblInfo = StringUtils.isBlank(tblComment) ? tblName : String.format("表 -%s %s表", i, tblComment.replace("表", ""));

            Paragraph ph = new Paragraph(tblInfo, new Font(chinaTxtFont, tbTitleFontSize));
            ph.setAlignment(Paragraph.ALIGN_CENTER);
            document.add(ph);
            Font font = new Font(chinaTxtFont, cellHeaderFontSize);

            //添加表头行
            addCell(table, "字段名", 0);
            addCell(table, "字段描述", 0);
            addCell(table, "数据类型", 0);
            addCell(table, "长度", 0);
            addCell(table, "可空", 0);
            addCell(table, "是否主键", 0);
            addCell(table, "备注", 0);
            table.endHeaders();
            int k;
            // 表格的主体
            for (k = 0; k < fileds.size(); k++) {
                TableFiled field = fileds.get(k);
                String lowerCase = field.getField().toLowerCase();
                String comment = field.getComment();
                String extra = field.getExtra();
                if (ObjectUtils.isNotEmpty(extra)){
                    if (extra.equals("auto_increment")){
                        extra="自增";
                    }
                }
                // 设置下框线
                if (k + 1 == fileds.size()) {
                    addCell(table, lowerCase, 1);
                    addCell(table, StringUtils.isEmpty(comment) ? lowerCase : comment, 1);
                    addCell(table, field.getType(), 1);
                    addCell(table, field.getLength(), 1);
                    addCell(table, field.isNull() ? "是" : "否", 1);
                    addCell(table, field.getKey().equals("PRI") ? "是" : "否", 1);
                    addCell(table, extra, 1);
                    break;
                }
                addCell(table, lowerCase);
                addCell(table, StringUtils.isEmpty(comment) ? lowerCase : comment);
                addCell(table, field.getType());
                addCell(table, field.getLength());
                addCell(table, field.isNull() ? "是" : "否");
                addCell(table, field.getKey().equals("PRI") ? "是" : "否");
                addCell(table, extra);
            }
            document.add(table);
            printMsg("处理%s表结束", tableInfo);
        }
    }

    private static void addCell(Table table, String content, int flag) {
        addCell(table, content, -1, Element.ALIGN_CENTER, flag);
    }


    private static void addCell(Table table, String content) {
        addCell(table, content, -1, Element.ALIGN_CENTER);
    }

    private static void addCell(Table table, String content, int width, int align) {
        Font font = new Font(chinaTxtFont, cellFontSize);
        try {
            Cell cell = new Cell(new Paragraph(content, font));
            cell.disableBorderSide(15);
            table.addCell(cell);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void addCell(Table table, String content, int width, int align, int flag) {
        try {
            Font font = new Font(chinaTxtFont, cellFontSize);
            Cell cell = new Cell(new Paragraph(content, font));
            if (width > 0) cell.setWidth(width);
            cell.setHorizontalAlignment(align);
            //0---header,有上下边界,1----有下边界
            if (flag == 0) {
                cell.disableBorderSide(10);
                cell.setBorderColorTop(new Color(0, 0, 0));
                cell.setBorderWidthTop(3f);
                cell.setBorderColorBottom(new Color(0, 0, 0));
                cell.setBorderWidthBottom(3f);
            } else {
                cell.disableBorderSide(13);
                cell.setBorderColorBottom(new Color(0, 0, 0));
                cell.setBorderWidthBottom(3f);
            }

            table.addCell(cell);
        } catch (Exception e) {
            e.printStackTrace();
        }

    }



    private static void printMsg(String format, Object... args) {
        System.out.printf((format) + "%n", args);
    }

    private static List<TableInfo> getTableInfos(DataSource ds, String databaseName) throws SQLException {
        Connection conn = null;
        PreparedStatement stmt = null;
        ResultSet rs = null;
        List<TableInfo> list = Lists.newArrayList();
        try {
            conn = ds.getConnection();
            String sql = "select TABLE_NAME,TABLE_TYPE,TABLE_COMMENT from information_schema.tables where table_schema =? order by table_name";

            stmt = conn.prepareStatement(sql);
            setParameters(stmt, Collections.<Object>singletonList(databaseName));

            rs = stmt.executeQuery();
            while (rs.next()) {
                TableInfo row = new TableInfo();
                row.setTblName(rs.getString(1));
                row.setTblType(rs.getString(2));
                row.setTblComment(rs.getString(3));
                list.add(row);
            }
        } finally {
            JdbcUtils.close(rs);
            JdbcUtils.close(stmt);
            JdbcUtils.close(conn);
        }
        return list;
    }

    private static List<TableFiled> getTableFields(DataSource ds, String tblName) throws SQLException {
        Connection conn = null;
        PreparedStatement stmt = null;
        ResultSet rs = null;
        List<TableFiled> list = Lists.newArrayList();
        try {
            conn = ds.getConnection();
            //返回的列顺序是: Field,Type,Collation,Null,Key,Default,Extra,Privileges,Comment
            String sql = "SHOW FULL FIELDS FROM " + tblName;
            //返回的列顺序是: Field,Type,Null,Key,Default,Extra
            // sql = "show columns FROM " + tblName;

            stmt = conn.prepareStatement(sql);

            rs = stmt.executeQuery();
            while (rs.next()) {
                TableFiled field = new TableFiled();
                field.setField(rs.getString(1));
                String type = rs.getString(2);
                String length = "";
                if (type.contains("(")) {
                    int idx = type.indexOf("(");
                    length = type.substring(idx + 1, type.length() - 1);
                    type = type.substring(0, idx);
                }
                field.setType(type);
                field.setLength(length);
                field.setNull(rs.getString(4).equalsIgnoreCase("YES"));
                field.setKey(rs.getString(5));
                field.setDefaultVal(rs.getString(6));
                field.setExtra(rs.getString(7));
                field.setComment(rs.getString(9));
                list.add(field);
            }
        } finally {
            JdbcUtils.close(rs);
            JdbcUtils.close(stmt);
            JdbcUtils.close(conn);
        }
        return list;
    }

    private static void setParameters(PreparedStatement stmt, List<Object> parameters) throws SQLException {
        for (int i = 0, size = parameters.size(); i < size; ++i) {
            Object param = parameters.get(i);
            stmt.setObject(i + 1, param);
        }
    }

    public static DataSource getDataSource() {
        DruidDataSource datasource = new DruidDataSource();
        datasource.setUrl("jdbc:mysql://" + dbHost + ":" + dbPort + "/" + dbName + "?useUnicode=true&characterEncoding=utf-8&useSSL=false&allowPublicKeyRetrieval=true&nullCatalogMeansCurrent=true&useInformationSchema=true");
        datasource.setUsername(userName);
        datasource.setPassword(password);
        datasource.setDriverClassName("com.mysql.cj.jdbc.Driver");
        datasource.setInitialSize(1);
        datasource.setMinIdle(1);
        datasource.setMaxActive(3);
        datasource.setMaxWait(60000);
        return datasource;
    }
}
