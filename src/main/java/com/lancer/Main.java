package com.lancer;

import com.lowagie.text.pdf.BaseFont;

import javax.sql.DataSource;
import java.util.Scanner;

public class Main {
    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);
        System.out.print("请输入生成路径：");
        Utils.out_path = scanner.nextLine();

        System.out.print("请输入数据库名：");
        Utils.dbName = scanner.nextLine();

        System.out.print("请输入密码：");
        Utils.password = scanner.nextLine();

        try {
            // 设置中文字体为宋体
            Utils.chinaTxtFont = BaseFont.createFont("simsun.ttc,0", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
            DataSource ds = Utils.getDataSource();
            Utils.mysqlToWordTb(ds);
        } catch (Exception e) {
            throw new RuntimeException(e);
        } finally {
            scanner.close();
        }
    }
}