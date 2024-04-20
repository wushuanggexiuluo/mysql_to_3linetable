package com.lancer;

import com.lowagie.text.pdf.BaseFont;
import org.springframework.boot.CommandLineRunner;
import org.springframework.stereotype.Component;

import javax.sql.DataSource;
import java.util.Scanner;

@Component
public class RunServe implements CommandLineRunner {
    @Override
    public void run(String... args) throws Exception {
        Scanner scanner = new Scanner(System.in);
        System.out.print("欢迎使用数据库生成三线表程序 \n注意：默认的用户名是root,如果更换请在WordUtils的userName中替换\n --请输入生成路径（路径示例：D://abc/a）\n");
        String path = scanner.nextLine();
        boolean res = PathUtils.checkPath(path);
        if (!res){
            return;
        }
        WordUtils.out_path = path;

        System.out.print("--请数据库输入密码\n");
        try{
            WordUtils.password = scanner.nextLine();

        }
        catch (Exception e){
            System.out.println("数据库密码不正确。");
        }
        DataSource source = WordUtils.getDataSource();
        WordUtils.showAllDbName(source);

        System.out.print("--请输入数据库名\n");
        WordUtils.dbName = scanner.nextLine();


        try {
            // 设置中文字体为宋体
            WordUtils.chinaTxtFont = BaseFont.createFont("simsun.ttc,0", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
            DataSource ds = WordUtils.getDataSource();
            WordUtils.mysqlToWordTb(ds);

        } catch (Exception e) {
            throw new RuntimeException(e);
        } finally {
            scanner.close();
            System.out.println("生成的文件已经保存在："+path+"路径下的 "+ WordUtils.dbName+".doc");
        }
    }
}
