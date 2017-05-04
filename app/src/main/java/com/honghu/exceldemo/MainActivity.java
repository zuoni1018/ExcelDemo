package com.honghu.exceldemo;

import android.os.Bundle;
import android.support.v7.app.AppCompatActivity;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

import static com.honghu.exceldemo.FileUtil.getSDPath;

public class MainActivity extends AppCompatActivity {


    private String[] tabName = {"姓名", "性别", "年龄", "职位", "籍贯", "备注"};
    private List<People> mList;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
//        File file;

        mList = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            People people = new People();

            people.setName("张三");
            people.setGender("男");
            people.setAge(i);
            people.setPosition("Android");
            people.setPlace("温州");
            people.setRemarks("无");
            mList.add(people);
        }
        File file = new File(getSDPath() + "/HongHuDate");
        FileUtil.makeDir(file);
        createExcel(file.toString() + "/Data.xls");


    }


    public void createExcel(String fileName) {
        WritableWorkbook mWritableWorkbook;
        WritableSheet ws;
        try {
            File file = new File(fileName);
            //每次都重新生成当前文件
            if (file.exists()) {
                file.delete();
            }

            mWritableWorkbook = Workbook.createWorkbook(file);
            // 创建表单,其中sheet表示该表格的名字,0表示第一个表格,
            ws = mWritableWorkbook.createSheet("测试表", 0);
            // 在指定单元格插入数据
            //备注： 第一个参数表示,0表示第一列,第二个参数表示行,同样0表示第一行,第三个参数表示想要添加到单元格里的数据.

            //插入第一行
            for (int i = 0; i < tabName.length; i++) {
                Label mLabel = new Label(i, 0, tabName[i]);
                ws.addCell(mLabel);
            }

            //插入后面几行

            for (int i = 0; i <mList.size() ; i++) {
                //5个为一行
                Label mLabel0 = new Label(0, i+1, mList.get(i).getName()+"");
                Label mLabel1 = new Label(1, i+1,mList.get(i).getGender());
                Label mLabel2 = new Label(2, i+1, mList.get(i).getAge()+"");
                Label mLabel3 = new Label(3, i+1,mList.get(i).getPosition());
                Label mLabel4 = new Label(4, i+1, mList.get(i).getPlace());
                Label mLabel5 = new Label(5, i+1, mList.get(i).getRemarks());

                ws.addCell(mLabel0);
                ws.addCell(mLabel1);
                ws.addCell(mLabel2);
                ws.addCell(mLabel3);
                ws.addCell(mLabel4);
                ws.addCell(mLabel5);
            }
            mWritableWorkbook.write();
            mWritableWorkbook.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
