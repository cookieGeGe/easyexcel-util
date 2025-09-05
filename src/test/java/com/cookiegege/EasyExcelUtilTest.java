package com.cookiegege;

import com.cookiegege.entity.DemoData;
import org.junit.Test;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @author JoSuper
 * @date 2025/9/5 10:43
 */
public class EasyExcelUtilTest {

    public List<DemoData> createDemoDataList() {
        List<DemoData> list = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            DemoData data = new DemoData();
            data.setString("字符串" + i);
            data.setDate(new Date());
            data.setDoubleData(0.56);
            list.add(data);
        }
        return list;
    }

    @Test
    public void testExport() {
        List<DemoData> list = createDemoDataList();
        try {
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            EasyExcelUtil.exportExcel(
                    list,
                    "测试",
                    DemoData.class,
                    true,
                    outputStream
            );

            FileOutputStream fileOutputStream = new FileOutputStream("D:\\test.xlsx");
            fileOutputStream.write(outputStream.toByteArray());
            fileOutputStream.flush();
            fileOutputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

}
