package com.hqt.excel;

import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.metadata.Sheet;
import com.hqt.util.FileUtil;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class Main {
    public static void main(String[] args) throws IOException {
        InputStream inputStream = FileUtil.getResourcesFileInputStream("2007.xlsx");
        List<Object> datas = EasyExcelFactory.read(inputStream, new Sheet(1, 1));
        //遍历Excel每行数据
        for (Object data : datas) {
            //获取每行数据
            ArrayList line = (ArrayList) data;
            //获取第二列数据
            String a = (String) line.get(1);
            a += "3";
            //更新第二列数据
            line.set(1,a);
        }
        inputStream.close();

    }

}
