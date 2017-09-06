package com.manjusaka.ybr.main;

import com.manjusaka.ybr.util.WordUtil;

/**
 * function 读取word文件数据
 * Created by 严彬荣 on 2017/9/6.
 */
public class ReadWordMain {

    public static void main(String[] arg) {
        String file = "C:\\Users\\ASUS\\Desktop\\文档\\工作\\园林大数据平台\\需求发布接口文档.docx";
        try {
            long start = System.currentTimeMillis();
            String data = WordUtil.readWord(file);
            long end = System.currentTimeMillis();
            double time = end - start;
            System.out.println("读取成功！用户：" + time / 1000 + " S");
            System.out.println("文档数据如下：");
            System.out.println(data);
        } catch (Exception e) {
            System.out.println("读取失败...");
            e.printStackTrace();
        }
    }
}
