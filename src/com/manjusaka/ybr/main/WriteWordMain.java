package com.manjusaka.ybr.main;

import com.manjusaka.ybr.util.WordUtil;

import java.util.*;

/**
 * function 程序入口
 * Created by 严彬荣 on 2017/9/6.
 */
public class WriteWordMain {
    public static void main(String[] arg) {
        String templateName = "template.ftl";
        String filePath = "D:/word";
        String imgs = "D:/图片/1492577537963.jpg";
        String fileName = new Date().getTime() + ".doc";
        Map<String, Object> map = new HashMap<>();
        List<Map<String, Object>> newsList = new ArrayList<>();
        map.put("name", "晏琦云");
        map.put("date", "2017-09-06 10:45");
        map.put("content", "测试内容测试内容测试内容测试内容测试内容测试内容测试内容测试内容测试内容测试内容测试内容测试内容测试内容");
        for (int i = 0; i < 100; ++i) {
            Map<String, Object> map1 = new HashMap<>();
            map1.put("title", "电商金融事业");
            map1.put("contents", "电商金融事业电商金融事业电商金融事业电商金融事业电商金融事业电商金融事业电商金融事业电商金融事业");
            map1.put("author", "晏琦云");
            newsList.add(map1);
        }
        map.put("newsList", newsList);
        map.put("imgs", WordUtil.getBase64ByImg(imgs));
        try {
            long start = System.currentTimeMillis();
            WordUtil.createWord(map, templateName, filePath, fileName);
            long end = System.currentTimeMillis();
            double time = end - start;
            System.out.println("word创建成功！用时：" + time / 1000 + " S");
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println("失败");
        }
    }
}
