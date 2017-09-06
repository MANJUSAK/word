package com.manjusaka.ybr.util;

import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;
import org.apache.poi.POIXMLDocument;
import org.apache.poi.POIXMLTextExtractor;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import sun.misc.BASE64Encoder;

import java.io.*;
import java.util.Map;

/**
 * function word操作工具类
 * Created by 严彬荣 on 2017/9/6.
 */
@SuppressWarnings("ALL")
public class WordUtil {
    /**
     * @param map      word中需要展示的动态数据，用map集合来保存
     * @param template word模板名称，例如：test.ftl
     * @param savePath 文件生成的目标路径，例如：D:/wordFile/
     * @param newFile  生成的文件名称，例如：test.doc
     */
    public static void createWord(Map<String, Object> map, String template, String savePath, String newFile) throws IOException, TemplateException {
        //创建配置实例
        Configuration configuration = new Configuration();
        //设置编码
        configuration.setDefaultEncoding("UTF-8");
        //ftl模板文件统一放至 com.manjusaka.ybr.template 包下面
        configuration.setClassForTemplateLoading(WordUtil.class, "/com/manjusaka/ybr/template/");
        //获取模板
        Template temp = configuration.getTemplate(template);
        //输出文件
        File outFile = new File(savePath + File.separator + newFile);
        //如果输出目标文件夹不存在，则创建
        if (!outFile.getParentFile().exists()) {
            outFile.getParentFile().mkdirs();
        }
        //将模板和数据模型合并生成文件
        Writer out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(outFile), "UTF-8"));
        //生成文件
        temp.process(map, out);
        //关闭流
        out.flush();
        out.close();
    }

    /**
     * 读取word文档数据
     *
     * @param path 文件路径
     * @return 读取数据
     * @throws Exception
     */
    public static String readWord(String path) throws Exception {
        //获取文档
        OPCPackage opcPackage = POIXMLDocument.openPackage(path);
        //打开docx文件xml
        XWPFDocument xwpf = new XWPFDocument(opcPackage);
        //读取xml
        POIXMLTextExtractor extractor = new XWPFWordExtractor(xwpf);
        //WordExtractor extractor = new WordExtractor(is);
//        CTBody ct = xwpf.getDocument().getBody();
        //读取内容
        String data = extractor.getText();
        /*String data = ct.xmlText();
        String prefix = data.substring(0, data.indexOf(">") + 1);
        String mainPart = data.substring(data.indexOf(">") + 1, data.lastIndexOf("<"));
        String sufix = data.substring(data.lastIndexOf("<"));*/
        return data;
    }

    /**
     * 将图片转换为base64编码字符
     *
     * @param path 图片路径
     * @return base64字符
     */
    public static String getBase64ByImg(String path) {
        File file = new File(path);
        FileInputStream is = null;
        try {
            is = new FileInputStream(file);
            //读取字节数组
            byte[] data = new byte[is.available()];
            is.read(data);
            //转码
            BASE64Encoder encoder = new BASE64Encoder();
            is.close();
            return encoder.encode(data);
        } catch (Exception e) {
            e.printStackTrace();
            return "";
        } finally {
            if (is != null) try {
                is.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }


    }
}
