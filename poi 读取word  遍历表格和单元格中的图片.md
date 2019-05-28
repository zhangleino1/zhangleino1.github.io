# 背景
项目需要解析word表格
![在这里插入图片描述](https://img-blog.csdnimg.cn/20190526214541805.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L2h1YW5nbWluZ2xlaWx1bw==,size_16,color_FFFFFF,t_70)
- 需要批量导入系统，并保存每行信息到数据库
- 并且要保存word中的图片，
- 并保持每条信息和图片的对应关系
- 一行数据可能有多条图片
# 解决办法
没有找到现成的代码，怎么办呐？看源码吧
# 分享快乐
给出代码

```java
package com.util;

import org.apache.poi.xwpf.usermodel.*;
import org.jeecgframework.core.common.model.json.AjaxJson;
import org.jeecgframework.poi.word.entity.MyXWPFDocument;
import org.jeecgframework.poi.word.parse.excel.ExcelEntityParse;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.imageio.stream.FileImageOutputStream;
import java.io.*;
import java.util.Iterator;
import java.util.List;

public class WordImportUtil {
    private static final Logger logger = LoggerFactory.getLogger(WordImportUtil.class);

    public static MyXWPFDocument getXWPFDocumen(InputStream is) {
        try {

            MyXWPFDocument doc = new MyXWPFDocument(is);
            return doc;
        } catch (Exception e) {
            logger.error(e.getMessage(), e);
        } finally {
            try {
                is.close();
            } catch (Exception e) {
                logger.error(e.getMessage(), e);
            }
        }
        return null;
    }

    public static AjaxJson parseThisTable(MyXWPFDocument doc){
        Iterator<XWPFTable> itTable = doc.getTablesIterator();
        XWPFTable table;
        while (itTable.hasNext()) {
            table = itTable.next();

            XWPFTableRow row;
            List<XWPFTableCell> cells;
            Object listobj;

            ExcelEntityParse excelEntityParse = new ExcelEntityParse();
            for (int i = 0; i < table.getNumberOfRows(); i++) {
                if(i ==0)
                    continue;
                row = table.getRow(i);
                cells = row.getTableCells();
                for (int j = 0; j < cells.size(); j++) {
                    XWPFTableCell cell = cells.get(j);
                    if(j == 10){
                        getCellImage(cell);
                    }

                    //输出当前的单元格的数据
                    System.out.print(cell.getText() + "\t");
                }

            }

        }
        return null;
    }

    public static  String getCellImage(XWPFTableCell cell){
        List<XWPFParagraph> xwpfParagraphs =  cell.getParagraphs();
        if(xwpfParagraphs == null) return null;
        for(XWPFParagraph xwpfParagraph:xwpfParagraphs){
            List<XWPFRun> xwpfRunList = xwpfParagraph.getRuns();
            if(xwpfRunList==null) return null;
            for(XWPFRun xwpfRun:xwpfRunList){
                List<XWPFPicture> xwpfPictureList =  xwpfRun.getEmbeddedPictures();
                if(xwpfParagraph==null) return null;
                for(XWPFPicture xwpfPicture:xwpfPictureList){
                    xwpfPicture.getPictureData().getData();
                    xwpfPicture.getPictureData().getFileName();
                    byte2image( xwpfPicture.getPictureData().getData(),"d:/"+ xwpfPicture.getPictureData().getFileName());
                }
            }
        }
        return "";
    }

    public static  void byte2image(byte[] data,String path){
        if(data.length<3||path.equals("")) return;
        FileImageOutputStream imageOutput = null;
        try{
            imageOutput = new FileImageOutputStream(new File(path));
            imageOutput.write(data, 0, data.length);
            System.out.println("Make Picture success,Please find image in " + path);
        } catch(Exception ex) {
            System.out.println("Exception: " + ex);
            ex.printStackTrace();
        }finally {
            try {
                imageOutput.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    public static void main(String[] args) throws Exception{
        MyXWPFDocument myXWPFDocument = getXWPFDocumen(new FileInputStream("d:/园艺作物加工副产物适宜性评价填写.docx"));
        parseThisTable(myXWPFDocument);
    }

}

```

