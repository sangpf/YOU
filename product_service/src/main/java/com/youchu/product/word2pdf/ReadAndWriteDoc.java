package com.youchu.product.word2pdf;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.hwpf.*;
import org.apache.poi.hwpf.model.FieldsDocumentPart;
import org.apache.poi.hwpf.usermodel.*;
import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFComment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class ReadAndWriteDoc {

    /**
     * 实现对word读取和修改操作(word2003.doc)
     */
    public static void readwriteWord1(String filePath, Map<String,String> map){
        //读取word模板
        FileInputStream in = null;
        try {
            in = new FileInputStream(new File(filePath));
        } catch (FileNotFoundException e1) {
            e1.printStackTrace();
        }
        HWPFDocument hdt = null;
        try {
            hdt = new HWPFDocument(in);
        } catch (IOException e1) {
            e1.printStackTrace();
        }

        Fields fields = hdt.getFields();
        Iterator<Field> it = fields.getFields(FieldsDocumentPart.MAIN).iterator();
        while(it.hasNext()){
            System.out.println(it.next().getType());
        }

        //读取word文本内容
        Range range = hdt.getRange();
        System.out.println(range.text());
        //替换文本内容
        for (Map.Entry<String,String> entry: map.entrySet()) {
            range.replaceText(entry.getKey() ,entry.getValue());
        }
        ByteArrayOutputStream ostream = new ByteArrayOutputStream();
//        String fileName = System.currentTimeMillis()+filePath.substring(filePath.lastIndexOf("/")+1, filePath.length());
        FileOutputStream out = null;
        try {
            out = new FileOutputStream("D:\\newdocModel.doc",true);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        try {
            hdt.write(ostream);
        } catch (IOException e) {
            e.printStackTrace();
        }
        //输出字节流
        try {
            out.write(ostream.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
        }
        try {
            out.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        try {
            ostream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 实现对word读取和修改操作(word2007.docx)
     */
    public static void readwriteWord2(String filePath, Map<String,String> map){
        try {
            OPCPackage pack = POIXMLDocument.openPackage(filePath);
            XWPFDocument doc = new XWPFDocument(pack);

            // -------------
            XWPFComment[] comments = doc.getComments();
            for (XWPFComment comment: comments){
                String author = comment.getAuthor();
                String id = comment.getId();
                String text = comment.getText();
            }

            //---输出书签信息

            //-------
            List<XWPFParagraph> paragraphs = doc.getParagraphs();
            System.out.println("paragraphs size : " + paragraphs.size());
            for (XWPFParagraph tmp : paragraphs) {
                System.out.println("ParagraphText : " + tmp.getParagraphText());
                List<XWPFRun> runs = tmp.getRuns();
                for (XWPFRun aa : runs) {
                    System.out.println("Text : " + aa.getText(0));
                    for (Map.Entry<String,String> entry: map.entrySet()) {
                        if (aa.getText(0) != null && aa.getText(0).contains(entry.getKey())) {
                            aa.setText(entry.getValue(), 0);
                        }
                    }
                }
            }
            String fileName = System.currentTimeMillis()+ ".docx";
            FileOutputStream fos = new FileOutputStream("D:\\"+fileName,true);
            doc.write(fos);
            fos.flush();
            fos.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 输出书签信息
     * @param bookmarks
     */
    private static void printInfo(Bookmarks bookmarks) {
        int count = bookmarks.getBookmarksCount();
        System.out.println("书签数量：" + count);
        Bookmark bookmark;
        for (int i=0; i<count; i++) {
            bookmark = bookmarks.getBookmark(i);
            System.out.println("书签" + (i+1) + "的名称是：" + bookmark.getName());
            System.out.println("开始位置：" + bookmark.getStart());
            System.out.println("结束位置：" + bookmark.getEnd());
        }
    }

    public static void main(String[] args) {
        String filePath = "D:\\123.docx";
        Map<String,String> map = new HashMap<String, String>();
        map.put("23", "hello!");
        map.put("123", "world!");
        map.put("name", "jake");
        map.put("age", "89");
        readwriteWord2(filePath, map);
    }

}
