package club.ljk.office.core;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.dom4j.Document;
import org.dom4j.DocumentFactory;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTDrawing;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTOneCellAnchor;
import org.xml.sax.InputSource;

import java.io.FileInputStream;
import java.io.StringReader;
import java.util.*;

/**
 * Created by xkey on 2018/1/8.
 */
public class ExcelReader {


    public static void reader() throws Exception {
        String excelPath = "/home/xkey/Downloads/case.xlsx";
        FileInputStream fis = new FileInputStream(excelPath);

        Workbook workbook = WorkbookFactory.create(fis);
        Sheet sheet = workbook.getSheetAt(0);

        Drawing drawing = sheet.getDrawingPatriarch();
        XSSFDrawing xssfDrawing = (XSSFDrawing) drawing;
        CTDrawing ctDrawing = xssfDrawing.getCTDrawing();
        CTOneCellAnchor[] ctDrawingOneCellAnchorArray = ctDrawing.getOneCellAnchorArray();

        for(CTOneCellAnchor ctOneCellAnchor : ctDrawingOneCellAnchorArray) {
            String xml = ctOneCellAnchor.xmlText();
            String omml = xml2Omml(xml);
            String mml = MathML.convertOMML2MML(omml);
            String latex = MathML.convertMML2Latex(mml);
            System.out.println(xml);
            System.out.println(omml);
            System.out.println(mml);
            System.out.println(latex);
        }


        /*Iterator iterator = drawing.iterator();
        while (iterator.hasNext()) {
            Object o = iterator.next();
            if (o instanceof XSSFPicture) {
                XSSFPicture picture = (XSSFPicture) o;
                XSSFClientAnchor clientAnchor = picture.getPreferredSize();
                CTMarker marker = clientAnchor.getFrom();

                System.out.println(marker.getRow());
                System.out.println(marker.getCol());
                System.out.println(picture.getPictureData().getData());
                FileUtils.writeByteArrayToFile(new File("/home/xkey/Documents/case.jpg"), picture.getPictureData().getData());
            }
        }*/
    }

    public static String xml2Omml(String xml) throws Exception {
        //dom4j解析器的初始化
        SAXReader reader = new SAXReader(new DocumentFactory());
        Map<String, String> map = new HashMap<String, String>();
        map.put("xdr","http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
        map.put("m","http://schemas.openxmlformats.org/officeDocument/2006/math");
        reader.getDocumentFactory().setXPathNamespaceURIs(map); //xml文档的namespace设置

        InputSource source = new InputSource(new StringReader(xml));
        source.setEncoding("utf-8");
        Document document = reader.read(source);
        Element rootElement = document.getRootElement();
        Element element = (Element)rootElement.selectSingleNode("//m:oMathPara");    //用xpath得到OMML节点
        return element.asXML();
    }



    public static void main(String[] arg) throws Exception {
        reader();
    }
}
