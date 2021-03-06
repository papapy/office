package club.ljk.office.core;

import javax.xml.transform.*;
import javax.xml.transform.stream.StreamResult;
import javax.xml.transform.stream.StreamSource;
import java.io.*;

/**
 * Created by xkey on 2018/1/15.
 */
public class MathML {

    /**
     * <p>Description: xsl转换器</p>
     */
    public static String xslConvert(String s, String xslpath, URIResolver uriResolver){
        TransformerFactory tFac = TransformerFactory.newInstance();
        if(uriResolver != null)  tFac.setURIResolver(uriResolver);
        StringWriter writer = new StringWriter();
        try {
            StreamSource xslSource = new StreamSource(MathML.class.getClassLoader().getResourceAsStream(xslpath));
            Transformer t = tFac.newTransformer(xslSource);
            Source source = new StreamSource(new StringReader(s));
            Result result = new StreamResult(writer);
            t.transform(source, result);
        } catch (Exception e) {
            System.out.println(e);
        }
        return writer.getBuffer().toString();
    }

    /**
     * <p>Description: 将mathml转为latx </p>
     * @param mml
     * @return
     */
    public static String convertMML2Latex(String mml){
        mml = mml.substring(mml.indexOf("?>")+2, mml.length()); //去掉xml的头节点
        URIResolver r = new URIResolver(){  //设置xls依赖文件的路径
            public Source resolve(String href, String base) {
                return new StreamSource(MathML.class.getClassLoader().getResourceAsStream("conventer/" + href));
            }
        };
        String latex = xslConvert(mml, "conventer/mmltex.xsl", r);
        if(latex != null && latex.length() > 1){
            latex = latex.substring(1, latex.length() - 1);
        }
        return latex;
    }

    /**
     * <p>Description: office mathml转为mml </p>
     * @param xml
     * @return
     */
    public static String convertOMML2MML(String xml){
        String result = xslConvert(xml, "conventer/omml2mml.xsl", null);
        return result;
    }
}
