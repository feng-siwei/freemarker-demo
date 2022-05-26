package cn.logicalthinking.freemarker.utils;

import cn.hutool.core.io.FileUtil;
import cn.logicalthinking.freemarker.bean.ImageInfo;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.Element;
import org.dom4j.io.OutputFormat;
import org.dom4j.io.SAXReader;
import org.dom4j.io.XMLWriter;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.List;

/**
 * @author lanping
 * @version 1.0
 * @date 2020/5/28
 */
public class XmlParseUtil {

    /**
     * 解析document.xml.rels
     *
     * @param imageInfos
     * @param sourcePath
     * @param outPath
     */
    public static void parseDocXmlres(List<ImageInfo> imageInfos, String sourcePath, String outPath) throws IOException, DocumentException {
        FileUtil.mkParentDirs(outPath);
        SAXReader reader = new SAXReader();
        Document doc = reader.read(new File(sourcePath));
        Element root = doc.getRootElement();

        if (imageInfos != null && imageInfos.size() > 0) {
            for (ImageInfo image : imageInfos) {
                Element element = root.addElement("Relationship");
                element.addAttribute("Id", image.getImageRid());
                element.addAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
                //Target="media/image1.png"
                element.addAttribute("Target", "media/" + image.getFileName());
            }
        }
        writer(doc, outPath);
    }

    private static void writer(Document doc, String outPath) throws IOException {
        FileUtil.mkParentDirs(outPath);
        // 紧凑的格式
        OutputFormat format = OutputFormat.createCompactFormat();
        // 排版缩进的格式
        //OutputFormat format = OutputFormat.createPrettyPrint();
        // 设置编码
        format.setEncoding("UTF-8");
        // 创建XMLWriter对象,指定了写出文件及编码格式
        // XMLWriter writer = new XMLWriter(new FileWriter(new
        // File("src//a.xml")),format);
        XMLWriter writer = new XMLWriter(new OutputStreamWriter(
                new FileOutputStream(new File(outPath)), "UTF-8"), format);
        // 写入
        writer.write(doc);
        // 立即写入
        writer.flush();
        // 关闭操作
        writer.close();
    }

    public static void createNewXml(String sourcePath, String outPath) throws DocumentException, IOException {
        FileUtil.mkParentDirs(outPath);
        SAXReader reader = new SAXReader();
        Document doc = reader.read(new File(sourcePath));
        writer(doc, outPath);
    }


    public static void main(String[] args) {

    }

}
