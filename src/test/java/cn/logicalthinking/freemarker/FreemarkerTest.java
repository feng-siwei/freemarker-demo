package cn.logicalthinking.freemarker;

import cn.hutool.core.date.DateUtil;
import cn.hutool.core.io.FileUtil;
import cn.hutool.core.util.StrUtil;
import cn.logicalthinking.freemarker.bean.ImageInfo;
import cn.logicalthinking.freemarker.bean.User;
import cn.logicalthinking.freemarker.utils.WordUtils;
import cn.logicalthinking.freemarker.utils.XmlParseUtil;
import org.junit.Test;

import java.io.File;
import java.util.*;

public class FreemarkerTest {

    private static String rootDir = "C:\\lanping\\logic-project\\freemarker-demo\\doc\\";
    /**
     * 模板目录
     */
    private static String TEMPLATE_DIR = rootDir + "template\\";

    /**
     * 生成word 目录
     */
    private static String TARGET_DIR = rootDir +"target\\";

    /**
     * 模板 document.xml.rels
     */
    private static String TEMPLATE_XML_RELS = "document.xml.rels";

    /**
     * 模板 document.xml
     */
    private static String TEMPLATE_DOCUMENET_XML = "document.xml";

    /**
     * 模板 freemarker_template.docx
     */
    private static String TEMPLATE_DOCX = "freemarker_template.docx";

    /**
     * 生成word文档测试
     * 替换word模板中的占位符
     */
    @Test
    public void createWordTest() throws Exception {
        Map<String,Object> map = new HashMap<>();

        String poemStr1 = "地势坤，君子以厚德载物";
        String poemStr2 = "莫使金樽空对月";

        map.put("poemStr1",poemStr1);
        map.put("poemStr2",poemStr2);

        List<User> users = new ArrayList<>();
        for(int i=0;i<6;i++){
            User user = new User("张"+(i+1),i%2==0?"男":"女",20+i);
            users.add(user);
        }
        map.put("users",users);

        String path = rootDir+"images\\1.jpg";

        List<ImageInfo> imageInfos = new ArrayList<>();

        ImageInfo image1 = new ImageInfo();
        writeFiles(imageInfos, image1,new File(path),"案例图片1");
        map.put("image1",image1);

        path = rootDir+"images\\2.jpg";
        ImageInfo image2 = new ImageInfo();
        writeFiles(imageInfos, image2,new File(path),"案例图片2");
        map.put("image2",image2);

        writeWordDocx(imageInfos,map);

    }

    /**
     * 替换书签
     */
    @Test
    public void replaceBookMarkTest() throws Exception {
        Map<String,String> map = new HashMap<>();

        String sourcePath = TARGET_DIR + "freemarker_1618386223987.docx";
        String destPath = TARGET_DIR +"freemarker_替换书签_"+System.currentTimeMillis()+".docx";
        String signImage = rootDir+"images\\tamp.png";

        map.put("signName","张三");
        map.put("signDate", DateUtil.formatChineseDate(new Date(),false,false));
        map.put("signImage", signImage);

        WordUtils.updateDocxMarks(map,sourcePath,destPath,"signImage");

        System.out.println("执行完成");

    }

    /**
     * 生成word
     * @param imageInfos
     * @param reportMap
     * @throws Exception
     */
    public void writeWordDocx(List<ImageInfo> imageInfos, Map<String,Object> reportMap) throws Exception {
        //解析 document.xml_template.rels
        String sourcePath = TEMPLATE_DIR + TEMPLATE_XML_RELS;
        String resOutPath = TARGET_DIR + System.currentTimeMillis()+"_document.xml.rels";
        XmlParseUtil.parseDocXmlres(imageInfos, sourcePath, resOutPath);

        //替换document.xml
        String outFormatDocXmlPath = TARGET_DIR + System.currentTimeMillis()+"_document.xml";
        String outDocXmlPath = TARGET_DIR + System.currentTimeMillis()+"_document_new.xml";
        WordUtils.exportWord(reportMap, TEMPLATE_DIR, TEMPLATE_DOCUMENET_XML, outFormatDocXmlPath);
        XmlParseUtil.createNewXml(outFormatDocXmlPath, outDocXmlPath);

        String docxFileName = "freemarker_"+System.currentTimeMillis()+".docx";

        //生成docx
        String docxTemp = TEMPLATE_DIR + TEMPLATE_DOCX;
        WordUtils.createDocx(outDocXmlPath, resOutPath, null, null,
                imageInfos, docxTemp, TARGET_DIR + docxFileName);

        System.out.println("生成Word文件："+TARGET_DIR + docxFileName);

        //删除临时文件
        new File(resOutPath).delete();
        new File(outFormatDocXmlPath).delete();
        new File(outDocXmlPath).delete();
        for (ImageInfo image : imageInfos) {
            if (StrUtil.isNotEmpty(image.getFilePath())) {
                new File(image.getFilePath()).delete();
            }
        }
    }


    /**
     * 读取文件
     * @param imageInfos
     * @param busiImage
     * @param file
     * @param remark
     */
    private void writeFiles(List<ImageInfo> imageInfos, ImageInfo busiImage, File file,String remark){
        String targetFilePath = TARGET_DIR + System.currentTimeMillis()+file.getName().substring(file.getName().lastIndexOf("."));
        FileUtil.copy(file,new File(targetFilePath),true);

        busiImage.setIndex(imageInfos.size()+1);
        busiImage.setImageId(String.valueOf(imageInfos.size()+1));
        busiImage.setFileName(file.getName());
        busiImage.setRemark(remark);
        busiImage.setImageRid("rId"+(100+imageInfos.size()));
        busiImage.setFilePath(targetFilePath);
        imageInfos.add(busiImage);
    }

}
