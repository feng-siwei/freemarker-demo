package cn.logicalthinking.freemarker.utils;

import cn.hutool.core.io.FileUtil;
import cn.hutool.core.util.StrUtil;
import cn.logicalthinking.freemarker.bean.ImageInfo;
import com.aspose.words.*;
import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;
import freemarker.template.Version;

import java.io.*;
import java.util.Enumeration;
import java.util.List;
import java.util.Map;
import java.util.UUID;
import java.util.zip.ZipEntry;
import java.util.zip.ZipException;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

/**
 * word工具类
 *
 * @author lanping
 * @version 1.0
 * @date 2020/4/28
 */
public class WordUtils {

    /**
     * 根据模板导出word
     *
     * @param params       参数
     * @param tempDir      模板目录
     * @param tempFileName 模板名称（ftl）
     * @param outFilePath  生成word路径
     * @throws IOException
     * @throws TemplateException
     */
    public static void exportWord(Map<String, Object> params,
                                  String tempDir, String tempFileName, String outFilePath) throws IOException, TemplateException {

        FileUtil.mkParentDirs(outFilePath);

        Configuration configuration = new Configuration(new Version("2.3.0"));
        configuration.setDefaultEncoding("utf-8");

        configuration.setDirectoryForTemplateLoading(new File(tempDir));

        //输出文档路径及名称
        File outFile = new File(outFilePath);

        //以utf-8的编码读取ftl文件
        Template template = configuration.getTemplate(tempFileName, "utf-8");
        Writer out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(outFile),
                "utf-8"), 10240);
        template.process(params, out);
        out.close();
    }

    /**
     * @param documentXmlFile 动态生成数据的docunment.xml文件
     * @param docResFile      word/_rels/document.xml.rels 动态生成资源引入文件
     * @param docxTemplate    docx的模板
     * @param toFilePath      需要导出的文件路径
     * @throws ZipException
     * @throws IOException
     */
    public static void createDocx(String documentXmlFile, String docResFile,
                                  String headerFile, String footerFile,
                                  List<ImageInfo> imageInfos,
                                  String docxTemplate, String toFilePath) throws Exception {

        FileUtil.mkParentDirs(toFilePath);

        //输出文档路径及名称
        ZipOutputStream zipout = null;
        ZipFile zipFile = null;
        try {
            File docxFile = new File(docxTemplate);
            zipFile = new ZipFile(docxFile);
            Enumeration<? extends ZipEntry> zipEntrys = zipFile.entries();
            zipout = new ZipOutputStream(new FileOutputStream(toFilePath));
            int len = -1;
            byte[] buffer = new byte[1024];

            if (imageInfos != null && imageInfos.size() > 0) {
                for (ImageInfo imageInfo : imageInfos) {
                    zipout.putNextEntry(new ZipEntry("word/media/" + imageInfo.getFileName()));
                    InputStream in = new FileInputStream(new File(imageInfo.getFilePath()));
                    while ((len = in.read(buffer)) != -1) {
                        zipout.write(buffer, 0, len);
                    }
                    in.close();
                }
            }

            while (zipEntrys.hasMoreElements()) {
                ZipEntry next = zipEntrys.nextElement();
                InputStream is = zipFile.getInputStream(next);
                // 把输入流的文件传到输出流中 如果是word/document.xml由我们输入
                zipout.putNextEntry(new ZipEntry(next.toString()));

                if ("word/document.xml".equals(next.toString()) && StrUtil.isNotEmpty(documentXmlFile)) {
                    File file = new File(documentXmlFile);
                    if (file.exists()) {
                        InputStream in = new FileInputStream(file);
                        while ((len = in.read(buffer)) != -1) {
                            zipout.write(buffer, 0, len);
                        }
                        in.close();
                    }
                } else if ("word/header1.xml".equals(next.toString()) && StrUtil.isNotEmpty(headerFile)) {
                    File file = new File(headerFile);
                    if (file.exists()) {
                        InputStream in = new FileInputStream(file);
                        while ((len = in.read(buffer)) != -1) {
                            zipout.write(buffer, 0, len);
                        }
                        in.close();
                    }
                } else if ("word/footer1.xml".equals(next.toString()) && StrUtil.isNotEmpty(footerFile)) {
                    File file = new File(footerFile);
                    if (file.exists()) {
                        InputStream in = new FileInputStream(file);
                        while ((len = in.read(buffer)) != -1) {
                            zipout.write(buffer, 0, len);
                        }
                        in.close();
                    }
                } else if ("word/_rels/document.xml.rels".equals(next.toString()) && StrUtil.isNotEmpty(docResFile)) {
                    File file = new File(docResFile);
                    if (file.exists()) {
                        InputStream in = new FileInputStream(file);
                        while ((len = in.read(buffer)) != -1) {
                            zipout.write(buffer, 0, len);
                        }
                        in.close();
                    }
                } else {
                    while ((len = is.read(buffer)) != -1) {
                        zipout.write(buffer, 0, len);
                    }
                }
                is.close();
            }
        } catch (Exception e) {
            throw new Exception("创建word文档失败", e);
        } finally {
            if (zipout != null) {
                zipout.close();
            }
            if (zipFile != null) {
                zipFile.close();
            }
        }
    }

    /**
     * xmlword转word
     *
     * @param sourceFile
     * @param destFile
     * @throws Exception
     */
    public static void xmlWord2Word(String sourceFile, String destFile) throws Exception {
        try {
            validLicense();
            FileUtil.mkParentDirs(destFile);
            File file = new File(destFile);

            FileOutputStream os = new FileOutputStream(file);
            Document doc = new Document(sourceFile);
            doc.save(os, SaveFormat.DOC);
            os.close();
            File soFile = new File(sourceFile);
            soFile.delete();
        } catch (Exception e) {
            throw new Exception("生成WORD文档失败!", e);
        }
    }

    /**
     * word转pdf
     *
     * @param sourceFile
     * @param destFile
     * @throws Exception
     */
    public static void word2Pdf(String sourceFile, String destFile) throws Exception {
        try {
            validLicense();
            FileUtil.mkParentDirs(destFile);
            File file = new File(destFile);  //新建一个空白pdf文档
            FileOutputStream os = new FileOutputStream(file);
            Document doc = new Document(sourceFile);//通过sourceFile创建word文档对象
            doc.save(os, SaveFormat.PDF);
            os.close();
        } catch (Exception e) {
            throw new Exception("生成PDF文档失败!", e);
        }
    }

    /**
     * 修改书签
     *
     * @param map
     * @param sourceFile
     * @param destFile
     * @param imageMarks
     * @throws Exception
     */
    public static void updateDocxMarks(Map<String, String> map, String sourceFile, String destFile,String ... imageMarks) throws Exception {
        try {
            validLicense();
            FileUtil.mkParentDirs(destFile);
            FileOutputStream os = new FileOutputStream(destFile);
            Document doc = new Document(sourceFile);
            BookmarkCollection bookmarks = doc.getRange().getBookmarks();
            for (Bookmark bookmark : bookmarks) {
                for (Map.Entry<String, String> entry : map.entrySet()) {
                    if (bookmark != null && entry.getKey().equals(bookmark.getName())) {
                        if(imageMarks!=null && imageMarks.length>0){
                            for(String imageMark:imageMarks){
                                if(imageMark.equals(bookmark.getName())){
                                    DocumentBuilder builder = new DocumentBuilder(doc);
                                    builder.moveToBookmark(bookmark.getName());
                                    Shape shape = builder.insertImage(entry.getValue(), RelativeHorizontalPosition.MARGIN, 260,
                                            RelativeVerticalPosition.MARGIN, 260, 140, 140, WrapType.NONE);
                                    shape.setBehindText(true);
                                }else{
                                    bookmark.setText(entry.getValue());
                                }
                            }
                        }
                    }
                }
            }
            doc.save(os, SaveFormat.DOCX);
            os.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * aspose授权
     *
     * @return
     */
    public static void validLicense() {
        try {
            String licenseXml = "<License><Data><Products><Product>Aspose.Total for Java</Product><Product>Aspose.Words for Java</Product></Products><EditionType>Enterprise</EditionType><SubscriptionExpiry>20991231</SubscriptionExpiry><LicenseExpiry>20991231</LicenseExpiry><SerialNumber>8bfe198c-7f0c-4ef8-8ff0-acc3237bf0d7</SerialNumber></Data><Signature>sNLLKGMUdF0r8O1kKilWAGdgfs2BvJb/2Xp8p5iuDVfZXmhppo+d0Ran1P9TKdjV4ABwAgKXxJ3jcQTqE/2IRfqwnPf8itN8aFZlV3TJPYeD3yWE7IT55Gz6EijUpC7aKeoohTb4w2fpox58wWoF3SNp6sK6jDfiAUGEHYJ9pjU=</Signature></License>";
            ByteArrayInputStream inputStream = new ByteArrayInputStream(licenseXml.getBytes());
            com.aspose.words.License license = new com.aspose.words.License();
            license.setLicense(inputStream);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    public static void main(String[] args) throws Exception {
        /*Map<String, Object> map = new HashMap<>();
        map.put("badType_A", "Y");
        String tempDir = "C:\\Users\\User\\Desktop";
        String tempFileName = "quality_book_document.xml";
        String destFile = "C:\\Users\\User\\Desktop\\quality_book_document2.xml";
        // updateDocxMarks(map,docxFile,destFile);
        exportWord(map,tempDir, tempFileName, destFile);*/


        WordUtils.word2Pdf("C:\\Users\\User\\Desktop\\fire_report.docx",
                "C:\\Users\\User\\Desktop\\" + UUID.randomUUID().toString() + ".pdf");
    }

}

