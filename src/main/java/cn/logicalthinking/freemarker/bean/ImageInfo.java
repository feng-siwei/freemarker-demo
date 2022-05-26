package cn.logicalthinking.freemarker.bean;

/**
 * word照片信息
 * @author lanping
 * @version 1.0
 * @date 2021/04/14
 **/
public class ImageInfo {

    /**
     * 序号
     */
    private int index;

    /**
     * 照片ID 数值
     */
    private String imageId;

    /**
     * 照片名称
     */
    private String fileName;

    /**
     * 备注
     */
    private String remark;

    /**
     * 照片rId
     */
    private String imageRid;

    /**
     * 文件全路径
     */
    private String filePath;

    public int getIndex() {
        return index;
    }

    public void setIndex(int index) {
        this.index = index;
    }

    public String getImageId() {
        return imageId;
    }

    public void setImageId(String imageId) {
        this.imageId = imageId;
    }

    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public String getRemark() {
        return remark;
    }

    public void setRemark(String remark) {
        this.remark = remark;
    }

    public String getImageRid() {
        return imageRid;
    }

    public void setImageRid(String imageRid) {
        this.imageRid = imageRid;
    }

    public String getFilePath() {
        return filePath;
    }

    public void setFilePath(String filePath) {
        this.filePath = filePath;
    }
}
