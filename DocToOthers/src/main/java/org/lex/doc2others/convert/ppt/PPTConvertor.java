package org.lex.doc2others.convert.ppt;

import java.io.File;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

/**
 * PPT -> 其他文件格式
 * 
 * @author Lex.Shaoli.Zhu
 *
 */
public class PPTConvertor {

    private static Logger log = LoggerFactory.getLogger(PPTConvertor.class);

    public static class SaveToType {
        //不支持
        //,12,13,14,20
        //其他原因
        //11,
        //版本原因
        //10,22
        //质量很差
        public static final int PPT_SAVEAS_WMF = 15;
        public static final int PPT_SAVEAS_GIF = 16;
        public static final int PPT_SAVEAS_JPG = 17;
        public static final int PPT_SAVEAS_PNG = 18;
        public static final int PPT_SAVEAS_BMP = 19;
        public static final int PPT_SAVEAS_TIF = 21;
        public static final int PPT_SAVEAS_ = 32;
        //2016不支持
        //public static final int PPT_SAVEAS_HTML = 20;
        /**
         * 存储路径和源文件同级, 名字同名, 只是后缀不同
         */
        public static final int PPT_SAVEAS_PDF = 32;
    }

    private ActiveXComponent ppt;
    private ActiveXComponent presentation;
    // 转换后的存储路径
    private String saveTo = null;

    /**
     * PPTConvertor构造方法
     * 
     * @param pptFilePath
     *            需要被转成图片的PPT文件路径(相对和绝对都行)
     * @param isVisble
     *            PowerPoint是否会弹出界面(ture时正常运行, false 程序报错)
     * @throws Exception
     */
    public PPTConvertor(String pptFilePath, boolean isVisble) throws Exception {
        this(new File(pptFilePath), isVisble);
    }

    /**
     * PPTConvertor构造方法
     * 
     * @param pptFile
     *            需要被转成图片的PPT文件
     * @param isVisble
     *            PowerPoint是否会弹出界面(ture时正常运行, false 程序报错)
     * @throws Exception
     */
    public PPTConvertor(File pptFile, boolean isVisble) throws Exception {
        if (!pptFile.exists()) {
            throw new Exception("文件不存在!");
        }

        // 根据ppt文件获取图片存储路径
        this.saveTo = getSavetoPath(pptFile);
        try {
            if (ppt == null || ppt.m_pDispatch == 0) {
                ppt = new ActiveXComponent("PowerPoint.Application");
                setIsVisble(ppt, isVisble);
            }
            // 打开一个现有的 Presentation 对象
            ActiveXComponent presentations = ppt.getPropertyAsComponent("Presentations");
            presentation = presentations.invokeGetComponent("Open", new Variant(pptFile.getAbsolutePath()),
                    new Variant(true));
        } catch (Exception e) {
            e.printStackTrace();
            pptClose();
        }
    }

    /**
     * 生成转换后的存储路径
     * 
     * @param pptFile
     *            需要转换的PPT文件
     * @return 最终存储路径 <br>
     *         1, 需要转换的PPT和存储路径的父目录平级; <br>
     *         2, 存储路径为PPT文件名字不带后缀;<br>
     *         如: <br>
     *         PPT : c:/simple/somthing.ppt<br>
     *         存储路径 : c:/simple/somthing/
     */
    private String getSavetoPath(File pptFile) {
        StringBuffer stringBuffer = new StringBuffer();
        String path = pptFile.getAbsolutePath().replaceAll("\\\\", "/");
        String fileName = path.substring(path.lastIndexOf("/"), path.lastIndexOf("."));
        stringBuffer.append(path.substring(0, path.lastIndexOf("/"))).append(fileName);
        path = null;
        fileName = null;
        return stringBuffer.toString();
    }

    /**
     * PPT -> JPG
     * 
     * @throws Exception
     */
    public void PPTToJPG(int saveToType) throws Exception {
        try {
            saveAs(presentation, this.saveTo, saveToType);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            pptClose();
        }
    }

    /**
     * 
     * @param presentation
     *            Dispatch实例, 实际上就是PowerPoint的句柄, 此时PPT文件都是处于打开状态
     * @param saveTo
     *            转换后的图片放到哪里
     * @param pptSaveAsFileType
     *            要转换成那种类型的图片
     * @throws Exception
     */
    private void saveAs(Dispatch presentation, String saveTo, int pptSaveAsFileType) throws Exception {
        File targetPath = new File(saveTo);
        if (!targetPath.exists()) {
            targetPath.mkdirs();
        }
        log.info("----目标路径:{}", targetPath.getAbsolutePath());
        Dispatch.call(presentation, "SaveAs", targetPath.getAbsolutePath(), new Variant(pptSaveAsFileType));
    }

    /**
     * 设置PowerPoint是否弹窗
     * 
     * @param presentation
     *            Dispatch实例, 实际上就是PowerPoint的句柄
     * @param visble
     *            true/false
     * @throws Exception
     */
    private void setIsVisble(Dispatch presentation, boolean visble) throws Exception {
        Dispatch.put(presentation, "Visible", new Variant(visble));
        // Dispatch.put(obj, "DisplayAlerts", new Variant(visble));
    }

    /**
     * 关闭PowerPoint
     * 
     * @throws Exception
     */
    private void pptClose() throws Exception {
        if (null != presentation) {
            Dispatch.call(presentation, "Close");
        }
        ppt.invoke("Quit", new Variant[] {});
        ComThread.Release();
    }

}
