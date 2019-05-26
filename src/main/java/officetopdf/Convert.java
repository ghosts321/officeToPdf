package main.java.officetopdf;

/**
 * Created by JiangJunpeng on 2019/1/3.<br>
 */
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import com.itextpdf.text.*;
import com.itextpdf.text.pdf.PdfWriter;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class Convert {
    private static final Integer WORD_TO_PDF_OPERAND = 17;
    private static final Integer PPT_TO_PDF_OPERAND = 32;
    private static final Integer EXCEL_TO_PDF_OPERAND = 0;

    private void doc2Pdf(String srcFilePath, String pdfFilePath) {
        ActiveXComponent app = null;
        Dispatch doc = null;
        try {
            ComThread.InitSTA();
            app = new ActiveXComponent("Word.Application");
            app.setProperty("Visible", false);
            Dispatch docs = app.getProperty("Documents").toDispatch();
            Object[] obj = new Object[]{
                    srcFilePath,
                    new Variant(false),
                    new Variant(false),//是否只读
                    new Variant(false),
                    new Variant("pwd")
            };
            doc = Dispatch.invoke(docs, "Open", Dispatch.Method, obj, new int[1]).toDispatch();
//          Dispatch.put(doc, "Compatibility", false);  //兼容性检查,为特定值false不正确
            Dispatch.put(doc, "RemovePersonalInformation", false);
            Dispatch.call(doc, "ExportAsFixedFormat", pdfFilePath, WORD_TO_PDF_OPERAND); // word保存为pdf格式宏，值为17

        }catch (Exception e) {
            e.printStackTrace();
            throw e;
        }
        finally {
            if (doc != null) {
                Dispatch.call(doc, "Close", false);
            }
            if (app != null) {
                app.invoke("Quit", 0);
            }
            ComThread.Release();
        }
    }

    private void ppt2Pdf(String srcFilePath, String pdfFilePath) {
        ActiveXComponent app = null;
        Dispatch ppt = null;
        try {
            ComThread.InitSTA();
            app = new ActiveXComponent("PowerPoint.Application");
            Dispatch ppts = app.getProperty("Presentations").toDispatch();

            /*
             * call
             * param 4: ReadOnly
             * param 5: Untitled指定文件是否有标题
             * param 6: WithWindow指定文件是否可见
             * */
            ppt = Dispatch.call(ppts, "Open", srcFilePath, true,true, false).toDispatch();
            Dispatch.call(ppt, "SaveAs", pdfFilePath, PPT_TO_PDF_OPERAND); // ppSaveAsPDF为特定值32

        } catch (Exception e) {
            e.printStackTrace();
            throw e;
        } finally {
            if (ppt != null) {
                Dispatch.call(ppt, "Close");
            }
            if (app != null) {
                app.invoke("Quit");
            }
            ComThread.Release();
        }
    }

    private void excel2Pdf(String inFilePath, String outFilePath) {
        ActiveXComponent ax = null;
        Dispatch excel = null;
        try {
            ComThread.InitSTA();
            ax = new ActiveXComponent("Excel.Application");
            ax.setProperty("Visible", new Variant(false));
            ax.setProperty("AutomationSecurity", new Variant(3)); // 禁用宏
            Dispatch excels = ax.getProperty("Workbooks").toDispatch();

            Object[] obj = new Object[]{
                    inFilePath,
                    new Variant(false),
                    new Variant(false)
            };
            excel = Dispatch.invoke(excels, "Open", Dispatch.Method, obj, new int[9]).toDispatch();

            // 转换格式
            Object[] obj2 = new Object[]{
                    new Variant(EXCEL_TO_PDF_OPERAND), // PDF格式=0
                    outFilePath,
                    new Variant(0)  //0=标准 (生成的PDF图片不会变模糊) ; 1=最小文件
            };
            Dispatch.invoke(excel, "ExportAsFixedFormat", Dispatch.Method,obj2, new int[1]);

        } catch (Exception es) {
            es.printStackTrace();
            throw es;
        } finally {
            if (excel != null) {
                Dispatch.call(excel, "Close", new Variant(false));
            }
            if (ax != null) {
                ax.invoke("Quit", new Variant[] {});
            }
            ComThread.Release();
        }

    }

    private void pic2Pdf(String imagePath, String pdfPath) {
        try {
            BufferedImage img = ImageIO.read(new File(imagePath));
            FileOutputStream fos = new FileOutputStream(pdfPath);
            Document doc = new Document(null, 0, 0, 0, 0);
            doc.setPageSize(new Rectangle(img.getWidth(), img.getHeight()));
            Image image = Image.getInstance(imagePath);
//            float scalePercentage = (72 / 300f) * 100.0f;
//            image.scalePercent(scalePercentage, scalePercentage);
            PdfWriter.getInstance(doc, fos);
            doc.open();
            doc.add(image);
            doc.close();
        } catch (IOException | DocumentException e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {

        SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        Date oldDate = new Date();
        new Convert().doc2Pdf("E:\\aaa.docx", "E:\\aaa.pdf");
        Date newDate = new Date();
        System.out.println("总耗时:" + new Convert().diffDate(oldDate, newDate));

//        if(args.length == 2){
//            inputFilePath = args[0];
//            outputFilePath = args[1];
//            String fix = inputFilePath.substring(inputFilePath.lastIndexOf(".") + 1);
//            Date oldDate = new Date();
//            System.out.println("开始生成文件"+ inputFilePath + "---- 当前时间:"+ df.format(oldDate));
//            switch (fix) {
//                case "docx":
//                case "doc":
//                    new Convert().doc2Pdf(inputFilePath, outputFilePath);
//                    break;
//                case "ppt":
//                case "pptx":
//                    new Convert().ppt2Pdf(inputFilePath, outputFilePath);
//                    break;
//                case "xls":
//                case "xlsx":
//                    new Convert().excel2Pdf(inputFilePath, outputFilePath);
//                    break;
//                default:
//                    new Convert().pic2Pdf(inputFilePath, outputFilePath);
//                    break;
//            }
//            Date newDate = new Date();
//            System.out.println("结束生成文件"+ outputFilePath + "---- 当前时间:"+ df.format(newDate));
//            System.out.println("总耗时:" + new Convert().diffDate(oldDate, newDate));
//        }else{
//            System.out.println("请输入文件输入地址及文件输出地址!!!");
//        }
    }

    private String diffDate(Date oldDate, Date newDate){
        long between = newDate.getTime() - oldDate.getTime();
        long day = between / (24 * 60 * 60 * 1000);
        long hour = (between / (60 * 60 * 1000) - day * 24);
        long min = ((between / (60 * 1000)) - day * 24 * 60 - hour * 60);
        long s = (between / 1000 - day * 24 * 60 * 60 - hour * 60 * 60 - min * 60);
        return min + "分" + s + "秒";
    }
}
