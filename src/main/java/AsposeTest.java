import com.aspose.cells.Workbook;
import com.aspose.words.*;
import javassist.ClassPool;
import javassist.CtClass;
import javassist.CtMethod;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;

/**
 * Aspose 测试
 *
 * @author
 * @since
 */
public class AsposeTest {

    public static void main(String[] args) throws Exception {
        registerWord2512();
        String path = "F:\\workplace\\aspose\\1.docx";
        Document doc = new Document(path);  // 替换为你的输入文件路径

        // 可以设置各种 PDF 选项，例如：
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
        pdfSaveOptions.setCompliance(PdfCompliance.PDF_17);
        pdfSaveOptions.setUseHighQualityRendering(true); // 高质量渲染
        pdfSaveOptions.setEmbedFullFonts(false);
        // pdfSaveOptions.setUseCoreFonts(true); // 使用基本字体减少体积
        pdfSaveOptions.setAdditionalTextPositioning(true); // 精确文本定位
        pdfSaveOptions.setUseBookFoldPrintingSettings(true);
        pdfSaveOptions.setExportDocumentStructure(true); // 导出文档结构
        pdfSaveOptions.setPreserveFormFields(true); // 保留表单字段

        // 设置PDF输出的大纲（书签）选项
        OutlineOptions outlineOptions = pdfSaveOptions.getOutlineOptions();
        // 1. 不自动创建缺失的大纲级别
        outlineOptions.setCreateMissingOutlineLevels(false);
        // 2. 设置默认书签的大纲级别为1（顶级）
        outlineOptions.setDefaultBookmarksOutlineLevel(1);
        // 3. 设置默认不展开任何大纲级别（0表示全部折叠）
        outlineOptions.setExpandedOutlineLevels(0);
        // 4. 设置标题大纲级别为0（不将标题显示为大纲）
        outlineOptions.setHeadingsOutlineLevels(0);
        // 将文档保存为 PDF
        doc.save("1.pdf", pdfSaveOptions);


        InputStream is = new FileInputStream("license.xml");
        com.aspose.cells.License license1 = new com.aspose.cells.License();
        license1.setLicense(is);
        String path1 = "F:\\workplace\\aspose\\1.xlsx";
        Workbook doc1 = new Workbook(path);  // 替换为你的输入文件路径
        // 创建 PDF 保存选项
        com.aspose.cells.PdfSaveOptions options = new com.aspose.cells.PdfSaveOptions();
        // 可以设置各种 PDF 选项，例如：
        options.setCompliance(PdfCompliance.PDF_17);// 设置 PDF 版本兼容性
        options.setOutputBlankPageWhenNothingToPrint(false);// 不打印空白页
        // 将文档保存为 PDF
        doc1.save("1.pdf", options);  // 替换为你想要的输出文件路径

        System.out.println("文档转换成功！");
    }


    /**
     * 修改cells.jar包里面的校验
     */
    public static void modifyExcelJar() {
        try {
            //获取指定的class文件对象  License的 new 对象类
            CtClass x6bClass = ClassPool.getDefault().getCtClass("com.aspose.cells.x6b");
            //从class对象中解析获取所有方法
            CtMethod[] methodA = x6bClass.getDeclaredMethods();
            for (CtMethod ctMethod : methodA) {
                //获取方法获取参数类型
                CtClass[] ps = ctMethod.getParameterTypes();
                //筛选同名方法，入参是Document
                if (ps.length == 1 && ctMethod.getName().equals("a") && ps[0].getName().equals("org.w3c.dom.Document")) {
                    System.out.println("ps[0].getName==" + ps[0].getName());
                    //替换指定方法的方法体
                    ctMethod.setBody("{ a = $0; com.aspose.cells.y6ab.g(-1); com.aspose.cells.u80.a(); return;}");
                }
            }
            //这一步就是将破译完的代码放在桌面上
            x6bClass.writeFile("C:\\Users\\uolds\\Desktop\\");

        } catch (Exception e) {
            System.out.println("错误==" + e);
        }
    }

    /**
     * aspose-words:jdk17:25.9 版本
     */
    public static void registerWord2512() {
        try {
            // License中的 new对象的第一个关键方法的返回对象
            Class<?> zzWmmClass = Class.forName("com.aspose.words.zzWmm");
            Class<?> zzMfClass = Class.forName("com.aspose.words.zzMf");
            Method method = zzMfClass.getDeclaredMethod(
                    "zznj",
                    ArrayList.class,
                    String.class,
                    String.class
            );
            method.setAccessible(true);
            ArrayList<String> var0 = new ArrayList<>(Collections.singletonList("success"));
            Object outInstance = method.invoke(null, var0, null, null);
            Field zzXFX = zzWmmClass.getDeclaredField("zzXFX");
            zzXFX.setAccessible(true);
            zzXFX.set(outInstance, 1);
            Field zzXow = zzWmmClass.getDeclaredField("zzXow");
            zzXow.setAccessible(true);
            zzXow.set(outInstance, 1);

            // License中的 new zz1V
            Class<?> zz1VClass = Class.forName("com.aspose.words.zz1V");
            Field zzVRB = zz1VClass.getDeclaredField("zzVRB");
            zzVRB.setAccessible(true);
            ArrayList<Object> innerInstance = new ArrayList<>();
            innerInstance.add(outInstance);
            zzVRB.set(null, innerInstance);

            // License中的 new对象方法的关键方法中的判断条件大于0
            Class<?> zz3NClass = Class.forName("com.aspose.words.zz3N");
            Field zzXl7 = zz3NClass.getDeclaredField("zzXl7");
            zzXl7.setAccessible(true);
            zzXl7.set(null, 128);
            Field zzxQ = zz3NClass.getDeclaredField("zzxQ");
            zzxQ.setAccessible(true);
            zzxQ.set(null, false);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
