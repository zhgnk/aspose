import com.aspose.cells.Workbook;
import com.aspose.words.License;
import com.aspose.words.PdfCompliance;
import com.aspose.words.PdfSaveOptions;
import javassist.ClassPool;
import javassist.CtClass;
import javassist.CtMethod;
import org.w3c.dom.Document;

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
        // registerWord259();
        // String path = "F:\\workplace\\aspose\\1.docx";
        // Document doc = new Document(path);  // 替换为你的输入文件路径
        // modifyExcelJar();
        // registerExcel259();
        InputStream is = new FileInputStream("license.xml");
        com.aspose.cells.License license1 = new com.aspose.cells.License();
        license1.setLicense(is);
        String path = "F:\\workplace\\aspose\\1.xlsx";
        Workbook doc = new Workbook(path);  // 替换为你的输入文件路径

        // // 创建 PDF 保存选项
        // PdfSaveOptions options = new PdfSaveOptions();
        // // 可以设置各种 PDF 选项，例如：
        // options.setCompliance(PdfCompliance.PDF_20); // 设置 PDF 版本兼容性
        // // 将文档保存为 PDF
        // doc.save("1.pdf", options);  // 替换为你想要的输出文件路径
        // System.out.println("文档转换成功！");
        // 创建 PDF 保存选项
        com.aspose.cells.PdfSaveOptions options = new com.aspose.cells.PdfSaveOptions();
        // 可以设置各种 PDF 选项，例如：
        options.setCompliance(PdfCompliance.PDF_20);// 设置 PDF 版本兼容性
        // 将文档保存为 PDF
        doc.save("1.pdf", options);  // 替换为你想要的输出文件路径

    }

    private static void registerExcel259() {
        try {
            Class<?> u65Class = Class.forName("com.aspose.cells.u65");
            Constructor<?> constructors = u65Class.getDeclaredConstructors()[0];
            constructors.setAccessible(true);
            Object instance = constructors.newInstance();
            Field b1 = u65Class.getDeclaredField("b");
            b1.setAccessible(true);
            b1.set(instance, "错误");
            Field c = u65Class.getDeclaredField("c");
            c.setAccessible(true);
            Date date = new Date(Long.MAX_VALUE);
            SimpleDateFormat formatter = new SimpleDateFormat("yyyyMMdd");
            String formattedDate = formatter.format(date);
            c.set(instance, formattedDate);
            // Class<?> f20Class = Class.forName("com.aspose.cells.f20");
            // Constructor<?> f20Classconstructors = f20Class.getDeclaredConstructors()[0];
            // f20Classconstructors.setAccessible(true);
            // Object f20instance = f20Classconstructors.newInstance(new Workbook("F:\\workplace\\aspose\\1.xlsx"));
            // Field b = f20Class.getDeclaredField("b");
            // b.setAccessible(true);
            // b.set(f20instance, 1);
            // Field c1 = f20Class.getDeclaredField("c");
            // c1.setAccessible(true);
            // c1.set(f20instance, 1);
            // Field k = f20Class.getDeclaredField("k");
            // k.setAccessible(true);
            // k.set(f20instance, 1);
            // Field l = f20Class.getDeclaredField("l");
            // l.setAccessible(true);
            // l.set(f20instance, 1);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 修改cells.jar包里面的校验
     */
    public static void modifyExcelJar() {
        try {
            //获取指定的class文件对象
            CtClass u65Class = ClassPool.getDefault().getCtClass("com.aspose.cells.u65");
            //从class对象中解析获取所有方法
            CtMethod[] methodA = u65Class.getDeclaredMethods();
            for (CtMethod ctMethod : methodA) {
                //获取方法获取参数类型
                CtClass[] ps = ctMethod.getParameterTypes();
                //筛选同名方法，入参是Document
                if (ps.length == 1 && ctMethod.getName().equals("a") && ps[0].getName().equals("org.w3c.dom.Document")) {
                    System.out.println("ps[0].getName==" + ps[0].getName());
                    //替换指定方法的方法体
                    ctMethod.setBody("{ a = $0; com.aspose.cells.f20.g(-1); com.aspose.cells.k1t.a(); return;}");
                }
            }
            //这一步就是将破译完的代码放在桌面上
            u65Class.writeFile("C:\\Users\\uolds\\Desktop\\");

        } catch (Exception e) {
            System.out.println("错误==" + e);
        }
    }

    /**
     * aspose-words:jdk17:25.9 版本
     */
    public static void registerWord259() {
        try {
            Class<?> zzWkUClass = Class.forName("com.aspose.words.zzWkU");
            Class<?> zz6LClass = Class.forName("com.aspose.words.zz6L");
            Method method = zz6LClass.getDeclaredMethod(
                    "zzW07",
                    ArrayList.class,
                    String.class,
                    String.class
            );
            method.setAccessible(true);
            ArrayList<String> var0 = new ArrayList<>(Collections.singletonList("success"));
            Object outInstance = method.invoke(null, var0, null, null);
            Field zzX4E = zzWkUClass.getDeclaredField("zzX4E");
            zzX4E.setAccessible(true);
            zzX4E.set(outInstance, 1);
            Field zzVY7 = zzWkUClass.getDeclaredField("zzVY7");
            zzVY7.setAccessible(true);
            zzVY7.set(outInstance, 1);

            Class<?> zzX4bClass = Class.forName("com.aspose.words.zzX4b");
            Field zzsI = zzX4bClass.getDeclaredField("zzsI");
            zzsI.setAccessible(true);
            ArrayList<Object> innerInstance = new ArrayList<>();
            innerInstance.add(outInstance);
            zzsI.set(null, innerInstance);

            Class<?> zzZgNClass = Class.forName("com.aspose.words.zzZgN");
            Field zzWtT = zzZgNClass.getDeclaredField("zzWtT");
            zzWtT.setAccessible(true);
            zzWtT.set(null, 128);
            Field zzYbl = zzZgNClass.getDeclaredField("zzYbl");
            zzYbl.setAccessible(true);
            zzYbl.set(null, false);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
