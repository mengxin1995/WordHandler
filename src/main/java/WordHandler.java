import java.io.*;
import java.util.HashMap;
import java.util.Map;

import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;
import sun.misc.BASE64Encoder;

public class WordHandler {

    private Configuration configuration = null;

    /****
     * 模板文件存放的目录
     */
    private static final String  baseDir = "/Users/mengxin/code/wordDemo";
    /***
     * 模板文件名称
     */
    private static final String  templateFile = "template.ftl";
    /***
     * word生成的输出目录
     */
    private static final String  outputDir = "/Users/mengxin/code/wordDemo/";

    public WordHandler(){
        configuration = new Configuration();
        configuration.setDefaultEncoding("UTF-8");
    }

    public static void main(String[] args) {
        WordHandler test = new WordHandler();
        test.createWord();
    }

    /*****
     *
     * Project Name: maventest
     * <p>转换成word<br>
     *
     * @author youqiang.xiong
     * @date 2019年2月21日  上午11:22:03
     * @version v1.0
     * @since
     */
    public void createWord(){
        Map<String,Object> dataMap=new HashMap<String,Object>();
        //构造参数
        getData(dataMap);

        configuration.setClassForTemplateLoading(this.getClass(), "");//模板文件所在路径
        Template t=null;
        try {
            configuration.setDirectoryForTemplateLoading(new File(baseDir));
            t = configuration.getTemplate(templateFile);
        } catch (IOException e) {
            e.printStackTrace();
        }
        File outFile = new File(outputDir+Math.random()*10000+".doc"); //导出文件
        Writer out = null;
        try {
            out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(outFile)));
        } catch (FileNotFoundException e1) {
            e1.printStackTrace();
        }

        try {
            t.process(dataMap, out); //将填充数据填入模板文件并输出到目标文件
            System.out.println("生成成功...");
        } catch (TemplateException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /****
     *
     * Project Name: maventest
     * <p>初始化数据map <br>
     *
     * @author youqiang.xiong
     * @date 2019年2月21日  上午11:26:34
     * @version v1.0
     * @since
     * @param dataMap
     * 		封装数据的map
     */
    private void getData(Map<String, Object> dataMap) {
        dataMap.put("userName", "刘德华");
        dataMap.put("signature", getImageStr());
    }

    private String getImageStr() {
         String imgFile = "/Users/mengxin/code/wordDemo/logo.png";
         InputStream in = null;
         byte[] data = null;
         try {
             in = new FileInputStream(imgFile);
             data = new byte[in.available()];
             in.read(data);
             in.close();
         } catch (IOException e) {
             e.printStackTrace();
         }
         BASE64Encoder encoder = new BASE64Encoder();
         return encoder.encode(data);
    }

}
