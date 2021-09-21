import java.io.File;
import java.io.IOException;


public class FormatJudger {
    public static void main(String[] args) throws IOException {

        String path = "src/main/testfile";
        File []  fileList = new File(path).listFiles();
        for (File file : fileList){
            if(file.isFile()){
                String fileName = file.getName();
                if (fileName.endsWith(".doc")) {
                    System.out.println(fileName);
                    DocParser docParser = new DocParser(path+"/"+fileName);
                    //process doc file
                }else if (fileName.endsWith(".docx")) {
                    System.out.println(fileName);
                    //process docx file
                }else if (fileName.endsWith(".pdf")) {
                    System.out.println(fileName);
                    //process pdf file
                }else if (fileName.endsWith(".wps")) {
                    System.out.println(fileName);
                    //process wps file
                }else {
                    System.out.println("无法解析文件："+fileName+",不支持该种格式文件解析");
                }
            }
        }

    }
}
