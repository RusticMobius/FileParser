import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.model.StyleDescription;
import org.apache.poi.hwpf.model.StyleSheet;
import org.apache.poi.hwpf.usermodel.*;
import org.apache.poi.hwpf.model.PicturesTable;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableCell;
import org.apache.poi.hwpf.usermodel.TableIterator;
import org.apache.poi.hwpf.usermodel.TableRow;
import org.apache.poi.ss.formula.functions.T;
import org.openxmlformats.schemas.drawingml.x2006.chart.STDepthPercentWithSymbol;

import java.io.*;
import java.util.List;
import java.util.Random;

public class DocParser  {
    InputStream is;
    HWPFDocument document;
    PicturesTable picturesTable;
    TableIterator tableIterator;
    Range range;


    DocParser(String filepath)throws IOException{
        is = new FileInputStream(filepath);
        document = new HWPFDocument(is);
        picturesTable = document.getPicturesTable();
        range = document.getRange();
        tableIterator = new TableIterator(range);
        parserInit();

    }

    public void parserInit(){
        int paraIndex=0;
        for ( int i = 0; i < range.numParagraphs() ; i++){
            paraIndex = i+1;
            Paragraph paragraph = range.getParagraph(i);
            System.out.println("段落:" + (i+1));
            //特殊标记，法院信息，标题，编号
            if(i==0){
                System.out.println("备注：法院信息");
            }else if(i==1){
                System.out.println("备注：文书类型");
            }else if(i==2){
                System.out.println("备注：文书编号");
            }
            paragraphParser(paragraph,paraIndex,true);
        }
        tableFormatParser(tableIterator);
    }

    public void paragraphParser(Paragraph paragraph,int paraIndex,boolean mode) {
        // mode是指是否解析字单元，true解析
        System.out.println("段落正文："+paragraph.text());
        //段落格式
        //段落等级
        StyleDescription style = document.getStyleSheet().getStyleDescription(paragraph.getStyleIndex());
        System.out.println("段落样式："+style.getName());
        System.out.println("大纲级别："+paragraph.getLvl());
        //todo 自动标号信息
        //getIlfo() Returns the ilfo, an index to the document's hpllfo, which describes the automatic number formatting of the paragraph. A value of zero means it isn't numbered.
        System.out.println("自动标号格式："+paragraph.getIlfo());
        //getIlvl() Returns the multi-level indent for the paragraph. Will be zero for non-list paragraphs, and the first level of any list. Subsequent levels in hold values 1-8.
        System.out.println("标号缩进："+paragraph.getIlvl());

        //对齐方式
        String justification = "";
        switch (paragraph.getJustification()){
            case 0 :
                justification="左对齐";
                break;
            case 1:
                justification="居中对齐";
                break;
            case 2:
                justification="右对齐";
                break;
            case 3:
                justification="两端对齐";
                break;
            default:
                justification="解析错误";
        }
        System.out.println("对齐方式："+justification);

        //缩进格式
        String specialIndentType = "";
        if(paragraph.getFirstLineIndent()==0){
            specialIndentType = "无";
        }else if(paragraph.getFirstLineIndent()>0){
            specialIndentType="首行缩进 "+paragraph.getFirstLineIndent();
        }else {
            specialIndentType="悬挂缩进 "+(-paragraph.getFirstLineIndent());
        }
        System.out.println("左缩进:"+paragraph.getIndentFromLeft());
        System.out.println("右缩进："+paragraph.getIndentFromRight());
        System.out.println("特殊格式："+specialIndentType);

        //段前段后段间距
        System.out.println("段前："+paragraph.getSpacingBefore());
        System.out.println("段后："+paragraph.getSpacingAfter());
        System.out.println("段间距："+paragraph.getLineSpacing().toInt());

        //字单元解析
        if(mode) {
            System.out.println("-----------字单元信息-------------");
            characterRunParser(paragraph, paraIndex);
        }
        //表格解析
        System.out.println("============段落分割线===========");

    }

    public void characterRunParser (Paragraph paragraph,int paraIndex) {
        //
        CharacterRun characterRun;
        int charaIndex;
        for (int j = 0; j < paragraph.numCharacterRuns(); j++) {
            characterRun = paragraph.getCharacterRun(j);
            charaIndex = j+1;
            //判断字单元是否为图片类型
            System.out.println("字单元"+charaIndex);
            if (picturesTable.hasPicture(characterRun)) {
                System.out.println("字单元类型：图片");
                pictureFormatParser(characterRun,charaIndex,paraIndex);
            }else{
                System.out.println("字单元类型：文字");
                System.out.println("文字内容："+characterRun.text());
                //字体信息
                System.out.println("字体名称："+characterRun.getFontName());
                System.out.println("字体大小："+characterRun.getFontSize());
                System.out.println("字体颜色："+characterRun.getColor());
                System.out.println("特殊格式："+
                                (characterRun.isBold() ? "加粗 ":"")+
                        (characterRun.isItalic() ? "斜体 ":"")+
                        (characterRun.getUnderlineCode()>0 ? "下划线 ":"" )+
                        (characterRun.isOutlined() ? "突出显示 ":" "));
                System.out.println("-----------------------------------");
            }
        }

    }

    public void pictureFormatParser (CharacterRun characterRun,int charaIndex,int paraIndex) {

        Picture picture = picturesTable.extractPicture(characterRun, true);
        //图片格式
        System.out.println("图片格式："+picture.suggestPictureType());
        //所属段落
        System.out.println("所属段落字："+"段落"+paraIndex+"，"+"字单元"+charaIndex);
        //图片大小
        System.out.println("图片大小："+picture.getSize());
        //图片高度
        System.out.println("图片高度："+picture.getDescription());
        //图片宽度
        System.out.println("图片宽度："+picture.getHeight());
        //todo
        //base64编码
        System.out.println("---------------------------------");

    }

    public void tableFormatParser (TableIterator tableIterator) {
        int tableCount = 0;
        while (tableIterator.hasNext()){
            tableCount++;
            System.out.println("=============表格分割线===========");
            System.out.println("表格"+tableCount+":");
            System.out.println("--------------------------------");
            Table table = (Table) tableIterator.next();
            for (int i = 0; i < table.numRows(); i++){
                TableRow tableRow = table.getRow(i);
                for ( int j = 0; j < tableRow.numCells(); j++){
                    TableCell tableCell = tableRow.getCell(j);
                    for ( int k = 0; k < tableCell.numParagraphs(); k++){
                        Paragraph paragraph = tableCell.getParagraph(k);
                        System.out.printf("第%d行，第%d个单元格，第%d段落\n",i+1,j+1,k+1);
                        //表格段落解析
                        paragraphParser(paragraph,k+1,false);
                    }
                }
            }
        }
    }

    public void headingParser (){}
}
