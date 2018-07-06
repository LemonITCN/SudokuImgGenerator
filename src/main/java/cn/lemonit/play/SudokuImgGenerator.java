package cn.lemonit.play;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.*;
import java.util.List;

public class SudokuImgGenerator {

    private static int width = 900;
    private static int height = 900;
    private static int font = 80;
    private static int outLineWidth = 5;
    private static int inLineWidth = 1;
    private static int numXDeviation = 20;
    private static int numYDeviation = 10;
    private static String source = "";
    private static String output = "";

    private static final String PNG = "png";
    private static final String DOC = "sudoku.doc";

//    private static String fileName = "1.png";

    public static void main(String[] args) {
        Map<String, String> config = new HashMap<String, String>();
        for (String str : args) {
            String[] param = str.split("=");
            if (param.length == 2) {
                config.put(param[0], param[1]);
            } else {
                System.err.println("无效参数：" + str);
            }
        }
        width = Integer.valueOf(config.getOrDefault("width", "900"));
        height = Integer.valueOf(config.getOrDefault("height", "900"));
        font = Integer.valueOf(config.getOrDefault("font", "80"));
        outLineWidth = Integer.valueOf(config.getOrDefault("outLineWidth", "5"));
        inLineWidth = Integer.valueOf(config.getOrDefault("inLineWidth", "1"));
        numXDeviation = Integer.valueOf(config.getOrDefault("numXDeviation", String.valueOf(font / 4)));
        numYDeviation = Integer.valueOf(config.getOrDefault("numYDeviation", String.valueOf(font / 8)));

        if (!config.containsKey("source")) {
            new Exception("请您必须要传有效的source参数，如e:/1.xlsx").printStackTrace();
        } else if (!config.containsKey("output")) {
            new Exception("请您必须要传有效的output参数，如e:/output/").printStackTrace();
        } else {
            // 参数完整!
            source = config.get("source");
            output = config.get("output");
            try {
                List<SudokuInfo> sudokuInfoList = readFromExcel();
                System.out.println("从Excel中读取到" + sudokuInfoList.size() + "条数据");
                for (SudokuInfo sudokuInfo : sudokuInfoList) {
                    generateImage(sudokuInfo);
                }
                createWord(sudokuInfoList);
                System.out.println("任务执行完毕，正常退出");
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 从EXCEL中读取数据
     *
     * @return 读取到的数据列表
     * @throws Exception 任何exception
     */
    private static List<SudokuInfo> readFromExcel() throws Exception {
        InputStream inputStream = new FileInputStream(source);
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        List<SudokuInfo> sudokuInfoList = new ArrayList<>();
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            Iterator<Row> rows = sheet.rowIterator();
            // 遍历所有行
            while (rows.hasNext()) {
                Row row = rows.next();
                row.getCell(0).setCellType(CellType.STRING);
                row.getCell(1).setCellType(CellType.STRING);
                String num = row.getCell(0).getStringCellValue();
                String content = row.getCell(1).getStringCellValue();
                SudokuInfo sudokuInfo = new SudokuInfo(num, content);
                sudokuInfoList.add(sudokuInfo);
            }
        }
        return sudokuInfoList;
    }

    private static void generateImage(SudokuInfo sudokuInfo) {
        BufferedImage image = new BufferedImage(width + outLineWidth, height + outLineWidth, BufferedImage.TYPE_INT_RGB);
        Graphics2D graphics = image.createGraphics();
        graphics.setColor(Color.white);
        graphics.fillRect(0, 0, width, height);
        graphics.setColor(Color.BLACK);
        if (sudokuInfo.getContent().length() == 81) {
            // 9宫
            int widthStep = (width) / 9;
            int heightStep = (height) / 9;
            for (int i = 0; i <= 9; i++) {
                graphics.fillRect(0, i * heightStep, width, i % 3 == 0 ? outLineWidth : inLineWidth);
                graphics.fillRect(widthStep * i, 0, i % 3 == 0 ? outLineWidth : inLineWidth, height);
            }
            for (int i = 0; i < sudokuInfo.getContent().length(); i++) {
                graphics.setFont(new Font("Arial", Font.PLAIN, font));
                String item = sudokuInfo.getContent().substring(i, i + 1);
                if (!item.equals("0")) {
                    graphics.drawString(item, i % 9 * widthStep + (widthStep - font) / 2 + numXDeviation, (i / 9 + 1) * heightStep - (heightStep - font) / 2 - numYDeviation);
                }
            }
        } else if (sudokuInfo.getContent().length() == 36) {
            // 9宫
            int widthStep = (width) / 6;
            int heightStep = (height) / 6;
            for (int i = 0; i <= 6; i++) {
                graphics.fillRect(0, i * heightStep, width, i % 2 == 0 ? outLineWidth : inLineWidth);
                graphics.fillRect(widthStep * i, 0, i % 3 == 0 ? outLineWidth : inLineWidth, height);
            }
            for (int i = 0; i < sudokuInfo.getContent().length(); i++) {
                graphics.setFont(new Font("Arial", Font.PLAIN, font));
                String item = sudokuInfo.getContent().substring(i, i + 1);
                if (!item.equals("0")) {
                    graphics.drawString(item, i % 6 * widthStep + (widthStep - font) / 2 + numXDeviation, (i / 6 + 1) * heightStep - (heightStep - font) / 2 - numYDeviation);
                }
            }
        }
        try {
            String imgPath = output + sudokuInfo.getNumber() + "." + PNG;
            FileOutputStream fos = new FileOutputStream(imgPath);
            ImageIO.write(image, PNG, fos);
            System.out.println("输出图片：" + imgPath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void createWord(List<SudokuInfo> sudokuInfoList) {
        CustomXWPFDocument document = new CustomXWPFDocument();
        try {
            for (int i = 0; i < sudokuInfoList.size(); i++) {
                XWPFParagraph paragraph = document.createParagraph();
                XWPFRun queNum = paragraph.createRun();
                queNum.setText(sudokuInfoList.get(i).getNumber());
                String picId = document.addPictureData(
                        new FileInputStream(output + sudokuInfoList.get(i).getNumber() + "." + PNG)
                        , XWPFDocument.PICTURE_TYPE_PNG
                );
                document.createPicture(picId, document.getAllPictures().size() - 1, 300, 300);
            }
            String wordPath = output + DOC;
            document.write(new FileOutputStream(wordPath));
            System.out.println("word文件已生成：" + wordPath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
