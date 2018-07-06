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
    private static int fontStyle = 0;
    private static String fontName = "";
    private static int outLineWidth = 5;
    private static int inLineWidth = 1;
    private static int numXDeviation = 20;
    private static int numYDeviation = 10;
    private static int wordMode = 2;
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
        fontStyle = Integer.valueOf(config.getOrDefault("fontStyle", String.valueOf(Font.PLAIN)));
        fontName = config.getOrDefault("fontName", "Arial");
        outLineWidth = Integer.valueOf(config.getOrDefault("outLineWidth", "5"));
        inLineWidth = Integer.valueOf(config.getOrDefault("inLineWidth", "1"));
        numXDeviation = Integer.valueOf(config.getOrDefault("numXDeviation", String.valueOf(font / 4)));
        numYDeviation = Integer.valueOf(config.getOrDefault("numYDeviation", String.valueOf(font / 8)));
        wordMode = Integer.valueOf(config.getOrDefault("wordMode", String.valueOf(2)));

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
        Font useFont = new Font(fontName, fontStyle, font);
        graphics.setFont(useFont);
        if (sudokuInfo.getContent().length() == 81) {
            // 9宫
            int widthStep = (width) / 9;
            int heightStep = (height) / 9;
            for (int i = 0; i <= 9; i++) {
                graphics.fillRect(0, i * heightStep, width, i % 3 == 0 ? outLineWidth : inLineWidth);
                graphics.fillRect(widthStep * i, 0, i % 3 == 0 ? outLineWidth : inLineWidth, height);
            }
            for (int i = 0; i < sudokuInfo.getContent().length(); i++) {
                String item = sudokuInfo.getContent().substring(i, i + 1);
                if (!item.equals("0")) {
                    graphics.drawString(item, i % 9 * widthStep + (widthStep - font) / 2 + numXDeviation, (i / 9 + 1) * heightStep - (heightStep - font) / 2 - numYDeviation);
                }
            }
        } else if (sudokuInfo.getContent().length() == 36) {
            // 6宫
            int widthStep = (width) / 6;
            int heightStep = (height) / 6;
            for (int i = 0; i <= 6; i++) {
                graphics.fillRect(0, i * heightStep, width, i % 2 == 0 ? outLineWidth : inLineWidth);
                graphics.fillRect(widthStep * i, 0, i % 3 == 0 ? outLineWidth : inLineWidth, height);
            }
            for (int i = 0; i < sudokuInfo.getContent().length(); i++) {
                String item = sudokuInfo.getContent().substring(i, i + 1);
                if (!item.equals("0")) {
                    graphics.drawString(item, i % 6 * widthStep + (widthStep - font) / 2 + numXDeviation, (i / 6 + 1) * heightStep - (heightStep - font) / 2 - numYDeviation);
                }
            }
        } else if (sudokuInfo.getContent().length() == 16) {
            // 4宫
            int widthStep = (width) / 4;
            int heightStep = (height) / 4;
            for (int i = 0; i <= 4; i++) {
                graphics.fillRect(0, i * heightStep, width, i % 2 == 0 ? outLineWidth : inLineWidth);
                graphics.fillRect(widthStep * i, 0, i % 2 == 0 ? outLineWidth : inLineWidth, height);
            }
            for (int i = 0; i < sudokuInfo.getContent().length(); i++) {
                String item = sudokuInfo.getContent().substring(i, i + 1);
                if (!item.equals("0")) {
                    graphics.drawString(item, i % 4 * widthStep + (widthStep - font) / 2 + numXDeviation, (i / 4 + 1) * heightStep - (heightStep - font) / 2 - numYDeviation);
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
        String numSpace = wordMode == 2
                ? ".                                                                        "
                : ".                                              ";
        String imgSpace = wordMode == 2 ? "       " : "           ";
        int imgSize = wordMode == 2 ? 260 : 150;
        CustomXWPFDocument document = new CustomXWPFDocument();
        try {
            XWPFParagraph numParagraph = null;
            XWPFParagraph imgParagraph = null;
            String numText = "";

            for (int i = 0; i < sudokuInfoList.size(); i++) {
                if (i % wordMode == 0) {
                    if (numParagraph != null) {
                        XWPFRun run = numParagraph.createRun();
                        run.setBold(true);
                        run.setText(numText);
                    }
                    numParagraph = document.createParagraph();
                    imgParagraph = document.createParagraph();
                    if (wordMode == 2 && i % 6 < 4) {
                        document.createParagraph();
                    }
                    numText = "";
                }
                numText += sudokuInfoList.get(i).getNumber() + numSpace;
                String picId = document.addPictureData(
                        new FileInputStream(output + sudokuInfoList.get(i).getNumber() + "." + PNG)
                        , XWPFDocument.PICTURE_TYPE_PNG
                );
                document.createPicture(imgParagraph, picId, document.getAllPictures().size() - 1, imgSize, imgSize);
                if (i % wordMode < (wordMode - 1)) {
                    imgParagraph.createRun().setText(imgSpace);
                }
            }

            if (numParagraph != null) {
                XWPFRun run = numParagraph.createRun();
                run.setBold(true);
                run.setText(numText);
            }

            String wordPath = output + DOC;
            document.write(new FileOutputStream(wordPath));
            System.out.println("word文件已生成：" + wordPath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
