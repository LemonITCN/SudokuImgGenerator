package cn.lemonit.play;

import com.sun.image.codec.jpeg.JPEGCodec;
import com.sun.image.codec.jpeg.JPEGImageEncoder;

import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.BufferedOutputStream;
import java.io.FileOutputStream;

public class sudoku_img_generator {

    public static void main(String[] args) {
        int width = 900;
        int height = 900;
        int font = 80;
        int outLineWidth = 5;
        int inLineWidth = 1;
        int numXDeviation = 20;
        int numYDeviation = 10;
        String outputPath = "e:/sudoku/";
        String fileName = "1.jpg";

        String content = "024806003000002070009004005086049000900020008000780950600400500030500000400207830";

        BufferedImage image = new BufferedImage(width + outLineWidth, height + outLineWidth, BufferedImage.TYPE_INT_RGB);
        Graphics graphics = image.getGraphics();
        graphics.setColor(Color.white);
        graphics.fillRect(0, 0, width, height);
        graphics.setColor(Color.BLACK);

        if (content.length() == 81) {
            // 9宫
            int widthStep = (width) / 9;
            int heightStep = (height) / 9;
            for (int i = 0; i <= 9; i++) {
                graphics.fillRect(0, i * heightStep, width, i % 3 == 0 ? outLineWidth : inLineWidth);
                graphics.fillRect(widthStep * i, 0, i % 3 == 0 ? outLineWidth : inLineWidth, height);
            }
            for (int i = 0; i < content.length(); i++) {
                graphics.setFont(new Font("Arial", Font.BOLD, font));
                String item = content.substring(i, i + 1);
                if (!item.equals("0")) {
                    graphics.drawString(item, i % 9 * widthStep + (widthStep - font) / 2 + numXDeviation, (i / 9 + 1) * heightStep - (heightStep - font) / 2 - numYDeviation);
                }
            }
        }
        createImage(image, outputPath + fileName);
    }

    private static void createImage(BufferedImage image, String fileLocation) {
        try {
            FileOutputStream fos = new FileOutputStream(fileLocation);
            BufferedOutputStream bos = new BufferedOutputStream(fos);
            JPEGImageEncoder encoder = JPEGCodec.createJPEGEncoder(bos);
            encoder.encode(image);
            bos.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
