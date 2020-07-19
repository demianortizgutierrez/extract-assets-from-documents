package main.lib;

import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;
import java.util.List;
import javax.imageio.ImageIO;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class DocExtractor implements Extractor {

    public void extract(String fileName) {

        String file = "";
        ClassLoader classLoader = getClass().getClassLoader();

        if (classLoader.getResource(fileName) != null) {
            file = classLoader.getResource(fileName).getFile();
        } else {
            file = "resources/" + fileName;
        }
        
        try {

            //create file inputstream to read from a binary file
            FileInputStream fs = new FileInputStream(file);
            //create office word 2007+ document object to wrap the word file
            XWPFDocument docx = new XWPFDocument(fs);
            //get al images from the document and store them in the list piclist
            List<XWPFPictureData> piclist = docx.getAllPictures();
            //traverse through the list and write each image to a file
            Iterator<XWPFPictureData> iterator = piclist.iterator();
            int i = 0;
            while (iterator.hasNext()) {
                XWPFPictureData pic = iterator.next();
                byte[] bytepic = pic.getData();
                BufferedImage imag = ImageIO.read(new ByteArrayInputStream(bytepic));
                ImageIO.write(imag, "jpg", new File("resources/doc-image-" + i + ".jpg"));
                i++;
            }

        } catch (Exception e) {
             System.out.println("Error on DOC images extraction");
             System.out.println(e.getMessage());
        }

    }

    @Override
    public void showFiles() {
        // TODO Auto-generated method stub
        
    }
}
