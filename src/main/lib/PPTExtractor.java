package main.lib;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hslf.usermodel.HSLFObjectData;
import org.apache.poi.hslf.usermodel.HSLFObjectShape;
import org.apache.poi.hslf.usermodel.HSLFPictureData;
import org.apache.poi.hslf.usermodel.HSLFPictureShape;
import org.apache.poi.hslf.usermodel.HSLFShape;
import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFSoundData;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;


/**
 * Demonstrates how you can extract misc embedded data from a ppt file
 */
public final class PPTExtractor implements Extractor {

    public void extract(String fileName) {
        
        String file = "";

        try {

            ClassLoader classLoader = getClass().getClassLoader();

            if (classLoader.getResource(fileName) != null) {
                file = classLoader.getResource(fileName).getFile();
            } else {
                file = "resources/" + fileName;
            }

            FileInputStream is = new FileInputStream(file);
            HSLFSlideShow ppt = new HSLFSlideShow(is);

            //extract all sound files embedded in this presentation
            HSLFSoundData[] sound = ppt.getSoundData();
            for (HSLFSoundData aSound : sound) {
                String type = aSound.getSoundType();  //*.wav
                String name = aSound.getSoundName();  //typically file name
                byte[] data = aSound.getData();       //raw bytes

                //save the sound  on disk
                try (FileOutputStream out = new FileOutputStream(name + type)) {
                    out.write(data);
                }
            }

            int oleIdx = -1, picIdx = -1;
            for (HSLFSlide slide : ppt.getSlides()) {
                //extract embedded OLE documents
                for (HSLFShape shape : slide.getShapes()) {
                    if (shape instanceof HSLFObjectShape) {
                        oleIdx++;
                        HSLFObjectShape ole = (HSLFObjectShape) shape;
                        HSLFObjectData data = ole.getObjectData();
                        String name = ole.getInstanceName();
                        if ("Worksheet".equals(name)) {

                            //read xls
                            @SuppressWarnings({"unused", "resource"})
                            HSSFWorkbook wb = new HSSFWorkbook(data.getInputStream());

                        } else if ("Document".equals(name)) {
                            try (HWPFDocument doc = new HWPFDocument(data.getInputStream())) {
                                //read the word document
                                Range r = doc.getRange();
                                for (int k = 0; k < r.numParagraphs(); k++) {
                                    Paragraph p = r.getParagraph(k);
                                    System.out.println(p.text());
                                }

                                //save on disk
                                try (FileOutputStream out = new FileOutputStream(name + "-(" + (oleIdx) + ").doc")) {
                                    doc.write(out);
                                }
                            }
                        } else {
                            try (FileOutputStream out = new FileOutputStream(ole.getProgId() + "-" + (oleIdx + 1) + ".dat");
                                InputStream dis = data.getInputStream()) {
                                byte[] chunk = new byte[2048];
                                int count;
                                while ((count = dis.read(chunk)) >= 0) {
                                    out.write(chunk, 0, count);
                                }
                            }
                        }
                    }

                    //Pictures
                    else if (shape instanceof HSLFPictureShape) {
                        picIdx++;
                        HSLFPictureShape p = (HSLFPictureShape) shape;
                        HSLFPictureData data = p.getPictureData();
                        String ext = data.getType().extension;
                        try (FileOutputStream out = new FileOutputStream("pict-" + picIdx + ext)) {
                            out.write(data.getData());
                        }
                    }
                }
            }
        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        
    }
    
    public void showFiles() {
        
    }
}
