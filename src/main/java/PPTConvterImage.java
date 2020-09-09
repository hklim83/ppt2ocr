import java.awt.Color;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import javax.imageio.ImageIO;

import org.apache.poi.hslf.model.Slide;
import org.apache.poi.hslf.usermodel.SlideShow;
import org.opencv.core.Core;
import org.opencv.core.Mat;
import org.opencv.imgcodecs.Imgcodecs;
import org.opencv.imgproc.Imgproc;

import net.sourceforge.tess4j.Tesseract;


public class PPTConvterImage {
 
	private File pptFile;
	private File cvtOrgImgFile;
	private File cvtGrayImgFile;
	static Tesseract tesseract = new Tesseract();
	static {
		System.loadLibrary(Core.NATIVE_LIBRARY_NAME);
		tesseract.setDatapath("C:/Program Files/Tesseract-OCR/tessdata");
		tesseract.setTessVariable("user_defined_dpi", "300");
	}
	
	public PPTConvterImage(File pptFile, File convertImgOrgFilePath, File convertImgGrayFilePath) {
		this.pptFile = pptFile;
		this.cvtOrgImgFile = convertImgOrgFilePath;
		this.cvtGrayImgFile = convertImgGrayFilePath;
	}
 
	public void convter(String type) throws Exception {
 
		FileInputStream is = new FileInputStream(pptFile);
 		SlideShow ppt = new SlideShow(is);
		is.close();
 
		Dimension pgsize = ppt.getPageSize();
		Slide[] slide = ppt.getSlides();
 
		for (int i = 0; i < slide.length; i++) {
 			BufferedImage img = new BufferedImage(pgsize.width, pgsize.height, BufferedImage.TYPE_INT_RGB);
			Graphics2D graphics = img.createGraphics();
			graphics.setPaint(Color.white);
			graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width, pgsize.height));
			
			slide[i].draw(graphics);
			
			String fileName = this.cvtOrgImgFile.getAbsolutePath()+"\\" + (i + 1)+ "."+type;
			FileOutputStream out = new FileOutputStream(fileName);
			ImageIO.write(img, type, out);
			out.close();
			
			Mat origin = Imgcodecs.imread(fileName);
			String result = extractString(origin, fileName, cvtGrayImgFile);
			System.out.println(result);
			
		}
	}
	
	public String extractString(Mat inputMat, String fileName, File cvtGrayImgFile) {
		String result = "";
		Mat gray = new Mat();
		String grayfilename =  fileName.substring(fileName.lastIndexOf("\\")+1, fileName.length());
		System.out.println(grayfilename);
		Imgproc.cvtColor(inputMat, gray, Imgproc.COLOR_BGR2GRAY);
		
		Imgcodecs.imwrite(cvtGrayImgFile+"\\" + grayfilename, gray); 
		  try { 
			  result = tesseract.doOCR(new File(cvtGrayImgFile +"\\"+ grayfilename)); 
		  } catch (Exception e) {
			  e.printStackTrace(); 
		  }
		return result;
	}
 

	public static void main(String[] args) {
		try {

			File pptFile = new File("c:/opencv_test/SILICON.ppt");
			String pptFilePath = pptFile.getAbsolutePath().substring(0, pptFile.getAbsolutePath().lastIndexOf("\\")+1);
			File convertImgOrgFilePath = new File( pptFilePath  +pptFile.getName()+"_org");
			File convertImgGrayFilePath = new File( pptFilePath  +pptFile.getName()+"_gray");
			
			if (!convertImgOrgFilePath.exists()) {
				convertImgOrgFilePath.mkdir();
			}
			if (!convertImgGrayFilePath.exists()) {
				convertImgGrayFilePath.mkdir();
			}
			
			PPTConvterImage cvtImage = new PPTConvterImage(pptFile, convertImgOrgFilePath, convertImgGrayFilePath);
			cvtImage.convter("png");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}