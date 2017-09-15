package util;

import java.awt.AlphaComposite;
import java.awt.Graphics2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.imageio.ImageIO;

/**
 * @author Administrator
 *
 */
public class StartMain {

	/**
	 * @param args
	 * @throws Exception 
	 */
	public void ImageToImage(String srcImagePath,String appendImagePath, 
			float alpha,int x,int y,int width,int height, 
			String imageFormat,String toPath) throws IOException{ 
		FileOutputStream fos = null; 
		try { 
			//读图 
			BufferedImage image = ImageIO.read(new File(srcImagePath)); 
			//创建java2D对象 
			Graphics2D g2d=image.createGraphics(); 
			//用源图像填充背景 
			g2d.drawImage(image, 0, 0, image.getWidth(), image.getHeight(), null, null); 

			//关键地方 
			AlphaComposite ac = AlphaComposite.getInstance(AlphaComposite.SRC_OVER, alpha); 
			g2d.setComposite(ac); 

			BufferedImage appendImage = ImageIO.read(new File(appendImagePath)); 
			g2d.drawImage(appendImage, x, y, width, height, null, null); 
			g2d.dispose(); 
			fos=new FileOutputStream(toPath); 
			ImageIO.write(image, imageFormat, fos); 
		} catch (Exception e) { 
			e.printStackTrace(); 
		}finally{ 
			if(fos!=null){ 
				fos.close(); 
			} 
		} 
	} 

	@SuppressWarnings("static-access")
	public static void main(String[] args) throws Exception {
		System.out.println("TestStart");
		String word = "C://Users/Administrator/Desktop/应收账款管理相关说明.docx";
		String image = "C://Users/Administrator/Desktop/1.png";
		new WordUtil().setImageWaterMarkAndProtect(word,image);
		System.out.println("TestEnd");
	}

}
