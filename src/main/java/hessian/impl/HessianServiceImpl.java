/**
 * 
 */
package hessian.impl;

import hessian.HessianService;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.MalformedURLException;

/**
 * @author Administrator
 *
 */
public class HessianServiceImpl implements HessianService {

	public String getNewMessage(String msg) {
		return "HESSIAN --> " + msg;
	}

	@SuppressWarnings("static-access")
	public void remoteWordUtil(String docPath) {
		try {
			String imgPath ="http://localhost:8081/EasyHessian/1.png";
			new util.WordUtil().setImageWaterMarkAndProtect(docPath, imgPath);
		} catch (Exception e) {
			e.printStackTrace();
		}	
	}

	public File returnToFile(File file,String FileFileName){
		//基于myFile创建一个文件输入流    		
		String prefix=FileFileName.substring(FileFileName.lastIndexOf(".")+1);
		String FileFileName2 = null;
		if(prefix.equals("doc")){
			FileFileName2 = "Upload.doc";
		}else{
			FileFileName2 = "Upload.docx";
		}
		File dirPath = new File("F:/WebServer/apache-tomcat-7.0.65/webapps/EasyHessian"+File.separator+"upload"+File.separator+"defineworkflow"+File.separator);
		File toFile = new File("F:/WebServer/apache-tomcat-7.0.65/webapps/EasyHessian"+File.separator+"upload"+File.separator+"defineworkflow"+File.separator, FileFileName2);    
		if(file!=null){
			boolean checkFolder = dirPath.exists(); // 检查文件夹是否存在
			if (checkFolder == false) {
				checkFolder = dirPath.mkdirs();
			}
			try {
				//				downLoadFromUrl.funcDownLoadFromUrl("http://192.168.130.10:8080/upload/"+FileFileName,FileFileName2,"E:/WebServer/apache-tomcat-7.0.65/webapps/EasyHessian"+File.separator+"upload"+File.separator);
				FileOutputStream fos = new FileOutputStream(dirPath+File.separator+FileFileName2);
				//以上传文件建立一个文件上传流
				FileInputStream fis = new FileInputStream(file);
				//将上传文件的内容写入服务器
				byte[] buffer = new byte[1024];
				int len = 0;
				while ((len = fis.read(buffer)) > 0)
				{
					fos.write(buffer , 0 , len);
				}
				fis.close();
				fos.close();

			} catch (MalformedURLException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		remoteWordUtil(dirPath+"/"+FileFileName2);
		return toFile;
	}

	public void upload(String filename, InputStream data) {
		File dirPath = new File("F:/WebServer/apache-tomcat-7.0.65/webapps/EasyHessian"+File.separator+"upload"+File.separator);
		boolean checkFolder = dirPath.exists(); // 检查文件夹是否存在
		if (checkFolder == false) {
			checkFolder = dirPath.mkdirs();
		}
		BufferedInputStream bis = null;
		BufferedOutputStream bos = null;
		try {
			//获取客户端传递的InputStream
			bis = new BufferedInputStream(data);
			//创建文件输出流
			bos = new BufferedOutputStream(new FileOutputStream(
					"F:/WebServer/apache-tomcat-7.0.65/webapps/EasyHessian"+File.separator+"upload"+File.separator + filename));
			byte[] buffer = new byte[8192];
			int r = bis.read(buffer, 0, buffer.length);
			while (r > 0) {
				bos.write(buffer, 0, r);
				r = bis.read(buffer, 0, buffer.length);
			}
			System.out.println("-------文件上传成功！-------------");
		} catch (IOException e) {
			throw new RuntimeException(e);
		} finally {
			if (bos != null) {
				try {
					bos.close();
				} catch (IOException e) {
					throw new RuntimeException(e);
				}
			}
			if (bis != null) {
				try {
					bis.close();
				} catch (IOException e) {
					throw new RuntimeException(e);
				}
			}
		}
	}

}
