/**
 * 
 */
package hessian;

import java.io.File;
import java.io.InputStream;

/**
 * @author Administrator
 *
 */
public interface HessianService {

	public String getNewMessage(String msg);

	public void remoteWordUtil(String docPath);
	
	public File returnToFile(File file,String FileFileName);
	
	public void upload(String filename, InputStream data);
}
