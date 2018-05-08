import java.io.File;
import java.util.Vector;

/**
 * 
 * This will act as a bridge between the agent that will want to create, upload,
 * or download a file. First, the site, the clientID, and the clientSecret must
 * be provided at the beginning. After, to be able to execute some methods the
 * initProxy() function needs to be called; else, nothing will happen. Setters
 * can be used to set the name of the file, a new file (newFileName var is use
 * only to modify a file), and the destination folder.
 * 
 * To delete a file,download, or upload a file, just call the appropriate
 * methods. If a different file needs to be modified in a different location,
 * and different site reset the folder, fileName, and site.
 * 
 * Usage Instructions:
 * 	Instantiate
 *	Set site - the folder path in sharepoint
 * 	Set clientID  - the authorized client ID
 * 	Set clientSecret - password from the authorized Client
 *  call initProxy - MUST be called in order after initializing the variables above
 *  listContent  - lists the content of the folder in sharepoint
 *  set path - set the path where the file should go
 *  set newFileName - use to rename a file
 *  set folder - use to know to what folder the file will go
 *  renameFile - Renames the file of the 
 *  confirmDownload  - This is used after download a file from sharepoint. It will check whether the file has been downloaded or not
 *  deleteFile  - This deletes a file in the local machine
 *  download  - Downloads a file from the folder, site, and the file name already defined previously
 *  upload  - Uploads a file to sharepoint to the folder already defined, and the file name also defined with the setters
 *  moveTo  - this will move a file from one directory to another or drive ( C:\, E:\, etc...)
 * 
 * 
 * @author Daniel.Diego-Garcia
 * 
 */
public class SharepointConnector
{

	private String site;
	private String clientID;
	private String clientSecret;
	private String folder;
	private String path;
	private String fileName;
	private String newFileName;
	private SharepointProxy proxy;

	final private static String BACKSLASH = "\\";

	public SharepointConnector()
	{
	}

	/**
	 * @return the site
	 */
	public String getSite()
	{
		return site;
	}

	/**
	 * @param site
	 *        the site to set
	 */
	public void setSite(String site)
	{
		this.site = site;
	}

	/**
	 * 
	 * @return
	 *
	 * Creation date: Feb 2, 2018
	 * author: Daniel Diego-Garcia
	 * copyright: <b><i>Daniel Diego-Garcia</b></i>
	 */
	public String getClientID()
	{
		return clientID;
	}

	/**
	 * @param clientID
	 *        the clientID to set
	 */
	public void setClientID(String clientID)
	{
		this.clientID = clientID;
	}

	/**
	 * @return the clientSecret
	 */
	public String getClientSecret()
	{
		return clientSecret;
	}

	/**
	 * @param clientSecret
	 *        the clientSecret to set
	 */
	public void setClientSecret(String clientSecret)
	{
		this.clientSecret = clientSecret;
	}

	/**
	 * @return the folderName
	 */
	public String getFolder()
	{
		return folder;
	}

	/**
	 * @param folderName
	 *        the folderName to set
	 */
	public void setFolderName(String folderName)
	{
		this.folder = folderName;
	}

	/**
	 * @return the fileName
	 */
	public String getFileName()
	{
		return fileName;
	}

	/**
	 * @param fileName
	 *        the fileName to set
	 */
	public void setFileName(String fileName)
	{
		this.fileName = fileName;
	}

	/**
	 * @return the newFileName
	 */
	public String getNewFileName()
	{
		return newFileName;
	}

	/**
	 * @param newFileName
	 *        the newFileName to set
	 */
	public void setNewFileName(String newFileName)
	{
		this.newFileName = newFileName;
	}

	/**
	 * @return the path
	 */
	public String getPath()
	{
		return path;
	}

	/**
	 * @param path
	 *        the path to set
	 */
	public void setPath(String path)
	{
		this.path = path;
	}

	/**
	 * 
	 * This will call the SharepointProxy object and 
	 * initialize it.
	 *
	 * Creation date: Feb 2, 2018
	 * author: Daniel Diego-Garcia
	 * copyright: <b><i>Daniel Diego-Garcia</b></i>
	 */
	public void initProxy()
	{
		proxy = new SharepointProxy(this.site, this.clientID, this.clientSecret);
	}

	/**
	 * 
	 * List the content of the sharepoint folder
	 *
	 * Creation date: Feb 2, 2018
	 * author: Daniel Diego-Garcia
	 * copyright: <b><i>Daniel Diego-Garcia</b></i>
	 */
	public void listContent()
	{
		try
		{
			Vector<String> files = proxy.getFolderList(this.folder);
			for (int i = 0; i < files.size(); i++)
			{
				System.out.println(files.get(i).toString());
			}

		}
		catch (Exception e)
		{
			e.printStackTrace();
		}

	}

	/**
	 * 
	 * Creates a reference to the file that has been downloaded in the local machine, and then it's
	 * going to check if the file is there first. After it calls a renameTo
	 * function to rename the file. If it's not successful, it will notify.
	 *
	 * Creation date: Feb 2, 2018
	 * author: Daniel Diego-Garcia
	 * copyright: <b><i>Daniel Diego-Garcia</b></i>
	 */
	public void renameFile()
	{
		File file = new File(this.path + BACKSLASH + this.fileName);
		String fullPath_newFile = this.path + BACKSLASH + this.newFileName;

		if (confirmDownload(file))
		{
			if (file.renameTo(new File(fullPath_newFile)))
			{
				this.fileName = this.newFileName;
				System.out.println("SUCCESS: File renamed successfully ");
			}
			else
			{
				System.out
						.println("FATAL: renameFile() failed : File couldn't be renamed");
			}
		}
		else
		{
			System.out.println("FATAL: file not found.");
		}
	}

	/**
	 * This only will check if the file has been downloaded and can be found
	 * 
	 * @param file
	 *        the that has to checked whether it's there or not
	 * @return true if file is found else false
	 *
	 * Creation date: Feb 2, 2018
	 * author: Daniel Diego-Garcia
	 * copyright: <b><i>Daniel Diego-Garcia</b></i>
	 */
	public boolean confirmDownload(File file)
	{
		if (!file.exists())
		{
			System.out.println("FATAL: file not found");
			return false;
		}
		else
		{
			System.out.println("SUCCESS: file downloaded...");
			return true;
		}
	}

	/**
	 * 
	 * Invokes the SharepointProxy object and deletes the file in the folder
	 * 
	 * Creation date: Feb 1, 2018 
	 * author: Daniel Diego-Garcia 
	 * copyright:<b><i>Daniel Diego-Garcia</b></i>
	 */
	public void deleteFile()
	{
		try
		{
			this.proxy.deleteFile(folder, fileName);
		}
		catch (Exception e)
		{
			e.printStackTrace();
		}
	}

	/**
	 * This downloads a file from a certain folder in the site; however, it
	 * first checks if the TACS folder exists.
	 * 
	 * Creation date: Feb 1, 2018 
	 * author: Daniel Diego-Garcia 
	 * copyright: <b><i>Daniel Diego-Garcia</b></i>
	 */
	public void download()
	{
		try
		{
			this.proxy
					.downloadBinaryFile(this.path, this.folder, this.fileName);
		}
		catch (Exception e)
		{
			e.printStackTrace();
		}
	}

	/**
	 * 
	 * This will invoke the SharepointProxy object to be able to upload to the
	 * site.
	 * 
	 * Creation date: Feb 1, 2018 
	 * author: Daniel Diego-Garcia 
	 * copyright: <b><i>Daniel Diego-Garcia</b></i>
	 */
	public void upload()
	{
		try
		{
			this.proxy.uploadBinaryFile(this.fileName, this.folder);
		}
		catch (Exception e)
		{
			e.printStackTrace();
		}
	}



}
