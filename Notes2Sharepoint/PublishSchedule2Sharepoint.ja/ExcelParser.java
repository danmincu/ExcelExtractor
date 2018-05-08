import java.text.SimpleDateFormat;
import java.util.Calendar;

import lotus.domino.NotesException;
import drnmain.Config;

/**
 * 
 * This allows to easily extract the information needed from the configuration file
 * and then, with that information, extract the data needed from the file queried 
 * in the beginning by using the download() method to get the file from sharepoint.
 * After, use the RunParser() method to trigger the ExcelExtractorWrapper class, and 
 * call the run method to retrieve the data and output it into a string format. After
 * that, call the upload() method to then upload the new string outputed to the 
 * sharepoint site.
 * 
 * 
 * 
 * Usage Instructions:
 * 	Instantiate
 *  call InitProxyCon - MUST be called after instantiating
 *  ======
 *  setScheduleFileName	- sets the file name that is needed
 *  setScheduleFolderName - sets the current folder that will be need
 *  setSchedulePublishedFileName - sets the file located in sharepoint
 *  setScheduledPublishedFolder - sets the folder located in sharepoint
 *  setExtractorFileName - sets the program that will extract the info  
 *  setSheetDetails -  sets the details for the extract part
 *  setNmbOfDays - sets the number of days to extract
 *  setLocalWorkFolder - sets the path of the local folder
 *  ======
 *  RunParser - use to extract the data from the file and output it into a string format
 *  download  - Downloads a file from the folder, site, and the file name already defined previously
 *  upload  - Uploads a file to sharepoint to the folder already defined, and the file name also defined 
 * 
 * @author Daniel.Diego-Garcia
 */
public class ExcelParser 
{
	

	Config config;
	String sourceFileName;
	String sourceFolderName;
	String destinationFileName;
	String destinationFolderName;
	String extractorFileName;
	String sheetDetails;
	String nmbOfDays;
	String localWorkFolder;
	String currentDate;
	SharepointConnector con;

	/**
	 * Constructor
	 * @param config The Config is empirical in order to have the other methods working
	 * as it will query the to the database.
	 */
	public ExcelParser(Config config)
	{
		this.config = config;
		this.con = new SharepointConnector();
		setCurrentDate();
	}
	
	/**
	 * @return the currentDate
	 */
	public String getCurrentDate()
	{
		return currentDate;
	}

	/**
	 * 
	 * Sets the currentData variable 
	 */
	public void setCurrentDate()
	{
		Calendar cal = Calendar.getInstance();
		SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd");
		this.currentDate = formatter.format(cal.getTime());
	}

	/**
	 * @return the config
	 */
	public Config getConfig()
	{
		return config;
	}

	/**
	 * @param config the config to set
	 */
	public void setConfig(Config config)
	{
		this.config = config;
	}


	/**
	 * @return the extractorFileName
	 */
	public String getExtractorFileName()
	{
		return extractorFileName;
	}

	/**
	 * @param extractorFileName the extractorFileName to set
	 */
	public void setExtractorFileName(String extractorFileName)
	{

		this.extractorFileName = getValue(extractorFileName);
	}

	/**
	 * @return the sheetDetails
	 */
	public String getSheetDetails()
	{
		return sheetDetails;
	}

	/**
	 * @param sheetDetails the sheetDetails to set
	 */
	public void setSheetDetails(String sheetDetails)
	{
		this.sheetDetails = getValue(sheetDetails);
	}

	/**
	 * @return the nmbOfDays
	 */
	public String getNmbOfDays()
	{
		return nmbOfDays;
	}

	/**
	 * @param nmbOfDays the nmbOfDays to set
	 */
	public void setNmbOfDays(String nmbOfDays)
	{
		this.nmbOfDays = getValue(nmbOfDays);
	}

	/**
	 * @return the localWorkFolder
	 */
	public String getLocalWorkFolder()
	{
		return localWorkFolder;
	}

	/**
	 * @param localWorkFolder the localWorkFolder to set
	 */
	public void setLocalWorkFolder(String localWorkFolder)
	{
		this.localWorkFolder = getValue(localWorkFolder);
	}

	/**
	 * @return the sourceFileName
	 */
	public String getSourceFileName()
	{
		return sourceFileName;
	}

	/**
	 * @param sourceFileName the sourceFileName to set
	 */
	public void setSourceFileName(String sourceFileName)
	{
		this.sourceFileName = getValue(sourceFileName);
	}

	/**
	 * @return the sourceFolderName
	 */
	public String getSourceFolderName()
	{
		return sourceFolderName;
	}

	/**
	 * @param sourceFolderName the sourceFolderName to set
	 */
	public void setSourceFolderName(String sourceFolderName)
	{
		this.sourceFolderName = getValue(sourceFolderName);
	}

	/**
	 * @return the destinationFileName
	 */
	public String getDestinationFileName()
	{
		return destinationFileName;
	}

	/**
	 * @param destinationFileName the destinationFileName to set
	 */
	public void setDestinationFileName(String destinationFileName)
	{
		this.destinationFileName = getValue(destinationFileName);
	}

	/**
	 * @return the destinationFolderName
	 */
	public String getDestinationFolderName()
	{
		return destinationFolderName;
	}

	/**
	 * @param destinationFolderName the destinationFolderName to set
	 */
	public void setDestinationFolderName(String destinationFolderName)
	{
		this.destinationFolderName = getValue(destinationFolderName);
	}

	/**
	 * @return the con
	 */
	public SharepointConnector getCon()
	{
		return con;
	}

	/**
	 * @param con the con to set
	 */
	public void setCon(SharepointConnector con)
	{
		this.con = con;
	}

	/**
	 * 
	 * This only initialized the proxy by giving the client id, client secret,
	 * and the site, and then it calls the initProxy method so downloading,
	 * and uploading can be performed.
	 * 
	 * @param site The sharepoint site
	 * @param clientId The authorized client
	 * @param clientSecret the access pass for the authorized client
	 *
	 * Creation date: Feb 5, 2018
	 * author: Daniel Diego-Garcia
	 * copyright: <b><i>Daniel Diego-Garcia</b></i>
	 */
	public void InitProxyCon(String site, String clientId, String clientSecret)
	{
		this.con.setClientID(getValue(site));
		this.con.setClientSecret(getValue(clientSecret));
		this.con.setSite(getValue(site));
		
		this.con.initProxy();
	}

	/**
	 * Sets the right path, the right folder, and the right file name 
	 * so the right file will be downloaded into the right directory
	 * 
	 * Creation date: Feb 5, 2018
	 * author: Daniel Diego-Garcia
	 * copyright: <b><i>Daniel Diego-Garcia</b></i>
	 */
	public void download()
	{
		this.con.setPath(this.localWorkFolder);
		this.con.setFolderName(this.sourceFolderName);
		this.con.setFileName(this.sourceFileName);
		this.con.download();
	}

	/**
	 * Takes care of setting the right file name and the right folder name
	 * where it will be uploaded
	 * 
	 * @param output The file name that has to be upload after the process was done
	 *
	 * Creation date: Feb 5, 2018
	 * author: Daniel Diego-Garcia
	 * copyright: <b><i>Daniel Diego-Garcia</b></i>
	 */
	public void upload(String output)
	{
		this.con.setFileName(output);
		this.con.setFolderName(this.destinationFolderName);
		this.con.upload();
	}

	/**
	 * 
	 * Invokes the ExcelExtractorWrapper giving the following arguments according to the 
	 * constructor.
	 * 
	 * After giving the variables, it will run the run() method and return the string
	 * which is the data into a string format
	 *
	 * Creation date: Feb 2, 2018
	 * author: Daniel Diego-Garcia
	 * copyright: <b><i>Daniel Diego-Garcia</b></i>
	 */
	public String RunParser() throws Exception
	{
		ExcelExtractorWrapper wrapper = new ExcelExtractorWrapper(
				this.extractorFileName, this.localWorkFolder + "\\"
						+ this.sourceFileName, this.sheetDetails,
				getCurrentDate(), this.nmbOfDays, this.destinationFileName);
		
		return wrapper.run();
	}

	/**
	 * 
	 * Use only within the class to get the value from the configuration file 
	 * by providing the key for what the user is looking for
	 * 
	 * @param key The key provided to look for
	 * @return a String that is the value it was looked for in the configuration file
	 *
	 * Creation date: Feb 2, 2018
	 * author: Daniel Diego-Garcia
	 * copyright: <b><i>Daniel Diego-Garcia</b></i>
	 */
	private String getValue(String key)
	{
		try
		{
			 return this.config.getKeyword(key);
		}
		catch(NotesException e)
		{
			e.printStackTrace();
		}
		
		return "";
	}
	
	public void moveTo()
	{

	}
	
	
}
