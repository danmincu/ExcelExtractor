import java.util.ArrayList;


public class ExcelFileInfo {
	
	private ArrayList sheetDetails;
	private String fileNameAndPath;
	
	public ExcelFileInfo(String excelFileToExtractFrom)
	{
		this.fileNameAndPath = excelFileToExtractFrom;
		this.sheetDetails = new ArrayList();
	}
	
	public String getFileNameAndPath()
	{
		return this.fileNameAndPath;
	}
	
	public SheetDetail CreateSheetDetail(String sheetName)
	{
		SheetDetail sheet = new SheetDetail(sheetName);
		this.sheetDetails.add(sheet);
		return sheet;
	}	
	
	public String sheetDetailsToParams()
	{
		StringBuilder sbStr = new StringBuilder();
	    for (int i = 0, il = this.sheetDetails.size(); i < il; i++) {
	        if (i > 0)
	            sbStr.append("|");
	       
	        SheetDetail sd = (SheetDetail)(this.sheetDetails.get(i));
	        
	        sbStr.append(sd.toParams());
	    }
	   return sbStr.toString();
	}
}