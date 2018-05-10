import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class ExcelExtractorWrapper
{
	private static String ExcelExtractorFilePath = "C:\\TACS\\ExcelExtractor2.exe";
	
	public static void SetExcelExtractorFilePath(String excelExtractorFilePath)
	{
		ExcelExtractorWrapper.ExcelExtractorFilePath = excelExtractorFilePath;
	}
	
	public static String ExtractFromExcel(ExcelFileInfo fileInfo, Date startingDate, int days, String outputFileName) throws Exception
	{
		SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd");
		String dateTimeString = formatter.format(startingDate);
		String sheetDetails = fileInfo.sheetDetailsToParams();
		
		if (sheetDetails == null || sheetDetails.equals(""))
		{
		    throw new Exception("Please provide at least 1 Sheet detail object!");	
		}
		
		ExcelExtractorWrapper wrapper = new ExcelExtractorWrapper(ExcelExtractorFilePath,
				fileInfo.getFileNameAndPath(), sheetDetails, dateTimeString, Integer.toString(days), outputFileName);
		try
		{
		  return wrapper.run();	
		}
		catch(ExcelExtractorException eex)
		{
			throw new Exception(eex.getExcelExtractorExceptionDetails());
		}
	}
	
	
	
    private String excelExtractorFileName;

    private String excelFilePath;

    private String outputFileName;

    private String positions;

    private String sheetDetails;

    private String yyyyMMdd;

    private String numberOfDaysToExtract;

    /**
     * 
     * @param excelExtractorFileName
     *            full file name of the excel extractor utility
     * @param excelFilePath
     *            full file path of the source excel file
     * @param sheetDetails
     *            "|" delimited [sheet name;airplane names start column;airplane
     *            names start row;stop column;stop row;calendar start
     *            row;calendar start column]
     * @param yyyyMMdd
     *            the start date string in the YYYY-MM-DD format
     * @param numberOFDaysToExtract
     *            the string representing the number days to extract from start
     *            date
     * @param outputFileName
     *            File name of the output
     */
    protected ExcelExtractorWrapper(String excelExtractorFileName,
            String excelFilePath, String sheetDetails, String yyyyMMdd,
            String numberOfDaysToExtract, String outputFileName)
    {
        this.excelExtractorFileName = excelExtractorFileName;
        this.excelFilePath = excelFilePath;
        this.outputFileName = outputFileName;
        this.sheetDetails = sheetDetails;
        this.yyyyMMdd = yyyyMMdd;
        this.numberOfDaysToExtract = numberOfDaysToExtract;
    }

    protected String run() throws Exception
    {

        this.positions = sheetDetails + "," + this.yyyyMMdd + ","
                + this.numberOfDaysToExtract;
        Object[] fArgs = new String[]
        { excelFilePath, this.positions, this.outputFileName };
        String[] args =
        { this.excelExtractorFileName, String.format("\"%s\",%s,\"%s\"", fArgs) };
        System.out.println("Executing:" + args[0] + " " + args[1]);
        Process p = Runtime.getRuntime().exec(args);
        String output = "";
        int i;
        while ((i = p.getInputStream().read()) != -1)
        {
            output = output + (char) i;
        }
        while ((i = p.getErrorStream().read()) != -1)
        {
            System.out.println("Error:");
            System.err.write(i);
        }
        p.waitFor();
        int exitValue = p.exitValue();

        // if the excel extraction was successful
        if (exitValue == 0)
        {
            String dataFileName = output.replace("\n", "").replace("\r", "")
                    + this.outputFileName;
            System.out.println("Extracted " + dataFileName);
            return dataFileName;
        }
        else
        {
            System.out.println("Error exit value:" + p.exitValue());
            String exceptionFileName = output.replace("\n", "").replace("\r", "") + "Exception.txt";            
            throw new ExcelExtractorException(readFile(exceptionFileName));
        }
    }
    
    private String readFile(String file) throws IOException {
        BufferedReader reader = new BufferedReader(new FileReader (file));
        String         line = null;
        StringBuilder  stringBuilder = new StringBuilder();
        String         ls = System.getProperty("line.separator");

        try {
            while((line = reader.readLine()) != null) {
                stringBuilder.append(line);
                stringBuilder.append(ls);
            }

            return stringBuilder.toString();
        } finally {
            reader.close();
        }
    }

}
