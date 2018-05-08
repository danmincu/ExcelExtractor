public class ExcelExtractorWrapper
{

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
     *            "|" delimited [sheet name;airplane names start row;airplane
     *            names start column;stop row;stop column;calendar start
     *            row;calendar start column]
     * @param yyyyMMdd
     *            the start date string in the YYYY-MM-DD format
     * @param numberOFDaysToExtract
     *            the string representing the number days to extract from start
     *            date
     * @param outputFileName
     *            File name of the output
     */
    public ExcelExtractorWrapper(String excelExtractorFileName,
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

    public String run() throws Exception
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
            throw new Exception(
                    "Error executing the ExcelExtractor utility. For details an Exception.txt file was created in the temporary folder!");
        }
    }

}
