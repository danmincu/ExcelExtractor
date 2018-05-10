public class ExcelExtractorException extends Exception {

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	private String excelExtractorExceptionDetails;
	
	public ExcelExtractorException(String details)
	{
		this.excelExtractorExceptionDetails = details;
	}
	
	public void setExcelExtractorExceptionDetails(String value)
	{
		this.excelExtractorExceptionDetails = value;
	}
	
	public String getExcelExtractorExceptionDetails()
	{
		return this.excelExtractorExceptionDetails;
	}
	
}
