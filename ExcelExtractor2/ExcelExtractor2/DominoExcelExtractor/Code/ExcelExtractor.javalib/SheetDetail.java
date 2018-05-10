public class SheetDetail {
	
	private String sheetName;
	private int anStartColumn;
	private int anStartRow;
	private int anStopColumn;
	private int anStopRow;
	private int calStartRow;
	private int calStopRow;
	
	public SheetDetail(String sheetName)
	{
		this.sheetName = sheetName.replace(" ", "%20");
	}
		
	public void setAirplaneNamesRange(int startColumn, int startRow, int endColumn, int endRow)
	{
		this.anStartColumn = startColumn;
		this.anStartRow = startRow;
		this.anStopColumn = endColumn;
		this.anStopRow = endRow;
		
	}
		
	public void setDatesRange(int startCalendarRow, int endCalendarRow)
	{
	  this.calStartRow = startCalendarRow;
	  this.calStopRow = endCalendarRow;
	}
	
	public String toParams()
	{
		//example T1 Jets only;2;1;17;1;1;4
		return this.sheetName + ";" +
		Integer.toString(this.anStartColumn) + ";" +
		Integer.toString(this.anStartRow) + ";" + 
		Integer.toString(this.anStopColumn) + ";" + 
		Integer.toString(this.anStopRow) + ";" + 
		Integer.toString(this.calStartRow) + ";" + 
		Integer.toString(this.calStopRow);
	}

}

