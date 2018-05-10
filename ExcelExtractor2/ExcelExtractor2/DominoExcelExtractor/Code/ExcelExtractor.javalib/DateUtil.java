import java.util.Calendar;
import java.util.Date;

public class DateUtil {

	public static Date DateFromYMD(int year, int month, int day)
	{		
		Calendar cal = Calendar.getInstance();
	    cal.set(Calendar.DAY_OF_MONTH, day);
		cal.set(Calendar.MONTH, month - 1);
		cal.set(Calendar.YEAR, year);
		return cal.getTime();		
	}
	
}
