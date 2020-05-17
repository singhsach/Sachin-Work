package automation_spire;

import java.text.SimpleDateFormat;
import java.util.Date;

public class InputBean {
	private String path = "D:\\saurabh\\";
	public final String daily1 = "daily1";
	public final String daily2 = "daily2";
	public final String weekly = "weekly";
	public final String excel = "Result";
	public final String monthly = "monthly";
	public final String imagelis = "imageLis";
	public final String imageRpa = "imageRpa";
	public final String rpa1 = "Rpa1";
	public final String rpa2 = "Rpa2";

	String date;

	public String getExcel() {
		return path + excel + getDate() + ".xlsx";
	}

	public String getDate() {
		String pattern = "dd-MM-yyyy";
		SimpleDateFormat simpleDateFormat = new SimpleDateFormat(pattern);
		date = simpleDateFormat.format(new Date());
		return date;
	}

	public String convert(String file) {

		return path + file + getDate() + ".png";

	}

}
