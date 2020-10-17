package com.create.format.run;

import java.util.Date;

public class FormatDate {

	public static void main(String[] args) {
		String sDate = "2010-05-10";
		Date date = DateUtil.getDateToStringYYYYMMDD(sDate);
        String strDate =DateUtil.getStringOfDateDDMMYYYY(date);
		System.out.println(strDate);
	}

}
