package com.create.format.run;


import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.time.LocalDate;
import java.time.ZoneId;

public class DateUtil {

    private DateUtil() {
    }


    public static Date getDateToStringDDMMYYYY(String sDate){
        try {
            return new SimpleDateFormat("dd-MM-yyyy").parse(sDate);
        } catch (Exception e) {
        	System.out.println(e.getMessage());
        }
        return null;
    }
    
    public static Date getDateToStringYYYYMMDD(String sDate){
        try {
            return new SimpleDateFormat("yyyy-MM-dd").parse(sDate);
        } catch (Exception e) {
        	System.out.println(e.getMessage());
        }
        return null;
    }
    
    public static String getStringOfDateDDMMYYYY(Date date){
    	String strDate = null;
        try {
        	DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");  
        	strDate = dateFormat.format(date);  
        } catch (Exception e) {
        	System.out.println(e.getMessage());
        }
        return strDate;
    }

    public static LocalDate convertDateToLocalDate(Date date) {
        return date.toInstant()
                .atZone(ZoneId.systemDefault())
                .toLocalDate();
    }

}