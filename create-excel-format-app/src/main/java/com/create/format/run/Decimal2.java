package com.create.format.run;

import java.math.BigDecimal;
import java.math.RoundingMode;

public class Decimal2 {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		Double PI = 12345674.1416787;
		//DecimalFormat df = new DecimalFormat("##############.##");
		
		
		//String snewPI = df.format(PI);
		//Double newPI = Double.parseDouble(df.format(PI));
		
		System.out.println(round(PI, 2));
		
	}
	
	private static double round(double value, int places) {
	    if (places < 0) throw new IllegalArgumentException();
	 
	    BigDecimal bd = new BigDecimal(Double.toString(value));
	    bd = bd.setScale(places, RoundingMode.HALF_UP);
	    return bd.doubleValue();
	}

}
