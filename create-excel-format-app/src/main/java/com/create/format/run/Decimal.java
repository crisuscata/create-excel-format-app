package com.create.format.run;

import org.decimal4j.util.DoubleRounder;

public class Decimal {

	public static void main(String[] args) {
		Double PI = 12345678912346.1416787;
		
		System.out.println(DoubleRounder.round(PI, 3));

	}
	

}
