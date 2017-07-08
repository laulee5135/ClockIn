package com.foorich.late;

public class MyTest {
	// 0 上班 1周末 2节假日
	public static void main(String[] args) {
		
		String httpArg = "20170501";
		String request = StatisticsServlet.request(httpArg);
		System.out.println(request);
		
	}

}
