package com.foorich.late;

public class MyTest {
	// 0 �ϰ� 1��ĩ 2�ڼ���
	public static void main(String[] args) {
		
		String httpArg = "20170501";
		String request = StatisticsServlet.request(httpArg);
		System.out.println(request);
		
	}

}
