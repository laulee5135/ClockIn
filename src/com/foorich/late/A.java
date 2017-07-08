package com.foorich.late;

import java.io.BufferedReader;
import java.io.File;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URL;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import jxl.Sheet;
import jxl.Workbook;
import net.sf.json.JSONObject;

public class A {
	static Date parse;
	public static void main(String[] args) throws Exception {
	
		
		StringBuilder all = new StringBuilder();
		//����Fileʵ����ָ����Ҫ��ȡ��Excel���ļ�  
        File file = null;  
        file = new File("C:/Users/LauLee/Desktop/66.xls");  
        Workbook wb = ExcelReader.getWorkBook(file);  
        Sheet sheet = ExcelReader.getWorkBookSheet(wb, 0);
		
		
	    //��ȡ3��3�е�����	2017-04-01 ~ 2017-04-28
        String celldata = ExcelReader.readExcelData(sheet, 3, 3);
        
        //��ȡ�����ܵ�������28,29,30,31��
		String lastDayStr = celldata.substring(celldata.length()-2, celldata.length());
		Integer lastDay = Integer.valueOf(lastDayStr);
		//��ȡ��Ҫ�������ݼ��·��ַ���
		String yearmonthStr = "2017-04-";
		
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        SimpleDateFormat f = new SimpleDateFormat("yyyyMMdd");
        SimpleDateFormat timeFormate = new SimpleDateFormat("HH:mm");
		//��ȡ�����ϰ��ռ���
        List<Object> workDayList = new ArrayList<>();
		 for (int j = 1; j <= lastDay; j++) {
			 String str = yearmonthStr + j;  
	        	
				parse = sdf.parse(str);  //2017-04-01
				
		        String httpArg=f.format(parse); //20170401
		       //�жϸ�����   �ϰ� ����ĩ ���ڼ��գ�
		        String jsonResult = request(httpArg);
		        JSONObject jsonObject = JSONObject.fromObject(jsonResult);
		        String object = (String) jsonObject.get(httpArg);
		        Integer result = Integer.valueOf(object);
		        //Integer result = Integer.valueOf(jsonResult);
		        if (result == 0) {
					workDayList.add(j);
				}
		 }
	        	
		 System.out.println("�ϰ��գ�" + workDayList.toString());
		
		
		int count = 0;//ȫ��û��
		int sum = 0; //�����������û��
		
		int late = 0;
		//��ȡÿ���˸��µ����д���Ϣ
		for (int i = 5; i <= 348; i++) {
        	//����&Ա������  
        	String jobno = ExcelReader.readExcelData(sheet, i, 3);
        	String staffname =  ExcelReader.readExcelData(sheet, i, 11);
        	i++;
        
		
	        //
	        Date timeMor = timeFormate.parse("09:01");
	        Date time06 = timeFormate.parse("09:06");
	        Date timeAft = timeFormate.parse("17:59");
	        Date noon = timeFormate.parse("12:59");
	      
	        Iterator it = workDayList.iterator();
	        while(it.hasNext()) {
	        	Integer j = (Integer) it.next();
	        	
	        	String singlePerson = ExcelReader.readExcelData(sheet, i, j);
	        	if (!"".equals(singlePerson)) { //����
	        		
	        		
	        		// for(int index=0;index < singlePerson.length();index+=5){
	        		//��һ��ʱ��
        			 String d1 = singlePerson.substring(0, 5);  
        			 Date firstTime = timeFormate.parse(d1);
        			//���һ��ʱ��
        			 String d2 = singlePerson.substring(singlePerson.length()-5, singlePerson.length());  
        			 Date lastTime = timeFormate.parse(d2);
        		
        			 //���ֻ����һ�ε����
        			if (compareDate(firstTime, lastTime)) {
						sum++;
						
					}else { //��������
						if (compareDate(firstTime, time06) || compareDate(timeAft, lastTime)) {
							late++;
						}
					}
					
        		
				}else { //û��
					count ++ ;
				}
	        }
	        	
			
	        String job = String.format("%-5s", jobno);
	        String staff = String.format("%-5s", staffname);
	        StringBuilder sb = new StringBuilder();
	        sb.append("���ţ�" + job + "����:" + staff );
	        if (count!=0) {
	        	
				sb.append("   " + count + "��ȫ��û��;" );
				
			}
	        
	        if (sum != 0 && sum > 3) {
				
				sb.append("  " + (sum-3) + "������;");
			}
	        
	        if (late != 0) {
				
				sb.append("  " + late + "�γٵ�������");
			}
	        
			if (count ==0 && sum <=3 && late == 0) {
				sb.append("��ͬ־��û�κβ�����¼!");
			}
				
			
	       
	     
	        
	        //System.out.println("���ţ�" + jobno + "	������" + staffname + "  " + count + "��ȫ��û��" + sum +"��û�� " + ";  �ٵ������ˣ�" + late );
	        System.out.println(sb.toString());
	     
	        count = 0;
	        sum = 0;
	        late = 0;
	        
	        all.append(sb+"<br>");
	        
    	//
		}
		
		System.exit(0);
	}
	

	/**
	 * ���ڱ��^����
	 * @param d1
	 * @param d2
	 * @return
	 */
	public static boolean compareDate(Date d1, Date d2) {  
	    Calendar c1 = Calendar.getInstance();  
	    Calendar c2 = Calendar.getInstance();  
	    c1.setTime(d1);  
	    c2.setTime(d2);  
	  
	    int result = c1.compareTo(c2);  
	    if (result >= 0)  
	        return true;  
	    else  
	        return false;  
	} 
	
	 /**
	  * �ж��·��еĹ����ա���ĩ���ڼ���
     * @param urlAll
     *            :����ӿ�
     * @param httpArg
     *            :����
     * @return ���ؽ��
     * 		   0 �ϰ� 1��ĩ 2�ڼ���
     */
    public static String request( String httpArg) {
        String httpUrl="http://www.easybots.cn/api/holiday.php";//���ڽӿ�
    	//String httpUrl = "http://tool.bitefu.net/jiari/";
        BufferedReader reader = null;
        String result = null;
        StringBuffer sbf = new StringBuffer();
        httpUrl = httpUrl + "?d=" + httpArg;

        try {
            URL url = new URL(httpUrl);
            HttpURLConnection connection = (HttpURLConnection) url
                    .openConnection();
            connection.setRequestMethod("GET");
            connection.connect();
            InputStream is = connection.getInputStream();
            reader = new BufferedReader(new InputStreamReader(is, "UTF-8"));
            String strRead = null;
            while ((strRead = reader.readLine()) != null) {
                sbf.append(strRead);
                sbf.append("\r\n");
            }
            reader.close();
            result = sbf.toString();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }

}
