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
		//生成File实例并指向需要读取的Excel表文件  
        File file = null;  
        file = new File("C:/Users/LauLee/Desktop/66.xls");  
        Workbook wb = ExcelReader.getWorkBook(file);  
        Sheet sheet = ExcelReader.getWorkBookSheet(wb, 0);
		
		
	    //获取3行3列的数据	2017-04-01 ~ 2017-04-28
        String celldata = ExcelReader.readExcelData(sheet, 3, 3);
        
        //获取该月总的天数（28,29,30,31）
		String lastDayStr = celldata.substring(celldata.length()-2, celldata.length());
		Integer lastDay = Integer.valueOf(lastDayStr);
		//获取需要计算的年份加月份字符串
		String yearmonthStr = "2017-04-";
		
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        SimpleDateFormat f = new SimpleDateFormat("yyyyMMdd");
        SimpleDateFormat timeFormate = new SimpleDateFormat("HH:mm");
		//获取本月上班日集合
        List<Object> workDayList = new ArrayList<>();
		 for (int j = 1; j <= lastDay; j++) {
			 String str = yearmonthStr + j;  
	        	
				parse = sdf.parse(str);  //2017-04-01
				
		        String httpArg=f.format(parse); //20170401
		       //判断该天是   上班 ？周末 ？节假日？
		        String jsonResult = request(httpArg);
		        JSONObject jsonObject = JSONObject.fromObject(jsonResult);
		        String object = (String) jsonObject.get(httpArg);
		        Integer result = Integer.valueOf(object);
		        //Integer result = Integer.valueOf(jsonResult);
		        if (result == 0) {
					workDayList.add(j);
				}
		 }
	        	
		 System.out.println("上班日：" + workDayList.toString());
		
		
		int count = 0;//全天没打
		int sum = 0; //上午或者下午没打
		
		int late = 0;
		//获取每个人该月的所有打卡信息
		for (int i = 5; i <= 348; i++) {
        	//工号&员工姓名  
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
	        	if (!"".equals(singlePerson)) { //打了
	        		
	        		
	        		// for(int index=0;index < singlePerson.length();index+=5){
	        		//第一个时间
        			 String d1 = singlePerson.substring(0, 5);  
        			 Date firstTime = timeFormate.parse(d1);
        			//最后一个时间
        			 String d2 = singlePerson.substring(singlePerson.length()-5, singlePerson.length());  
        			 Date lastTime = timeFormate.parse(d2);
        		
        			 //如果只打了一次的情况
        			if (compareDate(firstTime, lastTime)) {
						sum++;
						
					}else { //打了两次
						if (compareDate(firstTime, time06) || compareDate(timeAft, lastTime)) {
							late++;
						}
					}
					
        		
				}else { //没打
					count ++ ;
				}
	        }
	        	
			
	        String job = String.format("%-5s", jobno);
	        String staff = String.format("%-5s", staffname);
	        StringBuilder sb = new StringBuilder();
	        sb.append("工号：" + job + "姓名:" + staff );
	        if (count!=0) {
	        	
				sb.append("   " + count + "次全天没打;" );
				
			}
	        
	        if (sum != 0 && sum > 3) {
				
				sb.append("  " + (sum-3) + "次忘打卡;");
			}
	        
	        if (late != 0) {
				
				sb.append("  " + late + "次迟到或早退");
			}
	        
			if (count ==0 && sum <=3 && late == 0) {
				sb.append("好同志，没任何不良记录!");
			}
				
			
	       
	     
	        
	        //System.out.println("工号：" + jobno + "	姓名：" + staffname + "  " + count + "次全天没打；" + sum +"次没打卡 " + ";  迟到或早退：" + late );
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
	 * 日期比較方法
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
	  * 判断月份中的工作日、周末、节假日
     * @param urlAll
     *            :请求接口
     * @param httpArg
     *            :参数
     * @return 返回结果
     * 		   0 上班 1周末 2节假日
     */
    public static String request( String httpArg) {
        String httpUrl="http://www.easybots.cn/api/holiday.php";//日期接口
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
