package com.foorich.late;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URL;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;

import org.apache.commons.fileupload.FileItem;
import org.apache.commons.fileupload.disk.DiskFileItemFactory;
import org.apache.commons.fileupload.servlet.ServletFileUpload;

import jxl.Sheet;
import jxl.Workbook;
import net.sf.json.JSONObject;

public class StatisticsServlet extends HttpServlet{

	@Override
	protected void doPost(HttpServletRequest req, HttpServletResponse resp) throws IOException{
		 String excelnum = "";
		try {
			
			 //得到上传文件的保存目录
            String savePath = this.getServletContext().getRealPath("/upload");
            String filename = "";
            File file = new File(savePath);
            //判断上传文件的保存目录是否存在
            if (!file.exists() && !file.isDirectory()) {
                System.out.println(savePath+"目录不存在，需要创建");
                //创建目录
                file.mkdir();
            }
            System.out.println(savePath);
          
            try{
                //使用Apache文件上传组件处理文件上传步骤：
                //1、创建一个DiskFileItemFactory工厂
                DiskFileItemFactory factory = new DiskFileItemFactory();
                //2、创建一个文件上传解析器
                ServletFileUpload upload = new ServletFileUpload(factory);
                 //解决上传文件名的中文乱码
                upload.setHeaderEncoding("UTF-8"); 
                //3、判断提交上来的数据是否是上传表单的数据
                if(!ServletFileUpload.isMultipartContent(req)){
                    //按照传统方式获取数据
                    return;
                }
                //4、使用ServletFileUpload解析器解析上传数据，解析结果返回的是一个List<FileItem>集合，每一个FileItem对应一个Form表单的输入项
                List<FileItem> list = upload.parseRequest(req);
                for(FileItem item : list){
                    //如果fileitem中封装的是普通输入项的数据
                    if(item.isFormField()){
                        String name = item.getFieldName();
                        //解决普通输入项的数据的中文乱码问题
                        excelnum = item.getString("UTF-8");
                       
                       
                    }else{//如果fileitem中封装的是上传文件
                        //得到上传的文件名称，
                        filename = item.getName();
                        System.out.println(filename);
                        if(filename==null || filename.trim().equals("")){
                            continue;
                        }
                        //注意：不同的浏览器提交的文件名是不一样的，有些浏览器提交上来的文件名是带有路径的，如：  c:\a\b\1.txt，而有些只是单纯的文件名，如：1.txt
                        //处理获取到的上传文件的文件名的路径部分，只保留文件名部分
                        filename = filename.substring(filename.lastIndexOf("\\")+1);
                        //获取item中的上传文件的输入流
                        InputStream in = item.getInputStream();
                        //创建一个文件输出流
                        FileOutputStream out = new FileOutputStream(savePath + "\\" + filename);
                        //创建一个缓冲区
                        byte buffer[] = new byte[1024];
                        //判断输入流中的数据是否已经读完的标识
                        int len = 0;
                        //循环将输入流读入到缓冲区当中，(len=in.read(buffer))>0就表示in里面还有数据
                        while((len=in.read(buffer))>0){
                            //使用FileOutputStream输出流将缓冲区的数据写入到指定的目录(savePath + "\\" + filename)当中
                            out.write(buffer, 0, len);
                        }
                        //关闭输入流
                        in.close();
                        //关闭输出流
                        out.close();
                        //删除处理文件上传时生成的临时文件
                        item.delete();
                     
                    }
                }
            }catch (Exception e) {
              
                e.printStackTrace();
                
            }
            System.out.println("可读取的目录：" + savePath + "\\" + filename);
			String result = result(savePath + "\\" +filename,excelnum);
			
			HttpSession session = req.getSession();
			session.setAttribute("data", result);
			resp.setContentType("text/html;charset=utf-8");
			resp.sendRedirect("index.jsp");
			
			
			
		} catch (Exception e) {
			HttpSession session = req.getSession();
			String s = "<table tyle=\"width:1500px\" id=\"tblMain\"><tr><td>出错，可能文件无法解析!</td></tr><table>";
			session.setAttribute("data", s);
			resp.setContentType("text/html;charset=utf-8");
			resp.sendRedirect("index.jsp");
			
		}
		
		
		
		
	}
	
	
	public String result(String path,String excelnum) throws Exception{
		Date parse;
		
		//生成File实例并指向需要读取的Excel表文件  
        File file = null;  
        file = new File(path);  
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
		//获取本月上班日集合和非工作日集合
        List<Object> workDayList = new ArrayList<>();
        List<Object> holidayList = new ArrayList<>();
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
		        if (result == 1 || result ==2) {
		        	holidayList.add(j);//此功能暂时还未做
				}
		 }
	        	
		 System.out.println("上班日：" + workDayList.toString());
		 
		
		
		int count = 0;//全天没打
		int sum = 0; //上午或者下午没打
		int total = 0; //超过5分钟统计
		int late = 0;
	
		Date timeMor = timeFormate.parse("09:01");
        Date time06 = timeFormate.parse("09:06");
        Date timeAft = timeFormate.parse("17:59");
        Date noon = timeFormate.parse("12:59");
        
        StringBuilder sb = new StringBuilder();
		sb.append("<table style=\"width:1500px\" id=\"tblMain\"><tr><th style=\"width:10%\">工号</th><th style=\"width:10%\">姓名</th><th>详情</th></tr>");
		//获取每个人该月的所有打卡信息
		Integer exnum = Integer.valueOf(excelnum);
		for (int i = 5; i <= exnum; i++) {
			StringBuilder sbtime = new StringBuilder();
        	//工号&员工姓名  
        	String jobno = ExcelReader.readExcelData(sheet, i, 3);
        	String staffname =  ExcelReader.readExcelData(sheet, i, 11);
        	i++;
       
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
						
					}else { //至少打了两次的 
						//两次之 迟到或早退
						
						if ((compareDate(firstTime, time06)&&compareDate(noon, firstTime) ) ||  (compareDate(timeAft, lastTime)&&compareDate(lastTime, noon))    ) {
							late++;
							sbtime.append("&nbsp;"+j+"号&nbsp;");
						}
						//两次之都是上午或者都是下午
						
						if (compareDate(noon, lastTime) || compareDate(firstTime, noon)) {
							sum++;
						}
					}
					
        		
				}else { //没打
					count ++ ;
				}
			}
	        
	       
	        sb.append("<tr ><td align=\"center\" >" + jobno + "</td><td align=\"center\" >" + staffname +"</td>");
	        
	        if (count != 0 || (sum != 0 && sum > 3) || late != 0) {
				sb.append("<td>");
				 if (count!=0) {
					sb.append( count + "次全天没打；" );
				}
		        
		        if (sum != 0 && sum > 3) {
					sb.append( (sum-3) + "次忘打卡(上午或者下午)；");
				}
		        
		        if (late != 0) {
					sb.append(late + "次迟到或早退【"+sbtime+"】");
				} 
		        sb.append("</td>");
			}
	        
	       
	        
	        if (count ==0 && sum <=3 && late == 0) {
	        	sb.append("<td><font color=\"#00E5EE\">好同志，没任何不良记录!</font></td>");
			}
	        
	        //System.out.println("工号：" + jobno + "	姓名：" + staffname + "  " + count + "次全天没打；" + sum +"次没打卡 " + ";  迟到或早退：" + late );
	        System.out.println(sb.toString());
	        
	        sb.append("</tr>");
	       
	        count = 0;
	        sum = 0;
	        late = 0;
	        total = 0;
	     
		}
		
		sb.append("<table>");
		return sb.toString();
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
     *            :参数  格式：20170401
     * @return 返回结果
     * 		   0 上班 1周末 2节假日
     */
    public static String request( String httpArg) {
        String httpUrl="http://tool.bitefu.net/jiari/";
    	//String httpUrl="http://www.easybots.cn/api/holiday.php";//日期接口
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
                //sbf.append("\r\n");
            }
            reader.close();
            result = sbf.toString();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }
	
	

}
