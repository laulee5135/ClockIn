<%@ page language="java" contentType="text/html; charset=UTF-8"
    pageEncoding="UTF-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<style type="text/css">

table {

border-right: 1px solid #8B8B7A;
border-left: 1px solid #8B8B7A;
border-bottom: 1px solid #8B8B7A;
border-top: 1px solid #8B8B7A;
border-collapse:collapse;

}

table td {

border-left: 1px solid #8B8B7A;

border-top: 1px solid #8B8B7A;

border-bottom: 1px solid #8B8B7A;
border-right: 1px solid #8B8B7A;

}
table th {

border-left: 1px solid #8B8B7A;

border-top: 1px solid #8B8B7A;

border-bottom: 1px solid #8B8B7A;
border-right: 1px solid #8B8B7A;

}
</style>
<script type="text/javascript">  
        function check(){  
            var val = window.document.getElementById("excel").value;  
            if (!val) 
            {  
                window.alert("必须选一个考勤Excel文件!");  
                return false;  
            }else{
            	var suffix = val.substring(val.length-3,val.length);
            	//alert(suffix);
            	if(suffix!='xls'){
            		alert("只能处理xls格式文件，可在源文件的基础上右键另存为.xls的文件！");
            		return false;
            	}
            	
            }  
            
          
            return true;  
        }  
        
        function SetTableColor() {
        	  var tbl = document.getElementById("tblMain");
        	  var trs = tbl.getElementsByTagName("tr");
        	  for (var i = 0; i < trs.length; i++) {
        	 var j = i + 1;
        	 if (j % 2 == 0) { //偶数行
        	   trs[i].style.background = "#EEE9BF";
        	 }
        	/*  else {
        	   trs[i].style.background = "blue";
        	 } */
        	  }
        }
</script>  
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title>Punch statistics</title>
</head>
<body onload="SetTableColor()">
	<center><h2 style="color:#CDC673"> 工时统计表</h2><center>
	<form  action="${pageContext.request.contextPath}/servlet" method="post"  onsubmit="return check()" enctype="multipart/form-data">
		<table style="width:1500px" >
			<tr>
				<td style="text-align:center; ">Excel行数：<input type="text" value="348" name="excelnum"/></td>
				<td style="text-align:center; "><input type="file" name="file" id="excel"/><input type="submit" value="查询"/></td>
				
			</tr>
			<tr >
				<td colspan="2" > <strong>结果:</strong><br></td>
				${sessionScope.data}
				
			</tr>
			
		</table>
		 
	
	</form>
	<br>
	<br>
	
	<%
        session.invalidate();	
	%> 
	
	
	
</body>
</html>