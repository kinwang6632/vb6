<script language="jscript"" >
function delete_confirm(e){
	if (event.srcElement.type=="submit" && document.all(event.srcElement.name).value =="刪除" )
	   event.returnValue =confirm("確認刪除否?");
}
document.onclick=delete_confirm;
</script>

<script language="C#" runat=server>
const string MSG_REGULAR_COLOR = "blue";
const string MSG_ERROR_COLOR = "red";
const string MSG_SQL_INS_DONE = 
"<b>新增資料:</b><br>資料已成功的加入資料庫";
const string MSG_SQL_DEL_DONE = "<b>刪除資料:</b><br>資料已成功的由資料庫中刪除" ;
const string MSG_SQL_UPT_DONE = "<b>資料修改:</b><br>資料已成功的在資料庫中修改";
const string MSG_SQL_ERROR_DUPLICATED_KEY = "ERROR: 代碼重複, 請重新輸入"; 
const string MSG_SQL_ERROR_DATA_LOCK = "ERROR: 此筆資料已被其他資料使用, 不可刪除";
const string MSG_SQL_ERROR_OTHER_PREFIX = "ERROR: ";


private void ErrorHandler(string strError){
 try{
	Message.Style["color"] = "red";
	Message.InnerHtml = "ERROR:" + strError;
	Message.InnerHtml +="<BR><a href=\"/error_report.aspx?error_description="+Server.UrlEncode(strError)+"\">請跟站長講一聲!!</a>";
   }
   catch (System.Exception e){
        ErrorHandler(e.ToString());
	return ;
   }
}

private  void SQLErrorHandler(System.Data.SqlClient.SqlException E){
 try{
	Message.Style["color"] = MSG_ERROR_COLOR ;
	switch (E.Number) {
	case 2627:
	   Message.InnerHtml = MSG_SQL_ERROR_DUPLICATED_KEY;
	   break;
	case 547:
   	   Message.InnerHtml = MSG_SQL_ERROR_DATA_LOCK;	
	   break;
	default:
	   Message.InnerHtml = MSG_SQL_ERROR_OTHER_PREFIX +"[" +E.Number.ToString()+"]"+ E.ToString();	   
           break;
	}
	
   }
   catch (System.Exception e){
        ErrorHandler(e.ToString());
	return ;
   }

}

private void MessageSQLInsDone() {
   try{
	Message.Style["color"] = MSG_REGULAR_COLOR;
	Message.InnerHtml = MSG_SQL_INS_DONE;
   }
   catch (System.Exception e){
        ErrorHandler(e.ToString());
   }
}

private void MessageSQLUptDone() {
   try{
	Message.Style["color"] = MSG_REGULAR_COLOR;
	Message.InnerHtml = MSG_SQL_UPT_DONE;
   }
   catch (System.Exception e){
        ErrorHandler(e.ToString());
   }
}

private void MessageSQLDelDone() {
   try{
	Message.Style["color"] = MSG_REGULAR_COLOR;
	Message.InnerHtml = MSG_SQL_DEL_DONE;
   }
   catch (System.Exception e){
        ErrorHandler(e.ToString());
   }
}

private string GetClassIntervalName(string strClassInterval){
   try{
	string strOutput="每週";
	char[] class_interval_array=strClassInterval.ToCharArray();
	for (int i=0; i < 7; i++) {
	   if (class_interval_array[i].ToString()== "1") {
	      if(i==6)
	         strOutput+=GetDayOfWeekName(0);                  
	      else	
	         strOutput+=GetDayOfWeekName(i+1);                  
	   }    
	}
	return strOutput;
   }
   catch (System.Exception e){
        ErrorHandler(e.ToString());
	return "";
   }
}

private string GetDayOfWeekName(int i){
	string  DayOfWeekName="";
	switch (i) {
	case 0:
	     DayOfWeekName= "日";
	     break;
        case 1:
             DayOfWeekName="一";
	     break;	
        case 2:
            DayOfWeekName= "二";
	    break;
        case 3:
            DayOfWeekName= "三";
	    break;
        case 4:
            DayOfWeekName= "四";
	    break;		
        case 5:
            DayOfWeekName= "五";
	    break;
        case 6:
            DayOfWeekName= "六";
	    break;	
        }
        return DayOfWeekName;
}
</script>