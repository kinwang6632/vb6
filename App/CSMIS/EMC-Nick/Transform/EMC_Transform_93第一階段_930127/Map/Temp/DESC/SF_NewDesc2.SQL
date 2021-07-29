/*
  字串型態代碼之名稱有異動時，更新相關檔案內容: SF_NewDesc2()
  檔名: SF_NewDesc2.sql
 @c:\gird\V400\script\SF_NewDesc2.sql;  
  IN p_TableName VARCHAR2：代碼檔名稱
     p_CodeNo VARCHAR2：代碼
     p_OldDesc VARCHAR2：舊名稱
     p_NewDesc VARCHAR2：新名稱
  OUT p_RetMsg VARCHAR2：結果訊息
  Return NUMBER：處理結果代碼，其對應訊息存放於p_RetMsg中
           0：p_RetMsg='執行完畢'
          -1：p_RetMsg='參數錯誤'
         -99：p_RetMsg='其他錯誤'

  Date: 2000/08/17 
  By: Lawrence

  Edit: Lawrence by 2000.08.18
                               2002.09.27 加SO017.Name
                               2003.07.23 加CM003

  Stanley  2001.08.29 where條件加入 "and 名稱欄位 <> 傳來之名稱參數"
     
*/

CREATE OR REPLACE FUNCTION SF_NewDesc2
  (p_TableName VARCHAR2, p_CodeNo VARCHAR2, p_OldDesc VARCHAR2, p_NewDesc VARCHAR2, 
  p_RetMsg OUT VARCHAR2) RETURN Number

  as

begin
  -- 檢查參數
  if p_TableName is null or p_CodeNo is null or p_OldDesc is null or p_NewDesc is null then
    p_RetMsg := '參數錯誤';
    return -1;
  end if;

/*
  dbms_output.put_line('A1');
  dbms_output.put_line('# of rcd updated = ' || to_char(SQL%ROWCOUNT));
  dbms_output.put_line('A2');
*/

  /* 
    行政區代碼檔 CD001: 待改檔: SO014, SO016, SO019
    服務區代碼檔 CD002: 待改檔: SO001, SO014, SO016, SO019
    工程組代碼檔 CD003: 待改檔: SO007, SO008, SO009
    頻道代碼檔   CD024: 待改檔: SOAC0102
    收費區代碼檔 CD040: 待改檔: SO001, SO014, SO016
    業務區檔 CD050: 待改檔: SO007, SO008, SO009, SO014, SO016
    人事資料小檔 CM003: 待改檔: CD002CM003,CD003,SO004,SO005,SO006,SO007,SO008,SO009,SO014, SO016, SO017, SO018, SO019, SO020, 
                                                         SO022, SO033, SO034, SO035, SO036, SO086, SO096, SO100,SO121,SO126,SO127,SO132
    
  */
  if p_TableName = 'CD001' then
    update SO014 set AreaName=p_NewDesc where AreaCode=p_CodeNo and AreaName<>p_NewDesc;
    update SO016 set AreaName=p_NewDesc where AreaCode=p_CodeNo and AreaName<>p_NewDesc;
    update SO019 set AreaName=p_NewDesc where AreaCode=p_CodeNo and AreaName<>p_NewDesc;

  elsif p_TableName = 'CD002' then
    update SO001 set ServArea=p_NewDesc where ServCode=p_CodeNo and ServArea<>p_NewDesc;
    update SO014 set ServName=p_NewDesc where ServCode=p_CodeNo and ServName<>p_NewDesc;
    update SO016 set ServName=p_NewDesc where ServCode=p_CodeNo and ServName<>p_NewDesc;
    update SO019 set AreaName=p_NewDesc where AreaCode=p_CodeNo and AreaName<>p_NewDesc;

  elsif p_TableName = 'CD003' then
    update SO007 set GroupName=p_NewDesc where GroupCode=p_CodeNo and GroupName<>p_NewDesc;
    update SO008 set GroupName=p_NewDesc where GroupCode=p_CodeNo and GroupName<>p_NewDesc;
    update SO009 set GroupName=p_NewDesc where GroupCode=p_CodeNo and GroupName<>p_NewDesc;

  elsif p_TableName = 'CD024' then
    update SOAC0102 set ChName=p_NewDesc where ChCode=p_CodeNo and ChName<>p_NewDesc;

  elsif p_TableName = 'CD040' then
    update SO001 set ClctAreaName=p_NewDesc where ClctAreaCode=p_CodeNo and ClctAreaName<>p_NewDesc;
    update SO014 set ClctAreaName=p_NewDesc where ClctAreaCode=p_CodeNo and ClctAreaName<>p_NewDesc;
    update SO016 set ClctAreaName=p_NewDesc where ClctAreaCode=p_CodeNo and ClctAreaName<>p_NewDesc;

--  stanely 90/08/29 增加  業務區檔 CD050: 待改檔: SO007, SO008, SO009, SO014, SO016
  elsif p_TableName = 'CD050' then
    update SO007 set SalesName=p_NewDesc where SalesCode=p_CodeNo and SalesName<>p_NewDesc;    
    update SO008 set SalesName=p_NewDesc where SalesCode=p_CodeNo and SalesName<>p_NewDesc;    
    update SO009 set SalesName=p_NewDesc where SalesCode=p_CodeNo and SalesName<>p_NewDesc;    
    update SO014 set SalesName=p_NewDesc where SalesCode=p_CodeNo and SalesName<>p_NewDesc;    
    update SO016 set SalesName=p_NewDesc where SalesCode=p_CodeNo and SalesName<>p_NewDesc; 

-- Lawrence 2002.09.27
  elsif p_TableName = 'SO017' then
    update SO014 set MDUName=p_NewDesc where MDUID=p_CodeNo and MDUName<>p_NewDesc;
    update SO016 set MDUName=p_NewDesc where MDUID=p_CodeNo and MDUName<>p_NewDesc;  
/*
-- Lawrence 2003.07.23
  elsif p_TableName = 'CM003' then
    update CD002CM003 set EmpName=p_NewDesc where EmpNo=p_CodeNo and EmpName<>p_NewDesc;
    update SO121 set EmpName=p_NewDesc where EmpNo=p_CodeNo and EmpName<>p_NewDesc;          
    update SO126 set EmpName=p_NewDesc where EmpNo=p_CodeNo and EmpName<>p_NewDesc;      
    update SO127 set EmpName=p_NewDesc where EmpNo=p_CodeNo and EmpName<>p_NewDesc;      
    update SO132 set EmpName=p_NewDesc where EmpNo=p_CodeNo and EmpName<>p_NewDesc;      
    update CD003 set ServName1=p_NewDesc where ServId1=p_CodeNo and ServName1<>p_NewDesc;      
    update CD003 set ServName2=p_NewDesc where ServId2=p_CodeNo and ServName2<>p_NewDesc;      
    update CD003 set ServName3=p_NewDesc where ServId3=p_CodeNo and ServName3<>p_NewDesc;     
    update SO004 set InstName1=p_NewDesc where InstEn1=p_CodeNo and InstName1<>p_NewDesc;      
    update SO004 set InstName2=p_NewDesc where InstEn2=p_CodeNo and InstName2<>p_NewDesc;  
    update SO004 set PRName1=p_NewDesc where PREn1=p_CodeNo and PRName1<>p_NewDesc;  
    update SO004 set PRName2=p_NewDesc where PREn2=p_CodeNo and PRName2<>p_NewDesc;  
    update SO004 set IntroName=p_NewDesc where IntroId=p_CodeNo and IntroName<>p_NewDesc;  
    update SO005 set SetName=p_NewDesc where SetEn=p_CodeNo and SetName<>p_NewDesc;      
    update SO006 set AcceptName=p_NewDesc where AcceptEn=p_CodeNo and AcceptName<>p_NewDesc;      
    update SO006 set HandleName=p_NewDesc where HandleEn=p_CodeNo and HandleName<>p_NewDesc;   
    update SO007 set WorkerName1=p_NewDesc where WorkerEn1=p_CodeNo and WorkerName1<>p_NewDesc; 
    update SO007 set WorkerName2=p_NewDesc where WorkerEn2=p_CodeNo and WorkerName2<>p_NewDesc; 
    update SO007 set SignName=p_NewDesc where SignEn=p_CodeNo and SignName<>p_NewDesc; 
    update SO007 set AcceptName=p_NewDesc where AcceptEn=p_CodeNo and AcceptName<>p_NewDesc; 
    update SO007 set IntroName=p_NewDesc where IntroId=p_CodeNo and IntroName<>p_NewDesc; 
    update SO008 set WorkerName1=p_NewDesc where WorkerEn1=p_CodeNo and WorkerName1<>p_NewDesc; 
    update SO008 set WorkerName2=p_NewDesc where WorkerEn2=p_CodeNo and WorkerName2<>p_NewDesc; 
    update SO008 set SignName=p_NewDesc where SignEn=p_CodeNo and SignName<>p_NewDesc; 
    update SO008 set AcceptName=p_NewDesc where AcceptEn=p_CodeNo and AcceptName<>p_NewDesc; 
    update SO009 set WorkerName1=p_NewDesc where WorkerEn1=p_CodeNo and WorkerName1<>p_NewDesc; 
    update SO009 set WorkerName2=p_NewDesc where WorkerEn2=p_CodeNo and WorkerName2<>p_NewDesc; 
    update SO009 set SignName=p_NewDesc where SignEn=p_CodeNo and SignName<>p_NewDesc; 
    update SO009 set AcceptName=p_NewDesc where AcceptEn=p_CodeNo and AcceptName<>p_NewDesc; 
    update SO014 set ClctName=p_NewDesc where ClctEn=p_CodeNo and ClctName<>p_NewDesc; 
    update SO016 set ClctName=p_NewDesc where ClctEn=p_CodeNo and ClctName<>p_NewDesc; 
    update SO017 set DevName=p_NewDesc where DevEn=p_CodeNo and DevName<>p_NewDesc; 
    update SO017 set CSignName=p_NewDesc where CSignEn=p_CodeNo and CSignName<>p_NewDesc; 
    update SO018 set CSignName=p_NewDesc where CSignEn=p_CodeNo and CSignName<>p_NewDesc; 
    update SO019 set InstName=p_NewDesc where InstEn=p_CodeNo and InstName<>p_NewDesc; 
    update SO019 set PRName=p_NewDesc where PREn=p_CodeNo and PRName<>p_NewDesc; 
    update SO020 set HandleName=p_NewDesc where HandleEn=p_CodeNo and HandleName<>p_NewDesc; 
    update SO022 set AcceptName=p_NewDesc where AcceptEn=p_CodeNo and AcceptName<>p_NewDesc; 
    update SO033 set ClctName=p_NewDesc where ClctEn=p_CodeNo and ClctName<>p_NewDesc; 
    update SO033 set OldClctName=p_NewDesc where OldClctEn=p_CodeNo and OldClctName<>p_NewDesc; 
    update SO034 set ClctName=p_NewDesc where ClctEn=p_CodeNo and ClctName<>p_NewDesc; 
    update SO034 set OldClctName=p_NewDesc where OldClctEn=p_CodeNo and OldClctName<>p_NewDesc; 
    update SO035 set ClctName=p_NewDesc where ClctEn=p_CodeNo and ClctName<>p_NewDesc; 
    update SO035 set OldClctName=p_NewDesc where OldClctEn=p_CodeNo and OldClctName<>p_NewDesc; 
    update SO036 set ClctName=p_NewDesc where ClctEn=p_CodeNo and ClctName<>p_NewDesc; 
    update SO036 set OldClctName=p_NewDesc where OldClctEn=p_CodeNo and OldClctName<>p_NewDesc; 
    update SO086 set WorkerName1=p_NewDesc where WorkerEn1=p_CodeNo and WorkerName1<>p_NewDesc; 
    update SO096 set AcceptName=p_NewDesc where AcceptEn=p_CodeNo and AcceptName<>p_NewDesc; 
    update SO100 set SalesName=p_NewDesc where SalesEn=p_CodeNo and SalesName<>p_NewDesc;
*/ 
  end if;


--  commit;
  p_RetMsg := '執行完畢';
  return 0;

exception
  when others then
    rollback;
    p_RetMsg := '其他錯誤';
    return -99;
end;
/