/*
@c:\gird\V300\script\SF_NewDesc1.sql;
variable nn number;
variable msg varchar2(200);
exec :nn := SF_NewDesc1('CD017',1,'中壢市中園路','中園路',:msg)
print nn;
print msg;


  數字型態代碼之名稱有異動時，更新相關檔案內容: SF_NewDesc1()
  檔名: SF_NewDesc1.sql
  
  IN p_TableName VARCHAR2：代碼檔名稱
     p_CodeNo NUMBER：代碼
     p_OldDesc VARCHAR2：舊名稱
     p_NewDesc VARCHAR2：新名稱
  OUT p_RetMsg VARCHAR2：結果訊息
  Return NUMBER：處理結果代碼，其對應訊息存放於p_RetMsg中
           0：p_RetMsg='執行完畢'
          -1：p_RetMsg='參數錯誤'
         -99：p_RetMsg='其他錯誤'

  Date: 2000.08.16 
  By: Lawrence

  Edit: Lawrence by 2000.08.18
  2002.06.20 調: for New schema
  2003.12.22 加入新增的代碼檔
  
  Stanley  2001.08.29 where條件加入 "and 名稱欄位 <> 傳來之名稱參數"
           2002.11.05 加入新增的代碼檔
*/

CREATE OR REPLACE FUNCTION SF_NewDesc1
  (p_TableName VARCHAR2, p_CodeNo NUMBER, p_OldDesc VARCHAR2, p_NewDesc VARCHAR2, 
  p_RetMsg OUT VARCHAR2) RETURN Number

  as
v_return number;

begin
  -- 檢查參數
  if p_TableName is null or nvl(p_CodeNo,0)=0 or p_OldDesc is null or p_NewDesc is null then
    p_RetMsg := '參數錯誤';
    return -1;
  end if;

/*
  dbms_output.put_line('A1');
  dbms_output.put_line('# of rcd updated = ' || to_char(SQL%ROWCOUNT));
  dbms_output.put_line('A2');
*/

  /* 
    客戶類別代碼檔 CD004: 待改檔: SO001, SO019
    裝機類別代碼檔 CD005: 待改檔: SO007
    維修申告類別代碼檔 CD006: 待改檔: SO008
    停拆移機類別代碼檔 CD007: 待改檔: SO002, SO009
    電話申告類別代碼檔 CD008: 待改檔: SO006
    媒體介紹類別代碼檔 CD009: 待改檔: SO002, SO004, --SO006, SO007, --SO098
    銷售點行業類別代碼檔 CD010: 待改檔: SO013
    故障原因代碼檔 CD011: 待改檔: SO008, SO022
    註銷原因代碼檔 CD012: 待改檔: SO002
    未繳費原因代碼檔 CD013: 待改檔: SO032, SO033, SO034, SO036
    停拆機原因代碼檔 CD014: 待改檔: SO009
    退單原因代碼檔 CD015: 待改檔: SO007, SO008, SO009
    短收原因代碼檔 CD016: 待改檔: SO032, SO033, SO034, SO036
    街道編號代碼檔 CD017: 待改檔: SO014, SO016
    銀行代碼檔 CD018: 待改檔: SO002
    收費項目代碼檔 CD019: 待改檔: CD005, CD006, CD007, CD022, CD024, SO003, SO032, SO033, SO034, SO036,--SOAC0102
    付費意願代碼檔 CD020: 待改檔: SO001
    建築型態代碼檔 CD021: 待改檔: SO014, SO016
    品名編號代碼檔 CD022: 待改檔: CD005, CD006, CD007, SO004, SO019
    客戶來源代碼檔 CD025: 待改檔: SO001
    服務滿意度代碼檔 CD026: 待改檔: SO007, SO008, SO009
    收費方式代碼檔 CD031: 待改檔: SO002, SO032, SO033, SO034, SO036
    付款種類代碼檔 CD032: 待改檔: SO032, SO033, SO034, SO036
    稅率代碼檔 CD033: 待改檔: CD019
    買賣方式代碼檔 CD034: 待改檔: CD022, SO004
    客戶狀態代碼檔 CD035: 待改檔: SO002 
    派工狀態代碼檔 CD036: 待改檔: SO002
    公司別代碼檔 CD039: 待改檔: SO001, SO014, SO016--, SO062, SO083
    供電戶費用公式代碼檔 CD041: 待改檔: SO020
    業務活動代碼檔 CD042: 待改檔: SO006, SO007, SO095, SO096, SO098
    設備型號代碼檔 CD043: 待改檔: SO004
    管派類別代碼檔 CD044: 待改檔: CD005, CD006, CD007, SO083
    優待辦法代碼檔 CD048: 待改檔: CD042
    消息來源檔 CD049: 待改檔: --SO006, --SO007, SO095, SO096--,SO098
    作廢原因代碼檔 CD051: 待改檔: SO032, SO033, SO034, SO036

    套餐代碼檔 CD027: 待改檔: SO003, SO032, SO033, SO034, SO035, SO036, SO067, SO105A, SO105B, SO105C, SO105D
    STB配件代碼檔 CD028: Relation用 
    節目等級代碼檔 CD029: 待改檔: CD024
    權限組別名稱代碼檔 CD030: Relation用 

    切換資料區 CD052: 單獨使用
    營運預算分類代碼檔 CD053 單獨使用
    營運預算項目代碼檔 CD054 單獨使用
    管線類別代碼檔 CD055: 待改檔: SO014, SO017
    設備裝置點代碼檔 CD056 Relation用 
    訂單類別代碼檔 CD057 Relation用 
    村里代碼檔 CD058: 待改檔: SO014
    商贈品代碼檔 CD059 待改檔: CD027B, SO105C
    頻道結清原因代碼檔 CD060 待改檔: SO108
    E-Box結清原因代碼檔 CD061 Relation用 
    BISP優先順序代碼檔 CD062 Relation用 
    密碼提示代碼檔 CD063 待改檔: SO004
    CM速率代碼檔CD064 待改檔: CD005, CD057A, SO007
  */

  if p_TableName = 'CD004' then
    update SO001 set ClassName1=p_NewDesc where ClassCode1=p_CodeNo and ClassName1<>p_NewDesc;
    update SO019 set ClassName=p_NewDesc where ClassCode=p_CodeNo and ClassName<>p_NewDesc;

  elsif p_TableName = 'CD005' then
    update SO007 set InstName=p_NewDesc where InstCode=p_CodeNo and InstName<>p_NewDesc;

  elsif p_TableName = 'CD006' then
    update SO008 set ServiceName=p_NewDesc where ServiceCode=p_CodeNo and ServiceName<>p_NewDesc;

--  stanely 90/08/29 增加  停拆移機類別代碼檔 CD007: 待改檔: SO002, SO009
  elsif p_TableName = 'CD007' then
    update SO002 set PRName=p_NewDesc where PRCode=p_CodeNo and PRName<>p_NewDesc;
    update SO009 set PRName=p_NewDesc where PRCode=p_CodeNo and PRName<>p_NewDesc;

  elsif p_TableName = 'CD008' then
    update SO006 set ServiceName=p_NewDesc where ServiceCode=p_CodeNo and ServiceName<>p_NewDesc;

--  stanely 90/08/29 增加  媒體介紹類別代碼檔 CD009: 待改檔: SO098
  elsif p_TableName = 'CD009' then
    update SO002 set MediaName=p_NewDesc where MediaCode=p_CodeNo and MediaName<>p_NewDesc;
    update SO004 set MediaName=p_NewDesc where MediaCode=p_CodeNo and MediaName<>p_NewDesc;
    update SO006 set MediaName=p_NewDesc where MediaCode=p_CodeNo and MediaName<>p_NewDesc;
    update SO007 set MediaName=p_NewDesc where MediaCode=p_CodeNo and MediaName<>p_NewDesc;
    update SO098 set MediaName=p_NewDesc where MediaCode=p_CodeNo and MediaName<>p_NewDesc;

  elsif p_TableName = 'CD010' then
    update SO013 set ProfName=p_NewDesc where ProfCode=p_CodeNo and ProfName<>p_NewDesc;

  elsif p_TableName = 'CD011' then
    update SO008 set MFName1=p_NewDesc where MFCode1=p_CodeNo and MFName1<>p_NewDesc;
    update SO008 set MFName2=p_NewDesc where MFCode2=p_CodeNo and MFName2<>p_NewDesc;

  elsif p_TableName = 'CD012' then
    update SO002 set DelName=p_NewDesc where DelCode=p_CodeNo and DelName<>p_NewDesc;

  elsif p_TableName = 'CD013' then
    update SO032 set UCName=p_NewDesc where UCCode=p_CodeNo and UCName<>p_NewDesc;
    update SO033 set UCName=p_NewDesc where UCCode=p_CodeNo and UCName<>p_NewDesc;
    update SO034 set UCName=p_NewDesc where UCCode=p_CodeNo and UCName<>p_NewDesc;
    update SO036 set UCName=p_NewDesc where UCCode=p_CodeNo and UCName<>p_NewDesc;

  elsif p_TableName = 'CD014' then
    update SO009 set ReasonName=p_NewDesc where ReasonCode=p_CodeNo and ReasonName<>p_NewDesc;

  elsif p_TableName = 'CD015' then
    update SO007 set ReturnName=p_NewDesc where ReturnCode=p_CodeNo and ReturnName<>p_NewDesc;
    update SO008 set ReturnName=p_NewDesc where ReturnCode=p_CodeNo and ReturnName<>p_NewDesc;
    update SO009 set ReturnName=p_NewDesc where ReturnCode=p_CodeNo and ReturnName<>p_NewDesc;

  elsif p_TableName = 'CD016' then
    update SO032 set STName=p_NewDesc where STCode=p_CodeNo and STName<>p_NewDesc;
    update SO033 set STName=p_NewDesc where STCode=p_CodeNo and STName<>p_NewDesc;
    update SO034 set STName=p_NewDesc where STCode=p_CodeNo and STName<>p_NewDesc;
    update SO036 set STName=p_NewDesc where STCode=p_CodeNo and STName<>p_NewDesc;

  elsif p_TableName = 'CD017' then
--    update SO014 set StrtName=p_NewDesc where StrtCode=p_CodeNo and StrtName<>p_NewDesc;
--    update SO016 set StrtName=p_NewDesc where StrtCode=p_CodeNo and StrtName<>p_NewDesc;
    v_return := SF_Adjstrt(p_CodeNo,p_OldDesc,p_NewDesc,p_RetMsg);
  elsif p_TableName = 'CD018' then
    update SO002 set BankName=p_NewDesc where BankCode=p_CodeNo and BankName<>p_NewDesc;

--  stanely 90/08/29 增加  收費項目代碼檔 CD019: 待改檔: SOAC0102
  elsif p_TableName = 'CD019' then
    update CD005 set CitemName1=p_NewDesc where CitemCode1=p_CodeNo and CitemName1<>p_NewDesc;
    update CD005 set CitemName2=p_NewDesc where CitemCode2=p_CodeNo and CitemName2<>p_NewDesc;
    update CD005 set CitemName3=p_NewDesc where CitemCode3=p_CodeNo and CitemName3<>p_NewDesc;
    update CD005 set CitemName4=p_NewDesc where CitemCode4=p_CodeNo and CitemName4<>p_NewDesc;
    update CD005 set CitemName5=p_NewDesc where CitemCode5=p_CodeNo and CitemName5<>p_NewDesc;
    update CD006 set CitemName1=p_NewDesc where CitemCode1=p_CodeNo and CitemName1<>p_NewDesc;
    update CD006 set CitemName2=p_NewDesc where CitemCode2=p_CodeNo and CitemName2<>p_NewDesc;
    update CD006 set CitemName3=p_NewDesc where CitemCode3=p_CodeNo and CitemName3<>p_NewDesc;
    update CD006 set CitemName4=p_NewDesc where CitemCode4=p_CodeNo and CitemName4<>p_NewDesc;
    update CD006 set CitemName5=p_NewDesc where CitemCode5=p_CodeNo and CitemName5<>p_NewDesc;
    update CD007 set CitemName1=p_NewDesc where CitemCode1=p_CodeNo and CitemName1<>p_NewDesc;
    update CD007 set CitemName2=p_NewDesc where CitemCode2=p_CodeNo and CitemName2<>p_NewDesc;
    update CD007 set CitemName3=p_NewDesc where CitemCode3=p_CodeNo and CitemName3<>p_NewDesc;
    update CD007 set CitemName4=p_NewDesc where CitemCode4=p_CodeNo and CitemName4<>p_NewDesc;
    update CD007 set CitemName5=p_NewDesc where CitemCode5=p_CodeNo and CitemName5<>p_NewDesc;
    update CD022 set DefCitemName=p_NewDesc where DefCitemCode=p_CodeNo and DefCitemName<>p_NewDesc;
    update CD024 set CitemName=p_NewDesc where CitemCode=p_CodeNo and CitemName<>p_NewDesc;
    update SO003 set CitemName=p_NewDesc where CitemCode=p_CodeNo and CitemName<>p_NewDesc;
    update SO032 set CitemName=p_NewDesc where CitemCode=p_CodeNo and CitemName<>p_NewDesc;
    update SO033 set CitemName=p_NewDesc where CitemCode=p_CodeNo and CitemName<>p_NewDesc;
    update SO034 set CitemName=p_NewDesc where CitemCode=p_CodeNo and CitemName<>p_NewDesc;
    update SO036 set CitemName=p_NewDesc where CitemCode=p_CodeNo and CitemName<>p_NewDesc;
    update SOAC0102 set CitemName=p_NewDesc where CitemCode=p_CodeNo and CitemName<>p_NewDesc;

  elsif p_TableName = 'CD020' then
    update SO001 set PWName=p_NewDesc where PWCode=p_CodeNo and PWName<>p_NewDesc;

  elsif p_TableName = 'CD021' then
    update SO014 set BTName=p_NewDesc where BTCode=p_CodeNo and BTName<>p_NewDesc;
    update SO016 set BTName=p_NewDesc where BTCode=p_CodeNo and BTName<>p_NewDesc;

  elsif p_TableName = 'CD022' then
    update CD005 set FacilName1=p_NewDesc where FacilCode1=p_CodeNo and FacilName1<>p_NewDesc;
    update CD005 set FacilName2=p_NewDesc where FacilCode2=p_CodeNo and FacilName2<>p_NewDesc;
    update CD005 set FacilName3=p_NewDesc where FacilCode3=p_CodeNo and FacilName3<>p_NewDesc;
    update CD006 set FacilName1=p_NewDesc where FacilCode1=p_CodeNo and FacilName1<>p_NewDesc;
    update CD006 set FacilName2=p_NewDesc where FacilCode2=p_CodeNo and FacilName2<>p_NewDesc;
    update CD006 set FacilName3=p_NewDesc where FacilCode3=p_CodeNo and FacilName3<>p_NewDesc;
    update CD007 set FacilName1=p_NewDesc where FacilCode1=p_CodeNo and FacilName1<>p_NewDesc;
    update CD007 set FacilName2=p_NewDesc where FacilCode2=p_CodeNo and FacilName2<>p_NewDesc;
    update CD007 set FacilName3=p_NewDesc where FacilCode3=p_CodeNo and FacilName3<>p_NewDesc;
    update SO004 set FaciName=p_NewDesc where FaciCode=p_CodeNo and FaciName<>p_NewDesc;
    update SO019 set FaciName1=p_NewDesc where FaciCode1=p_CodeNo and FaciName1<>p_NewDesc;
    update SO019 set FaciName2=p_NewDesc where FaciCode2=p_CodeNo and FaciName2<>p_NewDesc;

  elsif p_TableName = 'CD025' then
    update SO001 set OrgName=p_NewDesc where OrgCode=p_CodeNo and OrgName<>p_NewDesc;

  elsif p_TableName = 'CD026' then
    update SO007 set SatiName=p_NewDesc where SatiCode=p_CodeNo and SatiName<>p_NewDesc;
    update SO008 set SatiName=p_NewDesc where SatiCode=p_CodeNo and SatiName<>p_NewDesc;
    update SO009 set SatiName=p_NewDesc where SatiCode=p_CodeNo and SatiName<>p_NewDesc;

-- Stanely 91/11/05 增加  套餐代碼檔 CD027: 待改檔: SO003, SO032, SO033, SO034, SO035, SO036, SO105A, SO105B, SO105C, SO105D
  elsif p_TableName = 'CD027' then
    update SO003 set PackageName=p_NewDesc where PackageNo=p_CodeNo and PackageName<>p_NewDesc;
    update SO032 set PackageName=p_NewDesc where PackageNo=p_CodeNo and PackageName<>p_NewDesc;
    update SO033 set PackageName=p_NewDesc where PackageNo=p_CodeNo and PackageName<>p_NewDesc;
    update SO034 set PackageName=p_NewDesc where PackageNo=p_CodeNo and PackageName<>p_NewDesc;
    update SO035 set PackageName=p_NewDesc where PackageNo=p_CodeNo and PackageName<>p_NewDesc;
    update SO036 set PackageName=p_NewDesc where PackageNo=p_CodeNo and PackageName<>p_NewDesc;
    update SO105A set PackageName=p_NewDesc where PackageNo=p_CodeNo and PackageName<>p_NewDesc;
    update SO105B set PackageName=p_NewDesc where PackageNo=p_CodeNo and PackageName<>p_NewDesc;
    update SO105C set PackageName=p_NewDesc where PackageNo=p_CodeNo and PackageName<>p_NewDesc;
    update SO105D set PackageName=p_NewDesc where PackageNo=p_CodeNo and PackageName<>p_NewDesc;

-- Stanely 91/11/05 節目等級代碼檔 CD029: 待改檔: CD024
  elsif p_TableName = 'CD029' then
    update CD024 set GradeName=p_NewDesc where GradeCode=p_CodeNo and GradeName<>p_NewDesc;

  elsif p_TableName = 'CD031' then
    update SO002 set CMName=p_NewDesc where CMCode=p_CodeNo and CMName<>p_NewDesc;
    update SO033 set CMName=p_NewDesc where CMCode=p_CodeNo and CMName<>p_NewDesc;
    update SO034 set CMName=p_NewDesc where CMCode=p_CodeNo and CMName<>p_NewDesc;
    update SO036 set CMName=p_NewDesc where CMCode=p_CodeNo and CMName<>p_NewDesc;

  elsif p_TableName = 'CD032' then
    update SO033 set PTName=p_NewDesc where PTCode=p_CodeNo and PTName<>p_NewDesc;
    update SO034 set PTName=p_NewDesc where PTCode=p_CodeNo and PTName<>p_NewDesc;
    update SO036 set PTName=p_NewDesc where PTCode=p_CodeNo and PTName<>p_NewDesc;

--  stanely 90/08/29 增加  稅率代碼檔 CD033: 待改檔: CD019
  elsif p_TableName = 'CD033' then
    update CD019 set TaxName=p_NewDesc where TaxCode=p_CodeNo and TaxName<>p_NewDesc;

  elsif p_TableName = 'CD034' then
    update CD022 Set DefBuyName=p_NewDesc where DefBuyCode=p_CodeNo and DefBuyName<>p_NewDesc;
    update SO004 set BuyName=p_NewDesc where BuyCode=p_CodeNo and BuyName<>p_NewDesc;
  
--  stanely 90/08/29 增加  客戶狀態代碼檔 CD035: 待改檔: SO002 
  elsif p_TableName = 'CD035' then
    update SO002 set CustStatusName=p_NewDesc where CustStatusCode=p_CodeNo and CustStatusName<>p_NewDesc;
   
--  stanely 90/08/29 增加  派工狀態代碼檔 CD036: 待改檔: SO002
  elsif p_TableName = 'CD036' then
    update SO002 set WipName1=p_NewDesc where WipCode1=p_CodeNo and WipName1<>p_NewDesc;
    update SO002 set WipName2=p_NewDesc where WipCode2=p_CodeNo and WipName2<>p_NewDesc;
    update SO002 set WipName3=p_NewDesc where WipCode3=p_CodeNo and WipName3<>p_NewDesc;
  
  elsif p_TableName = 'CD039' then
    update SO001 Set CompName=p_NewDesc where CompCode=p_CodeNo and CompName<>p_NewDesc;
    update SO014 Set CompName=p_NewDesc where CompCode=p_CodeNo and CompName<>p_NewDesc;
    update SO016 Set CompName=p_NewDesc where CompCode=p_CodeNo and CompName<>p_NewDesc;

  elsif p_TableName = 'CD041' then
    update SO020 Set CountName=p_NewDesc where CountCode=p_CodeNo and CountName<>p_NewDesc;

--  stanely 90/08/29 增加  業務活動代碼檔 CD042: 待改檔: SO098
  elsif p_TableName = 'CD042' then
    update SO006 Set PromName=p_NewDesc where PromCode=p_CodeNo and PromName<>p_NewDesc;
    update SO007 Set PromName=p_NewDesc where PromCode=p_CodeNo and PromName<>p_NewDesc;
    update SO095 Set PromName=p_NewDesc where PromCode=p_CodeNo and PromName<>p_NewDesc;
    update SO096 Set PromName=p_NewDesc where PromCode=p_CodeNo and PromName<>p_NewDesc;
    update SO098 Set PromName=p_NewDesc where PromCode=p_CodeNo and PromName<>p_NewDesc;

--  stanely 90/08/29 增加  設備型號代碼檔 CD043: 待改檔: SO004
  elsif p_TableName = 'CD043' then  
    update SO004 Set ModelName=p_NewDesc where ModelCode=p_CodeNo and ModelName<>p_NewDesc;

--  stanely 90/08/29 增加  管派類別代碼檔 CD044: 待改檔: CD005, CD006, CD007, SO083
  elsif p_TableName = 'CD044' then
    update CD005 Set GroupName=p_NewDesc where GroupNo=p_CodeNo and GroupName<>p_NewDesc;
    update CD006 Set GroupName=p_NewDesc where GroupNo=p_CodeNo and GroupName<>p_NewDesc;
    update CD007 Set GroupName=p_NewDesc where GroupNo=p_CodeNo and GroupName<>p_NewDesc;
    update SO083 Set MName=p_NewDesc where MCode=p_CodeNo and MName<>p_NewDesc;

--  stanely 90/08/29 增加  優待辦法代碼檔 CD048: 待改檔: CD042
  elsif p_TableName = 'CD048' then
    update CD042 Set DiscountName=p_NewDesc where DiscountCode=p_CodeNo and DiscountName<>p_NewDesc;

--  stanely 90/08/29 增加  消息來源檔 CD049: 待改檔: SO006, SO007, SO098
  elsif p_TableName = 'CD049' then
    update SO006 set BulletinName=p_NewDesc where BulletinCode=p_CodeNo and BulletinName<>p_NewDesc;
    update SO007 set BulletinName=p_NewDesc where BulletinCode=p_CodeNo and BulletinName<>p_NewDesc;
    update SO095 set BulletinName=p_NewDesc where BulletinCode=p_CodeNo and BulletinName<>p_NewDesc;
    update SO096 set BulletinName=p_NewDesc where BulletinCode=p_CodeNo and BulletinName<>p_NewDesc;
    update SO098 set BulletinName=p_NewDesc where BulletinCode=p_CodeNo and BulletinName<>p_NewDesc;

--  stanely 90/08/29 增加  作廢原因代碼檔 CD051: 待改檔: SO032, SO033, SO034, SO036
  elsif p_TableName = 'CD051' then
    update SO032 set CancelName=p_NewDesc where CancelCode=p_CodeNo and CancelName<>p_NewDesc; 
    update SO033 set CancelName=p_NewDesc where CancelCode=p_CodeNo and CancelName<>p_NewDesc; 
    update SO034 set CancelName=p_NewDesc where CancelCode=p_CodeNo and CancelName<>p_NewDesc; 
    update SO036 set CancelName=p_NewDesc where CancelCode=p_CodeNo and CancelName<>p_NewDesc;

--  Stanely 91/11/05 增加  管線類別代碼檔 CD055: 待改檔: SO014, SO017 
  elsif p_TableName = 'CD055' then
    update SO014 set PipelineName=p_NewDesc where PipelineCode=p_CodeNo and PipelineName<>p_NewDesc;
    update SO017 set PipelineName=p_NewDesc where PipelineCode=p_CodeNo and PipelineName<>p_NewDesc;

--  Stanely 91/11/05 增加  村里代碼檔 CD058: 待改檔: SO014
  elsif p_TableName = 'CD058' then
    update SO014 set BourgName=p_NewDesc where BourgCode=p_CodeNo and BourgName<>p_NewDesc;

--  Stanely 91/11/05 增加  商贈品代碼檔 CD059 待改檔: CD027B, SO105C
  elsif p_TableName = 'CD059' then
    update CD027B set Description=p_NewDesc where CodeNo=p_CodeNo and Description<>p_NewDesc;
    update SO105C set ArticleName=p_NewDesc where ArticleNo=p_CodeNo and ArticleName<>p_NewDesc;

--  Stanely 91/11/05 增加  頻道結清原因代碼檔 CD060 待改檔: SO108
  elsif p_TableName = 'CD060' then
    update SO108 set ChEvenName=p_NewDesc where ChEvenCode=p_CodeNo and ChEvenName<>p_NewDesc;

--  Lawrence 91/11/05 增加  密碼提示代碼檔 CD063 待改檔: SO004
  elsif p_TableName = 'CD063' then
    update SO004 set HintName=p_NewDesc where HintCodeNo=p_CodeNo and HintName<>p_NewDesc;
--  Lawrence 92/12/22 增加  CM速率代碼檔CD064 待改檔: CD005, CD057A, SO007
  elsif p_TableName = 'CD064' then
    update CD005 set CMName=p_NewDesc where CMCode=p_CodeNo and CMName<>p_NewDesc;
    update CD057A set CMName=p_NewDesc where CMCode=p_CodeNo and CMName<>p_NewDesc;
    update SO007 set SCMName=p_NewDesc where SCMCode=p_CodeNo and SCMName<>p_NewDesc;
    update SO007 set CMName=p_NewDesc where CMCode=p_CodeNo and CMName<>p_NewDesc;
--  Lawrence 92/12/22 增加  工期類別代碼檔CD065 待改檔: CD065A
  elsif p_TableName = 'CD065' then
    update CD065A set Description=p_NewDesc where CodeNo=p_CodeNo and Description<>p_NewDesc;

  end if;

  commit;
  p_RetMsg := '執行完畢';
  return 0;

exception
  when others then
    rollback;
    p_RetMsg := '其他錯誤';
    return -99;
end;
/