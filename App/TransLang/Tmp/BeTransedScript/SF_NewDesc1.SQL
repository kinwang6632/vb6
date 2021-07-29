
CREATE OR REPLACE FUNCTION SF_NewDesc1
  (p_TableName VARCHAR2, p_CodeNo NUMBER, p_OldDesc VARCHAR2, p_NewDesc VARCHAR2, 
  p_RetMsg OUT VARCHAR2) RETURN Number

  as

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
    停拆移機類別代碼檔 CD007: 待改檔: SO009
    電話申告類別代碼檔 CD008: 待改檔: SO006
    媒體介紹類別代碼檔 CD009: 待改檔: SO001, SO004, SO007
    銷售點行業類別代碼檔 CD010: 待改檔: SO013
    故障原因代碼檔 CD011: 待改檔: SO008, SO022
    註銷原因代碼檔 CD012: 待改檔: SO001
    未繳費原因代碼檔 CD013: 待改檔: SO032, SO033, SO034, SO036
    停拆機原因代碼檔 CD014: 待改檔: SO009
    退單原因代碼檔 CD015: 待改檔: SO007, SO008, SO009
    短收原因代碼檔 CD016: 待改檔: SO032, SO033, SO034, SO036
    街道編號代碼檔 CD017: 待改檔: SO014, SO016
    銀行代碼檔 CD018: 待改檔: SO002
    收費項目代碼檔 CD019: 待改檔: CD005, CD006, CD007, CD022, CD024, SO003, SO032, SO033, SO034, SO036
    付費意願代碼檔 CD020: 待改檔: SO001
    建築型態代碼檔 CD021: 待改檔: SO014, SO016
    品名編號代碼檔 CD022: 待改檔: CD005, CD006, CD007, SO004, SO019
    客戶來源代碼檔 CD025: 待改檔: SO001
    服務滿意度代碼檔 CD026: 待改檔: SO007, SO008, SO009
    收費方式代碼檔 CD031: 待改檔: SO001, SO032, SO033, SO034, SO036
    付款種類代碼檔 CD032: 待改檔: SO032, SO033, SO034, SO036
    買賣方式代碼檔 CD034: 待改檔: CD022, SO004
    公司別代碼檔 CD039: 待改檔: SO001, SO014, SO016
    供電戶費用公式代碼檔 CD041: 待改檔: SO020
  */
  if p_TableName = 'CD004' then
    update SO001 set ClassName1=p_NewDesc where ClassCode1=p_CodeNo;
    update SO019 set ClassName=p_NewDesc where ClassCode=p_CodeNo;

  elsif p_TableName = 'CD005' then
    update SO007 set InstName=p_NewDesc where InstCode=p_CodeNo;

  elsif p_TableName = 'CD006' then
    update SO008 set ServiceName=p_NewDesc where ServiceCode=p_CodeNo;

  elsif p_TableName = 'CD007' then
    update SO009 set PRName=p_NewDesc where PRCode=p_CodeNo;

  elsif p_TableName = 'CD008' then
    update SO006 set ServiceName=p_NewDesc where ServiceCode=p_CodeNo;

  elsif p_TableName = 'CD009' then
    update SO001 set MediaName=p_NewDesc where MediaCode=p_CodeNo;
    update SO004 set MediaName=p_NewDesc where MediaCode=p_CodeNo;
    update SO007 set MediaName=p_NewDesc where MediaCode=p_CodeNo;

  elsif p_TableName = 'CD010' then
    update SO013 set ProfName=p_NewDesc where ProfCode=p_CodeNo;

  elsif p_TableName = 'CD011' then
    update SO008 set MFName1=p_NewDesc where MFCode1=p_CodeNo;
    update SO008 set MFName2=p_NewDesc where MFCode2=p_CodeNo;

  elsif p_TableName = 'CD012' then
    update SO001 set DelName=p_NewDesc where DelCode=p_CodeNo;

  elsif p_TableName = 'CD013' then
    update SO032 set UCName=p_NewDesc where UCCode=p_CodeNo;
    update SO033 set UCName=p_NewDesc where UCCode=p_CodeNo;
    update SO034 set UCName=p_NewDesc where UCCode=p_CodeNo;
    update SO036 set UCName=p_NewDesc where UCCode=p_CodeNo;

  elsif p_TableName = 'CD014' then
    update SO009 set ReasonName=p_NewDesc where ReasonCode=p_CodeNo;

  elsif p_TableName = 'CD015' then
    update SO007 set ReturnName=p_NewDesc where ReturnCode=p_CodeNo;
    update SO008 set ReturnName=p_NewDesc where ReturnCode=p_CodeNo;
    update SO009 set ReturnName=p_NewDesc where ReturnCode=p_CodeNo;

  elsif p_TableName = 'CD016' then
    update SO032 set STName=p_NewDesc where STCode=p_CodeNo;
    update SO033 set STName=p_NewDesc where STCode=p_CodeNo;
    update SO034 set STName=p_NewDesc where STCode=p_CodeNo;
    update SO036 set STName=p_NewDesc where STCode=p_CodeNo;

  elsif p_TableName = 'CD017' then
    update SO014 set StrtName=p_NewDesc where StrtCode=p_CodeNo;
    update SO016 set StrtName=p_NewDesc where StrtCode=p_CodeNo;

  elsif p_TableName = 'CD018' then
    update SO002 set BankName=p_NewDesc where BankCode=p_CodeNo;

  elsif p_TableName = 'CD019' then
    update CD005 set CitemName1=p_NewDesc where CitemCode1=p_CodeNo;
    update CD005 set CitemName2=p_NewDesc where CitemCode2=p_CodeNo;
    update CD005 set CitemName3=p_NewDesc where CitemCode3=p_CodeNo;
    update CD005 set CitemName4=p_NewDesc where CitemCode4=p_CodeNo;
    update CD005 set CitemName5=p_NewDesc where CitemCode5=p_CodeNo;
    update CD006 set CitemName1=p_NewDesc where CitemCode1=p_CodeNo;
    update CD006 set CitemName2=p_NewDesc where CitemCode2=p_CodeNo;
    update CD006 set CitemName3=p_NewDesc where CitemCode3=p_CodeNo;
    update CD006 set CitemName4=p_NewDesc where CitemCode4=p_CodeNo;
    update CD006 set CitemName5=p_NewDesc where CitemCode5=p_CodeNo;
    update CD007 set CitemName1=p_NewDesc where CitemCode1=p_CodeNo;
    update CD007 set CitemName2=p_NewDesc where CitemCode2=p_CodeNo;
    update CD007 set CitemName3=p_NewDesc where CitemCode3=p_CodeNo;
    update CD007 set CitemName4=p_NewDesc where CitemCode4=p_CodeNo;
    update CD007 set CitemName5=p_NewDesc where CitemCode5=p_CodeNo;
    update CD022 set DefCitemName=p_NewDesc where DefCitemCode=p_CodeNo;
    update CD024 set CitemName=p_NewDesc where CitemCode=p_CodeNo;
    update SO003 set CitemName=p_NewDesc where CitemCode=p_CodeNo;
    update SO032 set CitemName=p_NewDesc where CitemCode=p_CodeNo;
    update SO033 set CitemName=p_NewDesc where CitemCode=p_CodeNo;
    update SO034 set CitemName=p_NewDesc where CitemCode=p_CodeNo;
    update SO036 set CitemName=p_NewDesc where CitemCode=p_CodeNo;

  elsif p_TableName = 'CD020' then
    update SO001 set PWName=p_NewDesc where PWCode=p_CodeNo;

  elsif p_TableName = 'CD021' then
    update SO014 set BTName=p_NewDesc where BTCode=p_CodeNo;
    update SO016 set BTName=p_NewDesc where BTCode=p_CodeNo;

  elsif p_TableName = 'CD022' then
    update CD005 set FacilName1=p_NewDesc where FacilCode1=p_CodeNo;
    update CD005 set FacilName2=p_NewDesc where FacilCode2=p_CodeNo;
    update CD005 set FacilName3=p_NewDesc where FacilCode3=p_CodeNo;
    update CD006 set FacilName1=p_NewDesc where FacilCode1=p_CodeNo;
    update CD006 set FacilName2=p_NewDesc where FacilCode2=p_CodeNo;
    update CD006 set FacilName3=p_NewDesc where FacilCode3=p_CodeNo;
    update CD007 set FacilName1=p_NewDesc where FacilCode1=p_CodeNo;
    update CD007 set FacilName2=p_NewDesc where FacilCode2=p_CodeNo;
    update CD007 set FacilName3=p_NewDesc where FacilCode3=p_CodeNo;
    update SO004 set FaciName=p_NewDesc where FaciCode=p_CodeNo;
    update SO019 set FaciName1=p_NewDesc where FaciCode1=p_CodeNo;
    update SO019 set FaciName2=p_NewDesc where FaciCode2=p_CodeNo;

  elsif p_TableName = 'CD025' then
    update SO001 set OrgName=p_NewDesc where OrgCode=p_CodeNo;

  elsif p_TableName = 'CD026' then
    update SO007 set SatiName=p_NewDesc where SatiCode=p_CodeNo;
    update SO008 set SatiName=p_NewDesc where SatiCode=p_CodeNo;
    update SO009 set SatiName=p_NewDesc where SatiCode=p_CodeNo;

  elsif p_TableName = 'CD031' then
    update SO001 set CMName=p_NewDesc where CMCode=p_CodeNo;
    update SO033 set CMName=p_NewDesc where CMCode=p_CodeNo;
    update SO034 set CMName=p_NewDesc where CMCode=p_CodeNo;
    update SO036 set CMName=p_NewDesc where CMCode=p_CodeNo;

  elsif p_TableName = 'CD032' then
    update SO033 set PTName=p_NewDesc where PTCode=p_CodeNo;
    update SO034 set PTName=p_NewDesc where PTCode=p_CodeNo;
    update SO036 set PTName=p_NewDesc where PTCode=p_CodeNo;

  elsif p_TableName = 'CD034' then
    update CD022 Set DefBuyName=p_NewDesc where DefBuyCode=p_CodeNo;
    update SO004 set BuyName=p_NewDesc where BuyCode=p_CodeNo;

  elsif p_TableName = 'CD039' then
    update SO001 Set CompName=p_NewDesc where CompCode=p_CodeNo;
    update SO014 Set CompName=p_NewDesc where CompCode=p_CodeNo;
    update SO016 Set CompName=p_NewDesc where CompCode=p_CodeNo;

  elsif p_TableName = 'CD041' then
    update SO020 Set CountName=p_NewDesc where CountCode=p_CodeNo;

  elsif p_TableName = 'CD042' then
    update SO007 Set PromName=p_NewDesc where PromCode=p_CodeNo;

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