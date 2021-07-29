 
create or replace function SF_ClsWorkOrder
  (p_StartDate varChar2, p_stopdate varChar2, p_ClsEn varChar2, p_ClsName VarChar2, 
   p_Option1 number, p_Option2 number ,p_option3 number ,P_RetMsg out varchar2) 
  return number 
  as
 
  v_Opt1 number;
  v_Opt2 number;
  v_Opt2 number;
  v_Count number :=0 ;
  v_BatchNo number;
  v_NowDate date;
  v_Para varchar2(500);
  v_NowStr varchar2(19);
  NoChangeData EXCEPTION;
 
  cursor C_So007 is 
	select SNo,InstCode WorkCode,InstName WorkName,CustId,CustName,AcceptTime,
	FinTime,ReturnName,SignDate,SignName from So007 Where SignDate >=
	to_date(p_startDate,'YYYYMMDD') and SignDate <=to_date(p_StopDate,'YYYYMMDD')
	and ClsTime is Null;
 
  cursor C_So008 is 
	select SNo,ServiceCode WorkCode,ServiceName WorkName,CustId,CustName,AcceptTime,
	FinTime,ReturnName,SignDate,SignName from So008 Where SignDate >=
	to_date(p_startDate,'YYYYMMDD') and SignDate <=to_date(p_StopDate,'YYYYMMDD')
	and ClsTime is Null;
 
  cursor C_So009 is 
	select SNo,PRCode WorkCode,PRName WorkName,CustId,CustName,AcceptTime,
	FinTime,ReturnName,SignDate,SignName from So009 Where SignDate >=
	to_date(p_startDate,'YYYYMMDD') and SignDate <=to_date(p_StopDate,'YYYYMMDD')
	and ClsTime is Null;
begin
	if p_StartDate is null or p_StopDate is null Or p_ClsEn is null or p_ClsName is null 
	 or p_Option1 is null or p_Option2 is null or p_Option3 is null then 
		p_RetMsg :='Wrong parameter.';
		return -1;
	end if;
 
	v_NowDate :=sysdate;
	v_NowStr  := GiPackage.GetDTString2(v_NowDate);
        select S_SO064_BatchNo.NextVal into v_BatchNo from dual;
 
	if p_option1 =1 then 
		for cr1 in C_So007 loop
			update So007 set ClsTime=v_NowDate ,ClsEn=p_ClsEn, UpdTime=v_NowStr, UpdEn=p_ClsName
				where Sno=cr1.Sno;
			insert into So064 (BatchNo,ClsEN,WoType,Sno,WorkCode,WorkName,CustId,
				CustName,AcceptTime,FinTime,ReturnName,SignDate,SignName,ClsTime)
			values(v_BatchNo,p_ClsEN,1,cr1.Sno,cr1.WorkCode,cr1.WorkName,cr1.CustId
			      ,cr1.CustName,cr1.AcceptTime,cr1.FinTime,cr1.ReturnName,cr1.SignDate,
			      cr1.SignName,v_NowDate);
			v_Count:=v_Count+1;
		end loop;		
	end if ;	
 
	if p_option2 =1 then 
		for cr1 in C_So008 loop
			update So008 set ClsTime=v_NowDate ,ClsEn=p_ClsEn, UpdTime=v_NowStr, UpdEn=p_ClsName
				where Sno=cr1.Sno;
			insert into So064 (BatchNo,ClsEN,WoType,Sno,WorkCode,WorkName,CustId,
				CustName,AcceptTime,FinTime,ReturnName,SignDate,SignName,ClsTime)
			values(v_BatchNo,p_ClsEN,2,cr1.Sno,cr1.WorkCode,cr1.WorkName,cr1.CustId
			      ,cr1.CustName,cr1.AcceptTime,cr1.FinTime,cr1.ReturnName,cr1.SignDate,
			      cr1.SignName,v_NowDate);
			v_Count:=v_Count+1;
		end loop;		
	end if ;	
 
	if p_option3 =1 then 
		for cr1 in C_So009 loop
			update So009 set ClsTime=v_NowDate ,ClsEn=p_ClsEn, UpdTime=v_NowStr, UpdEn=p_ClsName
				where Sno=cr1.Sno;
			insert into So064 (BatchNo,ClsEN,WoType,Sno,WorkCode,WorkName,CustId,
				CustName,AcceptTime,FinTime,ReturnName,SignDate,SignName,ClsTime)
			values(v_BatchNo,p_ClsEN,3,cr1.Sno,cr1.WorkCode,cr1.WorkName,cr1.CustId
			      ,cr1.CustName,cr1.AcceptTime,cr1.FinTime,cr1.ReturnName,cr1.SignDate,
			      cr1.SignName,v_NowDate);
			v_Count:=v_Count+1;
		end loop;		
	end if ;
	--設定變數
	v_para :='Start date='||GiPackage.ED2CDString(to_date(p_StartDate,'YYYYMMDD'))||Chr(13)||
		'End date='||GiPackage.ED2CDString(to_date(p_StopDate,'YYYYMMDD'))||Chr(13)||
		 '裝機類='||p_Option1||Chr(13)||'維修類='||p_Option2||Chr(13)||
		 'Disc/move= '||p_Option3;
 
	insert into So053 (Type,TranDate,UpdEn,UpdTime,para) 
		values(2,to_date(p_StopDate,'YYYYMMDD'),p_ClsEn,v_NowStr,v_para);
 
	Update So062 Set TranDate=to_date(p_StopDate,'YYYYMMDD'),UpdEn=p_ClsEn
			,UpdTime=v_NowStr,Para=v_Para Where type =2;
 
	if SQL%ROWCount =0 then 
		Insert into So062 (Type,TranDate,UpdEn,UpdTime,para) 
			values (2,to_date(p_StopDate,'YYYYMMDD'),p_ClsEn,v_NowStr,v_Para);
	end if ;
 
	p_RetMsg:='Execution OK, Rcd count='||to_char(v_count);
	commit;
 
	if v_Count =0 then 
		return 0;
	else 
		return v_BatchNo;
	end if;
exception 
	When Others Then 
             RollBack;
	     p_RetMsg:='Other error.';
	     return -99;	
end;
/
