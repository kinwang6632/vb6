/*
@C:\gird\v400\script\SF_TranInv2
@C:\Gird\v400\Script\�ݴ��{����2\SF_TranInv2

variable nn number
variable msg varchar2(4000)
--exec :nn := SF_TranInv2(11, 'IN (''C'')', 694406, 694406, null, null, null, null,'20080513', '20080513', null, null, null, null, null, 1, 2, 0, 0, null, null, null, 1, 1, :Msg);
  exec :nn := SF_TranInv2( 3, 'IN (''D'')',   NULL,   NULL, null, null, null, null,'20100228', '20100228', null, null, null, null, null, 1, 2, 1, 1, null, null, null, 1, 1, :Msg);
print NN
PRINT MSG

  �o�����ɧ@�~(�w��}�շs�o���t��):  SF_TranInv2()
  �ɦW:  SF_TranInv2.sql
  IN  
      p_CompCode number: ���q�O�N�X
      p_ServiceTypeSQL varchar2: �A�����O
      p_CustId1 number: �_�l�Ȥ�s��
      p_CustId2 number: �I��Ȥ�s��
      p_ServCodeSQL varchar2: �A�Ȱ�SQL����r��
      p_ClassSQL varchar2: �Ȥ����OSQL����r�� 
      p_BillTypeSQL varchar2: ��ں���SQL����r��
      p_CitemSQL varchar2: ���O����SQL����r��  
      p_RealDate1 varchar2: �_�l�J�b���, 'YYYYMMDD'
      p_RealDate2 varchar2: �I��J�b���, 'YYYYMMDD'
      p_ClctEnSQL varchar2: ���O�H���N��
      p_UpdDate1 varchar2: �_�l���ʤ��, 'YYYYMMDD'
      p_UpdDate2 varchar2: �I��ʤ��, 'YYYYMMDD'
      p_UpdEn varchar2: ���ʤH���m�W
      p_CMCodeSQL varchar2: ���O�覡�N�X
      p_InvType number: �o������(1=�Ҧ�, 2=�G�p, 3=�T�p)
      p_AddrType number: �H��a�}����(1=�˾�, 2=���O, 3=�l�H)
      p_AmtCheck number: ���B<0�O�_����(0=�_, 1=�O)
      p_AmtZero number: ���B=0�O�_����(0=�_, 1=�O)
      p_MduIdSQL varchar2:�˾��j��SQL����r��
      p_ShouldDate1 varchar2: �_�l�������, 'YYYYMMDD'
      p_ShouldDate2 varchar2: �I���������, 'YYYYMMDD'
      p_AreaCity number: �[������F�Ϧr��(0=�_, 1=�O)
      p_Mode number: 1=��SO4131, 2=��SO4132, 3=��SO4141, 4=��SO4142
  OUT  
      p_RetMsg varchar2: ���G�T�� (�ܤ�1000 bytes)
  Return number: �B�z���G�N�X, �����T���s��p_RetMsg��
      >=0: p_RetMsg='���ͧ���, �@��O<���>��, �B�z����=<����>' 
         -1: p_RetMsg='�Ѽƿ��~'
         -2: p_RetMsg='SQL���~: <SQL���O>'
         -3: p_RetMsg='�䤣�즬�O���إN�X'
         -4: p_RetMsg='�A�����O�S�]�w'
       -99: p_RetMsg='��L���~'

  By: Lawrence
  Date: 2003.08.18 ���gfor Jacy�W��(�h�b��B�z+�L�b��B�z) and InvoiceType��SO001����SO002
                              �e�ݶץX���ɮ榡��, �ݤUorder by CustId,SNO
           2003.08.22 z_MediaBillNo�[Nvl()�B�z; �վ�elsif p_AddrType=3 then..... and Else����;
                              Ū��SO002A IS NULL��, Ū��SO002 IS NULL��, Ū��SO001
           2003.08.25 �վ��ܼƤ����Y�P�_��
           2003.09.04 �]��select SQL.begin...end
           2003.09.17 �[z_InvNo := z_strInvNo and get SO002A������
           2003.09.22 �[v_ContName and v_ChkCustId�P�_��
           2003.09.26 for Jacy�W��##�W�[ 9.����H chargetitle,--�W�[ 16.�A�����O servicetype(CSMISV4.0_DMIST_�s���o������_�ץ�_JACY_920926.doc)
           2003.10.09 for Jacy�W��(CSMISV4.0_DMIST_�s���o������_�ץ�_JACY_921008.doc)
           2003.11.18 Mark delete�Ϊk,�אּTruncate
           2004.05.24 �վ��ܼ�z_VouString,z_VouString1��varchar2(305); z_Note��varchar2(750)
           2004.12.20 for �ݨD�s���G1342(���O����ۦP, �B�b���ɮɥu�n���O��ƪ��b���ۦP, �h�ݩߦb�P�@����ƤU)
           2005.05.17 for �ݨD�s���G1656 �վ�B�z�L�b�᪺����z_InvoiceType�P�_
           2005.09.03 �վ�CitemName, invtitle����
           2006.03.28 for 2205, 2206���D���W�[�˾��j��SQL����r��
           2006.08.07 for 2671 add ��������d��
           2006.11.08 for 2841
           2006.12.28 for 2930 �����a�}�[�꿤����F�Ϧr��
           2007.01.03 PM�S���a�}���n�[�j��142, �w��Casper update���D����󻡩�
           2007.12.06 for 3398
           2007.12.06 for 3436 ��ŪSO002A��ŪSO138
           2008.03.11 for 3785
           2008.04.17 add AccountNo is null and InvSeqNo>0��Ū��SO138
           2008.04.24 �վ�@�U�{���L�b���B�z��z_strInvNo, z_ContName�ܼƦW��
           2008.05.23 �վ�{����/�L�b���B�z��z_AddrNo����CityName, AreaName���, ���A�u�̸˾��a�}�Ӧ�(for PM Minchen and �y�pCasper)
           2008.08.25 3883 �pinvseqno = 0 or invseqno is null �h�̾ڭ즳�o�����ɪ����ǨӳB�z
           2008.08.29 3986 ���ަ��L�b��ҼW�[��InvSeqNo�ȳB�z
           2009.12.22 5439
           2010.02.10 for 5542 v_SQL�����ܼƥ[�j��10000
           2010.05.11 FOR 5438 Jacky��CombCitemCode and CombStartDate����, �h��Ū��CombAmount, CombStartDate, CombStopDate���

  �H�U��Pierre���ק����:
	2010.07.15 for 5668, ���ɤ����r��W�["�o���}�ߺ����B�`��E-MAIL�B��ʹq�ܡB���س��N�X�B���س��W��" 5�����
*/
CREATE OR REPLACE function SF_TranInv2(p_CompCode number, p_ServiceTypeSQL varchar2,
      p_CustId1 number, p_CustId2 number, p_ServCodeSQL varchar2, p_ClassSQL varchar2, 
      p_BillTypeSQL varchar2, p_CitemSQL varchar2, 
      p_RealDate1 varchar2, p_RealDate2 varchar2, 
      p_ClctEnSQL varchar2, p_UpdDate1 varchar2, p_UpdDate2 varchar2, 
      p_UpdEn varchar2, p_CMCodeSQL varchar2, p_InvType number, p_AddrType number, 
      p_AmtCheck number, p_AmtZero number, p_MduIdSQL varchar2, p_ShouldDate1 varchar2, 
      p_ShouldDate2 varchar2, p_AreaCity number, p_Mode number, p_RetMsg out varchar2) return number
  AS
  c39		char(1) := chr(39);
  v_Error       	boolean;
  v_Char1  		char(1);
  v_Char        	char(1);
  v_ChkRealDate 	date;
  v_StartExecTime 	date;
  v_TaxCode 	number;
  v_TaxFlag     	number;
  v_InvCompCode 	number;
  v_CustId		number;
  v_ChkCustId	number;
  v_Sno     		number :=1;
  v_RealAmt     	number :=0;
  v_RcdCount 	number :=0;
  v_CurrCitemCode 	number;
  v_ChkBillNo   	varchar2(15);
  v_AccountNo 	varchar2(16);
  v_ContName	varchar2(50);
  v_SQL 		varchar2(10000);
  v_SQL1 		varchar2(10000);
  v_SQL2 		varchar2(10000);
  v_ChkMediaBillNo 	varchar2(11);
  v_ChkAccountNo   	varchar2(16);
  z_TaxType	char(1);
  z_ServiceType	char(1);
  z_RealDate	date;
  z_UpdA		date;
  z_RealStopDate  	date;
  z_RealStartDate 	date;
  z_CustId 		number;
  z_RealAmt	number;
  z_CitemCode	number;
  z_AddrNo	number;
  z_TaxRate     	number;
  z_TaxMoney    	number;
  z_Qty         	number;
  z_TaxCode     	number;
  z_Item       	number;
  z_UpdB		number;
  z_InvoiceType	number;
  z_ZipCode	varchar2(5);
  z_strInvNo    	varchar2(8);
  z_InvNo		varchar2(8);
  z_MediaBillNo	varchar2(11);
  z_BillNo		varchar2(15);
  z_AccountNo	varchar2(16);
  z_UpdTime	varchar2(19);
  z_CitemName	varchar2(40);
  z_Tel1        	varchar2(20);
  z_ClctName    	varchar2(20);
  z_ContName	varchar2(50);
  z_InvTitle	varchar2(50);
  z_ChargeTitle	varchar2(50);
  z_Address	varchar2(142);
  z_InvAddress	varchar2(142);
  z_VouString   	varchar2(500);	-- 2007.03.11 by Lawrence ���380, 2010.07.15 by Pierre ���500
  z_VouString1  	varchar2(500);	-- 2007.03.11 by Lawrence ���380, 2010.07.15 by Pierre ���500
  z_Note		varchar2(750);
  z_AreaName 	varchar2(20);	-- 2006.12.28
  z_CityName 	varchar2(20);	-- 2006.12.28
  z_InvSeqNo 	number;		-- 2007.12.06 by Lawrence
  z_CurrInvSeqNo	number:=0;	-- 2008.08.29 by Lawrence
  z_InvPurposeCode 	number;		-- 2008.03.11 by Lawrence
  z_InvPurposeName 	varchar2(20);	-- 2008.03.11 by Lawrence
  z_CombCitemCode	number;		-- 2009.12.22 by Lawrence
  z_CombCitemName	varchar2(40);	-- 2009.12.22 by Lawrence
  z_CombAmount	number;		-- 2010.05.11 by Lawrence
  z_CombStartDate	date;		-- 2010.05.11 by Lawrence
  z_CombStopDate	date;		-- 2010.05.11 by Lawrence
  x_ContName	varchar2(50);

  TYPE CurTyp IS REF CURSOR;  --�ۭqcursor���A
  v_DyCursor CurTyp;          --�ѷs��dynamic SQL

  -- 2010.07.15 by Pierre
  z_FaciSeqNo	varchar2(15);		-- SO033/SO034���]�Ƭy����
  z_InvoiceKind	number;			-- �o���}�ߺ���
  z_Email	varchar2(50);		-- �`��E-MAIL
  z_ContMobile	varchar2(20);		-- ��ʹq��
  z_DenRecCode	number;			-- ���س��N�X
  z_DenRecName	varchar2(20);		-- ���س��W��
  x_RcdCnt1	number:=0;		-- ���b�᪺����
  x_RcdCnt2	number:=0;		-- �L�b�᪺����
BEGIN
  v_char := chr(01);   -- ��춡��Ÿ�
  -- �Ѽ��ˬd
  if p_InvType is null or (p_InvType not in (1,2,3)) or p_AddrType is null or (p_AddrType not in (1,2,3)) or (p_Mode not in (1,2)) then
     p_RetMsg := '�Ѽƿ��~';
     return -1;
  end if;
  -- �h�A�Ȯ�, InvCompCode�ݳ]�ۦP�_�h�|�߿��b
  v_SQL:='select InvCompCode from CD046 where CodeNo '||p_ServiceTypeSQL||' and InvCompCode is not null and RowNum=1';
  begin
     EXECUTE IMMEDIATE v_SQL INTO v_InvCompCode;     
  exception
    when others then
       p_RetMsg:='�A�����Oor�o���t�Τ��q�O�S�]�w';
       return -4;         
  end;
  -- �R���U�Ȧs��
  -- 2007.12.06 by Lawrence Mark Truncate�אּdelete...where...
  delete from Tmp004 where MFlag=p_Mode;
  delete from Tmp005 where MFlag=p_Mode;
  -- �]�w�U�ܼƪ��
  v_StartExecTime := sysdate;		-- �}�l����ɶ�
  v_SQL1 := null;
  -- �Y���Ȥ�s������
  if nvl(p_CustId1,0)>0 then
     v_SQL1 := v_SQL1 || ' and A.CustId>=' || p_CustId1 || ' and A.CustId<=' || p_CustId2;
  end if;  
  -- �Y�����O���ر���
  if p_CitemSQL is not null then
     v_SQL1 := v_SQL1 || ' and A.CitemCode ' || p_CitemSQL;
  end if;
  if p_ServCodeSQL is not null then
     v_SQL1 := v_SQL1 || ' and A.ServCode ' || p_ServCodeSQL;
  end if;
  -- �Y���Ȥ����O����
  if p_ClassSQL is not null then
     v_SQL1 := v_SQL1 || ' and A.ClassCode ' || p_ClassSQL;
  end if;
  -- 2006.03.28 by Lawrence
  -- �Y���˾��j�ӱ���
  if p_MduIdSQL is not null then
     v_SQL1 := v_SQL1 || ' and D.MduId ' || p_MduIdSQL;
  end if;
  -- �Y�����q�O����
  if p_CompCode is not null then
     v_SQL1 := v_SQL1 || ' and A.CompCode=' || p_CompCode;
  end if;
  -- �Y���J�b�������
  if p_RealDate1 is not null then
     v_SQL1 := v_SQL1 || ' and A.RealDate>='||'to_date('||chr(39)||p_RealDate1||chr(39)||','||chr(39)||'YYYYMMDD'||chr(39)||')'||
                       ' and A.RealDate<=to_date('||chr(39)||p_RealDate2||chr(39)||','||chr(39)||'YYYYMMDD'||chr(39)||')';
  end if;
  -- 2006.08.07 by Lawrence �Y�������������
  if p_ShouldDate1 is not null then
     v_SQL1 := v_SQL1 || ' and A.ShouldDate>='||'to_date('||chr(39)||p_ShouldDate1||chr(39)||','||chr(39)||'YYYYMMDD'||chr(39)||')'||
                       ' and A.ShouldDate<=to_date('||chr(39)||p_ShouldDate2||chr(39)||','||chr(39)||'YYYYMMDD'||chr(39)||')';
  end if;
  -- �Y�����O�H������
  if p_ClctEnSQL is not null then
     v_SQL1 := v_SQL1 || ' and A.ClctEn ' || p_ClctEnSQL;
  end if;
  -- �Y�����ʤ������
  if p_UpdDate1 is not null then
     v_SQL1 := v_SQL1 || ' And A.UpdTime>='||c39||ltrim(to_char(to_number(substr(p_UpdDate1,1,4))-1911,'99'))||
                       '/'||substr(p_UpdDate1,5,2)||'/'||substr(p_UpdDate1,7,2)||c39||
                       ' and A.UpdTime<='||c39||ltrim(to_char(to_number(substr(p_UpdDate2,1,4))-1911,'99'))||
                       '/'||substr(p_UpdDate2,5,2)||'/'||substr(p_UpdDate2,7,2)||' 23:59:59'||c39;
  end if;
  -- �Y�����ʤH������
  if p_UpdEn is not null then
     v_SQL1 := ' and A.UpdEn=' || c39 || p_UpdEn || c39;
  end if;
  -- �Y�����O�覡����
  -- 2006.11.08 Lawrence p_CMCode�令p_CMCodeSQL�ƿ�
  if p_CMCodeSQL is not null then
     v_SQL1 := v_SQL1 || ' and A.CMCode ' || p_CMCodeSQL;
  end if;
  v_SQL1 := v_SQL1 || ' and a.InvoiceTime is null and a.CancelFlag=0 and a.GUINo is null';
  -- �Y����ں�������
  if p_BillTypeSQL is not null then
     v_Char1 := substr(p_BillTypeSQL, 1, 1);
     if v_Char1 = '=' or v_Char1 = '!' then
        if p_BillTypeSQL like '%B%' or p_BillTypeSQL like '%T%'  then
           v_SQL1 := v_SQL1 || ' and substr(A.BillNo,7,1) '||p_BillTypeSQL;
        else
           v_SQL1 := v_SQL1 || ' and substr(A.BillNo,1,1) '||p_BillTypeSQL;
        end if;
     else
        if p_BillTypeSQL like '%B%' and p_BillTypeSQL like '%T%'  then
           v_SQL1 := v_SQL1 || ' and (substr(A.BillNo,7,1) '||p_BillTypeSQL||')';
        else
           v_SQL1 := v_SQL1 || ' and (substr(A.BillNo,1,1) '||p_BillTypeSQL||' or substr(A.BillNo,7,1) '||p_BillTypeSQL||')';
        end if;
     end if;
  end if;
  v_SQL2 := 'select A.CustId,A.BillNo,A.RealDate,A.RealAmt,A.CitemCode,A.CitemName,'||
                    'A.UpdTime,A.RealStartDate,A.RealStopDate,A.Note,A.Quantity,A.ClctName,A.Item,C.InvNo,';
  -- ���a�}���
  if p_AddrType=1 then
     v_SQL2 := v_SQL2 || 'B.InstAddress Address, B.InstAddrNo AddrNo';
  elsif p_AddrType=2 then
     v_SQL2 := v_SQL2 || 'B.ChargeAddress Address, B.ChargeAddrNo AddrNo';
  else
     v_SQL2 := v_SQL2 || 'B.MailAddress Address, B.MailAddrNo AddrNo';
  end if;
  --�B�z�L�b�᪺����(�̫Ȥ�s��,�o���Τ@�s��,�o�����Y,�p���H(�s�[���) GROUP BY ���,�N�ۦP�����ͦb�@�_)
  -- 2005.12.29 by Lawrence (2006) for Jacy �վ�order by�[RealDate
  -- 2010.07.15 by Pierre, �[��FaciSeqNo
  v_SQL := v_SQL2 || ',A.AccountNo,A.MediaBillNo,A.ServiceType,C.InvoiceType,C.InvTitle,C.ContName,C.InvAddress, D.AreaName, D.CityName, C.InvPurposeCode, C.InvPurposeName, A.InvSeqNo, A.CombCitemCode, A.CombCitemName, A.CombAmount, A.CombStartDate, A.CombStopDate, A.FaciSeqNo from SO034 A, SO001 B, SO002 C, SO014 D where A.CustId=B.CustId and B.CustId=C.CustId and A.ServiceType '||p_ServiceTypeSQL||' and A.ServiceType=C.ServiceType and C.PreInvoice=0 and B.CompCode=D.CompCode and B.InstAddrNo=D.AddrNo ' || 
	v_SQL1 || ' and A.AccountNo is null order by A.CustId,C.InvNo,C.InvTitle,C.ContName,A.RealDate';
  --�B�z���b�᪺����(��AccountNo,MediaBillNo GROUP BY���,�N�ۦP�����ͦb�@�_)
  -- 2004.12.20 by Lawrence order by A.RealDate,A.MediaBillNo���
  -- 2010.07.15 by Pierre, �[��FaciSeqNo
  v_SQL1 := v_SQL2 || ',A.AccountNo,A.MediaBillNo,A.ServiceType, D.AreaName, D.CityName, A.InvSeqNo, C.InvPurposeCode, C.InvPurposeName, A.CombCitemCode, A.CombCitemName, A.CombAmount, A.CombStartDate, A.CombStopDate, A.FaciSeqNo from SO034 A, SO001 B, SO002 C, SO014 D where A.CustId=B.CustId and B.CustId=C.CustId and A.ServiceType '||p_ServiceTypeSQL||' and A.ServiceType=C.ServiceType and C.PreInvoice=0 and B.CompCode=D.CompCode and B.InstAddrNo=D.AddrNo ' || 
                    v_SQL1 || ' and A.AccountNo is not null order by A.AccountNo,A.RealDate,A.MediaBillNo';
  
  --1.�B�z���b�᪺����
  begin
     OPEN v_DyCursor FOR v_SQL1;
  exception
    when others then
       rollback;
       p_RetMsg := 'SQL3���~: ' || v_SQL1;
       return -2;
  end;
  v_AccountNo :='xyz';
  v_CurrCitemCode := 0;
  v_ChkAccountNo := 'abc';
  v_ChkRealDate := to_Date('19000101','YYYYMMDD');
  v_ChkMediaBillNo := '19000100001';
  v_ChkCustId := 0;
  loop
    v_Error := False;
    -- 2010.07.15 by Pierre, �[��FaciSeqNo
    FETCH v_DyCursor INTO z_CustId,z_BillNo,z_RealDate,z_RealAmt,z_CitemCode,z_CitemName,z_UpdTime,z_RealStartDate,z_RealStopDate,z_Note,z_Qty,z_ClctName,z_Item,z_strInvNo,z_Address,z_AddrNo,z_AccountNo,z_MediaBillNo,z_ServiceType,z_AreaName,z_CityName,z_InvSeqNo,z_InvPurposeCode,z_InvPurposeName,z_CombCitemCode,z_CombCitemName, z_CombAmount, z_CombStartDate, z_CombStopDate, z_FaciSeqNo;
    EXIT WHEN v_DyCursor%NOTFOUND;
    DBMS_OUTPUT.PUT_LINE('v_AccountNo:'||v_AccountNo||' / z_AccountNo:'||z_AccountNo||' / z_MediaBillNo:'||z_MediaBillNo);
    -- 2010.05.11 by Lawrence
    if z_CombCitemCode is not null and z_CombStartDate is not null then
       z_RealAmt := z_CombAmount;
       z_RealStartDate := z_CombStartDate;
       z_RealStopDate := z_CombStopDate;
    end if;
    if (v_AccountNo != z_AccountNo) or (v_ChkCustId != z_CustId) or (z_CurrInvSeqNo != nvl(z_InvSeqNo,0)) then
       v_AccountNo := z_AccountNo;
       z_CurrInvSeqNo := nvl(z_InvSeqNo,0);
       -- 2006.12.28 by Lawrence
       if p_AreaCity=1 then
          -- 2008.05.23 by Lawrence for Minchen and �y�pCasper
          select CityName,AreaName into z_CityName,z_AreaName from so014 where AddrNo=z_AddrNo;
          z_Address:=Trim(z_CityName)||Trim(z_AreaName)||Trim(z_Address);
          DBMS_OUTPUT.PUT_LINE('2. Address='||z_Address||'/AddrNo='||z_AddrNo||'/ AreaName='||z_AreaName||'/ CityName='||z_CityName);
       end if;
       -- ���Ȥ᪺�o�����
       begin
          DBMS_OUTPUT.PUT_LINE('a.'||z_InvTitle);
          begin
             -- ��SO002A���
             -- 2003.08.22 by Lawrence �վ�elsif p_AddrType=3 then..... and Else����
             -- 2003.09.26 by Lawrence �[ChargeTitle
             if p_AddrType=2 and nvl(z_InvSeqNo,0)>0 then
                begin
                   -- 2007.12.06 by Lawrence
                   -- 2010.07.15 by Pierre, �[��3�����
                   select InvNo, InvTitle, InvAddress, InvoiceType,ChargeAddress, ChargeAddrNo, ChargeTitle, InvPurposeCode, InvPurposeName, InvoiceKind, DenRecCode, DenRecName into z_InvNo, z_InvTitle, z_InvAddress, z_InvoiceType, z_Address, z_AddrNo, z_ChargeTitle, z_InvPurposeCode, z_InvPurposeName, z_InvoiceKind, z_DenRecCode, z_DenRecName 
                      from SO138 where InvSeqNo=z_InvSeqNo; 
                   -- 2006.12.28 by Lawrence
                   if p_AreaCity=1 then
                      select CityName,AreaName into z_CityName,z_AreaName from so014 where AddrNo=z_AddrNo;
                      z_Address:=Trim(z_CityName)||Trim(z_AreaName)||Trim(z_Address);
                   end if;
                exception
                  when others then
                     p_RetMsg:='so138 not found! CustId : '||z_CustId||', p_AddrType='||p_AddrType;
                     return -99;
                end;
             elsif p_AddrType=3 and nvl(z_InvSeqNo,0)>0 then
                begin
                   -- 2007.12.06 by Lawrence
                   -- 2010.07.15 by Pierre, �[��3�����
                   select InvNo, InvTitle, InvAddress, InvoiceType,MailAddress, MailAddrNo, ChargeTitle, InvPurposeCode, InvPurposeName, InvoiceKind, DenRecCode, DenRecName into z_InvNo, z_InvTitle, z_InvAddress, z_InvoiceType, z_Address, z_AddrNo, z_ChargeTitle, z_InvPurposeCode, z_InvPurposeName, z_InvoiceKind, z_DenRecCode, z_DenRecName 
                      from SO138 where InvSeqNo=z_InvSeqNo; 
                   -- 2006.12.28 by Lawrence
                   if p_AreaCity=1 then
                      select CityName,AreaName into z_CityName,z_AreaName from so014 where AddrNo=z_AddrNo;
                      z_Address:=Trim(z_CityName)||Trim(z_AreaName)||Trim(z_Address);
                   end if;
                exception
                  when others then
                     p_RetMsg:='so138 not found! CustId : '||z_CustId||', p_AddrType='||p_AddrType;
                     return -99;
                end;
             elsif nvl(z_InvSeqNo,0)>0 then
                begin
                   -- 2007.12.06 by Lawrence
                   -- 2010.07.15 by Pierre, �[��3�����
                   select InvNo, InvTitle, InvAddress, InvoiceType, ChargeTitle, InvPurposeCode, InvPurposeName, InvoiceKind, DenRecCode, DenRecName into z_InvNo, z_InvTitle, z_InvAddress, z_InvoiceType, z_ChargeTitle, z_InvPurposeCode, z_InvPurposeName, z_InvoiceKind, z_DenRecCode, z_DenRecName
                      from SO138 where InvSeqNo=z_InvSeqNo; 
                exception
                  when others then
                     p_RetMsg:='so138 not found! CustId : '||z_CustId||', p_AddrType='||p_AddrType;
                     return -99;
                end;
             elsif nvl(z_InvSeqNo,0)=0 then
               -- 2010.07.15 by Pierre, �Y[�o���y����]�L��, �h�ܫȤ�U�A�Ȱ򥻸����(SO002)��, ��CustId+CompCode+ServiceType������ƪ��o���}�ߺ���(InvoiceKind)�B���س��N�X(DenRecCode)�B���س��W��(DenRecName)
               begin
                 select InvoiceKind, DenRecCode, DenRecName into z_InvoiceKind, z_DenRecCode, z_DenRecName from SO002 where CustId=z_CustId and CompCode=p_CompCode and ServiceType=z_ServiceType;
               exception
                 when others then
                   z_InvoiceKind := 0;
                   z_DenRecCode := null;
                   z_DenRecName  := null;
               end;
             end if;
          exception
            when others then
               z_InvNo:=null;
               z_InvTitle:=null;
               z_InvAddress:=null;
               z_ChargeTitle:=null;
               -- 2010.07.15 by Pierre, �[3�����
               z_InvoiceKind := 0;
               z_DenRecCode := null;
               z_DenRecName  := null;
               DBMS_OUTPUT.PUT_LINE('b1.'||z_InvTitle);         
          end;
          -- 2003.09.26 by Lawrence
          x_ContName:=z_ChargeTitle;
          if z_InvTitle is null then  -- �Ĥ@����
             z_InvTitle:=z_ChargeTitle;
          end if;
          DBMS_OUTPUT.PUT_LINE('c.'||z_InvTitle||'/z_InvoiceType='||z_InvoiceType||'/p_InvType='||p_InvType);
          -- check�o������, �ŦX�����͸��
          if p_InvType!=1 then
             if z_InvoiceType<>p_InvType then
                GOTO go_next;
             end if;
          else
             if z_InvoiceType=0 then
                GOTO go_next;
             end if;
          end if;
          if z_InvAddress is null then
             z_InvAddress := z_Address;
          end if;
       exception
         when others then
            if z_InvAddress is null then
               z_InvAddress := z_Address;
            end if;
            z_InvNo := null;
       end;
       -- ���H��a�}���l���ϸ�
       begin
          select ZipCode into z_ZipCode from SO014 where AddrNo=z_AddrNo and CompCode=p_CompCode;
       exception
         when others then
            z_ZipCode := null;
       end;
       -- ���Ȥ�q�ܤ@
       begin
          select tel1 into z_Tel1 from SO001 where CustId=z_CustId and CompCode=p_CompCode;
       exception
         when others then
            z_Tel1 := null;
       end;
       -- *****************************************************
       -- 2010.07.15 by Pierre, ���Ȥ᪺�`��E-MAIL�B��ʹq��
       begin
         if z_FaciSeqNo is not null then
           -- �]z_FaciSeqNo���e�i�ର�Ƚs, �y��SO004�L�ӵ��]�Ƹ��, �����p�N�n�hSO001,SO002��Email, ContMobile
           begin
             select Email, ContMobile into z_Email, z_ContMobile from SO004 where SeqNo=z_FaciSeqNo;
           exception
             when others then
               select Email into z_Email from SO002 where CustId=z_CustId and CompCode=p_CompCode and ServiceType=z_ServiceType;
               select Tel3 into z_ContMobile from SO001 where CustId=z_CustId and CompCode=p_CompCode;
           end;
         else
           select Email into z_Email from SO002 where CustId=z_CustId and CompCode=p_CompCode and ServiceType=z_ServiceType;
           select Tel3 into z_ContMobile from SO001 where CustId=z_CustId and CompCode=p_CompCode;
         end if;
       exception
         when others then
           z_Email := null;
           z_ContMobile := null;
       end;
       -- *****************************************************

    else
       if z_InvAddress is null then
          z_InvAddress := z_Address;
       end if;
    end if;
    -- �M�w�ҵ|�O
    if v_CurrCitemCode != z_CitemCode then
       v_CurrCitemCode := z_CitemCode;
       begin
          select TaxCode into v_TaxCode from CD019 where CodeNo=v_CurrCitemCode and (ServiceType=z_ServiceType or ServiceType is null);
          if v_TaxCode = 1 then
             z_TaxType := '1';
             -- ���|�v
             begin
                select Rate1 into z_TaxRate from CD033 where CodeNo=v_TaxCode;
             exception
               when others then
                  z_TaxRate := Nvl(z_TaxRate,0);
             end;
          elsif v_TaxCode = 3 then
             z_TaxType := '3';
             z_TaxRate:=0;
          elsif v_TaxCode is null then
             z_TaxType := null;
             z_TaxRate:=0;
          else
             z_TaxType := '2';
             z_TaxRate:=0;
          end if;
       exception
         when others then
            p_RetMsg := '�䤣�즬�O���إN�X';
            return -3;        
       end;
       --���|�v�N�X�ɤ���"�O�_�һ�"
       if v_TaxCode is not null then
          begin
             select TaxFlag into v_TaxFlag from CD033 where CodeNo=v_TaxCode;
          exception
            when others then
               p_RetMsg:='cd003.TaxFlag not found! CodeNo : '||v_TaxCode;
               return -99;
          end;
       end if;
       if v_TaxFlag=0 then 
          z_TaxType := null;
          z_TaxRate :=0;
       end if;
    end if;
    -- �����ʤ���P�ɶ�
    begin
       if substr(z_UpdTime,3,1)='/' then
          z_UpdA := to_date(to_char(to_number(substr(z_UpdTime,1,2))+1911)||substr(z_UpdTime,4,2)||substr(z_UpdTime,7,2),'YYYYMMDD');
          z_UpdB := to_number(substr(z_UpdTime, 10, 2)||substr(z_UpdTime, 13, 2));
       else
          if z_UpdTime is not null then
             z_UpdA := to_date(substr(z_UpdTime,1,4)||substr(z_UpdTime,6,2)||substr(z_UpdTime,9,2),'YYYYMMDD');
          else
             z_UpdA :=Null;
          end if;
          z_UpdB := to_number(substr(z_UpdTime, 12, 2)||substr(z_UpdTime, 15, 2));
       end if;
    exception
      when others then
         p_RetMsg := '���ʤ�����~! '||z_UpdTime||' '||z_UpdA;
         return -6;              
    end;
    -- �s�W�ܵo�����ɼȦs��
    -- 2003.08.22 by Lawrence z_MediaBillNo�[Nvl()
    -- 2004.12.20 by Lawrence �O�d���, �A�̻ݨD�վ�{��
    if (v_ChkAccountNo = v_AccountNo) and (v_ChkRealDate = z_RealDate) and (v_ChkCustId = z_CustId) and (z_CurrInvSeqNo = nvl(z_InvSeqNo,0)) then
       z_VouString1 := '';
       if z_TaxRate>0 then
          z_TaxMoney := z_RealAmt - round((z_RealAmt/(1+(z_TaxRate/100))),0);
       else
          z_TaxMoney := 0;
       end if;
       if z_RealAmt>0 then
          z_VouString := '--N'||v_char||z_BillNo||v_char||to_char(nvl(z_Item,0))||v_char||z_TaxType||v_char||to_char(z_RealDate,'YYYYMMDD')||v_char||to_char(z_CitemCode)||v_char||z_CitemName||v_char||'1'||v_char||
                                   to_char(nvl(z_RealAmt-z_TaxMoney,0))||v_char||to_char(nvl(z_TaxRate,0))||v_char||to_char(nvl(z_TaxMoney,0))||v_char||to_char(nvl(z_RealAmt,0))||v_char||to_char(z_RealStartDate,'YYYYMMDD')||
                                   v_char||to_char(z_RealStopDate,'YYYYMMDD')||v_char||z_ClctName||v_char||z_ServiceType||v_char||to_char(z_CombCitemCode)||v_char||z_CombCitemName;
       end if;
       if p_AmtCheck=1 and z_RealAmt<0 then
          z_VouString := '--N'||v_char||z_BillNo||v_char||to_char(nvl(z_Item,0))||v_char||z_TaxType||v_char||to_char(z_RealDate,'YYYYMMDD')||v_char||to_char(z_CitemCode)||v_char||z_CitemName||v_char||'1'||v_char||
                                   to_char(nvl(z_RealAmt-z_TaxMoney,0))||v_char||to_char(nvl(z_TaxRate,0))||v_char||to_char(nvl(z_TaxMoney,0))||v_char||to_char(nvl(z_RealAmt,0))||v_char||to_char(z_RealStartDate,'YYYYMMDD')||
                                   v_char||to_char(z_RealStopDate,'YYYYMMDD')||v_char||z_ClctName||v_char||z_ServiceType||v_char||to_char(z_CombCitemCode)||v_char||z_CombCitemName;
       end if;
       if p_AmtZero=1 and z_RealAmt=0 then
          z_VouString := '--N'||v_char||z_BillNo||v_char||to_char(nvl(z_Item,0))||v_char||z_TaxType||v_char||to_char(z_RealDate,'YYYYMMDD')||v_char||to_char(z_CitemCode)||v_char||z_CitemName||v_char||'1'||v_char||
                                   to_char(nvl(z_RealAmt-z_TaxMoney,0))||v_char||to_char(nvl(z_TaxRate,0))||v_char||to_char(nvl(z_TaxMoney,0))||v_char||to_char(nvl(z_RealAmt,0))||v_char||to_char(z_RealStartDate,'YYYYMMDD')||
                                   v_char||to_char(z_RealStopDate,'YYYYMMDD')||v_char||z_ClctName||v_char||z_ServiceType||v_char||to_char(z_CombCitemCode)||v_char||z_CombCitemName;
       end if;
       DBMS_OUTPUT.PUT_LINE('4.1-VouString '||z_VouString||' /p_AmtZero:'||p_AmtZero||' /p_AmtCheck:'||p_AmtCheck||' /RealAmt:'||z_RealAmt||' /TaxType:'||z_TaxType);
    else
       -- 2008.03.11 by Lawrence Add InvPurposeCode,InvPurposeName
       -- 2010.07.15 by Pierre, �[��"�o���}�ߺ����B�`��E-MAIL�B��ʹq�ܡB���س��N�X�B���س��W��" 5�����
       z_VouString1 := '##'||to_char(v_InvCompCode)||v_char||to_char(z_CustId)||v_char||rtrim(z_Tel1)||v_char||z_InvNo||v_char||z_InvTitle||v_char||z_Address||v_char||z_ZipCode||v_char||z_InvAddress||v_char||x_ContName||v_char||to_char(z_InvPurposeCode)||v_char||z_InvPurposeName||
                       v_char||to_char(z_InvoiceKind)||v_char||rtrim(z_Email)||v_char||rtrim(z_ContMobile)||v_char||to_char(z_DenRecCode)||v_char||rtrim(z_DenRecName) ;
       DBMS_OUTPUT.PUT_LINE('3-VouString '||z_VouString1||' /TaxRate:'||z_TaxRate);
       if z_TaxRate>0 then
          z_TaxMoney := z_RealAmt - round((z_RealAmt/(1+(z_TaxRate/100))),0);
       else
          z_TaxMoney := 0;
       end if;
       if z_RealAmt>0 then
          z_VouString := '--N'||v_char||z_BillNo||v_char||to_char(nvl(z_Item,0))||v_char||z_TaxType||v_char||to_char(z_RealDate,'YYYYMMDD')||v_char||to_char(z_CitemCode)||v_char||z_CitemName||v_char||'1'||v_char||
                                   to_char(nvl(z_RealAmt-z_TaxMoney,0))||v_char||to_char(nvl(z_TaxRate,0))||v_char||to_char(nvl(z_TaxMoney,0))||v_char||to_char(nvl(z_RealAmt,0))||v_char||to_char(z_RealStartDate,'YYYYMMDD')||
                                   v_char||to_char(z_RealStopDate,'YYYYMMDD')||v_char||z_ClctName||v_char||z_ServiceType||v_char||to_char(z_CombCitemCode)||v_char||z_CombCitemName;
       end if;
       if p_AmtCheck=1 and z_RealAmt<0 then
          z_VouString := '--N'||v_char||z_BillNo||v_char||to_char(nvl(z_Item,0))||v_char||z_TaxType||v_char||to_char(z_RealDate,'YYYYMMDD')||v_char||to_char(z_CitemCode)||v_char||z_CitemName||v_char||'1'||v_char||
                                   to_char(nvl(z_RealAmt-z_TaxMoney,0))||v_char||to_char(nvl(z_TaxRate,0))||v_char||to_char(nvl(z_TaxMoney,0))||v_char||to_char(nvl(z_RealAmt,0))||v_char||to_char(z_RealStartDate,'YYYYMMDD')||
                                   v_char||to_char(z_RealStopDate,'YYYYMMDD')||v_char||z_ClctName||v_char||z_ServiceType||v_char||to_char(z_CombCitemCode)||v_char||z_CombCitemName;
       end if;
       if p_AmtZero=1 and z_RealAmt=0 then
          z_VouString := '--N'||v_char||z_BillNo||v_char||to_char(nvl(z_Item,0))||v_char||z_TaxType||v_char||to_char(z_RealDate,'YYYYMMDD')||v_char||to_char(z_CitemCode)||v_char||z_CitemName||v_char||'1'||v_char||
                                   to_char(nvl(z_RealAmt-z_TaxMoney,0))||v_char||to_char(nvl(z_TaxRate,0))||v_char||to_char(nvl(z_TaxMoney,0))||v_char||to_char(nvl(z_RealAmt,0))||v_char||to_char(z_RealStartDate,'YYYYMMDD')||
                                   v_char||to_char(z_RealStopDate,'YYYYMMDD')||v_char||z_ClctName||v_char||z_ServiceType||v_char||to_char(z_CombCitemCode)||v_char||z_CombCitemName;
       end if;
       DBMS_OUTPUT.PUT_LINE('4-VouString '||z_VouString||' /p_AmtZero:'||p_AmtZero||' /p_AmtCheck:'||p_AmtCheck||' /RealAmt:'||z_RealAmt||' /TaxType:'||z_TaxType);
    end if;
    if z_TaxType is not null and z_TaxType<>' ' and z_RealAmt>0 then
       DBMS_OUTPUT.PUT_LINE('z_RealAmt>0');
       -- 2003.08.25 by Lawrence
       v_ChkAccountNo := v_AccountNo;
       v_ChkMediaBillNo := z_MediaBillNo;
       v_ChkRealDate := z_RealDate;
       v_ChkCustId := z_CustId;
       z_CurrInvSeqNo := nvl(z_InvSeqNo,0);
       -- Write Master
       if z_VouString1 is not null then
          begin
             insert into Tmp004
                (VouString,Sno,CustId,MFlag) 
             values
                (z_VouString1,v_Sno,z_CustId,p_Mode);
             v_Sno := v_Sno+1;
          exception
             when others then
                p_RetMsg := 'Write Master Error Code ='||SQLCODE||' '||'Error message='||SQLERRM; 
                return -99;
          end;
       end if;
       -- Write Detail 
       begin
          insert into Tmp004
             (VouString,Sno,CustId,MFlag) 
          values
             (z_VouString,v_Sno,z_CustId,p_Mode);
          v_Sno := v_Sno+1;
       exception
          when others then
             p_RetMsg := 'Write Detail'; 
             return -99;
       end;      
       begin
          insert into Tmp005 
             (Cust_Id, Invoice_No, Title, Address, Bill_No, Date_c, C_Real, CGName, Tax_Type, ItemNo, MediaBillNo, AccountNo, ChargeTitle, ServiceType, MFlag) 
          values 
             (z_CustId, z_InvNo, z_InvTitle, z_InvAddress,  z_BillNo, z_RealDate, z_RealAmt, z_CitemName, Z_TaxType, z_Item, z_MediaBillNo, z_AccountNo, x_ContName, z_ServiceType, p_Mode);
       exception
         when others then
            p_RetMsg := 'insert into Tmp005'; 
            return -99;
       end;
       v_RealAmt := v_RealAmt + z_RealAmt;
       v_RcdCount := v_RcdCount + 1;
       x_RcdCnt1 := x_RcdCnt1 + 1;		-- 2010.07.15 by Pierre
     elsif z_TaxType is not null and z_TaxType<>' ' and p_AmtCheck=1 and z_RealAmt<0 then
       DBMS_OUTPUT.PUT_LINE('z_RealAmt<0');
       -- 2003.08.25 by Lawrence
       v_ChkAccountNo := v_AccountNo;
       v_ChkMediaBillNo := z_MediaBillNo;
       v_ChkRealDate := z_RealDate;
       v_ChkCustId := z_CustId;
       z_CurrInvSeqNo := nvl(z_InvSeqNo,0);
       -- Write Master
       if z_VouString1 is not null then
          begin
             insert into Tmp004
                (VouString,Sno,CustId,MFlag) 
             values
                (z_VouString1,v_Sno,z_CustId,p_Mode);
             v_Sno := v_Sno+1;
          exception
            when others then
               p_RetMsg := 'Write Master_1'; 
               return -99;
          end;
       end if;
       -- Write Detail 
       begin
          insert into Tmp004
             (VouString,Sno,CustId,MFlag) 
          values
             (z_VouString,v_Sno,z_CustId,p_Mode);
          v_Sno := v_Sno+1;
       exception
          when others then
             p_RetMsg := 'Write Detail_1'; 
             return -99;
       end;      
       begin
          insert into Tmp005 
             (Cust_Id, Invoice_No, Title, Address, Bill_No, Date_c, C_Real, CGName, Tax_Type, ItemNo, MediaBillNo, AccountNo, ChargeTitle, ServiceType, MFlag) 
          values 
             (z_CustId, z_InvNo, z_InvTitle, z_InvAddress,  z_BillNo, z_RealDate, z_RealAmt, z_CitemName, Z_TaxType, z_Item, z_MediaBillNo, z_AccountNo, x_ContName, z_ServiceType, p_Mode);
       exception
         when others then
            p_RetMsg := 'insert into Tmp005_1'; 
            return -99;
       end;
       v_RealAmt := v_RealAmt + z_RealAmt;
       v_RcdCount := v_RcdCount + 1;
       x_RcdCnt1 := x_RcdCnt1 + 1;		-- 2010.07.15 by Pierre
     elsif z_TaxType is not null and z_TaxType<>' ' and p_AmtZero=1 and z_RealAmt=0 then
       DBMS_OUTPUT.PUT_LINE('z_RealAmt=0');
       -- 2003.08.25 by Lawrence
       v_ChkAccountNo := v_AccountNo;
       v_ChkMediaBillNo := z_MediaBillNo;
       v_ChkRealDate := z_RealDate;
       v_ChkCustId := z_CustId;
       z_CurrInvSeqNo := nvl(z_InvSeqNo,0);
       -- Write Master
       if z_VouString1 is not null then
          begin
             insert into Tmp004
                (VouString,Sno,CustId,MFlag) 
             values
                (z_VouString1,v_Sno,z_CustId,p_Mode);
             v_Sno := v_Sno+1;
          exception
            when others then
               p_RetMsg := 'Write Master_2'; 
               return -99;
          end;
       end if;
       -- Write Detail 
       begin
          insert into Tmp004
             (VouString,Sno,CustId,MFlag) 
          values
             (z_VouString,v_Sno,z_CustId,p_Mode);
          v_Sno := v_Sno+1;
       exception
         when others then
            p_RetMsg := 'Write Detail_2'; 
            return -99;
       end;      
       begin
          insert into Tmp005 
             (Cust_Id, Invoice_No, Title, Address, Bill_No, Date_c, C_Real, CGName, Tax_Type, ItemNo, MediaBillNo, AccountNo, ChargeTitle, ServiceType, MFlag) 
          values 
             (z_CustId, z_InvNo, z_InvTitle, z_InvAddress,  z_BillNo, z_RealDate, z_RealAmt, z_CitemName, Z_TaxType, z_Item, z_MediaBillNo, z_AccountNo, x_ContName, z_ServiceType, p_Mode);
       exception
         when others then
            p_RetMsg := 'insert into Tmp005_2'; 
            return -99;
       end;
       v_RealAmt := v_RealAmt + z_RealAmt;
       v_RcdCount := v_RcdCount + 1;
       x_RcdCnt1 := x_RcdCnt1 + 1;		-- 2010.07.15 by Pierre
     end if;
     <<go_next>>
         Null;
  end loop;
  CLOSE v_DyCursor;             --close cursor
  
  --2.�B�z�L�b�᪺����
  begin
     OPEN v_DyCursor FOR v_SQL;
  exception
    when others then
       rollback;
       p_RetMsg := 'SQL3_1���~: ' || v_SQL;
       return -2;
  end;
  v_CustId:=0;
  v_ContName:='xyz';
  v_CurrCitemCode := 0;
  v_ChkCustId := -2;
  v_ChkRealDate := to_Date('19000101','YYYYMMDD');
  z_CurrInvSeqNo := 0;
  loop
   v_Error := False;
   -- 2010.07.15 by Pierre, �[��FaciSeqNo
   FETCH v_DyCursor INTO z_CustId,z_BillNo,z_RealDate,z_RealAmt,z_CitemCode,z_CitemName,z_UpdTime,z_RealStartDate,z_RealStopDate,z_Note,z_Qty,z_ClctName,z_Item,z_strInvNo,z_Address,z_AddrNo,z_AccountNo,z_MediaBillNo,z_ServiceType,z_InvoiceType,z_InvTitle,z_ContName,z_InvAddress, z_AreaName, z_CityName, z_InvPurposeCode, z_InvPurposeName, z_InvSeqNo, z_CombCitemCode, z_CombCitemName, z_CombAmount, z_CombStartDate, z_CombStopDate, z_FaciSeqNo;
   EXIT WHEN v_DyCursor%NOTFOUND;
      DBMS_OUTPUT.PUT_LINE('1. Address='||z_Address||'/AddrNo='||z_AddrNo||'/ AreaName='||z_AreaName||'/ CityName='||z_CityName);
      -- 2010.05.11 by Lawrence
      if z_CombCitemCode is not null and z_CombStartDate is not null then
         z_RealAmt := z_CombAmount;
         z_RealStartDate := z_CombStartDate;
         z_RealStopDate := z_CombStopDate;
      end if;
      -- 2008.04.17 by Lawrence AddŪ��SO138
      if nvl(z_InvSeqNo,0)>0 then
         -- 2008.08.25 by Lawrence Add p_AddrType=2 and p_AddrType=3 �B�z
         if p_AddrType=2 then
            begin
               -- 2008.04.24 by Lawrence �վ�@�U�{���L�b���B�z��z_strInvNo, z_ContName�ܼƦW��
               -- 2010.07.15 by Pierre, �[��3�����
               select InvNo, InvTitle, InvAddress, InvoiceType,ChargeAddress, ChargeAddrNo, ChargeTitle, InvPurposeCode, InvPurposeName, InvoiceKind, DenRecCode, DenRecName into z_strInvNo, z_InvTitle, z_InvAddress, z_InvoiceType, z_Address, z_AddrNo, z_ContName, z_InvPurposeCode, z_InvPurposeName, z_InvoiceKind, z_DenRecCode, z_DenRecName
                  from SO138 where InvSeqNo=z_InvSeqNo; 
            exception
              when others then
                 p_RetMsg:='so138 not found! CustId : '||z_CustId||', InvSeqNo='||z_InvSeqNo;
                 return -99;
            end;
         elsif p_AddrType=3 then
            begin
               -- 2008.04.24 by Lawrence �վ�@�U�{���L�b���B�z��z_strInvNo, z_ContName�ܼƦW��
               -- 2010.07.15 by Pierre, �[��3�����
               select InvNo, InvTitle, InvAddress, InvoiceType,MailAddress, MailAddrNo, ChargeTitle, InvPurposeCode, InvPurposeName, InvoiceKind, DenRecCode, DenRecName into z_strInvNo, z_InvTitle, z_InvAddress, z_InvoiceType, z_Address, z_AddrNo, z_ContName, z_InvPurposeCode, z_InvPurposeName, z_InvoiceKind, z_DenRecCode, z_DenRecName
                  from SO138 where InvSeqNo=z_InvSeqNo; 
            exception
              when others then
                 p_RetMsg:='so138 not found! CustId : '||z_CustId||', InvSeqNo='||z_InvSeqNo;
                 return -99;
            end;
         else
            begin
               -- 2010.07.15 by Pierre, �[��3�����
               select InvNo, InvTitle, InvAddress, InvoiceType, ChargeTitle, InvPurposeCode, InvPurposeName, InvoiceKind, DenRecCode, DenRecName into z_strInvNo, z_InvTitle, z_InvAddress, z_InvoiceType, z_ContName, z_InvPurposeCode, z_InvPurposeName, z_InvoiceKind, z_DenRecCode, z_DenRecName
                  from SO138 where InvSeqNo=z_InvSeqNo; 
            exception
              when others then
                 p_RetMsg:='so138 not found! CustId : '||z_CustId||', InvSeqNo='||z_InvSeqNo;
                 return -99;
            end;
         end if;
      else
        -- 2010.07.15 by Pierre, �Y[�o���y����]�L��, �h�ܫȤ�U�A�Ȱ򥻸����(SO002)��, ��CustId+CompCode+ServiceType������ƪ��o���}�ߺ���(InvoiceKind)�B���س��N�X(DenRecCode)�B���س��W��(DenRecName)
        begin
          select InvoiceKind, DenRecCode, DenRecName into z_InvoiceKind, z_DenRecCode, z_DenRecName from SO002 where CustId=z_CustId and CompCode=p_CompCode and ServiceType=z_ServiceType;
        exception
          when others then
            z_InvoiceKind := 0;
            z_DenRecCode := null;
            z_DenRecName  := null;
        end;
      end if;
      -- 2006.12.28 by Lawrence
      if p_AreaCity=1 then
         -- 2008.05.23 by Lawrence for Minchen and �y�pCasper
         select CityName,AreaName into z_CityName,z_AreaName from so014 where AddrNo=z_AddrNo;
         z_Address:=Trim(z_CityName)||Trim(z_AreaName)||Trim(z_Address);
         DBMS_OUTPUT.PUT_LINE('2. Address='||z_Address||'/AddrNo='||z_AddrNo||'/ AreaName='||z_AreaName||'/ CityName='||z_CityName);
      end if;
      -- 2005.05.17 by Lawrence Add and ���ܦ�
      -- check�o������, �ŦX�����͸��
      DBMS_OUTPUT.PUT_LINE('p_InvType='||p_InvType||', z_InvoiceType='||z_InvoiceType||', Count='||v_RcdCount); 
      if p_InvType!=1 then
         if z_InvoiceType<>p_InvType then
            GOTO go_next1;
         end if;
      else
         if z_InvoiceType=0 then
            GOTO go_next1;
         end if;
      end if;
      -- 2003.10.23 by Lawrence
      x_ContName:='';
    if (v_CustId != z_CustId) or (nvl(v_ContName,'xyz') != nvl(z_ContName,'xya')) then
       DBMS_OUTPUT.PUT_LINE('v_CustId:'||v_CustId||' / z_CustId:'||z_CustId||' --- v_ContName:'||v_ContName||' / z_ContName:'||z_ContName);
       v_CustId := z_CustId;
       v_ContName := z_ContName;
       -- �P�_�o���B�z
       begin
         if z_InvTitle is null then  --�ĤG����
             begin
                select custname into z_InvTitle from so001 where CustId=z_CustId and CompCode=p_CompCode; 
             exception
               when others then
                  z_InvTitle := null;   
             end;
         end if;
         -- 2003.09.26 by Lawrence�Ĥ@����
         if x_ContName is null then
            begin
               select custname into x_ContName from so001 where CustId=z_CustId and CompCode=p_CompCode;
            exception
              when others then
                 x_ContName:=null; 
            end;
         end if;
         if z_InvAddress is null then
            z_InvAddress := z_Address;
         end if;
         -- 2003.09.17 by Lawrence
         z_InvNo := z_strInvNo;
       exception
         when others then
            if z_InvTitle is null then  --�ĤG����
               begin
                  select custname into z_InvTitle from so001 where CustId=z_CustId and CompCode=p_CompCode; 
               exception
                 when others then
                    z_InvTitle := null;   
               end;
            end if;
            -- 2003.09.26 by Lawrence�Ĥ@����
            if x_ContName is null then
               begin
                  select custname into x_ContName from so001 where CustId=z_CustId and CompCode=p_CompCode;
               exception
                 when others then
                    x_ContName:=null; 
               end;
            end if;
            if z_InvAddress is null then
               z_InvAddress := z_Address;
            end if;
            z_InvNo := null;
       end;
       -- ���H��a�}���l���ϸ�
       begin
          select ZipCode into z_ZipCode from SO014 where AddrNo=z_AddrNo and CompCode=p_CompCode;
       exception
         when others then
            z_ZipCode := null;
       end;
       -- ���Ȥ�q�ܤ@
       begin
          select tel1 into z_Tel1 from SO001 where CustId=z_CustId and CompCode=p_CompCode;
       exception
         when others then
            z_Tel1 := null;
       end;
       -- *****************************************************
       -- 2010.07.15 by Pierre, ���Ȥ᪺�`��E-MAIL�B��ʹq��
       begin
         if z_FaciSeqNo is not null then
           -- �]z_FaciSeqNo���e�i�ର�Ƚs, �y��SO004�L�ӵ��]�Ƹ��, �����p�N�n�hSO001,SO002��Email, ContMobile
           begin
             select Email, ContMobile into z_Email, z_ContMobile from SO004 where SeqNo=z_FaciSeqNo;
           exception
             when others then
               select Email into z_Email from SO002 where CustId=z_CustId and CompCode=p_CompCode and ServiceType=z_ServiceType;
               select Tel3 into z_ContMobile from SO001 where CustId=z_CustId and CompCode=p_CompCode;
           end;
         else
           select Email into z_Email from SO002 where CustId=z_CustId and CompCode=p_CompCode and ServiceType=z_ServiceType;
           select Tel3 into z_ContMobile from SO001 where CustId=z_CustId and CompCode=p_CompCode;
         end if;
       exception
         when others then
           z_Email := null;
           z_ContMobile := null;
       end;
       -- *****************************************************
    else
       if z_InvTitle is null then  --�ĤG����
          begin
             select custname into z_InvTitle from so001 where CustId=z_CustId and CompCode=p_CompCode; 
          exception
            when others then
               z_InvTitle := null;   
          end;
       end if;
       -- 2003.09.26 by Lawrence�Ĥ@����
       if x_ContName is null then
          begin
             select custname into x_ContName from so001 where CustId=z_CustId and CompCode=p_CompCode;
          exception
            when others then
               x_ContName:=null; 
          end;
       end if;
       if z_InvAddress is null then
          z_InvAddress := z_Address;
       end if;
    end if;
    -- �M�w�ҵ|�O
    if v_CurrCitemCode != z_CitemCode then
       v_CurrCitemCode := z_CitemCode;
       begin
          select TaxCode into v_TaxCode from CD019 where CodeNo=v_CurrCitemCode and (ServiceType=z_ServiceType or ServiceType is null);
          if v_TaxCode = 1 then
             z_TaxType := '1';
             -- ���|�v
             begin
                select Rate1 into z_TaxRate from CD033 where CodeNo=v_TaxCode;
             exception
               when others then
                  z_TaxRate := Nvl(z_TaxRate,0);
             end;
          elsif v_TaxCode = 3 then
             z_TaxType := '3';
             z_TaxRate:=0;
          elsif v_TaxCode is null then
             z_TaxType := null;
             z_TaxRate:=0;
          else
             z_TaxType := '2';
             z_TaxRate:=0;
          end if;
       exception
         when others then
            p_RetMsg := '�䤣�즬�O���إN�X';
            return -3;        
       end;
       --���|�v�N�X�ɤ���"�O�_�һ�"
       if v_TaxCode is not null then
          begin
             select TaxFlag into v_TaxFlag from CD033 where CodeNo=v_TaxCode;
          exception
            when others then
               p_RetMsg:='cd033.TaxFlag not found! CodeNo : '||v_TaxCode;
               return -99;
          end;
       end if;
       if v_TaxFlag=0 then 
          z_TaxType := null;
          z_TaxRate :=0;
       end if;
    end if;
    -- �����ʤ���P�ɶ�
    begin
       if substr(z_UpdTime,3,1)='/' then
          z_UpdA := to_date(to_char(to_number(substr(z_UpdTime,1,2))+1911)||substr(z_UpdTime,4,2)||substr(z_UpdTime,7,2),'YYYYMMDD');
          z_UpdB := to_number(substr(z_UpdTime, 10, 2)||substr(z_UpdTime, 13, 2));
       else
          if z_UpdTime is not null then
             z_UpdA := to_date(substr(z_UpdTime,1,4)||substr(z_UpdTime,6,2)||substr(z_UpdTime,9,2),'YYYYMMDD');
          else
             z_UpdA :=Null;
          end if;
          z_UpdB := to_number(substr(z_UpdTime, 12, 2)||substr(z_UpdTime, 15, 2));
       end if;
    exception
      when others then
         p_RetMsg := '���ʤ�����~! '||z_UpdTime||' '||z_UpdA;
         return -6;              
    end;
    -- �s�W�ܵo�����ɼȦs��
    -- DBMS_OUTPUT.PUT_LINE('--- v_ContName:'||nvl(v_ContName,'xyz')||' / z_ContName:'||nvl(z_ContName,'xyz'));
    if (v_ChkCustId = z_CustId) and (v_ChkRealDate = z_RealDate) and (nvl(v_ContName,'xyz') = nvl(z_ContName,'xyz')) then
        z_VouString1:='';
        if z_TaxRate>0 then
           z_TaxMoney := z_RealAmt - round((z_RealAmt/(1+(z_TaxRate/100))),0);
        else
           z_TaxMoney := 0;
        end if;
        if z_RealAmt>0 then
           z_VouString := '--N'||v_char||z_BillNo||v_char||to_char(nvl(z_Item,0))||v_char||z_TaxType||v_char||to_char(z_RealDate,'YYYYMMDD')||v_char||to_char(z_CitemCode)||v_char||z_CitemName||v_char||'1'||v_char||
                                    to_char(nvl(z_RealAmt-z_TaxMoney,0))||v_char||to_char(nvl(z_TaxRate,0))||v_char||to_char(nvl(z_TaxMoney,0))||v_char||to_char(nvl(z_RealAmt,0))||v_char||to_char(z_RealStartDate,'YYYYMMDD')||
                                    v_char||to_char(z_RealStopDate,'YYYYMMDD')||v_char||z_ClctName||v_Char||z_ServiceType||v_char||to_char(z_CombCitemCode)||v_char||z_CombCitemName;
        end if;
        if p_AmtCheck=1 and z_RealAmt<0 then
           z_VouString := '--N'||v_char||z_BillNo||v_char||to_char(nvl(z_Item,0))||v_char||z_TaxType||v_char||to_char(z_RealDate,'YYYYMMDD')||v_char||to_char(z_CitemCode)||v_char||z_CitemName||v_char||'1'||v_char||
                                    to_char(nvl(z_RealAmt-z_TaxMoney,0))||v_char||to_char(nvl(z_TaxRate,0))||v_char||to_char(nvl(z_TaxMoney,0))||v_char||to_char(nvl(z_RealAmt,0))||v_char||to_char(z_RealStartDate,'YYYYMMDD')||
                                    v_char||to_char(z_RealStopDate,'YYYYMMDD')||v_char||z_ClctName||v_Char||z_ServiceType||v_char||to_char(z_CombCitemCode)||v_char||z_CombCitemName;
        end if;
        if p_AmtZero=1 and z_RealAmt=0 then
           z_VouString := '--N'||v_char||z_BillNo||v_char||to_char(nvl(z_Item,0))||v_char||z_TaxType||v_char||to_char(z_RealDate,'YYYYMMDD')||v_char||to_char(z_CitemCode)||v_char||z_CitemName||v_char||'1'||v_char||
                                    to_char(nvl(z_RealAmt-z_TaxMoney,0))||v_char||to_char(nvl(z_TaxRate,0))||v_char||to_char(nvl(z_TaxMoney,0))||v_char||to_char(nvl(z_RealAmt,0))||v_char||to_char(z_RealStartDate,'YYYYMMDD')||
                                    v_char||to_char(z_RealStopDate,'YYYYMMDD')||v_char||z_ClctName||v_Char||z_ServiceType||v_char||to_char(z_CombCitemCode)||v_char||z_CombCitemName;
        end if;
        DBMS_OUTPUT.PUT_LINE('2.1-VouString '||z_VouString||' /p_AmtZero:'||p_AmtZero||' /p_AmtCheck:'||p_AmtCheck||' /RealAmt:'||z_RealAmt||' /TaxType:'||z_TaxType);
    else
       -- 2008.03.11 by Lawrence Add InvPurposeCode, InvPurposeName
       -- 2010.07.15 by Pierre, �[��"�o���}�ߺ����B�`��E-MAIL�B��ʹq�ܡB���س��N�X�B���س��W��" 5�����
       z_VouString1 := '##'||to_char(v_InvCompCode)||v_char||to_char(z_CustId)||v_char||rtrim(z_Tel1)||v_char||z_InvNo||v_char||z_InvTitle||v_char||z_Address||v_char||z_ZipCode||v_char||z_InvAddress||v_char||x_ContName||v_char||to_char(z_InvPurposeCode)||v_char||z_InvPurposeName||
                       v_char||to_char(z_InvoiceKind)||v_char||rtrim(z_Email)||v_char||rtrim(z_ContMobile)||v_char||to_char(z_DenRecCode)||v_char||rtrim(z_DenRecName) ;
       DBMS_OUTPUT.PUT_LINE('1-VouString '||z_VouString1||' /TaxRate:'||z_TaxRate);
       if z_TaxRate>0 then
          z_TaxMoney := z_RealAmt - round((z_RealAmt/(1+(z_TaxRate/100))),0);
       else
          z_TaxMoney := 0;
       end if;
       if z_RealAmt>0 then
          z_VouString := '--N'||v_char||z_BillNo||v_char||to_char(nvl(z_Item,0))||v_char||z_TaxType||v_char||to_char(z_RealDate,'YYYYMMDD')||v_char||to_char(z_CitemCode)||v_char||z_CitemName||v_char||'1'||v_char||
                                   to_char(nvl(z_RealAmt-z_TaxMoney,0))||v_char||to_char(nvl(z_TaxRate,0))||v_char||to_char(nvl(z_TaxMoney,0))||v_char||to_char(nvl(z_RealAmt,0))||v_char||to_char(z_RealStartDate,'YYYYMMDD')||
                                   v_char||to_char(z_RealStopDate,'YYYYMMDD')||v_char||z_ClctName||v_Char||z_ServiceType||v_char||to_char(z_CombCitemCode)||v_char||z_CombCitemName;
       end if;
       if p_AmtCheck=1 and z_RealAmt<0 then
          z_VouString := '--N'||v_char||z_BillNo||v_char||to_char(nvl(z_Item,0))||v_char||z_TaxType||v_char||to_char(z_RealDate,'YYYYMMDD')||v_char||to_char(z_CitemCode)||v_char||z_CitemName||v_char||'1'||v_char||
                                   to_char(nvl(z_RealAmt-z_TaxMoney,0))||v_char||to_char(nvl(z_TaxRate,0))||v_char||to_char(nvl(z_TaxMoney,0))||v_char||to_char(nvl(z_RealAmt,0))||v_char||to_char(z_RealStartDate,'YYYYMMDD')||
                                   v_char||to_char(z_RealStopDate,'YYYYMMDD')||v_char||z_ClctName||v_Char||z_ServiceType||v_char||to_char(z_CombCitemCode)||v_char||z_CombCitemName;
       end if;
       if p_AmtZero=1 and z_RealAmt=0 then
          z_VouString := '--N'||v_char||z_BillNo||v_char||to_char(nvl(z_Item,0))||v_char||z_TaxType||v_char||to_char(z_RealDate,'YYYYMMDD')||v_char||to_char(z_CitemCode)||v_char||z_CitemName||v_char||'1'||v_char||
                                   to_char(nvl(z_RealAmt-z_TaxMoney,0))||v_char||to_char(nvl(z_TaxRate,0))||v_char||to_char(nvl(z_TaxMoney,0))||v_char||to_char(nvl(z_RealAmt,0))||v_char||to_char(z_RealStartDate,'YYYYMMDD')||
                                   v_char||to_char(z_RealStopDate,'YYYYMMDD')||v_char||z_ClctName||v_Char||z_ServiceType||v_char||to_char(z_CombCitemCode)||v_char||z_CombCitemName;
       end if;
       DBMS_OUTPUT.PUT_LINE('2-VouString '||z_VouString||' /p_AmtZero:'||p_AmtZero||' /p_AmtCheck:'||p_AmtCheck||' /RealAmt:'||z_RealAmt||' /TaxType:'||z_TaxType);
    end if;
    if z_TaxType is not null and z_TaxType<>' ' and z_RealAmt>0 then 
       -- 2003.08.25 by Lawrence
       v_ChkCustId := z_CustId;
       v_ChkRealDate := z_RealDate;
       v_ContName := Nvl(z_ContName,'xyz');
       z_CurrInvSeqNo := nvl(z_InvSeqNo,0);
       -- Write Master
       if z_VouString1 is not null then
          begin
             insert into Tmp004
                (VouString,Sno,CustId,MFlag) 
             values
                (z_VouString1,v_Sno,z_CustId,p_Mode);
             v_Sno := v_Sno+1;
          exception
            when others then
               p_RetMsg := '1_Write Master';
               return -99;
          end;
       end if;
       -- Write Detail 
       begin
          insert into Tmp004
             (VouString,Sno,CustId,MFlag) 
          values
             (z_VouString,v_Sno,z_CustId,p_Mode);
          v_Sno := v_Sno+1;
       exception
         when others then
            p_RetMsg := '1_Write Detail';
            return -99;
       end;
       begin
          insert into Tmp005 
             (Cust_Id, Invoice_No, Title, Address, Bill_No, Date_c, C_Real, CGName, Tax_Type, ItemNo, MediaBillNo, AccountNo, ChargeTitle, ServiceType, MFlag) 
          values 
             (z_CustId, z_InvNo, z_InvTitle, z_InvAddress,  z_BillNo, z_RealDate, z_RealAmt, z_CitemName, Z_TaxType, z_Item, z_MediaBillNo, z_AccountNo, x_ContName, z_ServiceType, p_Mode);
       exception
         when others then
            p_RetMsg := '1_insert into Tmp005';
            return -99;
       end;
       v_RealAmt := v_RealAmt + z_RealAmt;
       v_RcdCount := v_RcdCount + 1;
       x_RcdCnt2 := x_RcdCnt2 + 1;		-- 2010.07.15 by Pierre
     elsif z_TaxType is not null and z_TaxType<>' ' and p_AmtCheck=1 and z_RealAmt<0 then
       -- 2003.08.25 by Lawrence
       v_ChkCustId := z_CustId;
       v_ChkRealDate := z_RealDate;
       v_ContName := Nvl(z_ContName,'xyz');
       z_CurrInvSeqNo := nvl(z_InvSeqNo,0);
       -- Write Master
       if z_VouString1 is not null then
          begin
             insert into Tmp004
                (VouString,Sno,CustId,MFlag) 
             values
                (z_VouString1,v_Sno,z_CustId,p_Mode);
             v_Sno := v_Sno+1;
          exception
            when others then
               p_RetMsg := '1_Write Master_1';
               return -99;
          end;
       end if;
       -- Write Detail 
       begin
          insert into Tmp004
             (VouString,Sno,CustId,MFlag) 
          values
             (z_VouString,v_Sno,z_CustId,p_Mode);
          v_Sno := v_Sno+1;
       exception
         when others then
            p_RetMsg := '1_Write Detail_1';
            return -99;
       end;
       begin
          insert into Tmp005 
             (Cust_Id, Invoice_No, Title, Address, Bill_No, Date_c, C_Real, CGName, Tax_Type, ItemNo, MediaBillNo, AccountNo, ChargeTitle, ServiceType, MFlag) 
          values 
             (z_CustId, z_InvNo, z_InvTitle, z_InvAddress,  z_BillNo, z_RealDate, z_RealAmt, z_CitemName, Z_TaxType, z_Item, z_MediaBillNo, z_AccountNo, x_ContName, z_ServiceType, p_Mode);
       exception
         when others then
            p_RetMsg := '1_insert into Tmp005_1';
            return -99;
       end;
       v_RealAmt := v_RealAmt + z_RealAmt;
       v_RcdCount := v_RcdCount + 1;
       x_RcdCnt2 := x_RcdCnt2 + 1;		-- 2010.07.15 by Pierre
     elsif z_TaxType is not null and z_TaxType<>' ' and p_AmtZero=1 and z_RealAmt=0 then
       -- 2003.08.25 by Lawrence
       v_ChkCustId := z_CustId;
       v_ChkRealDate := z_RealDate;
       v_ContName := Nvl(z_ContName,'xyz');
       z_CurrInvSeqNo := nvl(z_InvSeqNo,0);
       -- Write Master
       if z_VouString1 is not null then
          begin
             insert into Tmp004
                (VouString,Sno,CustId,MFlag) 
             values
                (z_VouString1,v_Sno,z_CustId,p_Mode);
             v_Sno := v_Sno+1;
          exception
            when others then
               p_RetMsg := '1_Write Master_2';
               return -99;
          end;
       end if;
       -- Write Detail 
       begin
          insert into Tmp004
             (VouString,Sno,CustId,MFlag) 
          values
             (z_VouString,v_Sno,z_CustId,p_Mode);
          v_Sno := v_Sno+1;
       exception
         when others then
            p_RetMsg := '1_Write Detail_2';
            return -99;
       end;
       begin
          insert into Tmp005 
             (Cust_Id, Invoice_No, Title, Address, Bill_No, Date_c, C_Real, CGName, Tax_Type, ItemNo, MediaBillNo, AccountNo, ChargeTitle, ServiceType, MFlag) 
          values 
             (z_CustId, z_InvNo, z_InvTitle, z_InvAddress,  z_BillNo, z_RealDate, z_RealAmt, z_CitemName, Z_TaxType, z_Item, z_MediaBillNo, z_AccountNo, x_ContName, z_ServiceType, p_Mode);
       exception
         when others then
            p_RetMsg := '1_insert into Tmp005_2';
            return -99;
       end;
       v_RealAmt := v_RealAmt + z_RealAmt;
       v_RcdCount := v_RcdCount + 1;
       x_RcdCnt2 := x_RcdCnt2 + 1;		-- 2010.07.15 by Pierre
     end if;
     <<go_next1>>
         Null;
  end loop;
  CLOSE v_DyCursor;	--close cursor
  p_RetMsg := '���ͧ���, �@��O'||to_char(trunc(86400*(sysdate-v_StartExecTime)))||'��, �B�z����='||to_char(v_RcdCount)||', �B�z���B='||to_char(v_RealAmt);
  DBMS_OUTPUT.PUT_LINE('���b�ᵧ��='||x_RcdCnt1||', �L�b�ᵧ��='||x_RcdCnt2);		-- 2010.07.15 by Pierre
  commit;
  return v_RcdCount;
EXCEPTION
  WHEN OTHERS THEN 
     rollback;
     CLOSE v_DyCursor;             --close cursor
     DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
     p_RetMsg := SQLERRM;
     return -99;
END;
/
