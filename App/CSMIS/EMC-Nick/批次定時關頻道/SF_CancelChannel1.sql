CREATE OR REPLACE FUNCTION SF_CANCELCHANNEL1(p_CompCode number, p_ProdIdList varchar2,
  p_ChcodeList varchar2,p_ExecTime varchar2, p_RetMsg out varchar2) return number
  as

/*
--@c:\SF_CancelChannel1
variable nn number
variable msg varchar2(100)
exec :nn := SF_CancelChannel1(7, '12,13', '200304010010', :msg);
print nn
print msg

  �妸�@�~: �s�W�@�����W�D���O, �H�����Y�@�ɶ����, �i�����o�ǥI�O�W�D
  �ت�: �i�w��Ҧ��ϥΤ���STB, �����լݦܬY�@������W�D
  �ɦW: SF_CancelChannel1.sql

  �`�N: ���禡�ت��]�p����92.03.31, EMC�|�n�D�ڭ̱N�Ҧ�STB�լݦܸӤ骺�����W�D�ե�������,
	���ڭ̪������\��å����ѸլݦܬY�骺�\��(�u���լݩT�w�Ѽƪ��\��), �����ɭ�, EMC
	�����t�Υx���n����@��.  ���禡�]�i�H���ڭ̨�N�����Y���W�D.

  IN p_CompCode number: ���q�O 
     p_ProdIdList varchar2: Nagra product id. list, �H','�Ϲj, ���n���ť�
     p_ExecTime varchar2: ���W�D�ɶ�, ���ɶ�10�X�r��, 'YYYYMMDDHH24MI', ���t��
  OUT p_RetMsg varchar2: �B�z���G�ۦ�
  Return NUMBER�G�B�z���G�N�X
	>=0:���槹��
	   -1:INSERT�ɵo�Ϳ��~
                     -2:Delete�ɵo�Ϳ��~
                     -3:INSERT��SOAC0202�ɵo�Ϳ��~�@�@�@�@
                   -99:��L���~

  �y�{:
	Loop�Ҧ��ϥΤ���STB
		    �s�W�@�����W�D���O(B2)�ܦ@�θ�ưϪ�Send_Nagra��, ���p�U:
				High_Level_Cmd_ID: 'B2'
			                   ICC_NO: �ӳ]�Ƥ�ICC�d��
			                   Notes: �����W�D��Product id. list, �H','�Ϲj
                                                                           CMD_STATUS:'W'
                                                                           OPERATOR:'NIGHT-RUN'
                                                                           UPDTIME:SYSDATE
                                                                           SEQNO:�q sequence object S_SendNagra_SeqNo ��select �XNextValue
                                                                           COMPCODE:�ӫȤ���ݤ����q�O�N�X
                                                                           ProcessingDate:�B�z���

                                         �R���Ȥ��W�D�����(SO005)�Q���W�D�����

                                         �����@����Ʀ�STB/ICC�]�w�O����(SOAC0202),�����Ȥ�Q���W�D�����

  �w�q: 
	. �Ҧ��ϥΤ���STB: �]�Ƹ����(SO004)����~�W�N�X(FaciCode), �����ܫ~�W�N�X
	  ��(CD022)���ѦҸ�(RefNo)��3(�YSTB)��, �B�ӳ]�ƪ��������Ū�(PRDate is null)

  Date:2003/01/23
  By: Morris
*/ 
 
I		                number;
v_Index	                                   number;
v_ChcodeList                             varchar2(30);
v_ChcodeCnt	                 number;
v_FirstChcode                            varchar2(20);
s_SQL                                          Varchar2(4000);
V_HIGHT_LEVEL_CMD_ID    Varchar2(5);
V_CMD_STATUS                     Varchar2(2);
V_OPERATOR                           Varchar2(15);
V_count                                       Number(3):=0;
V_ModeType                              Varchar2(5);
V_Time                                         Date;
V_ChName                                  Varchar2(20);
TYPE Varchar2Ary IS TABLE OF varchar2(20) INDEX BY BINARY_INTEGER;
ChcodeName Varchar2Ary;	-- array for Chcode

Cursor n_cancel is
 select distinct (a.smartcardno) smartcardno,a.compcode compcode,
                  a.custid CustID,b.cvtsno Cvtsno
  from so004 a,so005 b,cd022 c,cd024 d,so002 e
  where a.custid=b.custid
    and a.smartcardno=b.smartcardno   
    and a.facicode=c.codeno
    and c.refno='3'
    and a.prdate is null
    and b.chcode=d.codeno
    and a.compcode=p_CompCode
    and a.custid=e.custid
    and e.custstatuscode in (1,2,5,6)
 order by a.smartcardno;
        
BEGIN
  
  V_HIGHT_LEVEL_CMD_ID:='B2';
  V_CMD_STATUS:='W';
  V_OPERATOR:='NIGHT_RUN';
  V_ModeType:='B2';
  V_Time:=sysdate;  
        
    --����Chcode�r���Chcode�}�C��
     v_ChcodeCnt:= 0;
  if p_ChcodeList is not null then
     v_ChcodeList := p_ChcodeList;
     v_Index := 1;
     I := 1;
    while v_ChcodeList is not null loop
     v_Index := instr(v_ChcodeList, ',');
      if v_Index > 0 then
	begin
      ChcodeName(I) := ltrim(rtrim(substr(v_ChcodeList, 1, v_Index-1)));
	  v_ChcodeList := substr(v_ChcodeList, v_Index+1);
      I:=I+1;     
    exception
	  when others then
	    v_ChcodeList := null;
	end;   
      else
        ChcodeName(I) := rtrim(ltrim(v_ChcodeList));
        v_ChcodeList := null;     
      end if; 
           
    end loop;
    
    v_ChcodeCnt := I;
    v_FirstChcode := ChcodeName(1); -- �O��Ĥ@��Chcode
    
  end if; 
       
    for n_record in n_cancel loop
    
      begin               
        Insert into Com. Send_Nagra(HIGH_LEVEL_CMD_ID,ICC_NO,STB_NO,NOTES,CMD_STATUS,
               OPERATOR,UPDTIME,SEQNO,COMPCODE,RESENTTIMES,ProcessingDate)
                        Values(V_HIGHT_LEVEL_CMD_ID,n_record.smartcardno,n_record.Cvtsno,
                               p_ProdIdList,V_CMD_STATUS,V_OPERATOR,V_Time,Com.S_SENDNAGRA_SEQNO.NEXTVAL,n_record.compcode,'',to_date(p_ExecTime,'YYYYMMDDHH24MI'));
         
                    V_count:= V_count+1; 

        exception
          When OTHERS Then
           DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
           p_RetMsg := 'INSERT�ɵo�Ϳ��~: '||SQLERRM;
          return -1;                                
        end;
   
      for I in 1..v_ChcodeCnt loop 
      
      Select Description into V_ChName  from Com.CD024 Where CodeNo=ChcodeName(I);
  
     begin
        Delete  Com.SO005 Where CustID=n_record.CustID
                                                and Smartcardno=n_record.smartcardno 
                                                and Chcode=ChcodeName(I);
        exception
          When Others Then
           DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
           p_RetMsg := 'Delete�ɵo�Ϳ��~: '||SQLERRM;
           return -2;                                             
        end;
      
      begin
             Insert into  Com.SOAC0202(CompCode,CustId,STBSNo,
                            SmartCardNo,ModeType,ChCode,ChName,
                           UpdTime,UpdEn)
                     Values(n_record.Compcode,n_record.CustID,n_record.Cvtsno,
                            n_record.Smartcardno,V_ModeType,ChcodeName(i),V_ChName,
                        V_Time,'�t�Φ۰�');               
          exception
          When Others Then
           DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
           p_RetMsg := 'INSERT��SOAC0202�ɵo�Ϳ��~: '||SQLERRM;
           return -3;                                             
        end; 
              
      end loop;
      
    end loop;
          
    
     commit;
      
      p_RetMsg:='�۰������W�D���槹��,�@���� ' || V_count ||' �������W�D����� ';
      return 0;

 EXCEPTION
  WHEN OTHERS THEN
    DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
    p_RetMsg := '��L���~: '||SQLERRM;
    rollback;
    return -99;

END ; 

/
--modify by morris 2003/01/30
--Version 1.1

