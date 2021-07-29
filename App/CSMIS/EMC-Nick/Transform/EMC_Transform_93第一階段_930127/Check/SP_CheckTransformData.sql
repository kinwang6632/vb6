CREATE OR REPLACE PROCEDURE SP_CHECKTRANSFORMDATA(p_Flag NUMBER)
as
/*
--@C:\Transform\Check\SP_CheckTransformData.sql


exec SP_CheckTransformData(1);



  計算資料檔中的代碼是否皆有存在代碼檔中或是對照檔中
  注意: 1. 此程式於正式轉資料程式"前"執行
      	2. 做完此程式後, 才可以drop掉資料檔的index
      	3. 假設所有代碼對照表命名為: CDxxx_MAP

  1. 計算依據為Code_Check1, 結構應有下列欄位:
	CodeTable varchar2(30)
	DataTable varchar2(30)
	DataColumn varchar2(30)
  DataChName  varchar2(100)

  2. 計算結果儲存於Code_Check2, 結構應有下列欄位:
	CodeTable  varchar2(30)
	DataTable  varchar2(30)
	DataColumn varchar2(30)
  DataChColumn varchar2(100)
	CodeValue  varchar2(30)
  CodeName   varchar2(100)
	Description varchar2(500)
  RcdCnt  number(8) DEFAULT 0
  CheckTime  Date


  By: Jackal
  Date: 2003.12.24
	2003.12.25 Pierre: 改成若有該資料檔及該資料欄位, 才做檢查; 若無, 則不做檢查, 以免程式出錯
	2003.12.26 Jackal: 1.多Code_Check1.DataChName,Code_Check2.DataChColumn及CodeName三各欄位
	                   2.程式多檢核中文名稱是否有不同的
  2004.01.07 Jackal: 加傳一個參數
                     p_Flag=1-->檢查 [資料檔] 中的代碼是否皆有存在 [代碼檔] 中
                     p_Flag=2-->檢查 [資料檔] 中的代碼是否皆有存在 [對照檔] 中
                     p_Flag=3-->檢查 [對照檔] 中的代碼是否皆有存在 [新代碼檔] 中
                     p_Flag=4-->檢查p_Flag=1,2
                     p_Flag=5-->檢查p_Flag=1,2,3

  2004.01.09 Jackal: 加檢查新舊代碼有無重複
                     p_Flag=6-->檢查 [對照檔] 內的 [舊代碼] 是否有重複
                     p_Flag=7-->檢查 [新代碼檔] 內的 [新代碼] 是否有重複

*/
  v_SQL1 varchar2(2000);
  v_SQL2 varchar2(2000);
  v_SQL3 varchar2(2000);
  v_CodeNo number;
  v_CodeName varchar2(100);
  v_RcdCnt number;
  v_NewCodeNo number;
  v_OldName varchar2(30);
  v_NewName varchar2(30);
  v_TempCnt number :=0;
  v_Description varchar2(500);
  v_CodeTable varchar2(30);
  v_RetMsg varchar2(2000);

  TYPE CurTyp IS REF CURSOR;  --自訂cursor型態, 供新版dynamic SQL
  v_DyCursor1 CurTyp;
  v_DyCursor2 CurTyp;
  v_DyCursor3 CurTyp;

  v_Cnt1 number;

  cursor cc1 is
	select * from Code_Check1;
	
  cursor cc2 is
	select * from Code_Check3;	

begin
  -- 先清除結果檔
  delete from Code_Check2;

  -- 開始計算
  for cr1 in cc1 loop	-- loop Code_Check1 取到的每一筆資料
    -- 2003.12.25 若有該資料檔及該資料欄位, 才做檢查
    v_cnt1 := 0;
    IF cr1.DataChName IS NULL THEN
      IF (p_Flag=1) OR (p_Flag=2) OR (p_Flag=4) OR (p_Flag=5) THEN
        select count(*) into v_cnt1 from user_tab_columns where
         	table_name=upper(cr1.DataTable) and column_name=upper(cr1.DataColumn);

        if v_cnt1 = 0 then
         	v_RetMsg := '99: 沒有這些Table或欄位';
        	insert into Code_Check2 (CodeTable, DataTable, DataColumn, DataChColumn,CodeValue,CodeName, Description,RcdCnt,CheckTime) values
             (cr1.CodeTable, cr1.DataTable, cr1.DataColumn,cr1.DataChName, '','',v_RetMsg,0,SysDate);
          COMMIT;
          goto GO_NEXT;
        end if;
      END IF;


      IF (p_Flag=1) OR (p_Flag=4) OR (p_Flag=5) THEN
        --檢查資料檔的代碼是否皆存在代碼檔中
        v_SQL1 := 'select distinct '||cr1.DataColumn||' from '||cr1.DataTable
                  ||' WHERE '||cr1.DataColumn||' IS NOT NULL AND '||cr1.DataColumn||'<>0 '
                  ||' minus (select CodeNo from '||cr1.CodeTable||')';
        begin
          OPEN v_DyCursor1 FOR v_SQL1;
        exception
          when others then
         	v_RetMsg := 'SQL1錯誤: ' || v_SQL1;
        	insert into Code_Check2 (CodeTable, DataTable, DataColumn, DataChColumn,CodeValue,CodeName, Description,RcdCnt,CheckTime) values
             (cr1.CodeTable, cr1.DataTable, cr1.DataColumn,cr1.DataChName, '','',v_RetMsg,0,SysDate);
          COMMIT;
        end;


        loop
          FETCH v_DyCursor1 INTO v_CodeNo;
          EXIT WHEN v_DyCursor1%NOTFOUND;
          --DBMS_OUTPUT.PUT_LINE(cr1.CodeTable||', '||cr1.DataTable||', '||cr1.DataColumn||', '||v_CodeNo);

          if v_CodeNo is not null then
            v_Description := '1: [資料檔] 含有 [代碼檔] 沒有的值';

          	-- 取資料數 v_RcdCnt
          	v_SQL2 := 'select count(*) from '||cr1.DataTable||' where '||cr1.DataColumn||'='||v_CodeNo;
          	begin
    	        OPEN v_DyCursor2 FOR v_SQL2;
          	exception
    	      when others then
    	        v_RetMsg := 'SQL2錯誤: ' || v_SQL2;
            	insert into Code_Check2 (CodeTable, DataTable, DataColumn, DataChColumn,CodeValue,CodeName, Description,RcdCnt,CheckTime) values
                 (cr1.CodeTable, cr1.DataTable, cr1.DataColumn,cr1.DataChName, '','',v_RetMsg,0,SysDate);
               COMMIT;
           	end;
          	v_RcdCnt := 0;
    	      FETCH v_DyCursor2 INTO v_RcdCnt;
          	CLOSE v_DyCursor2;


          	insert into Code_Check2 (CodeTable, DataTable, DataColumn, DataChColumn,CodeValue,CodeName, Description,RcdCnt,CheckTime) values
    	         (cr1.CodeTable, cr1.DataTable, cr1.DataColumn,'', v_CodeNo,'',v_Description,v_RcdCnt,SysDate);
          end if;


        end loop;

        CLOSE v_DyCursor1;
      END IF;
--********************************************
      IF (p_Flag=2) OR (p_Flag=4) OR (p_Flag=5) THEN
        --檢查資料檔的代碼是否皆存在對照檔中
        v_SQL1 := 'select distinct '||cr1.DataColumn||' from '||cr1.DataTable
                  ||' WHERE '||cr1.DataColumn||' IS NOT NULL AND '||cr1.DataColumn||'<>0 '
                  ||' minus (select OldCode from '||cr1.CodeTable||'_MAP)';
        begin
          OPEN v_DyCursor1 FOR v_SQL1;
        exception
          when others then
         	v_RetMsg := 'SQL1錯誤: ' || v_SQL1;
        	insert into Code_Check2 (CodeTable, DataTable, DataColumn, DataChColumn,CodeValue,CodeName, Description,RcdCnt,CheckTime) values
             (cr1.CodeTable||'_MAP', cr1.DataTable, cr1.DataColumn,cr1.DataChName, '','',v_RetMsg,0,SysDate);
          COMMIT;
        end;


        loop
          FETCH v_DyCursor1 INTO v_CodeNo;
          EXIT WHEN v_DyCursor1%NOTFOUND;
          --DBMS_OUTPUT.PUT_LINE(cr1.CodeTable||', '||cr1.DataTable||', '||cr1.DataColumn||', '||v_CodeNo);

          if v_CodeNo is not null then
            v_Description := '2: [資料檔] 含有 [對照檔] 沒有的值';

          	-- 取資料數 v_RcdCnt
          	v_SQL2 := 'select count(*) from '||cr1.DataTable||' where '||cr1.DataColumn||'='||v_CodeNo;
          	begin
    	        OPEN v_DyCursor2 FOR v_SQL2;
          	exception
    	      when others then
    	        v_RetMsg := 'SQL2錯誤: ' || v_SQL2;
            	insert into Code_Check2 (CodeTable, DataTable, DataColumn, DataChColumn,CodeValue,CodeName, Description,RcdCnt,CheckTime) values
                 (cr1.CodeTable||'_MAP', cr1.DataTable, cr1.DataColumn,cr1.DataChName, '','',v_RetMsg,0,SysDate);
               COMMIT;
           	end;
          	v_RcdCnt := 0;
    	      FETCH v_DyCursor2 INTO v_RcdCnt;
          	CLOSE v_DyCursor2;


          	insert into Code_Check2 (CodeTable, DataTable, DataColumn, DataChColumn,CodeValue,CodeName, Description,RcdCnt,CheckTime) values
    	         (cr1.CodeTable||'_MAP', cr1.DataTable, cr1.DataColumn,'', v_CodeNo,'',v_Description,v_RcdCnt,SysDate);
          end if;

        end loop;

        CLOSE v_DyCursor1;
      END IF;

--含有中文名稱的檢查
--/////////////////////////////////////////////////////////////////
--/////////////////////////////////////////////////////////////////
    ELSE
      IF (p_Flag=1) OR (p_Flag=2) OR (p_Flag=4) OR (p_Flag=5) THEN
        select count(*) into v_cnt1 from user_tab_columns
          where	table_name=upper(cr1.DataTable) and
                column_name=upper(cr1.DataColumn);

        if v_cnt1 = 0 then
         	v_RetMsg := '99: 沒有這些Table或欄位';
        	insert into Code_Check2 (CodeTable, DataTable, DataColumn, DataChColumn,CodeValue,CodeName, Description,RcdCnt,CheckTime) values
             (cr1.CodeTable, cr1.DataTable, cr1.DataColumn,cr1.DataChName, '','',v_RetMsg,0,SysDate);
          COMMIT;
          goto GO_NEXT;
        end if;


        select count(*) into v_cnt1 from user_tab_columns
          where	table_name=upper(cr1.DataTable) and
                column_name=upper(cr1.DataChName);

        if v_cnt1 = 0 then
         	v_RetMsg := '99: 沒有這些Table或欄位';
        	insert into Code_Check2 (CodeTable, DataTable, DataColumn, DataChColumn,CodeValue,CodeName, Description,RcdCnt,CheckTime) values
             (cr1.CodeTable, cr1.DataTable, cr1.DataColumn,cr1.DataChName, '','',v_RetMsg,0,SysDate);
          COMMIT;
          goto GO_NEXT;
        end if;
      END IF;


      IF (p_Flag=1) OR (p_Flag=4) OR (p_Flag=5) THEN
        --檢查資料檔的代碼是否皆存在代碼檔中
        v_SQL1 := 'select distinct '||cr1.DataColumn||','||cr1.DataChName||' from '||cr1.DataTable
                  ||' WHERE '||cr1.DataColumn||' IS NOT NULL AND '||cr1.DataColumn||'<>0 '
                  ||' minus (select CodeNo,DESCRIPTION from '||cr1.CodeTable||')';
        begin
          OPEN v_DyCursor1 FOR v_SQL1;
        exception
          when others then
         	v_RetMsg := 'SQL1錯誤: ' || v_SQL1;
        	insert into Code_Check2 (CodeTable, DataTable, DataColumn, DataChColumn,CodeValue,CodeName, Description,RcdCnt,CheckTime) values
             (cr1.CodeTable, cr1.DataTable, cr1.DataColumn,cr1.DataChName, '','',v_RetMsg,0,SysDate);
          COMMIT;
        end;


        loop
          FETCH v_DyCursor1 INTO v_CodeNo,v_CodeName;
          EXIT WHEN v_DyCursor1%NOTFOUND;
          --DBMS_OUTPUT.PUT_LINE(cr1.CodeTable||', '||cr1.DataTable||', '||cr1.DataColumn||', '||v_CodeNo);

          if v_CodeNo is not null then
            v_Description := '1: [資料檔] 含有 [代碼檔] 沒有的值';

          	-- 取資料數 v_RcdCnt
          	v_SQL2 := 'select count(*) from '||cr1.DataTable||
                             ' where '||cr1.DataColumn||'='||v_CodeNo||
                             ' AND '||cr1.DataChName||'='''||v_CodeName||'''';
          	begin
    	        OPEN v_DyCursor2 FOR v_SQL2;
          	exception
    	      when others then
    	        v_RetMsg := 'SQL2錯誤: ' || v_SQL2;
            	insert into Code_Check2 (CodeTable, DataTable, DataColumn, DataChColumn,CodeValue,CodeName, Description,RcdCnt,CheckTime) values
                 (cr1.CodeTable, cr1.DataTable, cr1.DataColumn,cr1.DataChName, v_CodeNo,v_CodeName,v_RetMsg,0,SysDate);
               COMMIT;
           	end;
          	v_RcdCnt := 0;
    	      FETCH v_DyCursor2 INTO v_RcdCnt;
          	CLOSE v_DyCursor2;


          	insert into Code_Check2 (CodeTable, DataTable, DataColumn, DataChColumn,CodeValue,CodeName, Description,RcdCnt,CheckTime) values
    	         (cr1.CodeTable, cr1.DataTable, cr1.DataColumn,cr1.DataChName, v_CodeNo,v_CodeName,v_Description,v_RcdCnt,SysDate);
          end if;

        end loop;
        CLOSE v_DyCursor1;
      END IF;
--********************************************
      IF (p_Flag=2) OR (p_Flag=4) OR (p_Flag=5) THEN
        --檢查資料檔的代碼是否皆存在對照檔中
        v_SQL1 := 'select distinct '||cr1.DataColumn||' from '||cr1.DataTable
                  ||' WHERE '||cr1.DataColumn||' IS NOT NULL AND '||cr1.DataColumn||'<>0 '
                  ||' minus (select OldCode from '||cr1.CodeTable||'_MAP)';
        begin
          OPEN v_DyCursor1 FOR v_SQL1;
        exception
          when others then
         	v_RetMsg := 'SQL1錯誤: ' || v_SQL1;
        	insert into Code_Check2 (CodeTable, DataTable, DataColumn, DataChColumn,CodeValue,CodeName, Description,RcdCnt,CheckTime) values
             (cr1.CodeTable||'_MAP', cr1.DataTable, cr1.DataColumn,cr1.DataChName, '','',v_RetMsg,0,SysDate);
          COMMIT;
        end;


        loop
          FETCH v_DyCursor1 INTO v_CodeNo;
          EXIT WHEN v_DyCursor1%NOTFOUND;
          --DBMS_OUTPUT.PUT_LINE(cr1.CodeTable||', '||cr1.DataTable||', '||cr1.DataColumn||', '||v_CodeNo);

          if v_CodeNo is not null then
            v_Description := '2: [資料檔] 含有 [對照檔] 沒有的值';

          	-- 取資料數 v_RcdCnt
          	v_SQL2 := 'select count(*) from '||cr1.DataTable||
                             ' where '||cr1.DataColumn||'='||v_CodeNo;
          	begin
    	        OPEN v_DyCursor2 FOR v_SQL2;
          	exception
    	      when others then
    	        v_RetMsg := 'SQL2錯誤: ' || v_SQL2;
            	insert into Code_Check2 (CodeTable, DataTable, DataColumn, DataChColumn,CodeValue,CodeName, Description,RcdCnt,CheckTime) values
                 (cr1.CodeTable||'_MAP', cr1.DataTable, cr1.DataColumn,cr1.DataChName, v_CodeNo,'',v_RetMsg,0,SysDate);
               COMMIT;
           	end;
          	v_RcdCnt := 0;
    	      FETCH v_DyCursor2 INTO v_RcdCnt;
          	CLOSE v_DyCursor2;


          	insert into Code_Check2 (CodeTable, DataTable, DataColumn, DataChColumn,CodeValue,CodeName, Description,RcdCnt,CheckTime) values
    	         (cr1.CodeTable||'_MAP', cr1.DataTable, cr1.DataColumn,cr1.DataChName, v_CodeNo,'',v_Description,v_RcdCnt,SysDate);
          end if;

        end loop;
        CLOSE v_DyCursor1;
      END IF;
/*JACKAL930105
      IF (p_Flag=2) OR (p_Flag=4) OR (p_Flag=5) THEN
        --檢查資料檔的代碼是否皆存在對照檔中
        v_SQL1 := 'select distinct '||cr1.DataColumn||','||cr1.DataChName||' from '||cr1.DataTable
                  ||' WHERE '||cr1.DataColumn||' IS NOT NULL AND '||cr1.DataColumn||'<>0 '
                  ||' minus (select OldCode,OldName from '||cr1.CodeTable||'_MAP)';
        begin
          OPEN v_DyCursor1 FOR v_SQL1;
        exception
          when others then
         	v_RetMsg := 'SQL1錯誤: ' || v_SQL1;
        	insert into Code_Check2 (CodeTable, DataTable, DataColumn, DataChColumn,CodeValue,CodeName, Description,RcdCnt,CheckTime) values
             (cr1.CodeTable||'_MAP', cr1.DataTable, cr1.DataColumn,cr1.DataChName, '','',v_RetMsg,0,SysDate);
          COMMIT;
        end;


        loop
          FETCH v_DyCursor1 INTO v_CodeNo,v_CodeName;
          EXIT WHEN v_DyCursor1%NOTFOUND;
          --DBMS_OUTPUT.PUT_LINE(cr1.CodeTable||', '||cr1.DataTable||', '||cr1.DataColumn||', '||v_CodeNo);

          if v_CodeNo is not null then
            v_Description := '2: [資料檔] 含有 [對照檔] 沒有的值';

          	-- 取資料數 v_RcdCnt
          	v_SQL2 := 'select count(*) from '||cr1.DataTable||
                             ' where '||cr1.DataColumn||'='||v_CodeNo||
                             ' AND '||cr1.DataChName||'='''||v_CodeName||'''';
          	begin
    	        OPEN v_DyCursor2 FOR v_SQL2;
          	exception
    	      when others then
    	        v_RetMsg := 'SQL2錯誤: ' || v_SQL2;
            	insert into Code_Check2 (CodeTable, DataTable, DataColumn, DataChColumn,CodeValue,CodeName, Description,RcdCnt,CheckTime) values
                 (cr1.CodeTable||'_MAP', cr1.DataTable, cr1.DataColumn,cr1.DataChName, v_CodeNo,v_CodeName,v_RetMsg,0,SysDate);
               COMMIT;
           	end;
          	v_RcdCnt := 0;
    	      FETCH v_DyCursor2 INTO v_RcdCnt;
          	CLOSE v_DyCursor2;


          	insert into Code_Check2 (CodeTable, DataTable, DataColumn, DataChColumn,CodeValue,CodeName, Description,RcdCnt,CheckTime) values
    	         (cr1.CodeTable||'_MAP', cr1.DataTable, cr1.DataColumn,cr1.DataChName, v_CodeNo,v_CodeName,v_Description,v_RcdCnt,SysDate);
          end if;

        end loop;
        CLOSE v_DyCursor1;
      END IF;
*/
    END IF;

    <<GO_NEXT>>
    NULL;					-- 2003.12.25

  end loop;
  
  
--***************************************************************************
--***************************************************************************
  -- 開始計算
  IF (p_Flag=3) OR (p_Flag=5) OR (p_Flag=6) OR (p_Flag=7) THEN
    for cr2 in cc2 loop	-- loop Code_Check3 取到的每一筆資料
      v_cnt1 := 0;
      select count(*) into v_cnt1 from user_tab_columns
        where	table_name=upper(cr2.DataTable) and
              column_name=upper(cr2.DataColumn);

      if v_cnt1 = 0 then
       	v_RetMsg := '99: 沒有這些Table或欄位';
      	insert into Code_Check2 (CodeTable, DataTable, DataColumn, DataChColumn,CodeValue,CodeName, Description,RcdCnt,CheckTime) values
           (cr2.CodeTable, cr2.DataTable, cr2.DataColumn,cr2.DataChName, '','',v_RetMsg,0,SysDate);
        COMMIT;
        goto GO_NEXT2;
      end if;


      select count(*) into v_cnt1 from user_tab_columns
        where	table_name=upper(cr2.DataTable) and
              column_name=upper(cr2.DataChName);

      if v_cnt1 = 0 then
       	v_RetMsg := '99: 沒有這些Table或欄位';
      	insert into Code_Check2 (CodeTable, DataTable, DataColumn, DataChColumn,CodeValue,CodeName, Description,RcdCnt,CheckTime) values
           (cr2.CodeTable, cr2.DataTable, cr2.DataColumn,cr2.DataChName, '','',v_RetMsg,0,SysDate);
        COMMIT;
        goto GO_NEXT2;
      end if;


      IF (p_Flag=3) OR (p_Flag=5) THEN
        --檢查對照檔的代碼是否皆存在新代碼檔中
        v_SQL1 := 'select distinct NewCode,NewName from '||cr2.DataTable
                  ||' WHERE '||cr2.DataColumn||' IS NOT NULL AND '||cr2.DataColumn||'<>0 '
                  ||' minus (select CodeNo,DESCRIPTION from '||cr2.CodeTable||')';
        begin
          OPEN v_DyCursor1 FOR v_SQL1;
        exception
          when others then
         	v_RetMsg := 'SQL1錯誤: ' || v_SQL1;
        	insert into Code_Check2 (CodeTable, DataTable, DataColumn, DataChColumn,CodeValue,CodeName, Description,RcdCnt,CheckTime) values
             (cr2.CodeTable, cr2.DataTable, cr2.DataColumn,cr2.DataChName, '','',v_RetMsg,0,SysDate);
          COMMIT;
        end;


        loop
          FETCH v_DyCursor1 INTO v_CodeNo,v_CodeName;
          EXIT WHEN v_DyCursor1%NOTFOUND;
          --DBMS_OUTPUT.PUT_LINE(cr1.CodeTable||', '||cr1.DataTable||', '||cr1.DataColumn||', '||v_CodeNo);

          if v_CodeNo is not null then
            v_Description := '3: [對照檔] 含有 [新代碼檔] 沒有的值';

          	-- 取資料數 v_RcdCnt
          	v_SQL2 := 'select count(*) from '||cr2.DataTable||
                             ' where '||cr2.DataColumn||'='||v_CodeNo||
                             ' AND '||cr2.DataChName||'='''||v_CodeName||'''';
          	begin
    	        OPEN v_DyCursor2 FOR v_SQL2;
          	exception
    	      when others then
    	        v_RetMsg := 'SQL2錯誤: ' || v_SQL2;
            	insert into Code_Check2 (CodeTable, DataTable, DataColumn, DataChColumn,CodeValue,CodeName, Description,RcdCnt,CheckTime) values
                 (cr2.CodeTable, cr2.DataTable, cr2.DataColumn,cr2.DataChName, v_CodeNo,v_CodeName,v_RetMsg,0,SysDate);
               COMMIT;
           	end;
          	v_RcdCnt := 0;
    	      FETCH v_DyCursor2 INTO v_RcdCnt;
          	CLOSE v_DyCursor2;


          	insert into Code_Check2 (CodeTable, DataTable, DataColumn, DataChColumn,CodeValue,CodeName, Description,RcdCnt,CheckTime) values
    	         (cr2.CodeTable, cr2.DataTable, cr2.DataColumn,cr2.DataChName, v_CodeNo,v_CodeName,v_Description,v_RcdCnt,SysDate);
          end if;

        end loop;
        CLOSE v_DyCursor1;
      END IF;
      
      
--*******************************
      IF p_Flag=6 THEN
        --檢查對照檔的舊代碼是否有重複
        v_SQL1 := 'SELECT OLDCODE,COUNT(*) FROM '||cr2.DataTable
                  ||' GROUP BY OLDCODE';
        begin
          OPEN v_DyCursor1 FOR v_SQL1;
        exception
          when others then
         	v_RetMsg := 'SQL1錯誤: ' || v_SQL1;
        	insert into Code_Check2 (CodeTable, DataTable, DataColumn, DataChColumn,CodeValue,CodeName, Description,RcdCnt,CheckTime) values
             (cr2.CodeTable, cr2.DataTable, cr2.DataColumn,cr2.DataChName, '','',v_RetMsg,0,SysDate);
          COMMIT;
        end;


        loop
          v_RcdCnt := 0;
          FETCH v_DyCursor1 INTO v_CodeNo,v_RcdCnt;
          EXIT WHEN v_DyCursor1%NOTFOUND;

          if v_RcdCnt>=2 then
            v_Description := '6: [對照檔] 內的 [舊代碼檔]-['||v_CodeNo||'] 重複';


          	insert into Code_Check2 (CodeTable, DataTable, DataColumn, DataChColumn,CodeValue,CodeName, Description,RcdCnt,CheckTime) values
    	         (cr2.CodeTable, cr2.DataTable, cr2.DataColumn,cr2.DataChName, v_CodeNo,v_CodeName,v_Description,v_RcdCnt,SysDate);
          end if;

        end loop;
        CLOSE v_DyCursor1;
      END IF;
      
      
--*******************************
      IF p_Flag=7 THEN
        --檢查新代碼檔的新代碼是否有重複
        v_SQL1 := 'SELECT CODENO,COUNT(*) FROM '||cr2.CodeTable
                  ||' GROUP BY CODENO';
        begin
          OPEN v_DyCursor1 FOR v_SQL1;
        exception
          when others then
         	v_RetMsg := 'SQL1錯誤: ' || v_SQL1;
        	insert into Code_Check2 (CodeTable, DataTable, DataColumn, DataChColumn,CodeValue,CodeName, Description,RcdCnt,CheckTime) values
             (cr2.CodeTable, cr2.DataTable, cr2.DataColumn,cr2.DataChName, '','',v_RetMsg,0,SysDate);
          COMMIT;
        end;


        loop
          v_RcdCnt := 0;
          FETCH v_DyCursor1 INTO v_CodeNo,v_RcdCnt;
          EXIT WHEN v_DyCursor1%NOTFOUND;

          if v_RcdCnt>=2 then
            v_Description := '7: [新代碼檔] 內的 [新代碼]-['||v_CodeNo||'] 重複';


          	insert into Code_Check2 (CodeTable, DataTable, DataColumn, DataChColumn,CodeValue,CodeName, Description,RcdCnt,CheckTime) values
    	         (cr2.CodeTable, cr2.DataTable, cr2.DataColumn,cr2.DataChName, v_CodeNo,v_CodeName,v_Description,v_RcdCnt,SysDate);
          end if;

        end loop;
        CLOSE v_DyCursor1;
      END IF;




      <<GO_NEXT2>>
      NULL;			
    end loop;
  END IF;

  commit;
exception
  when others then
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    v_RetMsg := '-1:其他錯誤 SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM;
  	insert into Code_Check2 (CodeTable, DataTable, DataColumn, CodeValue, Description,RcdCnt,CheckTime) values
       ('', '', '', v_CodeNo,v_RetMsg,0,SysDate);
    COMMIT;
end;
/