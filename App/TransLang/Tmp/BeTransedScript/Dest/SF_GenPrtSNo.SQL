 
CREATE OR REPLACE Function SF_GENPRTSNO
   ( p_YM IN VARCHAR2,
     p_YM1 IN VARCHAR2,
     p_YM2 IN VARCHAR2,
     p_NOE IN NUMBER,
     p_ORDER IN VARCHAR2,
     p_GENOVERDUE IN NUMBER,
     p_GENPRCUST IN NUMBER,
     p_RANGE IN NUMBER,
     p_CompSQL IN VARCHAR2,
     p_CitemSQL IN VARCHAR2,
     p_ClassSQL IN VARCHAR2,
     p_ClctEnSQL IN VARCHAR2,
     p_ServSQL IN VARCHAR2,
     p_ClctAreaSQL IN VARCHAR2,
     p_MduIdSQL IN VARCHAR2,
     p_StrtSQL IN VARCHAR2,
     p_RETMSG OUT VARCHAR2,
     p_STARTSNO OUT VARCHAR2,
     p_ENDSNO OUT VARCHAR2,
     p_RCDCNT OUT VARCHAR2
   ) RETURN NUMBER
AS
   v_Where1 VARCHAR2(2000);
   v_Where2 VARCHAR2(2000);
   v_SQL VARCHAR2(3000);
   v_ORDER VARCHAR2(2000);
   v_NOWNO NUMBER;          -- 目前序號
   v_CUSTID NUMBER;
   v_CompCode number;	    -- 公司別代碼
   v_ChargType NUMBER;      -- 收費屬性
   v_CustStatus NUMBER;     -- 客戶狀態
   v_AddrNo NUMBER;         -- 地址編號
   v_MduID VARCHAR2(8);     -- 大樓編號
   v_SQLx VARCHAR2(200);
   v_Flag NUMBER;           -- 迴圈裡的旗標
   v_Para14 NUMBER;         -- 地址依據(收費地址或裝機地址)
   v_ClctEn VARCHAR2(10);   -- 收費人員代號
   v_StrtCode NUMBER;       -- 街道代碼
   v_ClctName VARCHAR2(20); -- 收費人員姓名
   v_TP VARCHAR2(1);        -- 由p_ORDER取出一個字元
 
   z_ROWID VARCHAR2(30);    -- 資料列編號
   z_CUSTID NUMBER;         -- 目前客編
   z_ClctEn VARCHAR2(10);   -- 收費人員代號
   z_ADDRNO NUMBER;         -- 地址編號
   z_StrtCode NUMBER;       -- 街道代碼
 
   v_Cursor NUMBER;
   v_Cursor2 NUMBER;
   v_RetCode NUMBER;
   I NUMBER;
   v_Count number;
   v_Flag2 number;
   v_ServCode varchar2(12);
   v_ClctAreaCode varchar2(12);
   v_AreaCode varchar2(12);
   v_StartExecTime date;
 
   v_STARTSNO number;
   v_ENDSNO number;
   v_RCDCNT number;
 
BEGIN
   -- 檢查參數
   IF (p_YM is NULL) or (p_YM1 is NULL) or (p_YM2 is NULL) or (p_NOE is NULL) or
   (p_ORDER is NULL) or (p_GENOVERDUE is NULL) or (p_GENPRCUST is NULL) or (p_RANGE is NULL) THEN
      p_RetMsg := 'Wrong parameter.';
      RETURN -1;
   END IF;
 
   If (p_Noe < 0) Or (p_Noe > 4) Or (p_GenOverdue < 0) Or (p_GenOverdue > 1) Or (p_RANGE < 0) Or (p_RANGE > 4) Then
      p_RetMsg := 'Wrong parameter.';
      RETURN -1;
   END IF;
 
   -- 至收費參數檔取地址依據Para14
   begin
     SELECT Para14 INTO v_PARA14 FROM SO043;
   exception
     when others then
      p_RetMsg := 'No data or too many data in SO043.';
      RETURN -4;
   end;
 
   -- 設定各變數初值
   v_StartExecTime := sysdate;		-- 開始執行時間
   v_SQL := 'SELECT A.Rowid, A.CustId, A.ClctEn, A.AddrNo, A.PrtSNo, B.Noe1 FROM SO033 A, SO014 B WHERE A.AddrNo = B.AddrNo And A.ClctYM >= ' || p_YM1 || ' And A.ClctYM <= ' || p_YM2 || ' ';
   v_Where1 := null;
   v_Where2 := null;
 
   -- 組合門號條件:
   -- 若門號=1: where1 = where1 + ' and B.Noe1=1'
   If p_Noe = 1 Then
      v_Where1 := v_Where1 || ' and B.Noe1=1';
   END IF;
 
   -- 若門號=2: where1 = where1 + ' and B.Noe1=2 '
   If p_Noe = 2 Then
      v_Where1 := v_Where1 || ' and B.Noe1=2';
   END IF;
 
   -- 若有公司別條件:
   -- where2 = where2 + ' and A.CompCode ' + <p_CompSQL>
   If p_CompSQL is not null Then
      v_Where2 := v_Where2 || ' and A.CompCode ' || p_CompSQL;
   END IF;
 
   -- 若有收費項目條件:
   -- where2 = where2 + ' and A.CitemCode ' + <p_CitemSQL>
   If p_CitemSQL is not null Then
      v_Where2 := v_Where2 || ' and A.CitemCode ' || p_CitemSQL;
   END IF;
 
   -- 若有客戶類別條件:
   -- where2 = where2 + ' and A.ClassCode ' + <p_ClassSQL>
   If p_ClassSQL Is Not Null Then
      v_Where2 := v_Where2 || ' and A.ClassCode ' || p_ClassSQL;
   END IF;
 
   -- 若有收費人員條件:
   -- where2 = where2 + ' and A.ClctEn ' + <p_ClctEnSQL>
   If p_ClctEnSQL Is Not Null Then
      v_Where2 := v_Where2 || ' and A.ClctEn' || p_ClctEnSQL || ' ';
   END IF;
 
   -- 若有服務區條件:
   -- where2 = where2 + ' and A.ServCode ' + <p_ServSQL>
   If p_ServSQL Is Not Null Then
      v_Where2 := v_Where2 || ' and A.ServCode ' || p_ServSQL;
   END IF;
 
   -- 若有收費區條件:
   -- where2 = where2 + ' and A.ClctAreaCode ' + <p_ClctAreaSQL>
   If p_ClctAreaSQL Is Not Null Then
      v_Where2 := v_Where2 || ' and A.ClctAreaCode ' || p_ClctAreaSQL;
   END IF;
 
   -- 若有大樓條件:
   -- where2 = where2 + ' and A.MduId ' + <p_MduSQL>
   If p_MduIdSQL Is Not Null Then
      v_Where2 := v_Where2 || ' and A.MduId ' || p_MduIdSQL;
   END IF;
 
   -- 若有街道條件:
   -- where2 = where2 + ' and A.StrtCode ' + <p_StrtSQL>
   If p_StrtSQL Is Not Null Then
      v_Where2 := v_Where2 || ' and A.StrtCode ' || p_StrtSQL;
   END IF;
 
/* -- 錯誤: '若重編前期欠費客戶'指前期若有欠費,則本期將重編,但前期欠費資料不會納入重編
   -- 若重編前期欠費客戶,
   -- 則: where2 = where2 + ' and (A.ClctYM is not null and A.ClctYM != <p_YM> and A.UCCode is not null or A.ClctYM = <p_YM> and A.PrtSNo is null and A.RealDate is null)'
   If p_GenOverdue = 1 Then
      v_Where2 := v_Where2 || ' and (A.ClctYM is not null and A.ClctYM!=' || p_YM || ' and A.UCCode is not null or A.ClctYM=' || p_YM || ' and A.PrtSNo is null and A.RealDate is null)';
   Else
      -- 否則: where2 = where2 + ' and A.ClctYM=<p_YM> and A.PrtSNo is null and A.RealDate is null'
      v_Where2 := v_Where2 || ' and A.ClctYM=' || p_YM || ' and A.PrtSNo is null and A.RealDate is null';
   END IF;
*/
   v_Where2 := v_Where2 || ' and A.ClctYM=' || p_YM || ' and A.PrtSNo is null and A.RealDate is null';
 
   -- 決定order方式:
   -- 1. loop3次:
   FOR I IN 1 .. 3 LOOP
      -- 取第i  byte 字元
      -- 若字元為'1', 則: order = order + ',A.ClctEn'
      v_TP := SUBSTR(p_ORDER, I, 1);
      IF v_TP = '1' THEN
         v_ORDER := v_ORDER || ', A.ClctEn';
      -- 若字元為'2', 則:
      ELSIF v_TP = '2' THEN
         -- 若門號=3, 則:  order = order + ',B.StrtCode, B.Noe1'
         If p_Noe = 3 Then
            v_ORDER := v_ORDER || ',B.StrtCode, B.Noe1';
         -- 若門號=4, 則:  order = order + ',B.StrtCode, B.Noe1 DESC'
         ELSIF p_Noe = 4 THEN
            v_ORDER := v_ORDER || ',B.StrtCode, B.Noe1 DESC';
         END IF;
         -- order = order + ',B.AddrSort'
         v_ORDER := v_ORDER || ',B.AddrSort';
      -- 若字元為'3', 則: order = order + ',A.ServCode'
      ELSIF v_TP = '3' THEN
         v_ORDER := v_ORDER || ',A.ServCode';
      END IF;
   END LOOP;
 
   v_ORDER := ltrim(v_ORDER, ',');			-- 去掉最左邊之,
   v_ORDER := ' order by ' || v_ORDER || ',A.CustId';	-- 將同一客戶放在一起
dbms_output.put_line(v_order);
 
   -- SQL = SQL + where1 + where2 + ‘ ‘ + order
   v_SQL := v_SQL || v_WHERE1 || ' ' || v_WHERE2 || ' ' || v_ORDER;
 
   -- Lock印單序號檔, 若error, 則return -2
   BEGIN
      lock table SO069 in exclusive mode nowait;
   EXCEPTION
      WHEN OTHERS THEN
         ROLLBACK;
         p_RetMsg := 'Print no. table may be locked by others.';
         RETURN -2;
   END;
 
   -- 取該年月之<目前序號>
   BEGIN
      SELECT PrtSNo INTO v_NOWNO FROM SO069 WHERE YM = TO_NUMBER(p_YM);
   EXCEPTION
      -- 若無該筆資料, 則<目前序號>=0, 且新增一筆:
      -- 出單年月=該年月, 目前已用序號=0
      WHEN OTHERS THEN
         v_NOWNO := 0;
         INSERT INTO SO069 (YM, PRTSNO) VALUES (TO_NUMBER(p_YM), 0);
   END;
 
   v_STARTSNO := v_NOWNO + 1;
   p_STARTSNO := to_char(v_STARTSNO);
 
   -- 執行該查詢, 並做相關處理, 參考流程圖
   z_CUSTID := 0;       -- 目前客編
   v_CustStatus := 0;   -- 客戶狀態
   v_ChargType := 0;    -- 收費屬性
   v_AddrNo := 0;       -- 地址編號
   v_Flag := 0;         -- 迴圈裡的旗標
   v_MduID := '';       -- 大樓編號
   v_ClctEn := '';      -- 收費人員代號
   v_CUSTID := 0;
 
   BEGIN
      v_Cursor := DBMS_SQL.Open_cursor;
      DBMS_SQL.Parse(v_Cursor, v_SQL, DBMS_SQL.V7);
      -- 定義取出的欄位
      DBMS_SQL.Define_Column(v_Cursor, 1, z_ROWID, 30);
      DBMS_SQL.Define_Column(v_Cursor, 2, z_CUSTID);
      DBMS_SQL.Define_Column(v_Cursor, 3, z_ClctEn, 10);
      DBMS_SQL.Define_Column(v_Cursor, 4, z_ADDRNO);
      -- DBMS_SQL.Define_Column(v_Cursor, 5, z_StrtCode);
      v_RetCode := DBMS_SQL.Execute(v_Cursor);
   EXCEPTION
      WHEN OTHERS THEN
	 DBMS_SQL.Close_cursor(v_Cursor);
         ROLLBACK;
         p_RetMsg := 'SQL錯誤: ' || v_SQL;
         RETURN -3;
   END;
 
   v_Cursor2 := DBMS_SQL.Open_cursor;
 
   LOOP                      -- 取到的每一筆資料
      IF DBMS_SQL.Fetch_Rows(v_Cursor) <= 0 THEN -- 取得一筆資料
         EXIT;                   -- 已經到了 EOF 則脫離迴圈
      END IF;
 
      -- 取各欄位值
      DBMS_SQL.Column_Value(v_Cursor, 1, z_ROWID);
      DBMS_SQL.Column_Value(v_Cursor, 2, z_CUSTID);
      DBMS_SQL.Column_Value(v_Cursor, 3, z_ClctEn);
      DBMS_SQL.Column_Value(v_Cursor, 4, z_ADDRNO);
      -- DBMS_SQL.Column_Value(v_Cursor, 5, z_StrtCode);
 
      -- 如果不同一客戶
      If z_CUSTID != v_CUSTID Then
         v_CUSTID := z_CUSTID;		-- <目前客編> = <該客戶編號>
         v_Flag := 0;
 
         -- 根據該筆客戶編號至客戶資料檔(SO001)取客戶狀態、收費屬性、xx地址編號，並存入<客戶狀態>
         If v_PARA14 = 1 Then
            SELECT CustStatusCode, ChargeType, ChargeAddrNo, MduID, CompCode
              Into v_CustStatus, v_ChargType, v_AddrNO, v_MDUID, v_CompCode
              From SO001
             WHERE CustID = z_CUSTID;
         Else
            SELECT CustStatusCode, ChargeType, InstAddrNo, MduID, CompCode
              Into v_CustStatus, v_ChargType, v_AddrNO, v_MDUID, v_CompCode
              From SO001
             WHERE CustID = z_CUSTID;
         END IF;
 
	 -- 若不重編前期欠費客戶, 則檢查該戶是否有前期欠費
	 v_Flag2 := 1;
	 If p_GenOverdue = 0 then
	    select count(*) into v_Count from SO033 where 
		CustId=z_CUSTID and ClctYM<to_number(p_YM) and UCCode is not null;
	    if v_Count > 0 then
		v_Flag2 := 0;
	    end if;
	 end if;
 
	 -- if p_GenPRCust=1 and (v_CustStatus=1 or v_CustStatus=6) or 
	 --    p_GenPRCust=0 and v_CustStatus=1 then 同以下一行
	 if v_Flag2=1 and (v_CustStatus=1 or p_GenPRCust=1 and v_CustStatus=6) then
            -- 若不是以下狀況, 則 FLAG = 1, 否則仍為0
            -- Range=1(一般戶)且該筆大樓編號有值(or <收費屬性>!=1)，或
            -- Range=2(大樓戶)且該筆大樓編號無值(or <收費屬性>=1)，或
            -- Range=3(大樓個收)且<收費屬性>!=2，或
            -- Range=4(大樓統收)且<收費屬性>!=3
            -- IF NOT ((p_RANGE=1 AND v_MduId is not null) OR (p_RANGE=2 AND v_MDUID IS NULL)
            --    OR (p_RANGE=3 AND v_ChargType<>2) OR (p_RANGE=4 AND v_ChargType<>3)) THEN
 
            IF NOT ((p_RANGE=1 AND v_ChargType<>1) OR (p_RANGE=2 AND v_ChargType=1)
               OR (p_RANGE=3 AND v_ChargType<>2) OR (p_RANGE=4 AND v_ChargType<>3)) THEN
               v_FLAG := 1;
 
	       -- (4)<目前序號>=取次一序號
	       v_NOWNO := v_NOWNO + 1;
 
	       -- (6)依地址編號取該客戶目前所屬的收費人員代號／姓名，街道代碼,服務區代碼,收費區代碼,行政區代碼
	       SELECT ClctEn, ClctName, StrtCode, ServCode, ClctAreaCode, AreaCode 
		INTO v_ClctEn, v_ClctName, v_StrtCode, v_ServCode, v_ClctAreaCode, v_AreaCode 
		FROM SO014 WHERE CompCode=v_CompCode and AddrNo=v_AddrNO;
            END IF;
         END IF;
      END IF;
 
      -- 是同一客戶且 FLAG = 1
      If v_FLAG = 1 Then
         -- (5)SQLx = "update SO033 set PrtSNo='<目前序號>'"
	 -- 印單序號 = 'YYYYMM999999'
         v_SQLx := 'UPDATE SO033 SET PrtSNo='||Chr(39)||p_YM||LTRIM(TO_CHAR(v_NowNo, '099999'))||Chr(39);
 
	 --** Q: 若已於過入前做過調單動作,則(7.1)會將其調回預設值,是否恰當?
         -- (7.1)若<ClctEn>!=<該筆收費人員代號>. 則:
         --      SQLx = SQLx+".ClctEn='<該筆收費人員代號>'"
         If v_ClctEn <> z_ClctEn Then
            v_SQLx := v_SQLx||',ClctEn='||CHR(39)||v_ClctEn||CHR(39)||',ClctName='||Chr(39)||v_ClctName||Chr(39);
         END IF;
 
	 --** Q: 
         -- (7.2)若<地址編號>!=<該筆地址編號>,
	 --      則更新該筆之街道代碼,服務區代碼,收費區代碼,行政區代碼,大樓編號
         If v_AddrNO <> z_AddrNO Then
            v_SQLx := v_SQLx||',AddrNo='||v_AddrNo||',StrtCode='||v_StrtCode||
		',ServCode='||chr(39)||v_ServCode||chr(39)||
		',ClctAreaCode='||chr(39)||v_ClctAreaCode||chr(39)||
		',AreaCode='||chr(39)||v_AreaCode||chr(39)||
		',MduId='||chr(39)||v_MduId||chr(39);
         END IF;
 
         -- (8)SQLx=SQLx+" where rowid='<該筆rowid>'"
         --    執行該筆應收資料(執行SQLx)
         v_SQLx := v_SQLx || ' WHERE ROWID = ' || Chr(39) || z_ROWID || Chr(39) ;
 
         BEGIN
            DBMS_SQL.Parse(v_Cursor2, v_SQLx, DBMS_SQL.V7);
            v_RetCode := DBMS_SQL.Execute(v_Cursor2);
         EXCEPTION
            WHEN OTHERS THEN
              DBMS_SQL.Close_cursor(v_Cursor2);
	      DBMS_SQL.Close_cursor(v_Cursor);
              ROLLBACK;
              p_RetMsg := 'SQL錯誤: ' || v_SQLx;
              RETURN -3;
         END;
      END IF;
   END LOOP;
 
   DBMS_SQL.Close_cursor(v_Cursor2);
   DBMS_SQL.Close_cursor(v_Cursor);
 
   IF v_NowNo >= v_StartSNo THEN
      v_EndSno := v_NowNo;
      p_EndSNo := to_char(v_NowNo);
      p_RcdCnt := to_char(v_EndSno - v_StartSno + 1);
      -- 15.填回<目前序號>於印單序號檔, SQL:
      --    UPDATE SO069 SET PrtSNo= <目前序號> where YM=該年月;
      UPDATE SO069 SET PrtSNO = v_NOWNO WHERE YM = TO_NUMBER(p_YM);
   ELSE
      p_StartSNo := '0';
      p_EndSNo := '0';
      p_RcdCnt := '0';
   END IF;
 
   COMMIT;
   p_RetMsg := 'Execution OK, it takes'||to_char(trunc(86400*(sysdate-v_StartExecTime)))||'second';
   RETURN 0;
 
EXCEPTION
   WHEN OTHERS THEN
      DBMS_SQL.Close_cursor(v_Cursor);
      DBMS_SQL.Close_cursor(v_Cursor2);
      ROLLBACK;
      p_RetMsg := 'Other error.';
      RETURN -99;
END;
/
