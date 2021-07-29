1.請將所有的資料檔案放於  C:\Transform\Best_Transform
  檔案包括:
  .. SetupData1.sql
  .. SetupData2.sql
  .. FillCompCode.SQL
  .. FillCompName.SQL
  .. FillServiceType.SQL


2.建立基本資料一: 內容來自北視CATV資料區, 屬於經營區相關檔案內容
  @C:\Transform\Best_Transform\SetupData1.SQL
  
  TableSpace 暫定為 Best , 請依實際 TableSpace 命名

3.建立基本資料二: 內容來自LSC CM/DVS資料區, 屬於服務別相關檔案內容
  @C:\Transform\Best_Transform\SetupData2.SQL

  TableSpace 暫定為 Lsc , 請依實際 TableSpace 命名

4.調整公司別 = 6
  @C:\Transform\Best_Transform\FillCompCode.SQL

  若改為其他公司代碼,v_Option 設為 1 ,變更 v_CompCode 的設定值即可

5.調整公司名稱 = 北視光速通訊
  @C:\Transform\Best_Transform\FillCompName.SQL

  若改為其他公司名稱,v_Option 設為 1 ,變更 v_CompName 的設定值即可

6.調整服務別 = I
  @C:\Transform\Best_Transform\FillServiceType.SQL

  若改為其他服務別,v_Option 設為 1 ,變更 v_SourceType 的設定值即可

7.於FTools 產生Seqence Script files,然後執行此 Script files.


