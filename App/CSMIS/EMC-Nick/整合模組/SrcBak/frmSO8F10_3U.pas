unit frmSO8F10_3U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Grids, DBGrids, DB ,ADODB ,IniFiles;


const
  SELECT_MODE = 1;
  IUD_MODE = 2;
    
type
  TfrmSO8F10_3 = class(TForm)
    Panel1: TPanel;
    cobCompInfo: TComboBox;
    StaticText6: TStaticText;
    btnClose: TButton;
    btnQuery: TButton;
    Memo1: TMemo;
    procedure FormShow(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure cobCompInfoCloseUp(Sender: TObject);
    procedure btnQueryClick(Sender: TObject);
  private
    { Private declarations }
    G_DbConnAry : array of TADOConnection;
    G_AdoQueryAry : array of TADOQuery;
    procedure getCompDBInfo(sI_CompName : String);
    function connectToDB : WideString;
    function getAreaDataNo(sI_CompName : String) : String;
    function getActiveQuery(sI_AreaDataNo : String):TADOQuery;
    function runSQL(nI_Mode:Integer; sI_AreaDataNo, sI_SQL:String):TADOQuery;
    procedure insertToCD067A(I_AdoQry: TADOQuery);
    procedure insertToCD067B(I_AdoQry: TADOQuery);
    procedure insertToCD067C(I_AdoQry: TADOQuery);
    function getComboBoxCompCode(sI_CompName : String) : String;
    function getCompDBIniInfo : WideString;
  public
    { Public declarations }
    sG_CompCode,sG_CompName,sG_Alias,sG_UserID,sG_Password,sG_DataAreaNo : String;
    G_CompInfoStrList,G_DbAreaNameStrList : TStringList;
    sG_ExePath : String;
    procedure SynchronizationToSo(sI_CompName : String);

  end;

var
  frmSO8F10_3: TfrmSO8F10_3;

implementation

uses frmMainMenuU, UCommonU, dtmMain4U, Ustru;

{$R *.dfm}

procedure TfrmSO8F10_3.FormShow(Sender: TObject);
begin
    self.Caption := frmMainMenu.setCaption('SO8F10','[二階代碼資料同步]');

    //回復原狀態
    TUCommonFun.setDefaultCursor;
end;

procedure TfrmSO8F10_3.btnCloseClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmSO8F10_3.FormCreate(Sender: TObject);
var
    ii : Integer;
    sL_CompName,sL_CompCode,sL_ErrMsg : String;
begin
    //執行等待狀態
    TUCommonFun.setWaitingCursor;
    
    //將加密的ini檔案解密
    sL_ErrMsg := frmMainMenu.TransTmpIniFile(SYS_COMPDBINFO_INI_FILE_NAME);
    Memo1.Clear;

    if sL_ErrMsg = '' then
    begin
      //取得Ini連資料庫資料
      getCompDBIniInfo;

      //與資料庫連線
      sL_ErrMsg := connectToDB;
      if sL_ErrMsg <> '' then
      begin
        dtmMain4.cdsDBError.First;
        while not dtmMain4.cdsDBError.Eof do
        begin
          sL_CompCode := dtmMain4.cdsDBError.FieldByName('CompCode').AsString;
          if dtmMain4.cdsIniInfo.Locate('CompCode',sL_CompCode,[]) then
            dtmMain4.cdsIniInfo.Delete;

          Memo1.Lines.Add(dtmMain4.cdsDBError.FieldByName('ErrorMsg').AsString);

          dtmMain4.cdsDBError.Next;
        end;
      end;

      //將有連線的資料顯示在畫面上
      dtmMain4.cdsIniInfo.First;
      while not dtmMain4.cdsIniInfo.Eof do
      begin
        sL_CompCode := dtmMain4.cdsIniInfo.FieldByName('CompCode').AsString;
        if sL_CompCode='0' then
          sL_CompName := ALL_COMPCODE + '-所有'
        else
          sL_CompName := dtmMain4.cdsIniInfo.FieldByName('CompName').AsString;

        cobCompInfo.Items.Add(sL_CompName);
        dtmMain4.cdsIniInfo.Next;
      end;

      cobCompInfo.ItemIndex := 0;
      sG_CompCode := ALL_COMPCODE;
      btnQuery.Enabled := true;
    end
    else
    begin
      btnQuery.Enabled := false;
      Memo1.Lines.Add(sL_ErrMsg);
    end;
end;

procedure TfrmSO8F10_3.FormClose(Sender: TObject;
  var Action: TCloseAction);
var
    sL_IniFileName : String;
begin
    G_CompInfoStrList.Free;
    G_DbAreaNameStrList.Free;

    sL_IniFileName := sG_ExePath + '\' + TMP_SYS_INI_FILE_NAME;
    DeleteFile(sL_IniFileName);
end;

procedure TfrmSO8F10_3.cobCompInfoCloseUp(Sender: TObject);
begin
    sG_CompName := Trim(cobCompInfo.Text);
end;

procedure TfrmSO8F10_3.btnQueryClick(Sender: TObject);
var
    nL_SelectedIndex : Integer;
begin
    //執行等待狀態
    TUCommonFun.setWaitingCursor;

    sG_CompCode := '';
    nL_SelectedIndex := cobCompInfo.ItemIndex;

    if nL_SelectedIndex = 0  then
    begin
      sG_CompName := cobCompInfo.Text;
      dtmMain4.cdsIniInfo.First;
      while not dtmMain4.cdsIniInfo.Eof do
      begin
        sG_CompName := dtmMain4.cdsIniInfo.FieldByName('CompName').AsString;
        getCompDBInfo(sG_CompName);
        sG_CompCode := getComboBoxCompCode(sG_CompName);
        if sG_CompCode <> '0' then
        begin
          SynchronizationToSo(sG_CompName);
        end;
        dtmMain4.cdsIniInfo.Next;
      end;
    end
    else
    begin
      SynchronizationToSo(sG_CompName);
    end;

    MessageDlg('同步完成', mtInformation, [mbok],0);
    //回復原狀態
    TUCommonFun.setDefaultCursor;
end;

procedure TfrmSO8F10_3.getCompDBInfo(sI_CompName: String);
begin
    with dtmMain4.cdsIniInfo do
    begin
      if Locate('COMPNAME', sI_CompName, []) then
      begin
        sG_DataAreaNo := FieldByName('DataAreaNo').AsString;
        sG_CompName := FieldByName('CompName').AsString;
        sG_CompCode := FieldByName('CompCode').AsString;
        sG_Alias := FieldByName('Alias').AsString;
        sG_UserID := FieldByName('UserID').AsString;
        sG_Password := FieldByName('Password').AsString;
      end;
    end;
end;

procedure TfrmSO8F10_3.SynchronizationToSo(sI_CompName : String);
var
    sL_SQL,sL_InsertSQL : String;
    L_AdoQry : TADOQuery;

begin
    getCompDBInfo(sG_CompName);

    sL_SQL := 'DELETE CD067A';
    runSQL(IUD_MODE,sG_DataAreaNo,sL_SQL);
    sL_SQL := 'DELETE CD067B';
    runSQL(IUD_MODE,sG_DataAreaNo,sL_SQL);
    sL_SQL := 'DELETE CD067C WHERE COMPCODE=' + sG_CompCode +
              ' OR COMPCODE=' + ALL_COMPCODE;
    runSQL(IUD_MODE,sG_DataAreaNo,sL_SQL);

    sL_SQL := 'SELECT * FROM CD067A';
    L_AdoQry := runSQL(SELECT_MODE,'0',sL_SQL);
    //複製至實體AdoQuery
    dtmMain4.adoQryCD067A.Clone(L_AdoQry,ltUnspecified);
    if dtmMain4.adoQryCD067A.RecordCount > 0 then
      insertToCD067A(dtmMain4.adoQryCD067A);



    sL_SQL := 'SELECT * FROM CD067B';
    L_AdoQry := runSQL(SELECT_MODE,'0',sL_SQL);
    //複製至實體AdoQuery
    dtmMain4.adoQryCD067B.Clone(L_AdoQry,ltUnspecified);
    if dtmMain4.adoQryCD067B.RecordCount > 0 then
      insertToCD067B(dtmMain4.adoQryCD067B);


    sL_SQL := 'SELECT * FROM CD067C WHERE COMPCODE=' + sG_CompCode +
              ' OR COMPCODE=' + ALL_COMPCODE;
    L_AdoQry := runSQL(SELECT_MODE,'0',sL_SQL);
    //複製至實體AdoQuery
    dtmMain4.adoQryCD067C.Clone(L_AdoQry,ltUnspecified);
    if dtmMain4.adoQryCD067C.RecordCount > 0 then
      insertToCD067C(dtmMain4.adoQryCD067C);


end;

function TfrmSO8F10_3.connectToDB : WideString;
var
    ii,nL_DataAreaNo : Integer;
    sL_DbConnStr,sL_ErrMsg : String;
    sL_DbPassword, sL_DbAlias, sL_DbUserName : String;
begin
    try
      sL_ErrMsg := '';
      if not dtmMain4.cdsDBError.Active then
        dtmMain4.cdsDBError.CreateDataSet;
      dtmMain4.cdsDBError.EmptyDataSet;
      
      setlength(G_DbConnAry, dtmMain4.cdsIniInfo.RecordCount);
      setlength(G_AdoQueryAry, dtmMain4.cdsIniInfo.RecordCount);

      dtmMain4.cdsIniInfo.First;
      while not dtmMain4.cdsIniInfo.Eof do
      begin
        sG_CompCode := dtmMain4.cdsIniInfo.FieldByName('CompCode').AsString;

        nL_DataAreaNo := dtmMain4.cdsIniInfo.FieldByName('DataAreaNo').AsInteger;
        sL_DbPassword := dtmMain4.cdsIniInfo.FieldByName('Password').AsString;
        sL_DbAlias := dtmMain4.cdsIniInfo.FieldByName('Alias').AsString;
        sL_DbUserName := dtmMain4.cdsIniInfo.FieldByName('UserID').AsString;

//    sL_DbConnStr := 'Provider=OraOLEDB.Oracle.1;Password=' + sL_DbPassword +';Persist Security Info=True;User ID='+sL_DbUserName+';Data Source='+sL_DbAlias; //=> 適用於my nb
        sL_DbConnStr := 'Provider=MSDAORA.1;Password=' + sL_DbPassword + ';User ID='+sL_DbUserName+';Data Source='+sL_DbAlias+';Persist Security Info=True'; //適用於其他 PC
        try
          G_DbConnAry[nL_DataAreaNo] := TADOConnection.Create(nil);
          G_DbConnAry[nL_DataAreaNo].ConnectionString := sL_DbConnStr;
          G_DbConnAry[nL_DataAreaNo].LoginPrompt := false;
          G_DbConnAry[nL_DataAreaNo].Connected := true;
          if G_DbConnAry[nL_DataAreaNo].Connected then
          begin
            try
              G_AdoQueryAry[nL_DataAreaNo] := TADOQuery.Create(nil);
              G_AdoQueryAry[nL_DataAreaNo].Connection := G_DbConnAry[nL_DataAreaNo];

            except
              sL_ErrMsg := '資料庫連線失敗, Alias=<'+sL_DbAlias+'>, DB User=<'+sL_DbUserName +'>';

              dtmMain4.cdsDBError.Append;
              dtmMain4.cdsDBError.FieldByName('ErrorMsg').AsString := sL_ErrMsg;
              dtmMain4.cdsDBError.FieldByName('CompCode').AsString := sG_CompCode;
              dtmMain4.cdsDBError.post;
            end;
          end;
          
        except
          sL_ErrMsg := '資料庫連線失敗, Alias=<'+sL_DbAlias+'>, DB User=<'+sL_DbUserName +'>';

          dtmMain4.cdsDBError.Append;
          dtmMain4.cdsDBError.FieldByName('ErrorMsg').AsString := sL_ErrMsg;
          dtmMain4.cdsDBError.FieldByName('CompCode').AsString := sG_CompCode;
          dtmMain4.cdsDBError.post;
        end;

        dtmMain4.cdsIniInfo.Next;
      end;
    except
      sL_ErrMsg := '資料庫連線失敗, Alias=<'+sL_DbAlias+'>, DB User=<'+sL_DbUserName +'>';

      dtmMain4.cdsDBError.Append;
      dtmMain4.cdsDBError.FieldByName('ErrorMsg').AsString := sL_ErrMsg;
      dtmMain4.cdsDBError.FieldByName('CompCode').AsString := sG_CompCode;
      dtmMain4.cdsDBError.post;
    end;
    result := sL_ErrMsg;

end;

function TfrmSO8F10_3.getAreaDataNo(sI_CompName: String): String;
var
    sL_DataAreaNo : String;
begin
    with dtmMain4.cdsIniInfo do
    begin
      if Locate('COMPNAME', sI_CompName, []) then
      begin
        sL_DataAreaNo := FieldByName('DataAreaNo').AsString;
        Result := sL_DataAreaNo;
      end;
    end;
end;

function TfrmSO8F10_3.getActiveQuery(sI_AreaDataNo: String): TADOQuery;
var
    ii, nL_Pos : Integer;
begin
    nL_Pos := -1;
    for ii:=0 to G_DbAreaNameStrList.Count -1 do
    begin
      if (DATA_AREA_HEADER + sI_AreaDataNo) = G_DbAreaNameStrList.Strings[ii] then
      begin
        nL_Pos := ii;
        break;
      end;
    end;
    if nL_Pos<>-1 then
    begin
      result := G_AdoQueryAry[nL_Pos];
      //savetofile('get query '+ sI_AreaDataNo, 'c:\bb4.txt');
    end
    else
    begin
      result := nil;
      //savetofile('can not get query '+ sI_AreaDataNo, 'c:\bb5.txt');
    end;

end;

function TfrmSO8F10_3.runSQL(nI_Mode: Integer; sI_AreaDataNo,
  sI_SQL: String): TADOQuery;
var
  L_AdoQuery : TADOQuery;
begin
  L_AdoQuery := getActiveQuery(sI_AreaDataNo);
  if L_AdoQuery<>nil then
  begin
    L_AdoQuery.SQL.Clear;
    L_AdoQuery.SQL.Add(sI_SQL);
    case nI_Mode of
      SELECT_MODE://select
        L_AdoQuery.Open;
      IUD_MODE://DML
        L_AdoQuery.ExecSQL;
    end;
  end;
  result := L_AdoQuery;


end;

procedure TfrmSO8F10_3.insertToCD067A(I_AdoQry: TADOQuery);
var
    sL_TableName,sL_TableDescription,sL_Operator,sL_UpdTime : String;
    sL_SQL : String;
begin
    I_AdoQry.First;
    while not I_AdoQry.Eof do
    begin
      sL_TableName := I_AdoQry.FieldByName('TableName').AsString;
      sL_TableDescription := I_AdoQry.FieldByName('TableDescription').AsString;
      sL_Operator := I_AdoQry.FieldByName('Operator').AsString;
      sL_UpdTime := I_AdoQry.FieldByName('UpdTime').AsString;

      sL_SQL := 'INSERT INTO CD067A(TABLENAME,TABLEDESCRIPTION,OPERATOR,UPDTIME) VALUES(''' +
                sL_TableName + ''',''' + sL_TableDescription + ''',''' + sL_Operator +
                ''',''' + sL_UpdTime + ''')';

      runSQL(IUD_MODE,sG_DataAreaNo,sL_SQL);

      I_AdoQry.Next;
    end;
end;

procedure TfrmSO8F10_3.insertToCD067B(I_AdoQry: TADOQuery);
var
    sL_CodeNo,sL_TableName,sL_Description,sL_RefNo,sL_ServiceType : String;
    sL_StopFlag,sL_Operator,sL_UpdTime : String;
    sL_SQL : String;
begin
    I_AdoQry.First;
    while not I_AdoQry.Eof do
    begin
      sL_CodeNo := I_AdoQry.FieldByName('CodeNo').AsString;
      sL_TableName := I_AdoQry.FieldByName('TableName').AsString;
      sL_Description := I_AdoQry.FieldByName('Description').AsString;
      sL_RefNo := I_AdoQry.FieldByName('RefNo').AsString;
      if sL_RefNo='' then
        sL_RefNo := 'null';
      sL_ServiceType := I_AdoQry.FieldByName('ServiceType').AsString;
      sL_StopFlag := I_AdoQry.FieldByName('StopFlag').AsString;
      sL_Operator := I_AdoQry.FieldByName('Operator').AsString;
      sL_UpdTime := I_AdoQry.FieldByName('UpdTime').AsString;


      sL_SQL := 'INSERT INTO CD067B(CODENO,TABLENAME,DESCRIPTION,REFNO,SERVICETYPE,STOPFLAG,OPERATOR,UPDTIME) VALUES(' +
                sL_CodeNo + ',''' + sL_TableName + ''',''' + sL_Description +
                ''',' + sL_RefNo + ',''' + sL_ServiceType + ''',' + sL_StopFlag +
                ',''' + sL_Operator + ''',''' + sL_UpdTime + ''')';

      runSQL(IUD_MODE,sG_DataAreaNo,sL_SQL);

      I_AdoQry.Next;
    end;

end;

procedure TfrmSO8F10_3.insertToCD067C(I_AdoQry: TADOQuery);
var
    sL_CompCode,sL_TableName,sL_MasterCodeNo,sL_DetailCodeNo : String;
    sL_StopFlag,sL_Operator,sL_UpdTime : String;
    sL_SQL : String;
begin
    I_AdoQry.First;
    while not I_AdoQry.Eof do
    begin
      sL_CompCode := I_AdoQry.FieldByName('CompCode').AsString;
      sL_TableName := I_AdoQry.FieldByName('TableName').AsString;
      sL_MasterCodeNo := I_AdoQry.FieldByName('MasterCodeNo').AsString;
      if sL_MasterCodeNo = '' then
        sL_MasterCodeNo := 'null';

      sL_DetailCodeNo := I_AdoQry.FieldByName('DetailCodeNo').AsString;
      if sL_DetailCodeNo = '' then
        sL_DetailCodeNo := 'null';

      sL_StopFlag := I_AdoQry.FieldByName('StopFlag').AsString;
      if sL_StopFlag = '' then
        sL_StopFlag := 'null';


      sL_Operator := I_AdoQry.FieldByName('Operator').AsString;
      sL_UpdTime := I_AdoQry.FieldByName('UpdTime').AsString;


      sL_SQL := 'INSERT INTO CD067C(COMPCODE,TABLENAME,MASTERCODENO,DETAILCODENO,STOPFLAG,OPERATOR,UPDTIME) VALUES(' +
                sL_CompCode + ',''' + sL_TableName + ''',' + sL_MasterCodeNo +
                ',' + sL_DetailCodeNo + ',' + sL_StopFlag +
                ',''' + sL_Operator + ''',''' + sL_UpdTime + ''')';

      runSQL(IUD_MODE,sG_DataAreaNo,sL_SQL);

      I_AdoQry.Next;
    end;


end;

function TfrmSO8F10_3.getComboBoxCompCode(sI_CompName: String): String;
var
    L_StrList : TStringList;
begin
    L_StrList := TStringList.Create;
    L_StrList := TUstr.ParseStrings(sI_CompName,'-');
    Result := L_StrList.Strings[0];
    L_StrList.Free;



end;

function TfrmSO8F10_3.getCompDBIniInfo: WideString;
var
    sL_IniFileName : String;
    L_IniFile : TIniFile;
    ii, nL_TotalDataArea : Integer;
    sL_DbAlias, sL_DbUserName, sL_DbPassword, sL_Tag : String;
    sL_CompCodeAndName,sL_CompCode,sL_CompName : String;
    L_Obj : TSnapShotIniData;
    sL_ExeFileName,sL_ExePath : String;
begin
    G_CompInfoStrList := TStringList.Create;
    G_DbAreaNameStrList := TStringList.Create;

    sL_ExeFileName := Application.ExeName;

    sL_ExePath := ExtractFileDir(sL_ExeFileName);

    sG_ExePath := sL_ExePath;

    sL_IniFileName := sL_ExePath + '\' + TMP_SYS_INI_FILE_NAME;

    if not FileExists(sL_IniFileName) then
    begin
      result := '-1: 讀取資料庫連線參數檔<'+sL_IniFileName +'>失敗';
      exit;
    end;
    L_IniFile := TIniFile.Create(sL_IniFileName);
    nL_TotalDataArea := L_IniFile.ReadInteger('SYSINFO','TotalDataAreaCounts',0);
    if nL_TotalDataArea=0 then
    begin
      result := '-1: 資料庫連線參數檔之資料區總數設定錯誤!<'+sL_IniFileName +'>';
      exit;
    end;


    if not dtmMain4.cdsIniInfo.Active then
      dtmMain4.cdsIniInfo.CreateDataSet;

    dtmMain4.cdsIniInfo.EmptyDataSet;

    try
      for ii:=0 to nL_TotalDataArea -1 do
      begin
        sL_Tag := DATA_AREA_HEADER + IntToStr(ii);

        sL_CompName := L_IniFile.ReadString(sL_Tag,'COMPNAME','');
        sL_CompCode := L_IniFile.ReadString(sL_Tag,'COMPCODE','');


        dtmMain4.cdsIniInfo.Append;
        dtmMain4.cdsIniInfo.FieldByName('DataAreaNo').AsString := IntToStr(ii);
        dtmMain4.cdsIniInfo.FieldByName('DataArea').AsString := sL_Tag;
        dtmMain4.cdsIniInfo.FieldByName('Alias').AsString := L_IniFile.ReadString(sL_Tag,'ALIAS','');;
        dtmMain4.cdsIniInfo.FieldByName('UserID').AsString := L_IniFile.ReadString(sL_Tag,'USERID','');
        dtmMain4.cdsIniInfo.FieldByName('Password').AsString := L_IniFile.ReadString(sL_Tag,'PASSWORD','');
        dtmMain4.cdsIniInfo.FieldByName('CompName').AsString := sL_CompCode + '-' + sL_CompName;
        dtmMain4.cdsIniInfo.FieldByName('CompCode').AsString := sL_CompCode;
        dtmMain4.cdsIniInfo.Post;

        G_DbAreaNameStrList.Add(sL_Tag);

        G_CompInfoStrList.AddObject(sL_CompName, L_Obj);
      end;

      L_IniFile.Free;
    except
      result := '-1: 資料庫連線參數檔之資料區設定錯誤!<'+sL_IniFileName +'>';
      exit;
    end;

    result := '';


end;

end.
