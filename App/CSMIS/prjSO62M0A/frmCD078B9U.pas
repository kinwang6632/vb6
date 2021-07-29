unit frmCD078B9U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ActnList, ExtCtrls, DB, ADODB, Menus, ComObj,

  cbDBController, cbStyleController,
  cxControls, cxContainer, cxEdit, cxTextEdit, cxMaskEdit, cxDropDownEdit,
  cxImageComboBox, cxDBLookupComboBox,
  cxCurrencyEdit, cxCheckBox, cxRadioGroup,
  cxLookAndFeelPainters, cxButtons, dxSkinsCore, dxSkinsDefaultPainters,
  dxSkinBlack, dxSkinBlue, dxSkinCaramel, dxSkinCoffee;

type
  TOldValue = record
    InstCode: String;
    ServiceType: String;
    GroupNo: String;
    RefNo: String;
  end;

  TfrmCD078B9 = class(TForm)
    Panel1: TPanel;
    Bevel1: TBevel;
    btnSave: TButton;
    btnCancel: TButton;
    Panel2: TPanel;
    btnBPCodeStr: TButton;
    cxButton1: TcxButton;
    edtBPCodeStr: TcxTextEdit;
    btnCompCodeStr: TButton;
    cxButton2: TcxButton;
    edtCompCodeStr: TcxTextEdit;
    DataConnection: TADOConnection;
    DataWriter: TADOQuery;
    DataReader: TADOQuery;
    txtMessage: TMemo;
    grpCopy: TGroupBox;
    radCopyNew: TcxRadioButton;
    radCopyExist: TcxRadioButton;
    radCopyOver: TcxRadioButton;
    Label1: TLabel;
    procedure btnBPCodeStrClick(Sender: TObject);
    procedure btnCompCodeStrClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure cxButton1Click(Sender: TObject);
    procedure cxButton2Click(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    FCloneCodeNo: String;
    FCloneDescription: String;
    BpCodeStr: String;
    BpNameStr: string;
    CompCodeStr: String;
    CompCodeStrCanSelect: String;
    CompNameStr: String;
    FReader: TADOQuery;
    FWriter: TADOQuery;
    FSourceLinkKeyList: TStringList;
    FCloneLinkKeyList: TStringList;
    FExists:Boolean;
    function VdCanInsert(const ATableOwner, ABPCode: String): Boolean;
    procedure DeleteCD078X(const ATableOwner, ABPCode: String);
    function CopyNew(const AReadOwner, ATableOwner, ABPCode, ACodeName: String;
      var AErrMessage: String): Boolean;
    function GetStepNo(const AOwner: String): String;
    function GetCD078G1Step(const AOwner:String): String;
    { Private declarations }
  public
    { Public declarations }
    property CloneCodeNo: String read FCloneCodeNo;
    property CloneDescription: String read FCloneDescription;
  end;

var
  frmCD078B9: TfrmCD078B9;

implementation

uses cbUtilis, frmMultiSelectU;

{$R *.dfm}

{ TfrmCD078B2 }

{ -------------------------------------------------------------------------------------- }

procedure TfrmCD078B9.btnBPCodeStrClick(Sender: TObject);
begin
  frmMultiSelect := TfrmMultiSelect.Create( Application );
  try
    frmMultiSelect.Connection := DBController.DataConnection;
    frmMultiSelect.KeyFields := 'CODENO';
    frmMultiSelect.KeyValues := BpCodeStr;
    frmMultiSelect.DisplayFields := 'CODENO, 優惠組合代碼,DESCRIPTION,優惠組合名稱';
    frmMultiSelect.ResultFields := 'DESCRIPTION';
    frmMultiSelect.SQL.Text := Format(
      ' SELECT CODENO, DESCRIPTION FROM %s.CD078 ' +
      '  WHERE STOPFLAG = 0 ORDER BY CODENO ',
      [DBController.LoginInfo.DbAccount] );
    if frmMultiSelect.ShowModal = mrOk then
      begin
        BpCodeStr := frmMultiSelect.SelectedValue;
        BpNameStr := frmMultiSelect.SelectedDisplay;
        edtBPCodeStr.Text := frmMultiSelect.SelectedDisplay;
      end;
  finally
    frmMultiSelect.Free;
  end;
end;

{ -------------------------------------------------------------------------------------- }

procedure TfrmCD078B9.btnCompCodeStrClick(Sender: TObject);
begin
  frmMultiSelect := TfrmMultiSelect.Create( Application );
  try
    frmMultiSelect.Connection := DBController.DataConnection;
    frmMultiSelect.KeyFields := 'CODENO';
    frmMultiSelect.KeyValues := CompCodeStr;
    frmMultiSelect.DisplayFields := 'CODENO, 代碼,DESCRIPTION,名稱';
    frmMultiSelect.ResultFields := 'DESCRIPTION';
    frmMultiSelect.SQL.Text := Format(
      ' SELECT CODENO, DESCRIPTION FROM %s.CD052 ' +
      ' WHERE CODENO <> %s                       ',
       [DBController.LoginInfo.DbAccount,
       DBController.LoginInfo.CompCode]);
    if CompCodeStrCanSelect <> EmptyStr then
      frmMultiSelect.SQL.Text := frmMultiSelect.SQL.Text +
        ' AND CODENO IN ( ' + CompCodeStrCanSelect +' ) '
    else
      frmMultiSelect.SQL.Text := frmMultiSelect.SQL.Text +
        ' And 1=0 ';
    frmMultiSelect.SQL.Text := frmMultiSelect.SQL.Text +
      '  ORDER BY CODENO ';
    if frmMultiSelect.ShowModal = mrOk then
      begin
        CompCodeStr := frmMultiSelect.SelectedValue;
        CompNameStr := frmMultiSelect.SelectedDisplay;
        edtCompCodeStr.Text := frmMultiSelect.SelectedDisplay;
      end;
  finally
    frmMultiSelect.Free;
  end;
end;

{ -------------------------------------------------------------------------------------- }

procedure TfrmCD078B9.FormCreate(Sender: TObject);
var
  aIndex: Integer;
begin

  FSourceLinkKeyList := TStringList.Create;
  FCloneLinkKeyList := TStringList.Create;

  FReader := DBController.DataReader;
  //#6644 不判斷權限,所以多加1=1 By Kin 2013/10/25
  if ( DBController.LoginInfo.UserId = 'A' ) or ( 1=1 )   then
    begin
      CompCodeStrCanSelect := '';
      for aIndex :=1 to 15 do
        CompCodeStrCanSelect := CompCodeStrCanSelect + IntToStr(aIndex)+',';
      CompCodeStrCanSelect := Copy (CompCodeStrCanSelect, 1, Length(CompCodeStrCanSelect)-1)
    end
  else
    begin
      FReader.Close;
      FReader.SQL.Text := Format (
        'SELECT COMPSTR FROM %s.SO026 WHERE USERID=''%s'' ',
        [DBController.LoginInfo.DbAccount, DBController.LoginInfo.UserId] );
      FReader.Open;
      CompCodeStrCanSelect := FReader.FieldByName('COMPSTR').AsString;
    end;
  edtCompCodeStr.Text := '(請選擇)';
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B9.FormDestroy(Sender: TObject);
begin
  FReader.Close;
  FSourceLinkKeyList.Free;
  FCloneLinkKeyList.Free;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B9.GetStepNo(const AOwner: String): String;
begin
  FWriter.Close;
  FWriter.SQL.Text := Format(
    ' SELECT %s.S_CD078A1.NEXTVAL FROM DUAL ', [AOwner] );
  FWriter.Open;
  Result := FWriter.Fields[0].AsString;
  FWriter.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B9.cxButton1Click(Sender: TObject);
begin
  BpCodeStr := EmptyStr;
  BpNameStr := EmptyStr;
  edtBPCodeStr.Text := EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B9.cxButton2Click(Sender: TObject);
begin
  CompCodeStr := EmptyStr;
  CompNameStr := EmptyStr;
  edtCompCodeStr.Text := EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B9.btnSaveClick(Sender: TObject);
var
  aIndex: Integer;
  aCompCodeStr, aCompCode: String;
  aCompNameStr, aCompName: String;
  aBPCodeStr, aBPNameStr, aBPCode, aBpName, aData, aOwner, aPass: String;
  aMessage, aErrMessage: String;
  aList: TStringList;
  EnCryptX: OleVariant;
  aErrCount, aTotalErrCount: Integer;
begin
  if (BpCodeStr = EmptyStr) or (CompCodeStr = EmptyStr) then
  begin
    WarningMsg( '尚未選擇【優惠組合代碼】或是【公司別】。' );
  end else
  begin
    txtMessage.Lines.Clear;
    {}
    EnCryptX := CreateOleObject( 'DevPower.Encrypt' );
    aList := TStringList.Create;
    try
      aList.LoadFromFile( ExtractFilePath(Application.ExeName) + 'GICMIS2.INI' );
      aCompCodeStr := CompCodeStr;
      aCompNameStr := CompNameStr;
      aTotalErrCount := 0;
      //公司別迴圈
      while ( aCompCodeStr <> EmptyStr ) do
      begin
        aCompCode := ExtractValue( aCompCodeStr );
        aCompCode := StringReplace( aCompCode, '''', EmptyStr, [rfReplaceAll] );
        aCompName := ExtractValue( aCompNameStr );
        if ( DBController.LoginInfo.CompCode <> aCompCode ) then
        begin
          //查 ConnectString 及 DB USER
          for aIndex := 0 to aList.Count -1 do
          begin
            if aList.Strings[aIndex] = '[' + aCompCode + ']' then
            begin
              aData := Copy( aList.Strings[aIndex + 1],3,Length(aList.Strings[aIndex + 1] ) - 2 );
              aData := EnCryptX.DeCrypt( aData );
              aOwner := Copy(aList.Strings[aIndex + 2],3,Length(aList.Strings[aIndex + 2] ) - 2 );
              aOwner := EnCryptX.DeCrypt(aOwner);
              aPass := Copy( aList.Strings[aIndex + 3],3,Length(aList.Strings[aIndex + 3] ) - 2 );
              aPass := EnCryptX.DeCrypt( aPass );
            end;
          end;
          DataConnection.ConnectionString := Format(
            'Provider=MSDAORA.1;Password=%s;User ID=%s;Data Source=%s;Persist Security Info=True',
            [aPass, aOwner, aData] );
          DataConnection.Open;
          aBPCodeStr := BPCodeStr;
          aBPNameStr := BpNameStr;
          {}
          aErrMessage := EmptyStr;
          aErrCount := 0;
          //優惠組合代碼迴圈
          while ( aBPCodeStr <> EmptyStr ) do
          begin
            aBPCode := ExtractValue( aBPCodeStr );
            aBPCode := StringReplace( aBPCode, '''', EmptyStr, [rfReplaceAll] );
            aBpName := ExtractValue( aBPNameStr );
            if not CopyNew( DBController.LoginInfo.DbAccount, aOwner, aBPCode, aBpName,
              aErrMessage ) then
            begin
              Inc( aErrCount );
              Inc( aTotalErrCount );
            end;  
          end;
          DataConnection.Close;
          if ( aErrCount <= 0 ) then
            aMessage := ( aMessage + #13 + Format( '公司別: %s, 複製完成。', [aCompName] ) )
          else begin
            aMessage := ( aMessage + #13 + Format( '公司別: %s, 複製失敗。', [aCompName] ) );
            aMessage := ( aMessage + aErrMessage );
          end;
        end;
      end;
    finally
      aList.Free;
    end;
    if ( aTotalErrCount <= 0 ) then
      InfoMsg( aMessage )
    else
      WarningMsg( aMessage );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B9.VdCanInsert(const ATableOwner, ABPCode: String): Boolean;
begin
//  Result := True;
//  if ( radCopyExist.Checked ) then
//  begin
//    DataReader.Close;
//    DataReader.SQL.Text := Format ( ' SELECT * FROM %s.CD078 WHERE CODENO=''%s'' ',
//      [ATableOwner, ABPCode] );
//    DataReader.Open;
//    Result := not DataReader.IsEmpty;
//    DataReader.Close;
//  end;
  DataWriter.Close;
  DataWriter.SQL.Text := Format (
    ' SELECT COUNT(1) FROM %s.CD078 WHERE CODENO=''%s'' ',
    [ATableOwner, ABPCode] );
  DataWriter.Open;
  //增加跨區服務的選項 By Kin 2009/08/06
  if DataWriter.Fields[0].AsInteger > 0 then
  begin
    if radCopyOver.Checked then
    begin
      FExists:=True;
      DataWriter.Close;
      DataWriter.SQL.Text:=Format('Select OnSaleStartDate From %s.CD078 ' +
                                  ' Where CODENO=''%s'' ',
                                  [ATableOwner,ABPCode]);
      DataWriter.Open;
      if DataWriter.Fields[0].AsString <> EmptyStr then
      begin
        if DataWriter.Fields[0].AsDateTime<=Now then
          Result:= False
        else
          Result := True;
      end else
      begin
        Result := True;
      end;
      DataWriter.Close;
      Exit;
    end;

  end;

  Result := ( DataWriter.Fields[0].AsInteger <= 0 );
  DataWriter.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B9.DeleteCD078X(const ATableOwner, ABPCode: String);
var
  aIndex:Integer;
begin
  {#7150 Modifing the funciont beacause It happend  pk error By Kin 2015/12/08}
  for aIndex := 65 to 70 do // A..F
  begin
    DataWriter.Close;
    DataWriter.SQL.Text := Format ( ' DELETE FROM %s.CD078%s WHERE BPCODE = ''%s'' ',
      [ATableOwner, Chr( aIndex ), ABPCode] );
    DataWriter.ExecSQL;
  end;
  DataWriter.Close;
  DataWriter.SQL.Text := Format (
    ' DELETE FROM %s.CD078 WHERE CODENO = ''%s'' ', [ATableOwner, ABPCode] );
  DataWriter.ExecSQL;
  {}
  DataWriter.Close;
  DataWriter.SQL.Text := Format(
    ' DELETE FROM %s.CD078A1 WHERE BPCODE = ''%s'' ', [ATableOwner, ABPCode] );
  DataWriter.ExecSQL;
  {}
  DataWriter.SQL.Text := Format(
    ' DELETE FROM %s.CD078A2 WHERE BPCODE = ''%s'' ', [ATableOwner, ABPCode] );
  DataWriter.ExecSQL;
  {}
  DataWriter.SQL.Text := Format(
    ' DELETE FROM %s.CD078A3 WHERE BPCODE = ''%s'' ', [ATableOwner, ABPCode] );
  DataWriter.ExecSQL;
  {}
  DataWriter.SQL.Text := Format(
    ' DELETE FROM %s.CD078A4 WHERE BPCODE = ''%s'' ', [ATableOwner, ABPCode] );
  DataWriter.ExecSQL;
  {}
  FWriter.Close;
  FWriter.SQL.Text := Format (
    ' DELETE FROM %s.CD078B1 WHERE BPCODE = ''%s'' ', [ATableOwner, ABPCode] );
  FWriter.ExecSQL;
  {}
  FWriter.Close;
  FWriter.SQL.Text := Format (
    ' DELETE FROM %s.CD078G  WHERE BPCODE = ''%s'' ', [ATableOwner, ABPCode] );
  FWriter.ExecSQL;
  {}
  FWriter.Close;
  FWriter.SQL.Text := Format (
    ' DELETE FROM %s.CD078G1 WHERE BPCODE = ''%s'' ', [ATableOwner, ABPCode] );
  FWriter.ExecSQL;
  {}
  FWriter.Close;
  FWriter.SQL.Text := Format (
    ' DELETE FROM %s.CD078I WHERE BPCODE = ''%s'' ', [ATableOwner, ABPCode] );
  FWriter.ExecSQL;
  {}
  FWriter.Close;
  FWriter.SQL.Text := Format (
    ' DELETE FROM %s.CD078J WHERE BPCODE = ''%s'' ', [ATableOwner, ABPCode] );
  FWriter.ExecSQL;
  {}
  FWriter.Close;
  FWriter.SQL.Text := Format (
    ' DELETE FROM %s.CD078K WHERE BPCODE = ''%s'' ', [ATableOwner, ABPCode] );
  FWriter.ExecSQL;
  {}
  FWriter.Close;
  FWriter.SQL.Text := Format (
    ' DELETE FROM %s.CD078L WHERE BPCODE = ''%s'' ', [ATableOwner, ABPCode] );
  FWriter.ExecSQL;
  FWriter.Close;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B9.CopyNew(const AReadOwner, ATableOwner, ABPCode, ACodeName: String;
  var AErrMessage: String): Boolean;
var
  aDateSt, aDateEd, aStepNo, aCD078G1Step, aLinkKey: String;
  aLinkIndex: Integer;
  aIndex: Integer;

begin

  Result := True;
  FExists := False; //#5232如果是True的話,代表是選刪除後重新覆蓋 By Kin
  if not VdCanInsert( ATableOwner, ABPCode ) then
  begin
      if FExists then
      begin
        txtMessage.Lines.Add( Format(
          '優惠組合:%s-%s  ---->  該優惠組合已上架開賣，不允許覆蓋代碼內容！',
          [ABPCode, ACodeName]));
        Application.ProcessMessages;
        Exit;
      end else
      begin
        txtMessage.Lines.Add( Format(
          '優惠組合:%s-%s  ---->  已存在,略過複製。', [ABPCode, ACodeName] ) );
        Application.ProcessMessages;
        Exit;
      end;
  end;

  FWriter := DataWriter;
//  FDelete := DataDelete;
  DataConnection.BeginTrans;
  Screen.Cursor := crHourGlass;
  try
    //#5232 如果是選擇覆蓋,要先刪除資料 By Kin 2009/08/11
    if FExists then
    begin
      DeleteCD078X( ATableOwner, ABPCode );
    end;
    //CD078
    FReader.Close;
    FReader.SQL.Text := Format (
      ' SELECT * FROM %s.CD078 WHERE CODENO=''%s'' ',
      [AReadOwner, ABPCode] );
    FReader.Open;
    {}
    aDateSt := EmptyStr;
    aDateEd := EmptyStr;
    if ( FReader.FieldByName( 'ONSALESTARTDATE' ).AsString <> EmptyStr ) then
      aDateSt := FormatDateTime( 'yyyymmdd', FReader.FieldByName( 'ONSALESTARTDATE' ).AsDateTime );
    if ( FReader.FieldByName( 'ONSALESTOPDATE' ).AsString <> EmptyStr ) then
      aDateEd := FormatDateTime( 'yyyymmdd', FReader.FieldByName( 'ONSALESTOPDATE' ).AsDateTime );
    {}
    { #4393 計價依據KINDFUNCTION 沒有同步複制 By Kin 2009/03/05 }
    { #6267 增加SYNCBPCODE 欄位 By Kin 2012/07/19}
    FWriter.Close;
    FWriter.SQL.Text := Format (
      ' INSERT INTO %s.CD078 (                                        ' +
      '   CODENO, DESCRIPTION, ONSALESTARTDATE, ONSALESTOPDATE,       ' +
      '   MASTERPROD, BUNDLE, BUNDLESAMEMON, BUNDLEMON, SAMEPERIOD,   ' +
      '   PERIOD, MERGEPRINT, PENAL, PENALTYPE, PENALSINGLE,          ' +
      '   MERGEWORK, BUNDLEPRODNOTE, WORKNOTE, REFNO, UPDTIME, UPDEN, ' +
      '   STOPFLAG, CONDITIONAL,KINDFUNCTION,CHOOSEKIND,CANPAYNOW,    ' +
      '   SYNCBPCODE,CANUSETYPE, EPGIMAGEURL, EPGORD )                ' +
      ' VALUES (                                                      ' +
      '   ''%s'', ''%s'', TO_DATE( ''%s'', ''YYYYMMDD'' ),            ' +
      '   TO_DATE( ''%s'', ''YYYYMMDD'' ), ''%s'', ''%s'',            ' +
      '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',     ' +
      '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',     ' +
      '   ''%s'', ''%s'',%d,%d,%d, ''%s'',                            ' +
      '   ''%s'', ''%s'', ''%s'' )                                    ',
      [ATableOwner,
      ABPCode,
      FReader.FieldByName('DESCRIPTION').AsString,
      aDateSt,
      aDateEd,
      FReader.FieldByName('MASTERPROD').AsString,
      FReader.FieldByName('BUNDLE').AsString,
      FReader.FieldByName('BUNDLESAMEMON').AsString,
      FReader.FieldByName('BUNDLEMON').AsString,
      FReader.FieldByName('SAMEPERIOD').AsString,
      FReader.FieldByName('PERIOD').AsString,
      FReader.FieldByName('MERGEPRINT').AsString,
      FReader.FieldByName('PENAL').AsString,
      FReader.FieldByName('PENALTYPE').AsString,
      FReader.FieldByName('PENALSINGLE').AsString,
      FReader.FieldByName('MERGEWORK').AsString,
      FReader.FieldByName('BUNDLEPRODNOTE').AsString,
      FReader.FieldByName('WORKNOTE').AsString,
      FReader.FieldByName('REFNO').AsString,
      FReader.FieldByName('UPDTIME').AsString,
      FReader.FieldByName('UPDEN').AsString,
      FReader.FieldByName('STOPFLAG').AsString,
      FReader.FieldByName('CONDITIONAL').AsString,
      FReader.FieldByName('KINDFUNCTION').AsInteger,
      FReader.FieldByName('CHOOSEKIND').AsInteger,
      FReader.FieldByName('CANPAYNOW').AsInteger,
      FReader.FieldByName('SYNCBPCODE').AsString,
      FReader.FieldByName('CANUSETYPE').AsString,
      FReader.FieldByName('EPGIMAGEURL').AsString,
      FReader.FieldByName('EPGORD').AsString ] );
    FWriter.ExecSQL;

    //CD078A
    //#5318 測試不OK 複製過去CD078A少了點數辦法代碼及名稱值 By Kin 2009/10/30
    //#7130 There was lost to update AuthTunerCount field of value By Kin 2015/11/09
    FReader.Close;
    FReader.SQL.Text := Format (
        ' SELECT * FROM %s.CD078A WHERE BPCODE=''%s'' ',
        [AReadOwner, ABPCode] );
    FReader.Open;
    
    while not FReader.Eof do
    begin

      {5913 增加PFlag2 By Kin 2011/04/14}
      FWriter.Close;
      FWriter.SQL.Text := Format (
        ' INSERT INTO %s.CD078A ( bpcode, citemcode, citemname, servicetype,' +
        '   faciitem, bundle, bundlemon, penaltype, expiretype, expirecut,  ' +
        '   depositattr, depositcode, depositname, monthamt,                ' +
        '   dayamt, pflag1, pflagdate, truncamt, pflagcode, pflagname,      ' +
        '   mon1, period1, ratetype1, discountamt1, monthamt1, dayamt1,     ' +
        '   punish1, mon2, period2, ratetype2, discountamt2, monthamt2,     ' +
        '   dayamt2, punish2, mon3, period3, ratetype3, discountamt3,       ' +
        '   monthamt3, dayamt3, punish3, mon4, period4, ratetype4,          ' +
        '   discountamt4, monthamt4, dayamt4, punish4, mon5, period5,       ' +
        '   ratetype5, discountamt5, monthamt5, dayamt5, punish5, mon6,     ' +
        '   period6, ratetype6, discountamt6, monthamt6, dayamt6, punish6,  ' +
        '   mon7, period7, ratetype7, discountamt7, monthamt7, dayamt7,     ' +
        '   punish7, mon8, period8, ratetype8, discountamt8, monthamt8,     ' +
        '   dayamt8, punish8, mon9, period9, ratetype9, discountamt9,       ' +
        '   monthamt9, dayamt9, punish9, mon10, period10, ratetype10,       ' +
        '   discountamt10, monthamt10, dayamt10, punish10, mon11, period11, ' +
        '   ratetype11, discountamt11, monthamt11, dayamt11, punish11,      ' +
        '   mon12, period12, ratetype12, discountamt12, monthamt12,         ' +
        '   dayamt12, punish12, penal, penalcode, penalname,                ' +
        '   depositcodestr, discountcalc, mcitemcode,                       ' +
        '   SalePointcode,SalePointName,PFlag2,AuthTunerCount  )            ' +
        ' VALUES ( ''%s'', ''%s'', ''%s'', ''%s'',                          ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'',                                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                         ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                         ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                         ' +
        '   ''%s'', ''%s'', ''%s'',''%s'',''%s'',''%s'',''%s'' )            ',
        [ATableOwner,
         ABPCode,
         FReader.FieldByName( 'citemcode' ).AsString,
         FReader.FieldByName( 'citemname' ).AsString,
         FReader.FieldByName( 'servicetype' ).AsString,
         FReader.FieldByName( 'faciitem' ).AsString,
         FReader.FieldByName( 'bundle' ).AsString,
         FReader.FieldByName( 'bundlemon' ).AsString,
         FReader.FieldByName( 'penaltype' ).AsString,
         FReader.FieldByName( 'expiretype' ).AsString,
         FReader.FieldByName( 'expirecut' ).AsString,
         FReader.FieldByName( 'depositattr' ).AsString,
         FReader.FieldByName( 'depositcode' ).AsString,
         FReader.FieldByName( 'depositname' ).AsString,
         FReader.FieldByName( 'monthamt' ).AsString,
         FReader.FieldByName( 'dayamt' ).AsString,
         FReader.FieldByName( 'pflag1' ).AsString,
         FReader.FieldByName( 'pflagdate' ).AsString,
         FReader.FieldByName( 'truncamt' ).AsString,
         FReader.FieldByName( 'pflagcode' ).AsString,
         FReader.FieldByName( 'pflagname' ).AsString,
         {}
         FReader.FieldByName( 'mon1').AsString,
         FReader.FieldByName( 'period1').AsString,
         FReader.FieldByName( 'ratetype1' ).AsString,
         FReader.FieldByName( 'discountamt1' ).AsString,
         FReader.FieldByName( 'monthamt1' ).AsString,
         FReader.FieldByName( 'dayamt1' ).AsString,
         FReader.FieldByName( 'punish1' ).AsString,
         {}
         FReader.FieldByName( 'mon2').AsString,
         FReader.FieldByName( 'period2').AsString,
         FReader.FieldByName( 'ratetype2' ).AsString,
         FReader.FieldByName( 'discountamt2' ).AsString,
         FReader.FieldByName( 'monthamt2' ).AsString,
         FReader.FieldByName( 'dayamt2' ).AsString,
         FReader.FieldByName( 'punish2' ).AsString,
         {}
         FReader.FieldByName( 'mon3').AsString,
         FReader.FieldByName( 'period3').AsString,
         FReader.FieldByName( 'ratetype3' ).AsString,
         FReader.FieldByName( 'discountamt3' ).AsString,
         FReader.FieldByName( 'monthamt3' ).AsString,
         FReader.FieldByName( 'dayamt3' ).AsString,
         FReader.FieldByName( 'punish3' ).AsString,
         {}
         FReader.FieldByName( 'mon4').AsString,
         FReader.FieldByName( 'period4').AsString,
         FReader.FieldByName( 'ratetype4' ).AsString,
         FReader.FieldByName( 'discountamt4' ).AsString,
         FReader.FieldByName( 'monthamt4' ).AsString,
         FReader.FieldByName( 'dayamt4' ).AsString,
         FReader.FieldByName( 'punish4' ).AsString,
         {}
         FReader.FieldByName( 'mon5').AsString,
         FReader.FieldByName( 'period5').AsString,
         FReader.FieldByName( 'ratetype5' ).AsString,
         FReader.FieldByName( 'discountamt5' ).AsString,
         FReader.FieldByName( 'monthamt5' ).AsString,
         FReader.FieldByName( 'dayamt5' ).AsString,
         Freader.FieldByName( 'punish5' ).AsString,
         {}
         FReader.FieldByName( 'mon6' ).AsString,
         FReader.FieldByName( 'period6' ).AsString,
         FReader.FieldByName( 'ratetype6' ).AsString,
         FReader.FieldByName( 'discountamt6' ).AsString,
         FReader.FieldByName( 'monthamt6' ).AsString,
         FReader.FieldByName( 'dayamt6' ).AsString,
         FReader.FieldByName( 'punish6' ).AsString,
         {}
         FReader.FieldByName( 'mon7').AsString,
         FReader.FieldByName( 'period7').AsString,
         FReader.FieldByName( 'ratetype7' ).AsString,
         FReader.FieldByName( 'discountamt7' ).AsString,
         FReader.FieldByName( 'monthamt7' ).AsString,
         FReader.FieldByName( 'dayamt7' ).AsString,
         FReader.FieldByName( 'punish7' ).AsString,
         {}
         FReader.FieldByName( 'mon8').AsString,
         FReader.FieldByName( 'period8').AsString,
         FReader.FieldByName( 'ratetype8' ).AsString,
         FReader.FieldByName( 'discountamt8' ).AsString,
         FReader.FieldByName( 'monthamt8' ).AsString,
         FReader.FieldByName( 'dayamt8' ).AsString,
         FReader.FieldByName( 'punish8' ).AsString,
         {}
         FReader.FieldByName( 'mon9').AsString,
         FReader.FieldByName( 'period9').AsString,
         FReader.FieldByName( 'ratetype9' ).AsString,
         FReader.FieldByName( 'discountamt9' ).AsString,
         FReader.FieldByName( 'monthamt9' ).AsString,
         FReader.FieldByName( 'dayamt9' ).AsString,
         FReader.FieldByName( 'punish9' ).AsString,
         {}
         FReader.FieldByName( 'mon10').AsString,
         FReader.FieldByName( 'period10').AsString,
         FReader.FieldByName( 'ratetype10' ).AsString,
         FReader.FieldByName( 'discountamt10' ).AsString,
         FReader.FieldByName( 'monthamt10' ).AsString,
         FReader.FieldByName( 'dayamt10' ).AsString,
         FReader.FieldByName( 'punish10' ).AsString,
         {}
         FReader.FieldByName( 'mon11').AsString,
         FReader.FieldByName( 'period11').AsString,
         FReader.FieldByName( 'ratetype11' ).AsString,
         FReader.FieldByName( 'discountamt11' ).AsString,
         FReader.FieldByName( 'monthamt11' ).AsString,
         FReader.FieldByName( 'dayamt11' ).AsString,
         FReader.FieldByName( 'punish11' ).AsString,
         {}
         FReader.FieldByName( 'mon12').AsString,
         FReader.FieldByName( 'period12').AsString,
         FReader.FieldByName( 'ratetype12' ).AsString,
         FReader.FieldByName( 'discountamt12' ).AsString,
         FReader.FieldByName( 'monthamt12' ).AsString,
         FReader.FieldByName( 'dayamt12' ).AsString,
         FReader.FieldByName( 'punish12' ).AsString,
         {}
         FReader.FieldByName( 'penal' ).AsString,
         FReader.FieldByName( 'penalcode' ).AsString,
         FReader.FieldByName( 'penalname' ).AsString,
         {}
         FReader.FieldByName( 'depositcodestr' ).AsString,
         FReader.FieldByName( 'discountcalc' ).AsString,
         FReader.FieldByName( 'mcitemcode' ).AsString,
         FReader.FieldByName( 'SalePointcode').AsString,
         FReader.FieldByName( 'SalePointName').AsString,
         FReader.FieldByName( 'PFlag2' ).AsString,
         FReader.FieldByName( 'AuthTunerCount' ).AsString ] );
        FWriter.ExecSQL;
        {}
        FWriter.SQL.Text := EmptyStr;
        FWriter.Parameters.Clear;
        {}
        {#5257 增加特定優惠截止日 By Kin 2010/05/05}
        FWriter.SQL.Text := Format(
          ' UPDATE %s.CD078A ' +
          '    SET INSTCODESTR = :INSTCODESTR,  ' +
          '        COMMENT1 = :COMMENT1,         ' +
          '        COMMENT2 = :COMMENT2,         ' +
          '        COMMENT3 = :COMMENT3,         ' +
          '        COMMENT4 = :COMMENT4,         ' +
          '        COMMENT5 = :COMMENT5,         ' +
          '        COMMENT6 = :COMMENT6,         ' +
          '        COMMENT7 = :COMMENT7,         ' +
          '        COMMENT8 = :COMMENT8,         ' +
          '        COMMENT9 = :COMMENT9,         ' +
          '        COMMENT10 = :COMMENT10,       ' +
          '        COMMENT11 = :COMMENT11,       ' +
          '        COMMENT12 = :COMMENT12,       ' +
          '        CMBAUDNO = :CMBAUDNO,         ' +
          '        CMBAUDRATE = :CMBAUDRATE,     ' +
          '        BPSTOPDATE = :BPSTOPDATE      ' +
          '  WHERE BPCODE = ''%s''               ' +
          '    AND CITEMCODE = ''%s''            ' +
          '    AND SERVICETYPE = ''%s''          ',
          [ATableOwner, ABPCode,
           FReader.FieldByName( 'CITEMCODE' ).AsString,
           FReader.FieldByName( 'SERVICETYPE' ).AsString] );
        {}
        FWriter.Parameters.ParamByName( 'INSTCODESTR' ).Value := FReader.FieldByName( 'INSTCODESTR' ).AsString;
        FWriter.Parameters.ParamByName( 'COMMENT1' ).Value := FReader.FieldByName( 'COMMENT1' ).AsString;
        FWriter.Parameters.ParamByName( 'COMMENT2' ).Value := FReader.FieldByName( 'COMMENT2' ).AsString;
        FWriter.Parameters.ParamByName( 'COMMENT3' ).Value := FReader.FieldByName( 'COMMENT3' ).AsString;
        FWriter.Parameters.ParamByName( 'COMMENT4' ).Value := FReader.FieldByName( 'COMMENT4' ).AsString;
        FWriter.Parameters.ParamByName( 'COMMENT5' ).Value := FReader.FieldByName( 'COMMENT5' ).AsString;
        FWriter.Parameters.ParamByName( 'COMMENT6' ).Value := FReader.FieldByName( 'COMMENT6' ).AsString;
        FWriter.Parameters.ParamByName( 'COMMENT7' ).Value := FReader.FieldByName( 'COMMENT7' ).AsString;
        FWriter.Parameters.ParamByName( 'COMMENT8' ).Value := FReader.FieldByName( 'COMMENT8' ).AsString;
        FWriter.Parameters.ParamByName( 'COMMENT9' ).Value := FReader.FieldByName( 'COMMENT9' ).AsString;
        FWriter.Parameters.ParamByName( 'COMMENT10' ).Value := FReader.FieldByName( 'COMMENT10' ).AsString;
        FWriter.Parameters.ParamByName( 'COMMENT11' ).Value := FReader.FieldByName( 'COMMENT11' ).AsString;
        FWriter.Parameters.ParamByName( 'COMMENT12' ).Value := FReader.FieldByName( 'COMMENT12' ).AsString;
        FWriter.Parameters.ParamByName( 'CMBAUDNO' ).Value := FReader.FieldByName( 'CMBAUDNO' ).AsString;
        FWriter.Parameters.ParamByName( 'CMBAUDRATE' ).Value := FReader.FieldByName( 'CMBAUDRATE' ).AsString;
        if FReader.FieldByName( 'BPSTOPDATE' ).AsString = EmptyStr then
          FWriter.Parameters.ParamByName( 'BPSTOPDATE' ).Value := EmptyStr
        else
          FWriter.Parameters.ParamByName( 'BPSTOPDATE' ).Value := FReader.FieldByName( 'BPSTOPDATE' ).AsDateTime;
        FWriter.ExecSQL;
        FReader.Next
    end;

    FSourceLinkKeyList.Clear;
    FCloneLinkKeyList.Clear;

    //CD078A1, PK 是 StepNo, Copy 功能要在取一次 StepNo, 不然會PK重覆
    FReader.Close;
    FReader.SQL.Text := Format (
        ' SELECT * FROM %s.CD078A1    ' +
        '  WHERE BPCODE = ''%s''      ' +
        '  ORDER BY BPCODE, CITEMCODE, SERVICETYPE, FACIITEM, STEPNO ',
        [AReadOwner, ABPCode] );
    FReader.Open;

    while not FReader.Eof do
    begin
      FSourceLinkKeyList.Add( FReader.FieldByName( 'STEPNO' ).AsString );
      {}
      aStepNo := GetStepNo( ATableOwner );
      {}
      FCloneLinkKeyList.Add( aStepNo );
      {}
      aLinkKey := EmptyStr;
      aLinkIndex := -1;
      if ( FReader.FieldByName( 'LINKKEY' ).AsString <> EmptyStr ) then
        aLinkIndex := FSourceLinkKeyList.IndexOf( FReader.FieldByName( 'LINKKEY' ).AsString );
      if ( aLinkIndex >= 0 ) then
        aLinkKey := FCloneLinkKeyList[aLinkIndex];
      {}
      {#6945 少了複制Period By Kin 2014/11/28}
      FWriter.SQL.Text := Format (
        ' INSERT INTO %s.CD078A1( bpcode, citemcode, servicetype,           ' +
        '   faciitem, StepNo, Description, MasterSale,                      ' +
        '   mon1, period1, ratetype1, discountamt1, monthamt1, dayamt1,     ' +
        '   punish1, mon2, period2, ratetype2, discountamt2, monthamt2,     ' +
        '   dayamt2, punish2, mon3, period3, ratetype3, discountamt3,       ' +
        '   monthamt3, dayamt3, punish3, mon4, period4, ratetype4,          ' +
        '   discountamt4, monthamt4, dayamt4, punish4, mon5, period5,       ' +
        '   ratetype5, discountamt5, monthamt5, dayamt5, punish5, mon6,     ' +
        '   period6, ratetype6, discountamt6, monthamt6, dayamt6, punish6,  ' +
        '   mon7, period7, ratetype7, discountamt7, monthamt7, dayamt7,     ' +
        '   punish7, mon8, period8, ratetype8, discountamt8, monthamt8,     ' +
        '   dayamt8, punish8, mon9, period9, ratetype9, discountamt9,       ' +
        '   monthamt9, dayamt9, punish9, mon10, period10, ratetype10,       ' +
        '   discountamt10, monthamt10, dayamt10, punish10, mon11, period11, ' +
        '   ratetype11, discountamt11, monthamt11, dayamt11, punish11,      ' +
        '   mon12, period12, ratetype12, discountamt12, monthamt12,         ' +
        '   dayamt12, punish12, linkkey,Period,StopFlag )                   ' +
        ' VALUES ( ''%s'', ''%s'', ''%s'',                                  ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'',                                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                         ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                         ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'' )                        ',
        [ATableOwner,
         ABPCode,
         FReader.FieldByName( 'citemcode' ).AsString,
         FReader.FieldByName( 'servicetype' ).AsString,
         FReader.FieldByName( 'faciItem' ).AsString,
         aStepNo,
         FReader.FieldByName( 'description' ).AsString,
         FReader.FieldByName( 'mastersale' ).AsString,
         {}
         FReader.FieldByName( 'mon1').AsString,
         FReader.FieldByName( 'period1').AsString,
         FReader.FieldByName( 'ratetype1' ).AsString,
         FReader.FieldByName( 'discountamt1' ).AsString,
         FReader.FieldByName( 'monthamt1' ).AsString,
         FReader.FieldByName( 'dayamt1' ).AsString,
         FReader.FieldByName( 'punish1' ).AsString,
         {}
         FReader.FieldByName( 'mon2').AsString,
         FReader.FieldByName( 'period2').AsString,
         FReader.FieldByName( 'ratetype2' ).AsString,
         FReader.FieldByName( 'discountamt2' ).AsString,
         FReader.FieldByName( 'monthamt2' ).AsString,
         FReader.FieldByName( 'dayamt2' ).AsString,
         FReader.FieldByName( 'punish2' ).AsString,
         {}
         FReader.FieldByName( 'mon3').AsString,
         FReader.FieldByName( 'period3').AsString,
         FReader.FieldByName( 'ratetype3' ).AsString,
         FReader.FieldByName( 'discountamt3' ).AsString,
         FReader.FieldByName( 'monthamt3' ).AsString,
         FReader.FieldByName( 'dayamt3' ).AsString,
         FReader.FieldByName( 'punish3' ).AsString,
         {}
         FReader.FieldByName( 'mon4').AsString,
         FReader.FieldByName( 'period4').AsString,
         FReader.FieldByName( 'ratetype4' ).AsString,
         FReader.FieldByName( 'discountamt4' ).AsString,
         FReader.FieldByName( 'monthamt4' ).AsString,
         FReader.FieldByName( 'dayamt4' ).AsString,
         FReader.FieldByName( 'punish4' ).AsString,
         {}
         FReader.FieldByName( 'mon5').AsString,
         FReader.FieldByName( 'period5').AsString,
         FReader.FieldByName( 'ratetype5' ).AsString,
         FReader.FieldByName( 'discountamt5' ).AsString,
         FReader.FieldByName( 'monthamt5' ).AsString,
         FReader.FieldByName( 'dayamt5' ).AsString,
         FReader.FieldByName( 'punish5' ).AsString,
         {}
         FReader.FieldByName( 'mon6').AsString,
         FReader.FieldByName( 'period6').AsString,
         FReader.FieldByName( 'ratetype6' ).AsString,
         FReader.FieldByName( 'discountamt6' ).AsString,
         FReader.FieldByName( 'monthamt6' ).AsString,
         FReader.FieldByName( 'dayamt6' ).AsString,
         FReader.FieldByName( 'punish6' ).AsString,
         {}
         FReader.FieldByName( 'mon7').AsString,
         FReader.FieldByName( 'period7').AsString,
         FReader.FieldByName( 'ratetype7' ).AsString,
         FReader.FieldByName( 'discountamt7' ).AsString,
         FReader.FieldByName( 'monthamt7' ).AsString,
         FReader.FieldByName( 'dayamt7' ).AsString,
         FReader.FieldByName( 'punish7' ).AsString,
         {}
         FReader.FieldByName( 'mon8').AsString,
         FReader.FieldByName( 'period8').AsString,
         FReader.FieldByName( 'ratetype8' ).AsString,
         FReader.FieldByName( 'discountamt8' ).AsString,
         FReader.FieldByName( 'monthamt8' ).AsString,
         FReader.FieldByName( 'dayamt8' ).AsString,
         FReader.FieldByName( 'punish8' ).AsString,
         {}
         FReader.FieldByName( 'mon9').AsString,
         FReader.FieldByName( 'period9').AsString,
         FReader.FieldByName( 'ratetype9' ).AsString,
         FReader.FieldByName( 'discountamt9' ).AsString,
         FReader.FieldByName( 'monthamt9' ).AsString,
         FReader.FieldByName( 'dayamt9' ).AsString,
         FReader.FieldByName( 'punish9' ).AsString,
         {}
         FReader.FieldByName( 'mon10').AsString,
         FReader.FieldByName( 'period10').AsString,
         FReader.FieldByName( 'ratetype10' ).AsString,
         FReader.FieldByName( 'discountamt10' ).AsString,
         FReader.FieldByName( 'monthamt10' ).AsString,
         FReader.FieldByName( 'dayamt10' ).AsString,
         FReader.FieldByName( 'punish10' ).AsString,
         {}
         FReader.FieldByName( 'mon11').AsString,
         FReader.FieldByName( 'period11').AsString,
         FReader.FieldByName( 'ratetype11' ).AsString,
         FReader.FieldByName( 'discountamt11' ).AsString,
         FReader.FieldByName( 'monthamt11' ).AsString,
         FReader.FieldByName( 'dayamt11' ).AsString,
         FReader.FieldByName( 'punish11' ).AsString,
         {}
         FReader.FieldByName( 'mon12').AsString,
         FReader.FieldByName( 'period12').AsString,
         FReader.FieldByName( 'ratetype12' ).AsString,
         FReader.FieldByName( 'discountamt12' ).AsString,
         FReader.FieldByName( 'monthamt12' ).AsString,
         FReader.FieldByName( 'dayamt12' ).AsString,
         FReader.FieldByName( 'punish12' ).AsString,
         aLinkKey,
         FReader.FieldByName( 'Period' ).AsString,
         FReader.FieldByName( 'StopFlag' ).AsString   ] );
      FWriter.ExecSQL;
      {}
      FWriter.SQL.Text := Format(
        ' UPDATE %s.CD078A1 ' +
        '    SET COMMENT1 = :COMMENT1,         ' +
        '        COMMENT2 = :COMMENT2,         ' +
        '        COMMENT3 = :COMMENT3,         ' +
        '        COMMENT4 = :COMMENT4,         ' +
        '        COMMENT5 = :COMMENT5,         ' +
        '        COMMENT6 = :COMMENT6,         ' +
        '        COMMENT7 = :COMMENT7,         ' +
        '        COMMENT8 = :COMMENT8,         ' +
        '        COMMENT9 = :COMMENT9,         ' +
        '        COMMENT10 = :COMMENT10,       ' +
        '        COMMENT11 = :COMMENT11,       ' +
        '        COMMENT12 = :COMMENT12,       ' +
        '        DISCOUNTRATE1 = :DISCOUNTRATE1,   ' +
        '        DISCOUNTRATE2 = :DISCOUNTRATE2,   ' +
        '        DISCOUNTRATE3 = :DISCOUNTRATE3,   ' +
        '        DISCOUNTRATE4 = :DISCOUNTRATE4,   ' +
        '        DISCOUNTRATE5 = :DISCOUNTRATE5,   ' +
        '        DISCOUNTRATE6 = :DISCOUNTRATE6,   ' +
        '        DISCOUNTRATE7 = :DISCOUNTRATE7,   ' +
        '        DISCOUNTRATE8 = :DISCOUNTRATE8,   ' +
        '        DISCOUNTRATE9 = :DISCOUNTRATE9,   ' +
        '        DISCOUNTRATE10 = :DISCOUNTRATE10, ' +
        '        DISCOUNTRATE11 = :DISCOUNTRATE11, ' +
        '        DISCOUNTRATE12 = :DISCOUNTRATE12  ' +
        '  WHERE BPCODE = ''%s''               ' +
        '    AND STEPNO = ''%s''               ',
        [ATableOwner, ABPCode, aStepNo] );
      {}  
      for aIndex := 1 to 12 do
      begin
        FWriter.Parameters.ParamByName( Format( 'COMMENT%d', [aIndex] ) ).Value :=
          FReader.FieldByName( Format( 'COMMENT%d', [aIndex] ) ).AsString;
        {}  
        FWriter.Parameters.ParamByName( Format( 'DISCOUNTRATE%d', [aIndex] ) ).Value :=
          FReader.FieldByName( Format( 'DISCOUNTRATE%d', [aIndex] ) ).AsString;
      end;
      FWriter.ExecSQL;
      {}
      FReader.Next;
    end;

    FSourceLinkKeyList.Clear;
    FCloneLinkKeyList.Clear;


    //CD078A2
    FReader.Close;
    FReader.SQL.Text := Format(
      ' SELECT * FROM %s.CD078A2 WHERE BPCODE = ''%s'' ', [AReadOwner, ABPCode] );
    FReader.Open;
    FReader.First;
    while not FReader.Eof do
    begin
      FWriter.Close;
      FWriter.SQL.Text := Format(
        ' INSERT INTO %s.CD078A2 (                               ' +
        '   BPCODE, CITEMCODE, SERVICETYPE, DEPOSITCODE,         ' +
        '   PTCODE, PTNAME, DEPOSITAMT, DEPOSITDEFAULT  )        ' +
        ' VALUES (                                               ' +
        '  ''%s'', ''%s'', ''%s'', ''%s'',                       ' +
        '  ''%s'', ''%s'', ''%s'', ''%s''  )                     ',
        [ATableOwner,
         ABPCode,
         FReader.FieldByName( 'CITEMCODE' ).AsString,
         FReader.FieldByName( 'SERVICETYPE' ).AsString,
         FReader.FieldByName( 'DEPOSITCODE' ).AsString,
         FReader.FieldByName( 'PTCODE' ).AsString,
         FReader.FieldByName( 'PTNAME' ).AsString,
         FReader.FieldByName( 'DEPOSITAMT' ).AsString,
         FReader.FieldByName( 'DEPOSITDEFAULT' ).AsString] );
      FWriter.ExecSQL;   
      FReader.Next;
    end;


    //CD078A3
    FReader.Close;
    FReader.SQL.Text := Format(
      ' SELECT * FROM %s.CD078A3 WHERE BPCODE = ''%s'' ', [AReadOwner, ABPCode] );
    FReader.Open;
    FReader.First;
    while not FReader.Eof do
    begin
      FWriter.Close;
      FWriter.SQL.Text := Format(
        ' INSERT INTO %s.CD078A3 (                               ' +
        '   BPCODE, CITEMCODE, SERVICETYPE, PENALCODE,            ' +
        '   ITEM, PENALMON, PENALAMT  )                          ' +
        ' VALUES (                                               ' +
        '  ''%s'', ''%s'', ''%s'', ''%s'',                       ' +
        '  ''%s'', ''%s'', ''%s''  )                             ',
        [ATableOwner,
         ABPCode,
         FReader.FieldByName( 'CITEMCODE' ).AsString,
         FReader.FieldByName( 'SERVICETYPE' ).AsString,
         FReader.FieldByName( 'PENALCODE' ).AsString,
         FReader.FieldByName( 'ITEM' ).AsString,
         FReader.FieldByName( 'PENALMON' ).AsString,
         FReader.FieldByName( 'PENALAMT' ).AsString] );
      FWriter.ExecSQL;   
      FReader.Next;
    end;

    //CD078A4
    FReader.Close;
    FReader.SQL.Text := Format(
      ' SELECT * FROM %s.CD078A4 WHERE BPCODE = ''%s'' ', [AReadOwner, ABPCode] );
    FReader.Open;
    FReader.First;
    while not FReader.Eof do
    begin
      FWriter.Close;
      FWriter.SQL.Text := Format(
        ' INSERT INTO %s.CD078A4 (                                 ' +
        '   BPCODE, CITEMCODE, SERVICETYPE, PENALCODE,             ' +
        '   ITEM, MONTHSTART, MONTHSTOP, MONTHAMT, DECREASEAMT )   ' +
        ' VALUES (                                                 ' +
        '  ''%s'', ''%s'', ''%s'', ''%s'',                         ' +
        '  ''%s'', ''%s'', ''%s'', ''%s'', ''%s''  )               ',
        [ATableOwner,
         ABPCode,
         FReader.FieldByName( 'CITEMCODE' ).AsString,
         FReader.FieldByName( 'SERVICETYPE' ).AsString,
         FReader.FieldByName( 'PENALCODE' ).AsString,
         FReader.FieldByName( 'ITEM' ).AsString,
         FReader.FieldByName( 'MONTHSTART' ).AsString,
         FReader.FieldByName( 'MONTHSTOP' ).AsString,
         FReader.FieldByName( 'MONTHAMT' ).AsString,
         FReader.FieldByName( 'DECREASEAMT' ).AsString] );
      FWriter.ExecSQL;
      FReader.Next;
    end;

    //CD078B
    FReader.Close;
    FReader.SQL.Text := Format (
      ' SELECT * FROM %s.CD078B WHERE BPCODE=''%s'' ', [AReadOwner, ABPCode] );
    FReader.Open;
    FReader.First;
    while not FReader.Eof do
    begin
      FWriter.Close;
      FWriter.SQL.Text := Format (
        ' INSERT INTO %s.CD078B (                                        ' +
        '   BPCODE, NEWBPCODE, INSTCODE, INSTNAME, SERVICETYPE, GROUPNO, ' +
        '   GROUPNAME, WORKUNIT, REFNO, INTERDEPEND,                     ' +
        '   PRCODE, PRNAME  )                                            ' +
        ' VALUES (                                                       ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',      ' +
        '   ''%s'', ''%s'', ''%s'',                                      ' +
        '   ''%s'', ''%s''  )                                            ',
        [ATableOwner,
         ABPCode,
         FReader.FieldByName('NEWBPCODE').AsString,
         FReader.FieldByName('INSTCODE').AsString,
         FReader.FieldByName('INSTNAME').AsString,
         FReader.FieldByName('SERVICETYPE').AsString,
         FReader.FieldByName('GROUPNO').AsString,
         FReader.FieldByName('GROUPNAME').AsString,
         FReader.FieldByName('WORKUNIT').AsString,
         FReader.FieldByName('REFNO').AsString,
         FReader.FieldByName('INTERDEPEND').AsString,
         FReader.FieldByName('PRCODE').AsString,
         FReader.FieldByName('PRNAME').AsString] );
      FWriter.ExecSQL;
      FReader.Next
    end;

    //CD078B1
    FReader.Close;
    FReader.SQL.Text := Format (
      ' SELECT * FROM %s.CD078B1 WHERE BPCODE=''%s'' ', [AReadOwner, ABPCode] );
    FReader.Open;
    FReader.First;
    while not FReader.Eof do
    begin
      FWriter.Close;
      FWriter.SQL.Text := Format (
        ' INSERT INTO %s.CD078B1 (                                       ' +
        '   BPCODE, INSTCODE, SERVICETYPE, INSTCODE2, INSTNAME2,         ' +
        '   PRCODE2, PRNAME2  )                                          ' +
        ' VALUES (                                                       ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s''  )    ',
        [ATableOwner,
         ABPCode,
         FReader.FieldByName( 'INSTCODE' ).AsString,
         FReader.FieldByName( 'SERVICETYPE' ).AsString,
         FReader.FieldByName( 'INSTCODE2' ).AsString,
         FReader.FieldByName( 'INSTNAME2' ).AsString,
         FReader.FieldByName( 'PRCODE2' ).AsString,
         FReader.FieldByName( 'PRNAME2' ).AsString] );
      FWriter.ExecSQL;
      FReader.Next
    end;


    //CD078C
    FReader.Close;
    FReader.SQL.Text := Format (
      ' SELECT * FROM %s.CD078C WHERE BPCODE=''%s'' ', [AReadOwner, ABPCode] );
    FReader.Open;
    FReader.First;
    while not FReader.Eof do
    begin
      FWriter.Close;
      FWriter.SQL.Text := Format (
        ' INSERT INTO %s.CD078C (                                        ' +
        '   BPCODE, NEWBPCODE, CITEMCODE, CITEMNAME, SERVICETYPE,        ' +
        '   FACIITEM, AMOUNT, DISCOUNTAMT, PUNISH, PENALTYPE,            ' +
        '   REFCITEMCODE, REFCITEMNAME, INSTCODESTR )                    ' +
        ' VALUES (                                                       ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',      ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'' )             ',
        [ATableOwner,
        ABPCode,
        FReader.FieldByName('NEWBPCODE').AsString,
        FReader.FieldByName('CITEMCODE').AsString,
        FReader.FieldByName('CITEMNAME').AsString,
        FReader.FieldByName('SERVICETYPE').AsString,
        FReader.FieldByName('FACIITEM').AsString,
        FReader.FieldByName('AMOUNT').AsString,
        FReader.FieldByName('DISCOUNTAMT').AsString,
        FReader.FieldByName('PUNISH').AsString,
        FReader.FieldByName('PENALTYPE').AsString,
        FReader.FieldByName('REFCITEMCODE').AsString,
        FReader.FieldByName('REFCITEMNAME').AsString,
        FReader.FieldByName('INSTCODESTR').AsString] );
      FWriter.ExecSQL;
      FReader.Next
    end;

    //CD078D
    FReader.Close;
    FReader.SQL.Text := Format (
      ' SELECT * FROM %s.CD078D WHERE BPCODE=''%s'' ', [AReadOwner, ABPCode] );
    FReader.Open;
    FReader.First;
    while not FReader.Eof do
    begin
      {#5859 增加CBPCHANGEFIXIP、CBPCHANGEDYNIP By Kin 2011/04/27}
      FWriter.Close;
      FWriter.SQL.Text := Format (
        ' INSERT INTO %s.CD078D (                                        ' +
        '   BPCODE, NEWBPCODE, FACIITEM, FACITYPE, FACICODE, FACINAME,   ' +
        '   SERVICETYPE, MODELCODE, MODELNAME, BUYCODE, BUYNAME,         ' +
        '   CMBAUDNO, CMBAUDRATE, FIXIPCOUNT, DYNIPCOUNT, EMCCM,         ' +
        '   EMCCMCP, CMCPDYNIP, CMCPFIXIP, INSTCODESTR,                  ' +
        '   CBPCHANGEFIXIP, CBPCHANGEDYNIP )                             ' +
        ' VALUES (                                                       ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',      ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',      ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',              ' +
        '   ''%s'', ''%s''     )                                         ',
        [ATableOwner,
        ABPCode,
        FReader.FieldByName('NEWBPCODE').AsString,
        FReader.FieldByName('FACIITEM').AsString,
        FReader.FieldByName('FACITYPE').AsString,
        FReader.FieldByName('FACICODE').AsString,
        FReader.FieldByName('FACINAME').AsString,
        FReader.FieldByName('SERVICETYPE').AsString,
        FReader.FieldByName('MODELCODE').AsString,
        FReader.FieldByName('MODELNAME').AsString,
        FReader.FieldByName('BUYCODE').AsString,
        FReader.FieldByName('BUYNAME').AsString,
        FReader.FieldByName('CMBAUDNO').AsString,
        FReader.FieldByName('CMBAUDRATE').AsString,
        FReader.FieldByName('FIXIPCOUNT').AsString,
        FReader.FieldByName('DYNIPCOUNT').AsString,
        FReader.FieldByName('EMCCM').AsString,
        FReader.FieldByName('EMCCMCP').AsString,
        FReader.FieldByName('CMCPDYNIP').AsString,
        FReader.FieldByName('CMCPFIXIP').AsString,
        FReader.FieldByName('INSTCODESTR').AsString,
        FReader.FieldByName( 'CBPCHANGEFIXIP' ).AsString,
        FReader.FieldByName( 'CBPCHANGEDYNIP' ).AsString ] );
      FWriter.ExecSQL;
      FReader.Next
    end;

    //CD078E
    FReader.Close;
    FReader.SQL.Text := Format ( ' SELECT * FROM %s.CD078E WHERE BPCODE=''%s'' ',
      [AReadOwner, ABPCode] );
    FReader.Open;
    while not FReader.Eof do
    begin
      DataWriter.Close;
      DataWriter.SQL.Text := Format (
        ' INSERT INTO %s.CD078E (                                               ' +
        '   BPCODE, NEWBPCODE, NEWBPNAME, CHANGESTARTDATE, CHANGESTOPDATE )     ' +
        ' VALUES (                                                              ' +
        '   ''%s'', ''%s'', ''%s'', TO_DATE( ''%s'', ''YYYYMMDD'' ),            ' +
        '   TO_DATE( ''%s'', ''YYYYMMDD'' ) )                                   ',
        [ATableOwner,
        FReader.FieldByName('BPCODE').AsString,
        FReader.FieldByName('NEWBPCODE').AsString,
        FReader.FieldByName('NEWBPNAME').AsString,
        aDateSt,
        aDateEd] );
      DataWriter.ExecSQL;
      FReader.Next
    end;

    //CD078F
    FReader.Close;
    FReader.SQL.Text := Format ( ' SELECT * FROM %s.CD078F WHERE BPCODE=''%s'' ',
      [AReadOwner, ABPCode] );
    FReader.Open;
    while not FReader.Eof do
    begin
      DataWriter.Close;
      DataWriter.SQL.Text := Format (
        ' INSERT INTO %s.CD078F (                                               ' +
        '   BPCODE, NEWBPCODE, SERVICETYPE, OLDCITEMCODE, OLDCITEMNAME,         ' +
        '   NEWCITEMCODE, NEWCITEMNAME, CHANGETYPE, PENALTYPE, PENAL,           ' +
        '   DEPOSITTYPE, FACIITEM )                                             ' +
        ' VALUES (                                                              ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',     ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'' )                                    ',
        [ATableOwner,
         FReader.FieldByName('BPCODE').AsString,
         FReader.FieldByName('NEWBPCODE').AsString,
         FReader.FieldByName('SERVICETYPE').AsString,
         FReader.FieldByName('OLDCITEMCODE').AsString,
         FReader.FieldByName('OLDCITEMNAME').AsString,
         FReader.FieldByName('NEWCITEMCODE').AsString,
         FReader.FieldByName('NEWCITEMNAME').AsString,
         FReader.FieldByName('CHANGETYPE').AsString,
         FReader.FieldByName('PENALTYPE').AsString,
         FReader.FieldByName('PENAL').AsString,
         FReader.FieldByName('DEPOSITTYPE').AsString,
         FReader.FieldByName('FACIITEM').AsString] );
      DataWriter.ExecSQL;
      FReader.Next
    end;
    FReader.Close;

    //CD078G
    FReader.Close;
    FReader.SQL.Text := Format (
        ' SELECT * FROM %s.CD078G WHERE BPCODE=''%s'' ',
        [AReadOwner, ABPCode] );
    FReader.Open;
    while not FReader.Eof do
    begin
      DataWriter.Close;
      DataWriter.SQL.Text := Format (
         ' INSERT INTO %s.CD078G ( BPCODE,CitemCode,CitemName,           ' +
         '   ServiceType,FaciItem,SalePointcode,                         ' +
          '   SalePointName,InstCodeStr )                                ' +
         ' VALUES ( ''%s'', %s , ''%s'', ''%s'',                         ' +
          '       ''%s'', ''%s'', ''%s'',''%s''   )                      ' ,
        [ ATableOwner,
        FReader.FieldByName('BPCODE').AsString,
        FReader.FieldByName( 'CitemCode' ).AsString,
        FReader.FieldByName( 'CitemName' ).AsString,
        FReader.FieldByName( 'ServiceType' ).AsString,
        FReader.FieldByName( 'FaciItem' ).AsString,
        FReader.FieldByName( 'SalePointCode' ).AsString,
        FReader.FieldByName( 'SalePointName').AsString,
        FReader.FieldByName( 'InstCodeStr').AsString ] );
      DataWriter.ExecSQL;
      FReader.Next
    end;

    //CD078G1
    FReader.Close;
    FReader.SQL.Text := Format (
        ' SELECT * FROM %s.CD078G1 WHERE BPCODE=''%s'' ',
        [AReadOwner, ABPCode] );
    FReader.Open;


    while not FReader.Eof do
    begin
      aCD078G1Step := GetCD078G1Step( ATableOwner );
      DataWriter.Close;
      DataWriter.SQL.Text := Format (
          ' INSERT INTO %s.CD078G1 ( BPCODE,CitemCode,                    ' +
          '   ServiceType,FaciItem,SalePointcode,SalePointName,           ' +
          '   StepNo,Description,MasterSale,DiscountAmt )                 ' +
          ' VALUES ( ''%s'', ''%s'' ,''%s'',                              ' +
          '       ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
          '       ''%s'', ''%s'' )                                        ' ,
        [ ATableOwner,
         FReader.FieldByName('BPCODE').AsString ,
         FReader.FieldByName( 'CitemCode' ).AsString ,
         FReader.FieldByName( 'ServiceType' ).AsString ,
         FReader.FieldByName( 'FaciItem' ).AsString ,
         FReader.FieldByName( 'SalePointCode' ).AsString ,
         FReader.FieldByName( 'SalePointName' ).AsString ,
         aCD078G1Step,
         FReader.FieldByName( 'Description' ).AsString ,
         FReader.FieldByName( 'MasterSale' ).AsString ,
         FReader.FieldByName( 'DiscountAmt' ).AsString ] );
      DataWriter.ExecSQL;
      FReader.Next
    end;

    //CD078I
    FReader.Close;
    FReader.SQL.Text := Format (
        ' SELECT * FROM %s.CD078I WHERE BPCODE=''%s'' ',
        [ AReadOwner, ABPCode ] );
    FReader.Open;
    while not FReader.Eof do
    begin
      DataWriter.Close;
      DataWriter.SQL.Text := Format (
         ' INSERT INTO %s.CD078I ( BPCODE,CitemCode,                     ' +
         '   ServiceType,FaciItem,DiscountCitemCode,                     ' +
          '   CASID,DiscountAmt,InstCodeStr )                            ' +
         ' VALUES ( ''%s'', %s , ''%s'', ''%s'',                         ' +
          '       ''%s'', ''%s'', ''%s'',''%s''   )                      ' ,
        [ ATableOwner,
        FReader.FieldByName('BPCODE').AsString ,
        FReader.FieldByName( 'CitemCode' ).AsString,
        FReader.FieldByName( 'ServiceType' ).AsString,
        FReader.FieldByName( 'FaciItem' ).AsString,
        FReader.FieldByName( 'DiscountCitemCode' ).AsString,
        FReader.FieldByName( 'CASID').AsString,
        FReader.FieldByName( 'DiscountAmt').AsString,
        FReader.FieldByName( 'InstCodeStr').AsString ] );
      DataWriter.ExecSQL;
      FReader.Next
    end;

    //CD078J
    FReader.Close;
    FReader.SQL.Text := Format (
        ' SELECT * FROM %s.CD078J WHERE BPCODE=''%s'' ',
        [ AReadOwner, ABPCode ] );
    FReader.Open;
    while not FReader.Eof do
    begin
      FWriter.Close;
      FWriter.SQL.Text := Format (
         ' INSERT INTO %s.CD078J ( BPCODE,PRODUCTCODE,                   ' +
         '   PRODUCTNAME,SERVICETYPE,FACIITEM,                           ' +
          '   INSTCODESTR )                                              ' +
         ' VALUES ( ''%s'', %s , ''%s'', ''%s'',                         ' +
          '       ''%s'', ''%s''               )                         ' ,
        [ ATableOwner,
        FReader.FieldByName('BPCODE').AsString ,
        FReader.FieldByName( 'PRODUCTCODE' ).AsString,
        FReader.FieldByName( 'PRODUCTNAME' ).AsString,
        FReader.FieldByName( 'SERVICETYPE' ).AsString,
        FReader.FieldByName( 'FACIITEM' ).AsString,
        FReader.FieldByName( 'INSTCODESTR').AsString ] );
      FWriter.ExecSQL;
      FReader.Next
    end;

     //CD078K
    FReader.Close;
    FReader.SQL.Text := Format (
        ' SELECT * FROM %s.CD078K WHERE BPCODE=''%s'' ',
        [ AReadOwner, ABPCode ] );
    FReader.Open;
    while not FReader.Eof do
    begin
      FWriter.Close;
      FWriter.SQL.Text := Format (
         ' INSERT INTO %s.CD078K ( BPCODE,CITEMCODE,                   ' +
         '   MASTERITEM,AMOUNT,IFRSPERCENT,                            ' +
          '   IFRSAMT )                                                ' +
         ' VALUES ( ''%s'', %s , ''%s'', ''%s'',                       ' +
          '       ''%s'', ''%s''               )                       ' ,
        [ ATableOwner,
        FReader.FieldByName('BPCODE').AsString,
        FReader.FieldByName( 'CITEMCODE' ).AsString,
        FReader.FieldByName( 'MASTERITEM' ).AsString,
        FReader.FieldByName( 'AMOUNT' ).AsString,
        FReader.FieldByName( 'IFRSPERCENT' ).AsString,
        FReader.FieldByName( 'IFRSAMT').AsString ] );
      FWriter.ExecSQL;
      FReader.Next
    end;

     //CD078L
    FReader.Close;
    FReader.SQL.Text := Format (
        ' SELECT * FROM %s.CD078L WHERE BPCODE=''%s'' ',
        [ AReadOwner, ABPCode ] );
    FReader.Open;
    while not FReader.Eof do
    begin
      FWriter.Close;
      FWriter.SQL.Text := Format (
         ' INSERT INTO %s.CD078L ( BPCODE,MONTH,                           ' +
         '   MASTERSERVICETYPE,CASHADJ,SUBTOTAL,                           ' +
         '   CMAMT,CMTAX,CMADJ ,                                           ' +
         '   DTVAMT,DTVTAX,DTVADJ,                                         ' +
         '   PRVAMT,PRVTAX,PRVADJ)                                         ' +
         ' VALUES ( ''%s'', ''%s'' , ''%s'', ''%s'',                       ' +
         '       ''%s'', ''%s'',''%s'', ''%s'',                            ' +
         '       ''%s'', ''%s'',''%s'', ''%s'',                            ' +
         '       ''%s'', ''%s''  ) ',
        [ ATableOwner,
        FReader.FieldByName('BPCODE').AsString,
        FReader.FieldByName( 'MONTH' ).AsString,
        FReader.FieldByName( 'MASTERSERVICETYPE' ).AsString,
        FReader.FieldByName( 'CASHADJ' ).AsString,
        FReader.FieldByName( 'SUBTOTAL' ).AsString,
        FReader.FieldByName( 'CMAMT').AsString ,
        FReader.FieldByName( 'CMTAX').AsString ,
        FReader.FieldByName( 'CMADJ').AsString ,
        FReader.FieldByName( 'DTVAMT').AsString ,
        FReader.FieldByName( 'DTVTAX').AsString ,
        FReader.FieldByName( 'DTVADJ').AsString ,
        FReader.FieldByName( 'PRVAMT').AsString ,
        FReader.FieldByName( 'PRVTAX').AsString ,
        FReader.FieldByName( 'PRVADJ').AsString,
        FReader.FieldByName( 'STOPFLAG').AsString  ]
         );
      FWriter.ExecSQL;
      FReader.Next
    end;
    FReader.Close;
    FWriter.Close;


    DataConnection.CommitTrans;
    {}
    txtMessage.Lines.Add( Format(
      '優惠組合:%s-%s  ---->  複製完成。', [ABPCode, ACodeName] ) );
    Application.ProcessMessages;  
  except
    on E: Exception do
    begin
      Result := False;
      DataConnection.RollbackTrans;
      txtMessage.Lines.Add( Format(
        '優惠組合 %s-%s  ---->  複製發生錯誤, 訊息:%s。', [ABPCode, ACodeName, E.Message] ) );
      AErrMessage := ( AErrMessage + Rpad( EmptyStr, 6, #32 ) +
        Format( '名稱%s , 原因:%s。'#13, [ACodeName, E.Message] ) );
      Application.ProcessMessages;  
    end;
  end;
  Screen.Cursor := crDefault;
end;

{ -------------------------------------------------------------------------------------- }

function TfrmCD078B9.GetCD078G1Step(const AOwner: String): String;
begin
  FWriter.Close;
  FWriter.SQL.Text := Format(
    ' SELECT %s.S_CD078G1.NEXTVAL FROM DUAL ', [AOwner] );
  FWriter.Open;
  Result := FWriter.Fields[0].AsString;
  FWriter.Close;
end;

end.
