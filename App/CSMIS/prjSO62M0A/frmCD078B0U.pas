unit frmCD078B0U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ActnList, StdCtrls, ExtCtrls, Grids, DBGrids, Buttons, Provider,
  DB, DBClient, Mask, ShellAPI,
  cbDBController, cbStyleController,
  cxClasses, cxGridStrs, cxEditConsts, dxSkinsCore, dxSkinBlack,
  dxSkinBlue, dxSkinCaramel, dxSkinCoffee, dxSkinsDefaultPainters,
  cxControls, cxContainer, cxEdit, cxTextEdit;

type

  TQueryParam = record
    Param1: String;
    Param2: String;
    Param3: String;
    Param4: String;
    Param5: String;
    Param6: String;
  end;

  TfrmCD078B0 = class(TForm)
    Panel1: TPanel;
    btnSearch: TButton;
    Panel2: TPanel;
    btnClear: TButton;
    btnPrint: TButton;
    btnInsert: TButton;
    btnModify: TButton;
    btnCopyLocal: TButton;
    btnCopyOther: TButton;
    btnQueryOther: TButton;
    Panel3: TPanel;
    DBGrid1: TDBGrid;
    CD078DataSet: TClientDataSet;
    CD078DataSetCODENO: TStringField;
    CD078DataSetDESCRIPTION: TStringField;
    CD078DataSetONSALESTARTDATE: TDateTimeField;
    CD078DataSetONSALESTOPDATE: TDateTimeField;
    CD078DataSetMASTERPROD: TIntegerField;
    CD078DataSetBUNDLE: TIntegerField;
    CD078DataSetBUNDLESAMEMON: TIntegerField;
    CD078DataSetBUNDLEMON: TIntegerField;
    CD078DataSetSAMEPERIOD: TIntegerField;
    CD078DataSetPERIOD: TIntegerField;
    CD078DataSetMERGEPRINT: TIntegerField;
    CD078DataSetPENAL: TIntegerField;
    CD078DataSetPENALTYPE: TIntegerField;
    CD078DataSetMERGEWORK: TIntegerField;
    CD078DataSetBUNDLEPRODNOTE: TStringField;
    CD078DataSetWORKNOTE: TStringField;
    CD078DataSetREFNO: TIntegerField;
    CD078DataSetUPDTIME: TStringField;
    CD078DataSetUPDEN: TStringField;
    CD078DataSetSTOPFLAG: TIntegerField;
    DataSource1: TDataSource;
    Panel4: TPanel;
    Bevel1: TBevel;
    Panel5: TPanel;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    Button5: TButton;
    Button6: TButton;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    SpeedButton3: TSpeedButton;
    SpeedButton4: TSpeedButton;
    SpeedButton5: TSpeedButton;
    SpeedButton6: TSpeedButton;
    edtBpCode: TEdit;
    edtPeriod: TEdit;
    edtServiceType: TEdit;
    edtBundleMon: TEdit;
    edtCItemCode: TEdit;
    edtInstCode: TEdit;
    GroupBox1: TGroupBox;
    chkMasterProd: TCheckBox;
    chkStopFlag: TCheckBox;
    Label1: TLabel;
    edtDateSt: TMaskEdit;
    Label2: TLabel;
    edtDateEd: TMaskEdit;
    CD078DataSetONSALESTARTCDATE: TStringField;
    CD078DataSetONSALESTOPCDATE: TStringField;
    ActionList1: TActionList;
    actSearch: TAction;
    actClear: TAction;
    CD078DataProvider: TDataSetProvider;
    actInsert: TAction;
    actPrint: TAction;
    actUpdate: TAction;
    Button7: TButton;
    btnFTP: TButton;
    CD078DataSetCanUseType: TIntegerField;
    lbl1: TLabel;
    edtCodeNo: TcxTextEdit;
    lbl2: TLabel;
    edtDescription: TcxTextEdit;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure edtDateStExit(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
    procedure SpeedButton4Click(Sender: TObject);
    procedure SpeedButton5Click(Sender: TObject);
    procedure SpeedButton6Click(Sender: TObject);
    procedure actSearchExecute(Sender: TObject);
    procedure actClearExecute(Sender: TObject);
    procedure actInsertExecute(Sender: TObject);
    procedure actPrintExecute(Sender: TObject);
    procedure actUpdateExecute(Sender: TObject);
    procedure btnCopyLocalClick(Sender: TObject);
    procedure CD078DataSetCalcFields(DataSet: TDataSet);
    procedure btnCopyOtherClick(Sender: TObject);
    procedure btnQueryOtherClick(Sender: TObject);
    procedure CD078DataSetAfterOpen(DataSet: TDataSet);
    procedure CD078DataSetAfterScroll(DataSet: TDataSet);
    procedure Button7Click(Sender: TObject);
    procedure DBGrid1DblClick(Sender: TObject);
    procedure btnFTPClick(Sender: TObject);
  private
    { Private declarations }
    FQueryParam: TQueryParam;
    FKeyValue: String;
    FCanUseEPG: Boolean;
    function RecvParamString: Boolean;
    function GetQueryParamString(const aIndex: Integer): String;
    procedure ChangeDevexComponentLanguage;
    procedure LoadData;
    procedure ShowEditorDataForm(const aMode: Integer; const aKeyValue: String);
  public
    { Public declarations }
    property KeyValue: String read FKeyValue;
  end;

var
  frmCD078B0: TfrmCD078B0;

implementation

uses cbUtilis, frmMultiSelectU, frmCD078B1U, frmCD078B8U, frmCD078B9U, frmCD078BAU ;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B0.FormCreate(Sender: TObject);
var
  aCreateError: Boolean;
begin

  DBController := TDBController.Create( nil );
  StyleController := TStyleController.Create( nil );
  aCreateError := ( not RecvParamString );
  if not aCreateError then DBController.Connected := True;
  aCreateError := ( not DBController.Connected );
  if ( aCreateError ) then
  begin
    Application.ShowMainForm := False;
    Application.Terminate;
    Exit;
  end;
  btnFTP.Enabled := DBController.GetUserPriv('SO62M06');
  FCanUseEPG := DBController.GetUserPriv( 'SO62M021' );
  {#7222 Checking button in autohority by Kin 2015/03/28 }
  btnInsert.Enabled := DBController.GetUserPriv('SO62M01');
  DBController.CanModiFy := DBController.GetUserPriv('SO62M02');
  btnModify.Enabled := DBController.CanModiFy;
  FKeyValue := EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B0.FormShow(Sender: TObject);
begin
  ChangeDevexComponentLanguage;
  Panel4.Caption := '0/0';
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B0.FormDestroy(Sender: TObject);
begin
  DBController.Free;
  StyleController.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B0.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  CanClose := True;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B0.RecvParamString: Boolean;
begin
  Result := ( ParamCount > 0 );
  if not Result then
  begin
    WarningMsg( '傳入之程式啟動參數有誤。' );
    Exit;
  end;
  { 傳入之參數 }
  { N 1 Nick 12 大安文山 emcdw emcdw mis }
  DBController.LoginInfo.StarPVOD := False;
  DBController.LoginInfo.IsSuperVisor := ( ParamStr( 1 ) = 'Y' );
  DBController.LoginInfo.UserId := ParamStr( 2 );
  DBController.LoginInfo.UserName := ParamStr( 3 );
  DBController.LoginInfo.CompCode := ParamStr( 4 );
  DBController.LoginInfo.CompName := ParamStr( 5 );
  DBController.LoginInfo.DbAccount := WideUpperCase( ParamStr( 6 ) );
  DBController.LoginInfo.DbPassword := ParamStr( 7 );
  DBController.LoginInfo.DbAliase := ParamStr( 8 );
  DBController.LoginInfo.ServiceType := ParamStr( 9 );
  {#6035 增加COM區的連線資訊 By Kin 2011/06/08}
  if ParamCount > 9 then
  begin
    DBController.LoginInfo.Com_DbAccount := ParamStr( 10 );
    DBController.LoginInfo.Com_DbPassword := ParamStr( 11 );
    DBController.LoginInfo.Com_DbAliase := ParamStr( 12 );
  end;
  if ( DBController.LoginInfo.Com_DbAccount <> EmptyStr ) and
        ( DBController.LoginInfo.Com_DbPassword <> EmptyStr ) and
        ( DBController.LoginInfo.Com_DbAliase <> EmptyStr ) then
  begin
    DBController.LoginInfo.StarPVOD := True;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B0.GetQueryParamString(const aIndex: Integer): String;
var
  aSource: String;
begin
  Result := EmptyStr;
  if not ( aIndex in [1..6] ) then Exit;
  case aIndex of
    { 組合產品 }
    1: aSource := FQueryParam.Param1;
    { 繳費期數 }
    2: aSource := FQueryParam.Param2;
    { 綁約服務類別 }
    3: aSource := FQueryParam.Param3;
    { 綁約月數 }
    4: aSource := FQueryParam.Param4;
    { 組合產品--收費項目 }
    5: aSource := FQueryParam.Param5;
    { 組合產品--派工類別 }
    6: aSource := FQueryParam.Param6;
  end;
  Result := aSource;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B0.ChangeDevexComponentLanguage;
begin
  cxSetResourceString( @scxGridDeletingConfirmationCaption, '確認' );
  cxSetResourceString( @scxGridDeletingFocusedConfirmationText, '刪除此筆資料?' );
  cxSetResourceString( @scxGridDeletingSelectedConfirmationText, '刪除所選取的資料?' );
  cxSetResourceString( @scxGridCustomizationFormCaption, '欄位選擇' );
  cxSetResourceString( @scxGridCustomizationFormBandsPageCaption, '項目' );
  cxSetResourceString( @scxGridCustomizationFormColumnsPageCaption, '欄位' );
  cxSetResourceString( @scxGridNoDataInfoText, '無資料可顯示' );
  cxSetResourceString( @cxSDatePopupToday, '今天' );
  cxSetResourceString( @cxSDatePopupClear, '清除' );
  cxSetResourceString( @cxSDatePopupNow, '現在' );
  cxSetResourceString( @cxSDateError, '日期有誤' );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B0.LoadData;

  { ----------------------------------------------------- }

  function AssemblyCD078ASql: String;
  var
    aValue: String;
  begin
    Result := EmptyStr;
    { 組合產品參數查詢條件, CD078A }
    { 繳費期數 }
    if ( edtPeriod.Text <> EmptyStr ) then
    begin
      aValue := GetQueryParamString( 2 );
      Result := Format(
        ' AND ( PERIOD1 IN ( %s ) OR PERIOD2 IN ( %s ) OR   ' +
        '       PERIOD3 IN ( %s ) OR PERIOD4 IN ( %s ) OR   ' +
        '       PERIOD5 IN ( %s ) OR PERIOD6 IN ( %s ) OR   ' +
        '       PERIOD7 IN ( %s ) OR PERIOD8 IN ( %s ) OR   ' +
        '       PERIOD9 IN ( %s ) OR PERIOD10 IN ( %s ) OR  ' +
        '       PERIOD11 IN ( %s ) OR PERIOD12 IN ( %s ) )  ',
        [aValue, aValue, aValue, aValue, aValue, aValue,
         aValue, aValue, aValue, aValue, aValue, aValue] );
    end;
    { 綁約服務類別 }
    if ( edtServiceType.Text <> EmptyStr ) then
    begin
      Result := ( Result + Format( ' AND SERVICETYPE IN ( %s ) ',
        [GetQueryParamString( 3 )] ) );
    end;
    { 綁約月數 }
    if ( edtBundleMon.Text <> EmptyStr ) then
    begin
      Result := ( Result + Format( ' AND BUNDLEMON IN ( %s ) ',
        [GetQueryParamString( 4 )] ) );
    end;
    { 收費項目 }
    if ( edtCItemCode.Text <> EmptyStr ) then
    begin
      Result := ( Result + Format( ' AND CITEMCODE IN ( %s ) ',
        [GetQueryParamString( 5 )] ) );
    end;

    if ( Result <> EmptyStr ) then
      Result := ( Format( '  SELECT DISTINCT BPCODE FROM %s.CD078A WHERE 1 = 1 ',
        [DBController.LoginInfo.DbAccount] ) + Result);
  end;

  { ----------------------------------------------------- }

  function AssemblyCD078BSql: String;
  begin
    Result := EmptyStr;
    { 裝機類別 }
    if ( edtInstCode.Text <> EmptyStr ) then
    begin
      Result := ( Result + Format( ' AND INSTCODE IN ( %s ) ',
        [GetQueryParamString( 6 )] ) );
    end;
    if ( Result <> EmptyStr ) then
    begin
      Result := ( Format( '  SELECT DISTINCT BPCODE FROM %s.CD078B WHERE 1 = 1 ',
        [DBController.LoginInfo.DbAccount] ) + Result );
    end;
  end;

  { ----------------------------------------------------- }

var
  aSelect, aFrom, aWhere, aDateStr1, aDateStr2: String;
  aSubSelectCD078A, aSubSelectCD078B: String;
begin
  {#5683 增加繳付類別 By Kin 2010/08/16}
  aSubSelectCD078A := AssemblyCD078ASql;
  aSubSelectCD078B := AssemblyCD078BSql;
  aSelect :=
    ' SELECT A.CODENO,              ' +
    '        A.DESCRIPTION,         ' +
    '        A.ONSALESTARTDATE,     ' +
    '        A.ONSALESTOPDATE,      ' +
    '        A.MASTERPROD,          ' +
    '        A.BUNDLE,              ' +
    '        A.BUNDLESAMEMON,       ' +
    '        A.BUNDLEMON,           ' +
    '        A.SAMEPERIOD,          ' +
    '        A.PERIOD,              ' +
    '        A.MERGEPRINT,          ' +
    '        A.PENAL,               ' +
    '        A.PENALTYPE,           ' +
    '        A.MERGEWORK,           ' +
    '        A.BUNDLEPRODNOTE,      ' +
    '        A.WORKNOTE,            ' +
    '        A.REFNO,               ' +
    '        A.UPDTIME,             ' +
    '        A.UPDEN,               ' +
    '        A.STOPFLAG,            ' +
    '        A.CANPAYNOW,           ' +
    '        A.CANUSETYPE,          ' +
    '        A.EPGIMAGEURL,         ' +
    '        A.EPGORD               ';
    
  aFrom := Format( ' FROM %s.CD078 A ', [DBController.LoginInfo.DbAccount] );
  aWhere := ' WHERE 1 = 1 ';

  if ( aSubSelectCD078A <> EmptyStr ) then
  begin
    aFrom := ( Format( aFrom + ', ( %s ) B ', [aSubSelectCD078A] ) );
    aWhere := ( aWhere + ' AND A.CODENO = B.BPCODE  ' );
  end;

  if ( aSubSelectCD078B <> EmptyStr ) then
  begin
    aFrom := ( Format( aFrom + ', ( %s ) C ', [aSubSelectCD078B] ) );
    aWhere := ( aWhere + ' AND A.CODENO = C.BPCODE ' );
  end;

  { 主資料表的查詢條件, CD078 }
  { 組合產品 }
  if ( edtBpCode.Text <> EmptyStr ) then
  begin
    aWhere := ( aWhere + Format( ' AND A.CODENO IN ( %s ) ',
      [GetQueryParamString( 1 )] ) );
  end;
  { 主推商品 }
  if chkMasterProd.Checked then
  begin
    aWhere := ( aWhere + ' AND A.MASTERPROD = 1 ' );
  end;
  { 停用是否顯示 }
  if ( chkStopFlag.Checked ) then
    aWhere := ( aWhere + ' AND A.STOPFLAG IN ( 0, 1 ) ' )
  else
    aWhere := ( aWhere + ' AND A.STOPFLAG = 0 ' );

  { 上架日 }
  if edtDateSt.Text <> EmptyStr then
  begin
    aDateStr1 := TrimChar( edtDateSt.Text, ['/'] );
    aWhere := ( aWhere + Format( ' AND A.ONSALESTARTDATE >= TO_DATE( ''%s'', ''YYYYMMDD'' ) ', [aDateStr1] ) );
  end;
  { #7673 優惠代碼 }
  if edtCodeNo.Text <> EmptyStr then
  begin
    aWhere := ( aWhere + Format( ' AND INSTR(A.CODENO,''%s'',-1,1) > 0 ',[edtCodeNo.Text]));
  end;
  { 優惠名稱 }

  if edtDescription.Text <> EmptyStr then
  begin
    aWhere := ( aWhere + Format( ' AND INSTR(A.Description,''%s'',-1,1) > 0 ',[edtDescription.Text]));
  end;

  { 上架日 }
  if edtDateEd.Text <> EmptyStr then
  begin
    aDateStr2 := TrimChar( edtDateEd.Text, ['/'] );
    aWhere := ( aWhere + Format( ' AND A.ONSALESTOPDATE <= TO_DATE( ''%s'', ''YYYYMMDD'' ) ', [aDateStr2] ) );
  end;

  DBController.DataReader.Close;
  CD078DataSet.Close;
  CD078DataSet.CommandText := ( aSelect + aFrom + aWhere + ' ORDER BY A.CODENO ' );
  Screen.Cursor := crSQLWait;
  try
    try
      CD078DataSet.Open;
      if ( FKeyValue <> EmptyStr ) and ( not CD078DataSet.IsEmpty ) then
        CD078DataSet.Locate( 'CODENO', FKeyValue, [] );
    except
      on E: Exception do
      begin
        ErrorMsg( Format( '查詢優惠組合有誤, 原因:%s。', [E.Message] ) );
      end;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B0.ShowEditorDataForm(const aMode: Integer;
  const aKeyValue: String);
begin
  frmCD078B1 := TfrmCD078B1.Create( TDML( aMode ), aKeyValue );
  try
    frmCD078B1.ShowModal;
    FKeyValue := frmCD078B1.KeyValue;
  finally
    frmCD078B1.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B0.edtDateStExit(Sender: TObject);
var
  aDate, aErrMsg: String;
begin
  aDate := TMaskEdit( Sender ).EditText;
  if not DateTextIsValidEx( aDate ) then
  begin
    if Sender = edtDateSt then
      aErrMsg := '您輸入的產品上架日(起)有誤。'
    else if Sender = edtDateEd then
      aErrMsg := '您輸入的產品上架日(迄)有誤。';
  end;
  if ( aErrMsg <> EmptyStr ) then
  begin
    WarningMsg( aErrMsg );
    TMaskEdit( Sender ).SetFocus;
  end else
  begin
    TMaskEdit( Sender ).EditText := aDate;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B0.Button1Click(Sender: TObject);
var
  aWhere: string;
begin
  if ( chkStopFlag.Checked ) then
    aWhere := ' STOPFLAG IN ( 0, 1 ) '
  else
    aWhere := ' STOPFLAG = 0 ';
  frmMultiSelect := TfrmMultiSelect.Create( Application );
  try
    frmMultiSelect.Connection := DBController.DataConnection;
    frmMultiSelect.KeyFields := 'CODENO';
    frmMultiSelect.KeyValues := FQueryParam.Param1;
    frmMultiSelect.DisplayFields := 'CODENO,優惠組合代碼,DESCRIPTION,優惠組合名稱';
    frmMultiSelect.ResultFields := 'DESCRIPTION';
    frmMultiSelect.SQL.Text := Format(
      ' SELECT CODENO, DESCRIPTION FROM %s.CD078 WHERE ' + aWhere + ' ORDER BY CODENO ',
      [DBController.LoginInfo.DbAccount] );

    if frmMultiSelect.ShowModal = mrOk then
    begin
      FQueryParam.Param1 := frmMultiSelect.SelectedValue;
      edtBpCode.Text := frmMultiSelect.SelectedDisplay;
    end;
  finally
    frmMultiSelect.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B0.Button2Click(Sender: TObject);
var
  aOwner: String;
begin
  aOwner := DBController.LoginInfo.DbAccount;
  frmMultiSelect := TfrmMultiSelect.Create( Application );
  try
    frmMultiSelect.Connection := DBController.DataConnection;
    frmMultiSelect.KeyFields := 'PERIOD';
    frmMultiSelect.KeyValues := FQueryParam.Param2;
    frmMultiSelect.DisplayFields := 'PERIOD,繳費期數';
    frmMultiSelect.ResultFields := 'PERIOD';
    frmMultiSelect.SQL.Text := Format(
      ' SELECT PERIOD1 AS PERIOD FROM (      ' +
      '    SELECT PERIOD1 FROM %s.CD078A     ' +
      '     UNION                            ' +
      '    SELECT PERIOD2 FROM %s.CD078A     ' +
      '     UNION                            ' +
      '    SELECT PERIOD3 FROM %s.CD078A     ' +
      '     UNION                            ' +
      '    SELECT PERIOD4 FROM %s.CD078A     ' +
      '     UNION                            ' +
      '    SELECT PERIOD5 FROM %s.CD078A     ' +
      '     UNION                            ' +
      '    SELECT PERIOD6 FROM %s.CD078A     ' +
      '     UNION                            ' +
      '    SELECT PERIOD7 FROM %s.CD078A     ' +
      '     UNION                            ' +
      '    SELECT PERIOD8 FROM %s.CD078A     ' +
      '     UNION                            ' +
      '    SELECT PERIOD9 FROM %s.CD078A     ' +
      '     UNION                            ' +
      '    SELECT PERIOD10 FROM %s.CD078A    ' +
      '     UNION                            ' +
      '    SELECT PERIOD11 FROM %s.CD078A    ' +
      '     UNION                            ' +
      '    SELECT PERIOD12 FROM %s.CD078A  ) ' +
      ' WHERE PERIOD1 IS NOT NULL            ',
      [aOwner, aOwner, aOwner, aOwner, aOwner, aOwner,
       aOwner, aOwner, aOwner, aOwner, aOwner, aOwner] );
    if frmMultiSelect.ShowModal = mrOk then
    begin
      FQueryParam.Param2 := frmMultiSelect.SelectedValue;
      edtPeriod.Text := frmMultiSelect.SelectedDisplay;
    end;
  finally
    frmMultiSelect.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B0.Button3Click(Sender: TObject);
begin
  frmMultiSelect := TfrmMultiSelect.Create( Application );
  try
    frmMultiSelect.Connection := DBController.DataConnection;
    frmMultiSelect.KeyFields := 'CODENO';
    frmMultiSelect.KeyValues := FQueryParam.Param3;
    frmMultiSelect.DisplayFields := 'CODENO,服務別,DESCRIPTION,代碼';
    frmMultiSelect.ResultFields := 'DESCRIPTION';
    frmMultiSelect.SQL.Text := Format(
      ' SELECT CODENO, DESCRIPTION FROM %s.CD046 ORDER BY CODENO ',
      [DBController.LoginInfo.DbAccount] );
    if frmMultiSelect.ShowModal = mrOk then
    begin
      FQueryParam.Param3 := frmMultiSelect.SelectedValue;
      edtServiceType.Text := frmMultiSelect.SelectedDisplay;
    end;
    edtCItemCode.Text := EmptyStr;
    FQueryParam.Param5 := EmptyStr;
    edtInstCode.Text := EmptyStr;
    FQueryParam.Param6 := EmptyStr;
  finally
    frmMultiSelect.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B0.Button4Click(Sender: TObject);
begin
  frmMultiSelect := TfrmMultiSelect.Create( Application );

  try
    frmMultiSelect.Connection := DBController.DataConnection;
    frmMultiSelect.KeyFields := 'BUNDLEMON';
    frmMultiSelect.KeyValues := FQueryParam.Param4;
    frmMultiSelect.DisplayFields := 'BUNDLEMON,綁約月數';
    frmMultiSelect.ResultFields := 'BUNDLEMON';
    frmMultiSelect.SQL.Text := Format(
      ' SELECT DISTINCT BUNDLEMON FROM %s.CD078A ' +
      '  WHERE BUNDLEMON IS NOT NULL             ',
      [DBController.LoginInfo.DbAccount] );
    if frmMultiSelect.ShowModal = mrOk then
    begin
      FQueryParam.Param4 := frmMultiSelect.SelectedValue;
      edtBundleMon.Text := frmMultiSelect.SelectedDisplay;
    end;
  finally
    frmMultiSelect.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B0.Button5Click(Sender: TObject);
begin
  frmMultiSelect := TfrmMultiSelect.Create( Application );
  try
    frmMultiSelect.Connection := DBController.DataConnection;
    frmMultiSelect.KeyFields := 'CODENO';
    frmMultiSelect.KeyValues := FQueryParam.Param5;
    frmMultiSelect.DisplayFields := 'CODENO,項目代碼,DESCRIPTION,收費項目';
    frmMultiSelect.ResultFields := 'DESCRIPTION';
    frmMultiSelect.SQL.Text := Format(
      ' SELECT CODENO, DESCRIPTION FROM %s.CD019 ' +
      '  WHERE PERIODFLAG = 1 AND SIGN = ''+'' AND STOPFLAG = 0 ',
      [DBController.LoginInfo.DbAccount] );
    //加入服務類別條件
    if ( edtServiceType.Text <> EmptyStr ) then
      frmMultiSelect.SQL.Text := frmMultiSelect.SQL.Text +
        Format( ' AND SERVICETYPE IN ( %s ) ',
        [GetQueryParamString( 3 )]  );
    if frmMultiSelect.ShowModal = mrOk then
    begin
      FQueryParam.Param5 := frmMultiSelect.SelectedValue;
      edtCItemCode.Text := frmMultiSelect.SelectedDisplay;
    end;
  finally
    frmMultiSelect.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B0.Button6Click(Sender: TObject);
begin
  frmMultiSelect := TfrmMultiSelect.Create( Application );
  try
    frmMultiSelect.Connection := DBController.DataConnection;
    frmMultiSelect.KeyFields := 'CODENO';
    frmMultiSelect.KeyValues := FQueryParam.Param6;
    frmMultiSelect.DisplayFields := 'CODENO,類別代碼,DESCRIPTION,裝機類別';
    frmMultiSelect.ResultFields := 'DESCRIPTION';
    frmMultiSelect.SQL.Text := Format(
      ' SELECT CODENO, DESCRIPTION FROM %s.CD005 ' +
      '  WHERE STOPFLAG = 0 ',
      [DBController.LoginInfo.DbAccount] );
    //加入服務類別條件
    if ( edtServiceType.Text <> EmptyStr ) then
      frmMultiSelect.SQL.Text := frmMultiSelect.SQL.Text +
        Format( ' AND SERVICETYPE IN ( %s ) ',
        [GetQueryParamString( 3 )]  );
    frmMultiSelect.SQL.Text := frmMultiSelect.SQL.Text +
      ' ORDER BY CODENO ';
    if frmMultiSelect.ShowModal = mrOk then
    begin
      FQueryParam.Param6 := frmMultiSelect.SelectedValue;
      edtInstCode.Text := frmMultiSelect.SelectedDisplay;
    end;
  finally
    frmMultiSelect.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B0.SpeedButton1Click(Sender: TObject);
begin
  edtBpCode.Text := EmptyStr;
  FQueryParam.Param1 := EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B0.SpeedButton2Click(Sender: TObject);
begin
  edtPeriod.Text := EmptyStr;
  FQueryParam.Param2 := EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B0.SpeedButton3Click(Sender: TObject);
begin
  edtServiceType.Text := EmptyStr;
  FQueryParam.Param3 := EmptyStr;
  edtCItemCode.Text := EmptyStr;
  FQueryParam.Param5 := EmptyStr;
  edtInstCode.Text := EmptyStr;
  FQueryParam.Param6 := EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B0.SpeedButton4Click(Sender: TObject);
begin
  edtBundleMon.Text := EmptyStr;
  FQueryParam.Param4 := EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B0.SpeedButton5Click(Sender: TObject);
begin
  edtCItemCode.Text := EmptyStr;
  FQueryParam.Param5 := EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B0.SpeedButton6Click(Sender: TObject);
begin
  edtInstCode.Text := EmptyStr;
  FQueryParam.Param6 := EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B0.actSearchExecute(Sender: TObject);
begin
  LoadData;
  ActiveControl := DBGrid1;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B0.actClearExecute(Sender: TObject);
begin
  edtBpCode.Text := EmptyStr;
  FQueryParam.Param1 := EmptyStr;
  edtPeriod.Text := EmptyStr;
  FQueryParam.Param2 := EmptyStr;
  edtServiceType.Text := EmptyStr;
  FQueryParam.Param3 := EmptyStr;
  edtBundleMon.Text := EmptyStr;
  FQueryParam.Param4 := EmptyStr;
  edtCItemCode.Text := EmptyStr;
  FQueryParam.Param5 := EmptyStr;
  edtInstCode.Text := EmptyStr;
  FQueryParam.Param6 := EmptyStr;
  edtDateSt.Text := EmptyStr;
  edtDateEd.Text := EmptyStr;
  edtCodeNo.Text := EmptyStr;
  edtDescription.Text := EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B0.actPrintExecute(Sender: TObject);
begin
  {}
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B0.actInsertExecute(Sender: TObject);
begin
  ShowEditorDataForm( Ord( dmInsert ), EmptyStr );
  if ( FKeyValue <> EmptyStr ) then
    actSearch.Execute;
  ActiveControl := DBGrid1;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B0.actUpdateExecute(Sender: TObject);
  var edtMode : Integer;
begin
  edtMode := Ord( dmUpdate );
  if not DBController.CanModify then edtMode := Ord( dmBrowser );
  if ( CD078DataSet.IsEmpty ) then Exit;
  ShowEditorDataForm( edtMode ,
    CD078DataSet.FieldByName( 'CODENO' ).AsString );
  ActiveControl := DBGrid1;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B0.btnCopyLocalClick(Sender: TObject);
begin
  frmCD078B8 := TfrmCD078B8.Create(Application);
  try
    frmCD078B8.ShowModal;
    FKeyValue := frmCD078B8.CloneCodeNo;
  finally
    frmCD078B8.Free;
  end;
  if ( FKeyValue <> EmptyStr ) then actSearch.Execute;
  ActiveControl := DBGrid1;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B0.btnCopyOtherClick(Sender: TObject);
begin
  frmCD078B9 := TfrmCD078B9.Create(Application);
  try
    frmCD078B9.ShowModal;
  finally
    frmCD078B9.Free;
  end;
  if ( FKeyValue <> EmptyStr ) then actSearch.Execute;
  ActiveControl := DBGrid1;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B0.btnQueryOtherClick(Sender: TObject);
var
  aPath, aParmText: String;
  aIndex: Integer;
  aFirstEnter: Boolean;
  aParaCount : Integer;
begin
  aPath := IncludeTrailingPathDelimiter( ExtractFilePath( ParamStr( 0 ) ) );
  aFirstEnter := True;
  aParaCount := 9;
  for aIndex := 1 to aParaCount do
  begin
    if aFirstEnter then
    begin
      aParmText := ParamStr( aIndex );
      aFirstEnter := False;
    end else
    begin
      aParmText := ( aParmText + ' ' + ParamStr( aIndex ) );
    end;
  end;
  ShellExecute( 0, 'open', PChar( aPath + 'prjSO64N0A.exe' ), PChar( aParmText ),
    PChar( aPath ), SW_SHOWNORMAL );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B0.Button7Click(Sender: TObject);
begin
  Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B0.DBGrid1DblClick(Sender: TObject);
begin
  actUpdate.Execute;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B0.CD078DataSetAfterOpen(DataSet: TDataSet);
begin
  Panel4.Caption := Format( '%d/%d', [CD078DataSet.RecNo, CD078DataSet.RecordCount] );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B0.CD078DataSetAfterScroll(DataSet: TDataSet);
begin
  Panel4.Caption := Format( '%d/%d', [CD078DataSet.RecNo, CD078DataSet.RecordCount] );
  btnModify.Enabled := True;
  btnModify.Enabled := DBController.CanModiFy;
  {
  if ( CD078DataSet.FieldByName( 'CanUseType' ).AsInteger = 2 ) then
  begin
    if ( not FCanUseEPG ) then
    begin
      btnModify.Enabled := False;
    end;
  end;
  }
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B0.CD078DataSetCalcFields(DataSet: TDataSet);
begin
  if CD078DataSet.FieldByName( 'onsalestartdate' ).AsString <> EmptyStr then
    CD078DataSet.FieldByName( 'onsalestartcdate' ).AsString := FormatDateTime(
      'yyyy/mm/dd', CD078DataSet.FieldByName( 'onsalestartdate' ).AsDateTime );
  {}
  if CD078DataSet.FieldByName( 'onsalestopdate' ).AsString <> EmptyStr then
    CD078DataSet.FieldByName( 'onsalestopcdate' ).AsString := FormatDateTime(
      'yyyy/mm/dd', CD078DataSet.FieldByName( 'onsalestopdate' ).AsDateTime );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B0.btnFTPClick(Sender: TObject);
begin
  frmCD078BA := TfrmCD078BA.Create(Application);
  try
    frmCD078BA.ShowModal;
  finally
    frmCD078BA.Free;
  end;
  if ( FKeyValue <> EmptyStr ) then actSearch.Execute;
  ActiveControl := DBGrid1;
end;



end.
