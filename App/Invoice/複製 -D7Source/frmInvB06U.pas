unit frmInvB06U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, cxLookAndFeelPainters, StdCtrls, cxButtons, ExtCtrls,
  cxStyles, dxSkinsCore, dxSkinBlack, dxSkinBlue, dxSkinCaramel,
  dxSkinCoffee, dxSkinsDefaultPainters, dxSkinscxPCPainter, cxCustomData,
  cxGraphics, cxFilter, cxData, cxDataStorage, cxEdit, DB, cxDBData,
  cxContainer, cxTextEdit, cxMaskEdit, cxDropDownEdit, cxDBEdit,
  cxGridLevel, cxClasses, cxControls, cxGridCustomView,
  cxGridCustomTableView, cxGridTableView, cxGridDBTableView, cxGrid,
  Provider, DBClient, cxLabel, cxDBLabel,ADODB;

type
  TDMLMode = ( dmBrowser, dmInsert, dmUpdate, dmSave, dmCancel, dmDelete, dmLock );
  TfrmInvB06 = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    btnAdd: TcxButton;
    btnEdit: TcxButton;
    btnCancel: TcxButton;
    btnSave: TcxButton;
    Panel3: TPanel;
    btnQuit: TcxButton;
    cdsMaster: TClientDataSet;
    dspEinv: TDataSetProvider;
    dsMaster: TDataSource;
    cxLabel1: TcxLabel;
    mskInvUseYear: TcxMaskEdit;
    mskSPECIALPRIZE1: TcxDBMaskEdit;
    Bevel1: TBevel;
    cboInvUseMonth: TcxComboBox;
    cxLabel2: TcxLabel;
    cxLabel3: TcxLabel;
    cxDBLabel1: TcxDBLabel;
    cxLabel4: TcxLabel;
    Bevel2: TBevel;
    mskSPECIALPRIZE2: TcxDBMaskEdit;
    mskSPECIALPRIZE3: TcxDBMaskEdit;
    cxLabel5: TcxLabel;
    mskFIRSTPRIZE1: TcxDBMaskEdit;
    mskFIRSTPRIZE2: TcxDBMaskEdit;
    mskFIRSTPRIZE3: TcxDBMaskEdit;
    Bevel3: TBevel;
    cxLabel6: TcxLabel;
    mskEXTRAPRIZE1: TcxDBMaskEdit;
    mskEXTRAPRIZE2: TcxDBMaskEdit;
    lblVerifyEn: TcxDBLabel;
    cxLabel7: TcxLabel;
    cxLabel8: TcxLabel;
    lblVerifyTime: TcxDBLabel;
    mskSPECIALPRIZE4: TcxDBMaskEdit;
    mskSPECIALPRIZE5: TcxDBMaskEdit;
    mskFIRSTPRIZE4: TcxDBMaskEdit;
    mskFIRSTPRIZE5: TcxDBMaskEdit;
    mskEXTRAPRIZE3: TcxDBMaskEdit;
    mskEXTRAPRIZE4: TcxDBMaskEdit;
    mskEXTRAPRIZE5: TcxDBMaskEdit;
    procedure btnAddClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure cdsMasterAfterScroll(DataSet: TDataSet);
    procedure btnSaveClick(Sender: TObject);
    procedure mskSPECIALPRIZE1PropertiesValidate(Sender: TObject;
      var DisplayValue: Variant; var ErrorText: TCaption;
      var Error: Boolean);
    procedure btnCancelClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure btnQuitClick(Sender: TObject);
    procedure cboInvUseMonthPropertiesEditValueChanged(Sender: TObject);

  private
    { Private declarations }
    FYearMonth : String;
    FMode: TDMLMode;
    function BuildINV035SQL(const aWhere: String): String;
    function DataSetStateChange(const aMode: TDMLMode): Boolean;
    function ValidateINV035DataSet(var aErrMsg: String): Boolean;
    function ApplyDataSet(var aErrMsg: String): Boolean;
    procedure EditorStateChange(const aMode: TDMLMode);
    procedure ButtonStateChange(const aMode: TDMLMode);
    procedure CreatePrizeTxt;
    function ValidateCanModify(var aErrMsg: String): Boolean;
    function GetDefPrizeInvDate:String;
  public
    { Public declarations }
    constructor Create(AYearMonth: String); reintroduce;
    procedure DMLModeChange(const aMode: TDMLMode);
  end;

var
  frmInvB06: TfrmInvB06;

implementation
uses cbUtilis,  dtmMainU, dtmMainJU, dtmMainHU;

{$R *.dfm}

procedure TfrmInvB06.btnAddClick(Sender: TObject);
begin
  Screen.Cursor := crSQLWait;
  try
    DMLModeChange( dmInsert );
    
  finally
    Screen.Cursor := crDefault;
  end;
end;

function TfrmInvB06.BuildINV035SQL(const aWhere: String): String;
  var
    aSQL : String;
begin
  aSQL := 'SELECT * FROM INV035 ';
  aSQL := Format(aSQL + ' WHERE YEARMONTH =''%s''' ,
              [ aWhere ]);
  Result := aSQL;
end;

constructor TfrmInvB06.Create(AYearMonth: String);
begin
  inherited Create( Application );
  FYearMonth := AYearMonth;
end;

function TfrmInvB06.DataSetStateChange(const aMode: TDMLMode): Boolean;
var
  aErrMsg: String;
begin
  Result := True;
  aErrMsg := '';
  case aMode of
    dmBrowser, dmLock:
      begin
        cdsMaster.ReadOnly := True;
        cdsMasterAfterScroll( cdsMaster );
      end;
    dmCancel:
      begin
        cdsMaster.Cancel;
        cdsMaster.ReadOnly := True;
        cdsMasterAfterScroll( cdsMaster );
      end;
    dmInsert:
      begin
        cdsMaster.ReadOnly := False;
        cdsMaster.Append;
        mskInvUseYear.Text := copy(GetDefPrizeInvDate,1,4);
        cboInvUseMonth.Text := copy(GetDefPrizeInvDate,5,2);
      end;
    dmUpdate:
      begin
        cdsMaster.ReadOnly := False;
        cdsMaster.Edit;
      end;
    dmSave:
      begin
        Result := ValidateINV035DataSet( aErrMsg );
        if not Result then
        begin
          WarningMsg( aErrMsg );
          Exit;
        end;
        cdsMaster.Post;
        { ?s?? }
        Result := ApplyDataSet( aErrMsg );
        if not Result then
        begin
          WarningMsg( aErrMsg );
          Exit;
        end;
        { ???s Open DataSet }
        FYearMonth := mskInvUseYear.Text + cboInvUseMonth.Text;
        Self.FormShow( Self );
      end;
  end;
end;

procedure TfrmInvB06.DMLModeChange(const aMode: TDMLMode);
begin
  if not DataSetStateChange( aMode ) then Exit;
  FMode := aMode;
  if ( aMode in [dmSave, dmCancel] ) then
    FMode := dmBrowser;
  if ( aMode in [dmInsert, dmUpdate] ) then
    mskSPECIALPRIZE1.SetFocus;
  ButtonStateChange( FMode );
  EditorStateChange( FMode );
end;

procedure TfrmInvB06.FormShow(Sender: TObject);
var
  aEvent: TDataSetNotifyEvent;
  aMode: TDMLMode;
  
begin
  aMode := dmBrowser;
  Screen.Cursor := crSQLWait;
  try
    Application.ProcessMessages;
    aEvent := cdsMaster.AfterScroll;
    cdsMaster.AfterScroll := nil;
    try
      cdsMaster.Close;
      cdsMaster.CommandText :=BuildINV035SQL( FYearMonth ) ;
      cdsMaster.Open;

      DMLModeChange( aMode );
    finally
      cdsMaster.AfterScroll := aEvent;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TfrmInvB06.cdsMasterAfterScroll(DataSet: TDataSet);
begin
  Screen.Cursor := crSQLWait;
  //FYearMonth := cdsMaster.FieldByName('YEARMONTH').AsString;
  if FYearMonth <> '' then
  begin
    mskInvUseYear.Text := Copy(FYearMonth,1,4);
    cboInvUseMonth.Text := Copy(FYearMonth,5,2);
  end;
  if cdsMaster.IsEmpty then
  begin
    mskInvUseYear.Text := '';
    cboInvUseMonth.Text := '';
  end;
  Screen.Cursor := crDefault;
end;

procedure TfrmInvB06.btnSaveClick(Sender: TObject);
begin
  Screen.Cursor := crSQLWait;
  try
    DMLModeChange( dmSave );
  finally
    Screen.Cursor := crDefault;
  end;
end;

function TfrmInvB06.ValidateINV035DataSet(var aErrMsg: String): Boolean;
var
  aFocusControl: TWinControl;
begin
  Result := False;
  aFocusControl := nil;
  try
    if ( mskInvUseYear.Text = EmptyStr ) then
    begin
      aErrMsg := '?????J?o?????X?????~???C';
      aFocusControl := mskInvUseYear;
      Exit;
    end;
    if ( cboInvUseMonth.Text = EmptyStr ) then
    begin
      aErrMsg := '?????J?o?????X????????';
      aFocusControl := mskInvUseYear;
      Exit;
    end;
    if ( cdsMaster.FieldByName( 'SPECIALPRIZE1' ).AsString = EmptyStr ) then
    begin
      aErrMsg := '?????J?S?????X1?C';
      aFocusControl := mskSPECIALPRIZE1;
      Exit;
    end;
    if ( cdsMaster.FieldByName( 'SPECIALPRIZE2' ).AsString = EmptyStr ) then
    begin
      aErrMsg := '?????J?S?????X2?C';
      aFocusControl := mskSPECIALPRIZE2;
      Exit;
    end;
    if ( cdsMaster.FieldByName( 'SPECIALPRIZE3' ).AsString = EmptyStr ) then
    begin
      aErrMsg := '?????J?S?????X3?C';
      aFocusControl := mskSPECIALPRIZE3;
      Exit;
    end;
    {}
    if ( cdsMaster.FieldByName( 'FIRSTPRIZE1' ).AsString = EmptyStr ) then
    begin
      aErrMsg := '?????J?Y?????X1?C';
      aFocusControl := mskFIRSTPRIZE1;
      Exit;
    end;
    if ( cdsMaster.FieldByName( 'FIRSTPRIZE2' ).AsString = EmptyStr ) then
    begin
      aErrMsg := '?????J?Y?????X2?C';
      aFocusControl := mskFIRSTPRIZE2;
      Exit;
    end;
    if ( cdsMaster.FieldByName( 'FIRSTPRIZE3' ).AsString = EmptyStr ) then
    begin
      aErrMsg := '?????J?Y?????X3?C';
      aFocusControl := mskFIRSTPRIZE3;
      Exit;
    end;

  finally
    Result := ( aErrMsg = EmptyStr );
    if not Result then
    begin
      if Assigned( aFocusControl ) and
        ( aFocusControl.CanFocus ) then
      begin
        aFocusControl.SetFocus;
        if aFocusControl is TcxCustomEdit then
          TcxCustomEdit( aFocusControl ).SelectAll;
      end;
    end;
  end;
end;

function TfrmInvB06.ApplyDataSet(var aErrMsg: String): Boolean;
var
  aConnection: TADOConnection;
  aDataSet : TADOQuery;
  function ApplyInsertData:Boolean;
    var aSQL : string;
  begin
    aSQL := Format( 'INSERT INTO INV035 ( YEARMONTH , SPECIALPRIZE1,                ' +
            'SPECIALPRIZE2 , SPECIALPRIZE3 ,                                        ' +
            'FIRSTPRIZE1 , FIRSTPRIZE2 , FIRSTPRIZE3 ,                              ' +
            'EXTRAPRIZE1 , EXTRAPRIZE2 ,                                            ' +
            'UPTTIME ,  UPTEN, SPECIALPRIZE4,SPECIALPRIZE5,                         ' +
            'FIRSTPRIZE4,FIRSTPRIZE5, EXTRAPRIZE3,                                  ' +
            'EXTRAPRIZE4,EXTRAPRIZE5 )                                              ' +
            ' VALUES (''%s'',''%s'',''%s'',''%s'',                                  ' +
            '''%s'',''%s'',''%s'',                                                  ' +
            '''%s'',''%s'',                                                         ' +
            'SysDate,''%s'',''%s'',''%s'',''%s'',''%s'',''%s'',''%s'',''%s'' )      ' ,
            [mskInvUseYear.Text + cboInvUseMonth.Text,
            cdsMaster.FieldByName( 'SPECIALPRIZE1' ).AsString,
            cdsMaster.FieldByName( 'SPECIALPRIZE2' ).AsString ,
            cdsMaster.FieldByName( 'SPECIALPRIZE3' ).AsString,
            cdsMaster.FieldByName( 'FIRSTPRIZE1' ).AsString ,
            cdsMaster.FieldByName( 'FIRSTPRIZE2' ).AsString ,
            cdsMaster.FieldByName( 'FIRSTPRIZE3' ).AsString ,
            cdsMaster.FieldByName( 'EXTRAPRIZE1' ).AsString ,
            cdsMaster.FieldByName( 'EXTRAPRIZE2' ).AsString ,
            dtmMain.getLoginUser,
            cdsMaster.FieldByName( 'SPECIALPRIZE4' ).AsString,
            cdsMaster.FieldByName( 'SPECIALPRIZE5' ).AsString,
            cdsMaster.FieldByName( 'FIRSTPRIZE4' ).AsString ,
            cdsMaster.FieldByName( 'FIRSTPRIZE5' ).AsString ,
            cdsMaster.FieldByName( 'EXTRAPRIZE3' ).AsString ,
            cdsMaster.FieldByName( 'EXTRAPRIZE4' ).AsString ,
            cdsMaster.FieldByName( 'EXTRAPRIZE5' ).AsString  ]);
    aDataSet.Close;
    aDataSet.SQL.Text := aSQL;
    aDataSet.ExecSQL;
    Result := True;
  end;
  function ApplyUpdateData:Boolean;
    var aSQL : string;
  begin
    aSQL :=Format('UPDATE INV035 SET  SPECIALPRIZE1=''%s'',SPECIALPRIZE2=''%s'',            ' +
                'SPECIALPRIZE3=''%s'',FIRSTPRIZE1=''%s'', FIRSTPRIZE2=''%s'',               ' +
                'FIRSTPRIZE3=''%s'', EXTRAPRIZE1=''%s'', EXTRAPRIZE2=''%s'',                ' +
                'UPTTIME=sysdate,UPTEN=''%s'',                                              ' +
                'SPECIALPRIZE4=''%s'',SPECIALPRIZE5=''%s'',                                 ' +
                'FIRSTPRIZE4=''%s'', FIRSTPRIZE5=''%s'',                                    ' +
                'EXTRAPRIZE3=''%s'', EXTRAPRIZE4=''%s'', EXTRAPRIZE5=''%s''                 ' +
                'WHERE YEARMONTH = ''%s''                                                   ',
                [cdsMaster.FieldByName( 'SPECIALPRIZE1' ).AsString,
                cdsMaster.FieldByName( 'SPECIALPRIZE2' ).AsString ,
                cdsMaster.FieldByName( 'SPECIALPRIZE3' ).AsString,
                cdsMaster.FieldByName( 'FIRSTPRIZE1' ).AsString ,
                cdsMaster.FieldByName( 'FIRSTPRIZE2' ).AsString ,
                cdsMaster.FieldByName( 'FIRSTPRIZE3' ).AsString ,
                cdsMaster.FieldByName( 'EXTRAPRIZE1' ).AsString ,
                cdsMaster.FieldByName( 'EXTRAPRIZE2' ).AsString ,
                dtmMain.getLoginUser,
                cdsMaster.FieldByName( 'SPECIALPRIZE4' ).AsString ,
                cdsMaster.FieldByName( 'SPECIALPRIZE5' ).AsString,
                cdsMaster.FieldByName( 'FIRSTPRIZE4' ).AsString ,
                cdsMaster.FieldByName( 'FIRSTPRIZE5' ).AsString ,
                cdsMaster.FieldByName( 'EXTRAPRIZE3' ).AsString ,
                cdsMaster.FieldByName( 'EXTRAPRIZE4' ).AsString ,
                cdsMaster.FieldByName( 'EXTRAPRIZE5' ).AsString ,
                mskInvUseYear.Text + cboInvUseMonth.Text]);
    aDataSet.Close;
    aDataSet.SQL.Text := aSQL;
    aDataSet.ExecSQL;
    Result := True;
  end;

begin
  aConnection := dtmMain.EInvConnection;
  aDataSet := dtmMain.adoEComm;
  try
    aConnection.BeginTrans;
    if ( FMode in [dmInsert] ) then
    begin
      ApplyInsertData;
    end else
    begin
     ApplyUpdateData;
    end;
    aConnection.CommitTrans;
    Result := True;
  except
    on E: Exception do
    begin
      TranslateError( E.Message );
      { PKey ????, PM ?n?D???q?T?? }
      if OracleErros.ErrorCode = 1 then
      begin
        aErrMsg := '???????????X?w?Q???J,????????????!';
      end
      else begin
        aErrMsg := E.Message;
      end;

      aConnection.RollbackTrans;
    end;
  end;
end;

procedure TfrmInvB06.EditorStateChange(const aMode: TDMLMode);
begin
  if ( aMode in [dmBrowser, dmLock, dmUpdate] ) then
  begin
    mskInvUseYear.Enabled := False;
    cboInvUseMonth.Enabled := False;
  end
  else if ( aMode in [dmInsert] ) then
  begin
    mskInvUseYear.Enabled := True;
    cboInvUseMonth.Enabled := True

  end;
end;

procedure TfrmInvB06.ButtonStateChange(const aMode: TDMLMode);
begin
  case aMode of
    dmBrowser, dmCancel:
      begin
        btnAdd.Enabled := True;
        btnEdit.Enabled := True;
        btnSave.Enabled := False;
        btnCancel.Enabled := False;
      end;
    dmInsert, dmUpdate:
      begin
        btnAdd.Enabled := False;
        btnEdit.Enabled := False;
        btnSave.Enabled := True;
        btnCancel.Enabled := True;

      end;
    dmLock:
      begin
        btnAdd.Enabled := False;
        btnEdit.Enabled := False;
        btnSave.Enabled := False;
        btnCancel.Enabled := False;
      end;
  end;
end;

procedure TfrmInvB06.mskSPECIALPRIZE1PropertiesValidate(Sender: TObject;
  var DisplayValue: Variant; var ErrorText: TCaption; var Error: Boolean);
begin

  if (Sender as TcxDBMaskEdit).Text = EmptyStr then Exit;
  if Length((Sender as TcxDBMaskEdit).Text) <> 8 then
  begin
    ErrorText := '?o??????8?X?I';
    Error := False;
    WarningMsg( ErrorText );
    (Sender as TcxDBMaskEdit).SetFocus;
  end;

end;

procedure TfrmInvB06.btnCancelClick(Sender: TObject);
begin
  Screen.Cursor := crSQLWait;
  try
    DMLModeChange( dmCancel );
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TfrmInvB06.btnEditClick(Sender: TObject);
  var
    aErrMsg : string;
begin
  Screen.Cursor := crSQLWait;
  try
    if not ValidateCanModify( aErrMsg ) then
    begin
      WarningMsg( aErrMsg );
      Exit;
    end;
    DMLModeChange( dmUpdate );
  finally
    Screen.Cursor := crDefault;
  end;
end;

function TfrmInvB06.ValidateCanModify(var aErrMsg: String): Boolean;
begin
  aErrMsg := EmptyStr;
  Result := True;
  if cdsMaster.IsEmpty then
  begin
    aErrMsg := '?L?o???????????I?L?k?????????C';
    Result := False;
    Exit;
  end;
  if ( cdsMaster.FieldByName( 'VerifyEn' ).AsString <> '' ) and
    ( cdsMaster.FieldByName( 'VerifyTime' ).AsString <> '' ) then
  begin
    aErrMsg := '?????????w?f???I?????\?????????C';
    Result := False;
    Exit;
  end;
end;

procedure TfrmInvB06.btnQuitClick(Sender: TObject);
begin
   Self.Close;
end;

function TfrmInvB06.GetDefPrizeInvDate: String;
var
  aYear,aMonth :String;
begin
  aYear := FormatDateTime('yyyy',Now);
  aMonth := FormatDateTime('MM',Now);
  if StrToInt(aMonth) > 2 then
  begin
    if ( StrToInt( aMonth ) mod 2 ) = 0 then
    begin
      aMonth := IntToStr( StrToInt( aMonth ) - 3 )
    end else
      aMonth := IntToStr( StrToInt( aMonth ) - 2 );
    if Length( aMonth ) = 1 then
      aMonth := '0' + aMonth;
  end else
  begin
    aMonth := '11';
    aYear := IntToStr( StrToInt( aYear ) - 1 );
  end;
  Result := aYear + aMonth;
end;

procedure TfrmInvB06.cboInvUseMonthPropertiesEditValueChanged(
  Sender: TObject);
  var aMonth : String;
begin
  aMonth := cboInvUseMonth.Text;
  if aMonth = '' then Exit;
  if StrToInt( aMonth ) >= 2 then
  begin
    if ( StrToInt( aMonth ) mod 2 ) = 0 then
    begin
      aMonth := IntToStr( StrToInt( aMonth ) - 1 );
      if Length( aMonth ) = 1 then
        aMonth := '0' + aMonth;

    end;
  end;
  cboInvUseMonth.Text := aMonth;
end;

procedure TfrmInvB06.CreatePrizeTxt;
var aSQL,aFileName: string;
begin
  aSQL := ' SELECT INVPARAM FROM INV003 WHERE COMPID = ' +
      QuotedStr( dtmMain.getCompID );

  aFileName := dtmMainJ.GetXMLAttribute( aSQL, 'MediaFilePath' );
  if not DirectoryExists( aFileName ) then ForceDirectories( aFileName );

  {
  if ( chkDebugFile.Checked ) then
    begin
      FDebugList.SaveToFile( IncludeTrailingPathDelimiter( aDebugPath ) + aDebugFileName );
      FDebugList.Clear;
    end;
    }
end;

end.
