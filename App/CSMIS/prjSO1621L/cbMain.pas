unit cbMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Buttons, DB, DBClient, ADODB, Menus,
  ActnList, ImgList, ShellApi,
  cbDataController, Crypt32DLL_TLB,
  cxGraphics, cxContainer, cxClasses, cxControls, cxDataStorage, cxDBData,
  cxGrid, cxGridLevel, cxGridCustomView, cxGridCustomTableView, cxGridTableView,
  cxGridBandedTableView, cxGridDBBandedTableView, cxGridStrs,
  cxEdit, cxImageComboBox, cxTextEdit, cxCheckBox, cxRadioGroup, cxButtons, cxProgressBar,
  dxSkinsCore, cxLookAndFeels, cxLookAndFeelPainters, dxSkinscxPCPainter,
  dxSkinsDefaultPainters, cxPC, cxGroupBox, cxMaskEdit, cxDropDownEdit,
  cxStyles, cxCustomData, cxFilter, cxData, dxSkinBlack, dxSkinBlue,
  dxSkinCaramel, dxSkinCoffee;

const CmdA1: Integer = 1;
const CmdB1: Integer = 2;
const CmdB3: Integer = 3;
const CmdB1_1: Integer = 4;

type
  TGuiInfo = record
    RecNo: Integer;
    SeqNo: String;
    CmdType: Integer;
    CmdStatus: String;
    ErrMsg: String;
  end;

type
  TQueryParam = class
  private
    FParam1: String;
  public
    property Param1: String read FParam1 write FParam1;
  end;

type
  TCD022Info = record
    IccSNoLength: Integer;
    StbSNoLength: Integer;
  end;

  
type
  TfmMain = class(TForm)
    Panel3: TPanel;
    dsDVN: TDataSource;
    cdsDVN: TClientDataSet;
    cxLookAndFeelController1: TcxLookAndFeelController;
    ActionList1: TActionList;
    actDvnSetup: TAction;
    actCancel: TAction;
    cdsDvnChannel: TClientDataSet;
    CmdImageList: TImageList;
    OpenDialog: TOpenDialog;
    actDvnClear: TAction;
    btnImport: TcxButton;
    btnSetup: TcxButton;
    btnClear: TcxButton;
    btnCancel: TcxButton;
    MainPage: TcxPageControl;
    cxTabDVN: TcxTabSheet;
    cxTabNAGRA: TcxTabSheet;
    Panel1: TPanel;
    GroupBox1: TGroupBox;
    Label2: TLabel;
    Label3: TLabel;
    Label1: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Panel5: TPanel;
    rdoA1: TcxRadioButton;
    rdoA7: TcxRadioButton;
    txtDvnStb: TcxTextEdit;
    txtDvnIcc: TcxTextEdit;
    btnDvnCannel: TcxButton;
    btnChannelClear: TcxButton;
    txtDvnChannel: TcxTextEdit;
    txtDays: TcxTextEdit;
    chkDvnReSendErr: TcxCheckBox;
    Panel2: TPanel;
    Bevel1: TBevel;
    gdDVN: TcxGrid;
    gvDVN: TcxGridDBBandedTableView;
    gvDVNSeqNo1: TcxGridDBBandedColumn;
    gvDVNSeqNo2: TcxGridDBBandedColumn;
    gvDVNStbNo: TcxGridDBBandedColumn;
    gvDVNIccNo: TcxGridDBBandedColumn;
    gvDVNCmdA1: TcxGridDBBandedColumn;
    gvDVNCmdB1: TcxGridDBBandedColumn;
    gvDVNErrMsg: TcxGridDBBandedColumn;
    glDVN: TcxGridLevel;
    Panel4: TPanel;
    lbDvnRecords: TLabel;
    DvnProgressBar: TcxProgressBar;
    cxGroupBox1: TcxGroupBox;
    cmbCAS: TcxComboBox;
    Label6: TLabel;
    GroupBox2: TGroupBox;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    txtNagraStb: TcxTextEdit;
    txtNagraIcc: TcxTextEdit;
    btnNagraChannel: TcxButton;
    btnChannelClear2: TcxButton;
    txtNagraChannel: TcxTextEdit;
    txtNagraDays: TcxTextEdit;
    chkNagraReSendErr: TcxCheckBox;
    Label7: TLabel;
    cmbNagraOpeartion: TcxComboBox;
    lbOpeartion: TLabel;
    cdsNagraChannel: TClientDataSet;
    cdsNagra: TClientDataSet;
    gdNagra: TcxGrid;
    gvNagra: TcxGridDBBandedTableView;
    gdNagraCol1: TcxGridDBBandedColumn;
    gdNagraCol2: TcxGridDBBandedColumn;
    gdNagraCol3: TcxGridDBBandedColumn;
    gdNagraCol4: TcxGridDBBandedColumn;
    gdNagraCol5: TcxGridDBBandedColumn;
    gdNagraCol6: TcxGridDBBandedColumn;
    gdNagraCol10: TcxGridDBBandedColumn;
    glNagra: TcxGridLevel;
    Panel6: TPanel;
    lbNagraRecords: TLabel;
    NagraProgressBar: TcxProgressBar;
    dsNagra: TDataSource;
    gdNagraCol7: TcxGridDBBandedColumn;
    gdNagraCol8: TcxGridDBBandedColumn;
    actNagraSetup: TAction;
    actNagraClear: TAction;
    Label13: TLabel;
    txtZipCode: TcxTextEdit;
    Label14: TLabel;
    txtBatId: TcxTextEdit;
    qryAutoFill: TADOQuery;
    chkCHANCEDAYS: TCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormDestroy(Sender: TObject);
    procedure cdsDVNAfterScroll(DataSet: TDataSet);
    procedure txtDvnStbKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure txtDvnIccKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure txtDaysKeyPress(Sender: TObject; var Key: Char);
    procedure btnDvnCannelClick(Sender: TObject);
    procedure btnChannelClearClick(Sender: TObject);
    procedure btnImportClick(Sender: TObject);
    procedure actDvnSetupExecute(Sender: TObject);
    procedure actDvnClearExecute(Sender: TObject);
    procedure actCancelExecute(Sender: TObject);
    procedure actNagraClearExecute(Sender: TObject);
    procedure gvDVNCustomDrawCell(Sender: TcxCustomGridTableView;
      ACanvas: TcxCanvas; AViewInfo: TcxGridTableDataCellViewInfo;
      var ADone: Boolean);
    procedure Panel3Resize(Sender: TObject);
    procedure cmbCASPropertiesChange(Sender: TObject);
    procedure cmbOpeartionPropertiesChange(Sender: TObject);
    procedure actNagraSetupExecute(Sender: TObject);
    procedure txtZipCodeExit(Sender: TObject);
    procedure txtBatIdExit(Sender: TObject);
    procedure txtNagraDaysExit(Sender: TObject);
    procedure chkCHANCEDAYSClick(Sender: TObject);
  private
    { Private declarations }
    FCasMode: Integer;
    FNagraOperation: Integer;
    FNagraStep: Integer;
    FDvnQueryParam: TQueryParam;
    FNagraQueryParam: TQueryParam;
    FImportList: TStringList;
    FLogList: TStringList;
    FImportLogFileName: String;
    FCmdLogFileName: String;
    FDvnRecNo: Integer;
    FNagraRecNo: Integer;
    FCmdHasErr: Boolean;
    FThreadIsRuning: Boolean;
    FDvnThread: TThread;
    FNagraThread: TThread;
    FDvnCD022Info: TCD022Info;
    FNagraCD022Info: TCD022Info;
    FAutoFill:Boolean;
    procedure ChangeDevexComponentLanguage;
    procedure PrepareDataSet;
    procedure UnPrepareDataSet;
    procedure ResetNagraCommandData;
    procedure SetNagraSkipFlag;
    procedure LoadChannelDataSet(const ACodeNo: String);
    procedure SetChannelBufferDays;
    procedure LoadDefaultZipCode;
    procedure LoadDefaultBatId;
    procedure Load_DVN_CD022Info;
    procedure Load_Nagra_CD022Info;
    procedure ResetProgressFlag;
    procedure BegindDvnSet;
    procedure EndDvnSet;
    procedure BeginNagraSet;
    procedure EndNagraSet;
    procedure DisplayNagraProgressBar;
    procedure DisplayDvnProgressBar;
    procedure SwitchingCAS(ACasText: String); overload;
    procedure SwitchingCAS(ACasMode: Integer); overload;
    procedure SwitchNagraOperation(AOperationText: String);
    procedure SwitchNagraStep;
    procedure RunDvnSet;
    procedure RunNagraSet;
    procedure DoDvnProgress(const ACmdType, ARecNo: Integer;
      const  ACmdStatus, AErrMsg: String);
    procedure DoNagraProgress(const ACmdType, ARecNo: Integer;
      const ACmdStatus, AErrMsg: String);

    procedure GetAutoFill;
    function RecvParamString: Boolean;
    function LoadCasItem: Integer;
    function AddToDvnGrid(const AStbNo, AIccNo: String; var AMsg: String): Boolean;
    function AddToNagraGrid(const AStbNo, AIccNo: String; var AMsg: String): Boolean;
    function LoadFromFile(const AFileName: String): Boolean;
    {}
    function VdIccNo(const AIccNo: String; var AMsg: String): Boolean;
    function VdStbNo(const AStbNo: String; var AMsg: String): Boolean;
    {}
    function VdIccNoLength(const AIccNo: String; var AMsg: String): Boolean;
    function VdStbNoLength(const AStbNo: String; var AMsg: String): Boolean;
    {}
    function VdMustInput: Boolean;
    function GetImportLogFileName: String;
    function GetCmdLogFileName: String;
    function GetRecNo: Integer;
    function GetCmdA1Text: String;
    function GetNagraStepText(const AStep: Integer): String;
  public
    { Public declarations }
    procedure ProgressReport(const aCmdType, aRecNo: Integer;
      const  aCmdStatus, aErrMsg: String);
    procedure DvnThreadTerminate(Sender: TObject);
    procedure NagraThreadTerminate(Sender: TObject);
  end;

var
  fmMain: TfmMain;

implementation

{$R *.dfm}

uses cbUtilis, cbMultiSelect, cbDvnThread, cbNagraThread;

{ ---------------------------------------------------------------------------- }

function IndexOfStringReplace(const AString: String; const AIndex: Integer;
  const AReplace: String ): String;
var
  aIndex2: Integer;
begin
  Result := EmptyStr;
  for aIndex2 := 1 to Length( AString ) do
  begin
    if ( aIndex2 = AIndex ) then
      Result := ( Result + AReplace )
    else
      Result := ( Result + AString[aIndex2] );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormCreate(Sender: TObject);
var
  aCreateError: Boolean;
begin
  FDvnRecNo := 0;
  { 1 -> 開機, 2 -> 開試看頻道, 3 -> 取消所有授權, 4 -> 開試看頻道 }
  FNagraStep := 1;
  FThreadIsRuning := False;
  FDvnThread := nil;
  FDvnQueryParam := TQueryParam.Create;
  FNagraQueryParam := TQueryParam.Create;
  DBController := TDBController.Create( nil );
  FImportList := TStringList.Create;
  FLogList := TStringList.Create;
  lbDvnRecords.Caption := Format( ' %d / %d ', [0,0] );
  ChangeDevexComponentLanguage;
  aCreateError := ( not RecvParamString );

  if not aCreateError then DBController.Connected := True;
  aCreateError := ( not DBController.Connected );
  if ( aCreateError ) then
  begin
    Application.ShowMainForm := False;
    Application.Terminate;
    Exit;
  end;
  PrepareDataSet;
  {}
  FCasMode := 2;
  LoadChannelDataSet( EmptyStr );
  FCasMode := 5;
  LoadChannelDataSet( EmptyStr );
  {}
  FCasMode := LoadCasItem;
  if ( FCasMode = 2 ) then
    Application.Title := 'NAGRA配對燒機作業[SO1621L]';
  Self.Caption := Application.Title;  
  {}
  LoadDefaultZipCode;
  LoadDefaultBatId;
  {}
  Load_DVN_CD022Info;
  Load_Nagra_CD022Info;
  {}
  FImportLogFileName := GetImportLogFileName;
  FCmdLogFileName := GetCmdLogFileName;
  DvnProgressBar.Visible := False;
  NagraProgressBar.Visible := False;
  MainPage.HideTabs := True;
  SwitchingCAS( FCasMode );
  cmbNagraOpeartion.ItemIndex := 0;
  GetAutoFill; //取CD039.NagraType 判斷是否要自動帶出資料或是Upd SOAC012AC By Kin 2009/07/30
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  if ( FThreadIsRuning ) then
  begin
    try
      if ( FCasMode = 2 ) then
      begin
        CanClose := ConfirmMsg( Format(
          '機上盒【%s】作業尚未完成,確認結束程式?', [cmbNagraOpeartion.Text] ) );
      end else
      begin
        CanClose := ConfirmMsg( '確認結束程式?' );
      end;
      if ( CanClose ) then
      begin
        if Assigned( FDvnThread ) then
        begin
          FDvnThread.Terminate;
          FDvnThread.WaitFor;
        end;
        if Assigned( FNagraThread ) then
        begin
          FNagraThread.Terminate;
          FNagraThread.WaitFor;
        end;
      end;
    except
      { ... }
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormDestroy(Sender: TObject);
begin
  FImportList.Free;
  FLogList.Free;
  DBController.Free;
  FDvnQueryParam.Free;
  FNagraQueryParam.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.ChangeDevexComponentLanguage;
begin
  cxSetResourceString( @scxGridDeletingConfirmationCaption, '確認' );
  cxSetResourceString( @scxGridDeletingFocusedConfirmationText, '刪除此筆資料?' );
  cxSetResourceString( @scxGridDeletingSelectedConfirmationText, '刪除所選取的資料?' );
  cxSetResourceString( @scxGridCustomizationFormCaption, '欄位選擇' );
  cxSetResourceString( @scxGridCustomizationFormBandsPageCaption, '項目' );
  cxSetResourceString( @scxGridCustomizationFormColumnsPageCaption, '欄位' );
  cxSetResourceString( @scxGridNoDataInfoText, '無資料可顯示' );
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.LoadCasItem: Integer;

  { ---------------------------------------- }

  procedure AddItem(const AItemText: String);
  begin
    if cmbCAS.Properties.Items.IndexOf( AItemText ) < 0 then
      cmbCAS.Properties.Items.Add( AItemText );
  end;

  { ---------------------------------------- }

begin
  Result := 5;
  DBController.DataReader.Close;
  DBController.DataReader.SQL.Text := Format(
    ' SELECT A.ACCTYPE, A.ACCTYPE2 ' +
    '   FROM CD039 A               ' +
    '  WHERE CODENO = ''%s''       ',
    [DBController.LoginInfo.CompCode] );
  try
    DBController.DataReader.Open;
  except
    on E: Exception do
    begin
      ErrorMsg( E.Message );
      Exit; 
    end;  
  end;          
  if ( not DBController.DataReader.IsEmpty ) then
  begin
    cmbCAS.Properties.Items.Clear;
    if ( DBController.DataReader.FieldByName( 'ACCTYPE' ).AsInteger = 2 ) or
       ( DBController.DataReader.FieldByName( 'ACCTYPE2' ).AsInteger = 2 ) then
      AddItem( 'NAGRA' );
    {}
    if ( DBController.DataReader.FieldByName( 'ACCTYPE' ).AsInteger = 5 ) or
       ( DBController.DataReader.FieldByName( 'ACCTYPE2' ).AsInteger = 5 ) then
      AddItem( 'DVN' );
    Result := DBController.DataReader.FieldByName( 'ACCTYPE' ).AsInteger;
  end;
  DBController.DataReader.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.PrepareDataSet;
begin
  if ( cdsDvn.FieldDefs.Count <= 0 ) then
  begin
    cdsDvn.FieldDefs.Add( 'RECNO', ftInteger );
    cdsDvn.FieldDefs.Add( 'PROGRESS', ftString, 10 );
    cdsDvn.FieldDefs.Add( 'SEQNO1', ftInteger );
    cdsDvn.FieldDefs.Add( 'SEQNO2', ftInteger );
    cdsDvn.FieldDefs.Add( 'STBNO', ftString, 20 );
    cdsDvn.FieldDefs.Add( 'ICCNO', ftString, 20 );
    cdsDvn.FieldDefs.Add( 'CMDA1', ftString, 1 );
    cdsDvn.FieldDefs.Add( 'CMDB1', ftString, 1 );
    cdsDvn.FieldDefs.Add( 'ERRMSG', ftString, 50 );
    cdsDvn.FieldDefs.Add( 'SENDTIME1', ftDateTime );
    cdsDvn.FieldDefs.Add( 'SENDTIME2', ftDateTime );
    cdsDvn.FieldDefs.Add( 'HASLOG1', ftString, 1 );
    cdsDvn.FieldDefs.Add( 'HASLOG2', ftString, 1 );
    cdsDvn.CreateDataSet;
  end;
  cdsDvn.EmptyDataSet;
  {}
  if ( cdsNagra.FieldDefs.Count <= 0 ) then
  begin
    cdsNagra.FieldDefs.Add( 'RECNO', ftInteger );
    cdsNagra.FieldDefs.Add( 'PROGRESS', ftString, 10 );
    cdsNagra.FieldDefs.Add( 'SEQNO1', ftInteger );
    cdsNagra.FieldDefs.Add( 'SEQNO2', ftInteger );
    cdsNagra.FieldDefs.Add( 'SEQNO3', ftInteger );
    cdsNagra.FieldDefs.Add( 'SEQNO4', ftInteger );
    cdsNagra.FieldDefs.Add( 'STBNO', ftString, 20 );
    cdsNagra.FieldDefs.Add( 'ICCNO', ftString, 20 );
    cdsNagra.FieldDefs.Add( 'CMDA1', ftString, 1 );
    cdsNagra.FieldDefs.Add( 'CMDB1', ftString, 1 );
    cdsNagra.FieldDefs.Add( 'CMDB3', ftString, 1 );
    cdsNagra.FieldDefs.Add( 'CMDB1_1', ftString, 1 );
    cdsNagra.FieldDefs.Add( 'ERRMSG', ftString, 50 );
    cdsNagra.FieldDefs.Add( 'SENDTIME1', ftDateTime );
    cdsNagra.FieldDefs.Add( 'SENDTIME2', ftDateTime );
    cdsNagra.FieldDefs.Add( 'SENDTIME3', ftDateTime );
    cdsNagra.FieldDefs.Add( 'SENDTIME4', ftDateTime );
    cdsNagra.FieldDefs.Add( 'HASLOG1', ftString, 1 );
    cdsNagra.FieldDefs.Add( 'HASLOG2', ftString, 1 );
    cdsNagra.FieldDefs.Add( 'HASLOG3', ftString, 1 );
    cdsNagra.FieldDefs.Add( 'HASLOG4', ftString, 1 );
    cdsNagra.CreateDataSet;
  end;
  cdsNagra.EmptyDataSet;
  {}
  if ( cdsDvnChannel.FieldDefs.Count <= 0 ) then
  begin
    cdsDvnChannel.FieldDefs.Add( 'SELECTED', ftString, 1 );
    cdsDvnChannel.FieldDefs.Add( 'CODENO', ftString, 20 );
    cdsDvnChannel.FieldDefs.Add( 'DESCRIPTION', ftString, 20 );
    cdsDvnChannel.FieldDefs.Add( 'CHANNELID', ftString, 20 );
    cdsDvnChannel.FieldDefs.Add( 'CHANCEDAYS', ftInteger );
    cdsDvnChannel.FieldDefs.Add( 'BUFFERDAYS', ftInteger );
    cdsDvnChannel.CreateDataSet;
  end;
  cdsDvnChannel.EmptyDataSet;
  {}
  cdsNagraChannel.FieldDefs.Assign( cdsDvnChannel.FieldDefs );
  cdsNagraChannel.CreateDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.UnPrepareDataSet;
var
  aBuffer: TClientDataSet;
begin
  aBuffer := cdsDVN;
  if ( FCasMode = 2 ) then
    aBuffer := cdsNagra;
  if not VarIsNull( aBuffer.Data ) then
    aBuffer.EmptyDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.ResetNagraCommandData;
begin
  cdsNagra.First;
  cdsNagra.DisableControls;
  try
    while not cdsNagra.Eof do
    begin
      cdsNagra.Edit;
      if ( FNagraOperation = 1 ) then
      begin
        cdsNagra.FieldByName( 'CMDA1' ).AsString := EmptyStr;
        cdsNagra.FieldByName( 'CMDB1' ).AsString := EmptyStr;
        cdsNagra.FieldByName( 'CMDB3' ).AsString := 'S';
        cdsNagra.FieldByName( 'CMDB1_1' ).AsString := 'S';
      end else
      if ( FNagraOperation = 2 ) then
      begin
        cdsNagra.FieldByName( 'CMDA1' ).AsString := EmptyStr;
        cdsNagra.FieldByName( 'CMDB1' ).AsString := EmptyStr;
        cdsNagra.FieldByName( 'CMDB3' ).AsString := EmptyStr;
        cdsNagra.FieldByName( 'CMDB1_1' ).AsString := EmptyStr;
      end else
      if ( FNagraOperation = 3 ) then
      begin
        cdsNagra.FieldByName( 'CMDA1' ).AsString := EmptyStr;
        cdsNagra.FieldByName( 'CMDB1' ).AsString := EmptyStr;
        cdsNagra.FieldByName( 'CMDB3' ).AsString := EmptyStr;
        cdsNagra.FieldByName( 'CMDB1_1' ).AsString := 'S';
      end;
      cdsNagra.Post;
      cdsNagra.Next;
    end;
    cdsNagra.First;
  finally
    cdsNagra.EnableControls;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.SetNagraSkipFlag;
begin
  cdsNagra.First;
  cdsNagra.DisableControls;
  try
    while not cdsNagra.Eof do
    begin
      cdsNagra.Edit;
      if ( FNagraQueryParam.Param1 <> EmptyStr ) then
      begin
        if ( FNagraStep <= 2 ) then
            cdsNagra.FieldByName( 'CMDB1' ).AsString := EmptyStr;
        if ( FNagraStep <= 4 ) and ( FNagraOperation in [2,3] ) then
            cdsNagra.FieldByName( 'CMDB1_1' ).AsString := EmptyStr;
      end else
      begin
        cdsNagra.FieldByName( 'CMDB1' ).AsString := 'S';
        cdsNagra.FieldByName( 'CMDB1_1' ).AsString := 'S';
        cdsNagra.FieldByName( 'ERRMSG' ).AsString := EmptyStr;
      end;
      cdsNagra.Post;
      cdsNagra.Next;
    end;
    cdsNagra.First;
  finally
    cdsNagra.EnableControls;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.LoadChannelDataSet(const ACodeNo: String);
var
  aValue, aText, aSelected: String;
  aRefNo: Integer;
  aChBuffer: TClientDataSet;
  aParam: TQueryParam;
  aEditor: TcxTextEdit;
begin
  aRefNo := 10;
  aChBuffer := cdsDvnChannel;
  aParam := FDvnQueryParam;
  aEditor := txtDvnChannel;
  if ( FCasMode = 2 ) then
  begin
    aRefNo := 11;
    aChBuffer := cdsNagraChannel;
    aParam := FNagraQueryParam;
    aEditor := txtNagraChannel;
  end;
  {}
  DBController.DataReader.Close;
  DBController.DataReader.SQL.Text := Format(
    ' SELECT CODENO, DESCRIPTION, CHANNELID,  ' +
    '        CHANCEDAYS                       ' +
    '   FROM CD024  WHERE STOPFLAG = 0        ' +
    '    AND PAYFLAG IN ( 0, 1 )              ' +
    '    AND REFNO = %d                       ' +
    '    AND CHANNELID IS NOT NULL            ',
    [aRefNo] );
  //#5241 有勾選的時候,不要判斷CHANCEDAYS By Kin 2009/08/18
  if (not chkCHANCEDAYS.Checked) or (FCasMode <> 2) then
    DBController.DataReader.SQL.Text := DBController.DataReader.SQL.Text +
                                        ' AND CHANCEDAYS > 0 ';

  try
    DBController.DataReader.Open;
  except
    on E: Exception do
    begin
      ErrorMsg( E.Message );
      Exit;
    end;
  end;
  aChBuffer.EmptyDataSet;
  aValue := EmptyStr;
  aText := EmptyStr;
  DBController.DataReader.First;
  while not DBController.DataReader.Eof do
  begin
    aChBuffer.Append;
    aSelected := 'Y';
    if ( ACodeNo <> EmptyStr ) or ( ACodeNo = 'X' ) then
    begin
      if not ( Pos( DBController.DataReader.FieldByName( 'CODENO' ).AsString,
        ACodeNo ) > 0 ) then
       aSelected := 'N';
    end;
    aChBuffer.FieldByName( 'SELECTED' ).AsString := aSelected;
    aChBuffer.FieldByName( 'CODENO' ).AsString :=
      DBController.DataReader.FieldByName( 'CODENO' ).AsString;
    aChBuffer.FieldByName( 'DESCRIPTION' ).AsString :=
      DBController.DataReader.FieldByName( 'DESCRIPTION' ).AsString;
    aChBuffer.FieldByName( 'CHANNELID' ).AsString :=
      DBController.DataReader.FieldByName( 'CHANNELID' ).AsString;
    aChBuffer.FieldByName( 'CHANCEDAYS' ).AsInteger :=
      StrToIntDef( DBController.DataReader.FieldByName( 'CHANCEDAYS' ).AsString, 0 );
    aChBuffer.Post;
    {}
    if ( aSelected = 'Y' ) then
    begin
      aValue := ( aValue + aChBuffer.FieldByName( 'CODENO' ).AsString + ','  );
      aText := ( aText + aChBuffer.FieldByName( 'DESCRIPTION' ).AsString + ',' );
    end;
    {}
    DBController.DataReader.Next;
  end;
  {}
  if IsDelimiter( ',', aValue, Length( aValue ) ) then
    Delete( aValue, Length( aValue ), 1 );
  {}
  if IsDelimiter( ',', aText, Length( aText ) ) then
    Delete( aText, Length( aText ), 1 );
  {}  
  DBController.DataReader.Close;
  aParam.Param1 := aValue;
  aEditor.Text := aText;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.SetChannelBufferDays;
var
  aChBuffer: TClientDataSet;
  aDays: Integer;
begin
  aChBuffer := cdsDvnChannel;
  aDays := StrToIntDef( txtNagraDays.Text, 0 );
  if ( FCasMode = 2 ) then
  begin
    aChBuffer := cdsNagraChannel;
    StrToIntDef( txtNagraDays.Text, 0 );
  end;
  {}
  aChBuffer.First;
  while not aChBuffer.Eof do
  begin
    aChBuffer.Edit;
    aChBuffer.FieldByName( 'BUFFERDAYS' ).AsInteger := aDays;
    aChBuffer.Post;
    aChBuffer.Next;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.LoadDefaultZipCode;
begin
  DBController.DataReader.Close;
  DBController.DataReader.SQL.Text :=
    ' SELECT A.ZIPCODE FROM CD098 A   ' +
    '  WHERE A.DEFAULTFLAG = 1        ' +
    '    AND A.STOPFLAG = 0           ';
  DBController.DataReader.Open;
  txtZipCode.Text := DBController.DataReader.Fields[0].AsString;
  DBController.DataReader.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.LoadDefaultBatId;
begin
  DBController.DataReader.Close;
  DBController.DataReader.SQL.Text :=
    ' SELECT A.BATID FROM CD099 A   ' +
    '  WHERE A.DEFAULTFLAG = 1      ' +
    '    AND A.STOPFLAG = 0         ';
  DBController.DataReader.Open;
  txtBatId.Text := DBController.DataReader.Fields[0].AsString;
  DBController.DataReader.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.Load_DVN_CD022Info;
begin
  FDvnCD022Info.IccSNoLength := 20;
  FDvnCD022Info.StbSNoLength := 20;
  {}
  DBController.DataReader.Close;
  DBController.DataReader.SQL.Text :=
    ' SELECT A.SNOLENGTH FROM CD022 A   ' +
    '  WHERE A.REFNO = 3                ' +
    '    AND A.SUBTYPE = 5              ' +
    '    AND A.STOPFLAG = 0             ';
  DBController.DataReader.Open;
  if ( not DBController.DataReader.IsEmpty ) then
    FDvnCD022Info.StbSNoLength := DBController.DataReader.Fields[0].AsInteger;
  DBController.DataReader.Close;
  {}
  DBController.DataReader.Close;
  DBController.DataReader.SQL.Text :=
    ' SELECT A.SNOLENGTH FROM CD022 A   ' +
    '  WHERE A.REFNO = 4                ' +
    '    AND A.SUBTYPE = 5              ' +
    '    AND A.STOPFLAG = 0             ';
  DBController.DataReader.Open;
  if ( not DBController.DataReader.IsEmpty ) then
    FDvnCD022Info.IccSNoLength := DBController.DataReader.Fields[0].AsInteger;
  DBController.DataReader.Close;
  {}
  txtDvnStb.Properties.MaxLength := FDvnCD022Info.StbSNoLength;
  txtDvnIcc.Properties.MaxLength := FDvnCD022Info.IccSNoLength;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.Load_Nagra_CD022Info;
begin
  FNagraCD022Info.IccSNoLength := 20;
  FNagraCD022Info.StbSNoLength := 20; 
  {}
  DBController.DataReader.Close;
  DBController.DataReader.SQL.Text :=
    ' SELECT A.SNOLENGTH FROM CD022 A   ' +
    '  WHERE A.REFNO = 3                ' +
    '    AND A.SUBTYPE = 2              ' +
    '    AND A.STOPFLAG = 0             ';
  DBController.DataReader.Open;
  if ( not DBController.DataReader.IsEmpty ) then
    FNagraCD022Info.StbSNoLength := DBController.DataReader.Fields[0].AsInteger;
  DBController.DataReader.Close;
  {}
  DBController.DataReader.Close;
  DBController.DataReader.SQL.Text :=
    ' SELECT A.SNOLENGTH FROM CD022 A   ' +
    '  WHERE A.REFNO = 4                ' +
    '    AND A.SUBTYPE = 2              ' +
    '    AND A.STOPFLAG = 0             ';
  DBController.DataReader.Open;
  if ( not DBController.DataReader.IsEmpty ) then
    FNagraCD022Info.IccSNoLength := DBController.DataReader.Fields[0].AsInteger;
  DBController.DataReader.Close;
  {}
  txtNagraStb.Properties.MaxLength := FNagraCD022Info.StbSNoLength;
  txtNagraIcc.Properties.MaxLength := FNagraCD022Info.IccSNoLength;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.ResetProgressFlag;
var
  aBuffer: TClientDataSet;
  aMark: String;
begin
  aBuffer := cdsDVN;
  if ( FCasMode = 2 ) then
    aBuffer := cdsNagra;
  {}
  aBuffer.DisableControls;
  try
    aBuffer.First;
    while not aBuffer.Eof do
    begin
      aBuffer.Edit;
      if ( FCasMode = 2 ) then
      begin
        aMark := aBuffer.FieldByName( 'PROGRESS' ).AsString;
        aMark := IndexOfStringReplace( aMark, FNagraStep, '0' );
        aBuffer.FieldByName( 'PROGRESS' ).AsString := aMark;
      end else
      begin
        aBuffer.FieldByName( 'PROGRESS' ).AsString := aMark;
      end;
      aBuffer.Post;
      aBuffer.Next;
    end;
  finally
    aBuffer.EnableControls;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.BegindDvnSet;
begin
  FThreadIsRuning := True;
  FCmdHasErr := False;
  btnImport.Enabled := False;
  actDvnSetup.Enabled := False;
  actDvnClear.Enabled := False;
  txtDvnStb.Properties.ReadOnly := True;
  txtDvnIcc.Properties.ReadOnly := True;
  txtDays.Properties.ReadOnly := True;
  Panel5.Enabled := False;
  btnDvnCannel.OnClick := nil;
  btnChannelClear.OnClick := nil;
  gvDVN.DataController.DataModeController.SyncMode := False;
  cmbCAS.Enabled := False;
  {}
  Panel4.BevelInner := bvNone;
  Panel4.BevelOuter := bvNone;
  DvnProgressBar.Visible := True;
  DvnProgressBar.Position := 0;
  DvnProgressBar.Properties.Max := cdsDVN.RecordCount;
  if ( FDvnQueryParam.Param1 <> EmptyStr ) then
   DvnProgressBar.Properties.Max := cdsDVN.RecordCount * 2;
  {}
  FLogList.Clear;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.EndDvnSet;
begin
  FThreadIsRuning := False; 
  btnImport.Enabled := True;
  actDvnSetup.Enabled := True;
  actDvnClear.Enabled := True;
  txtDvnStb.Properties.ReadOnly := False;
  txtDvnIcc.Properties.ReadOnly := False;
  txtDays.Properties.ReadOnly := False;
  Panel5.Enabled := True;
  btnDvnCannel.OnClick := btnDvnCannelClick;
  btnChannelClear.OnClick := btnChannelClearClick;
  gvDVN.DataController.DataModeController.SyncMode := True;
  cmbCAS.Enabled := True;
  {}
  Panel4.BevelInner := bvRaised;
  Panel4.BevelOuter := bvLowered;
  DvnProgressBar.Visible := False;
  DvnProgressBar.Position := 0;
  {}
  if ( FCmdHasErr ) then
  begin
    FLogList.Insert( 0, Format(
      '***************************** %s *****************************',
      [FormatDateTime('yyyy/mm/dd hh:nn:ss', Now)] ) );
    FLogList.Add(
      '*******************************************************************************' );
    FLogList.SaveToFile( FCmdLogFileName );
  end;
  FDvnThread := nil;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.BeginNagraSet;
begin
  FThreadIsRuning := True;
  FCmdHasErr := False;
  btnImport.Enabled := False;
  actNagraSetup.Enabled := False;
  actNagraClear.Enabled := False;
  txtNagraStb.Properties.ReadOnly := True;
  txtNagraIcc.Properties.ReadOnly := True;
  txtNagraDays.Properties.ReadOnly := True;
  gvNagra.DataController.DataModeController.SyncMode := False;
  {}
  Panel6.BevelInner := bvNone;
  Panel6.BevelOuter := bvNone;
  NagraProgressBar.Visible := True;
  NagraProgressBar.Position := 0;
  NagraProgressBar.Properties.Max := cdsNagra.RecordCount;
  {}
  cmbCAS.Enabled := False;
  cmbNagraOpeartion.Enabled := False;
  txtZipCode.Properties.ReadOnly := True;
  txtBatId.Properties.ReadOnly := True;
  btnNagraChannel.OnClick := nil;
  btnChannelClear2.OnClick := nil;
  {}
  FLogList.Clear;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.EndNagraSet;
begin
  FThreadIsRuning := False; 
  btnImport.Enabled := True;
  actNagraSetup.Enabled := True;
  actNagraClear.Enabled := True;
  txtNagraStb.Properties.ReadOnly := False;
  txtNagraIcc.Properties.ReadOnly := False;
  txtNagraDays.Properties.ReadOnly := False;
  gvNagra.DataController.DataModeController.SyncMode := True;
  {}
  Panel6.BevelInner := bvRaised;
  Panel6.BevelOuter := bvLowered;
  NagraProgressBar.Visible := False;
  NagraProgressBar.Position := 0;
  {}
  cmbCAS.Enabled := True;
  cmbNagraOpeartion.Enabled := True;
  txtZipCode.Properties.ReadOnly := False;
  txtBatId.Properties.ReadOnly := False;
  btnNagraChannel.OnClick := btnDvnCannelClick;
  btnChannelClear2.OnClick := btnChannelClearClick;
  {}
  if ( FCmdHasErr ) then
  begin
    FLogList.Insert( 0, Format(
      '***************************** %s *****************************',
      [FormatDateTime('yyyy/mm/dd hh:nn:ss', Now)] ) );
    FLogList.Add(
      '*******************************************************************************' );
    FLogList.SaveToFile( FCmdLogFileName );
  end;
  FNagraThread := nil;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.DisplayNagraProgressBar;
var
  aCmds: array [1..4] of string;
  aMark, aCmdStatus: String;
begin
  aCmds[1] := 'CMDA1';
  aCmds[2] := 'CMDB1';
  aCmds[3] := 'CMDB3';
  aCmds[4] := 'CMDB1_1';
  aMark := cdsNagra.FieldByName( 'PROGRESS' ).AsString;
  aCmdStatus := cdsNagra.FieldByName( aCmds[FNagraStep] ).AsString;
  if ( ( aCmdStatus = 'C'  ) or
       ( aCmdStatus = 'E'  ) or
       ( aCmdStatus = 'S'  ) ) and
     ( Copy( aMark, FNagraStep, 1 ) = '0' ) then
  begin
    NagraProgressBar.Position := ( NagraProgressBar.Position + 1 );
    aMark := IndexOfStringReplace( aMark, FNagraStep, '1' );
    cdsNagra.Edit;
    cdsNagra.FieldByName( 'PROGRESS' ).AsString := aMark;
    cdsNagra.Post;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.DisplayDvnProgressBar;
var
  aCmdA1, aCmdB1, aMark: String;
begin
  aCmdA1 := cdsDVN.FieldByName( 'CMDA1' ).AsString;
  aCmdB1 := cdsDVN.FieldByName( 'CMDB1' ).AsString;
  aMark := cdsDVN.FieldByName( 'PROGRESS' ).AsString;
  if ( ( aCmdA1 = 'C' ) or ( aCmdA1 = 'E' ) ) and
     ( ( aMark = '00' ) or ( aMark = '01' ) ) then
  begin
    NagraProgressBar.Position := ( NagraProgressBar.Position + 1 );
    aMark := '1' + Copy( aMark, 2, 1 );
    cdsDVN.Edit;
    cdsDVN.FieldByName( 'PROGRESS' ).AsString := aMark;
    cdsDVN.Post;
  end;
  {}
  if ( FDvnQueryParam.Param1 <> EmptyStr ) then
  begin
    if ( ( aCmdB1 = 'C' ) or ( aCmdB1 = 'E' ) ) and
       ( ( aMark = '00' ) or ( aMark = '10' ) ) then
    begin
      NagraProgressBar.Position := ( NagraProgressBar.Position + 1 );
      aMark := Copy( aMark, 1, 1 ) + '1';
      cdsDVN.Edit;
      cdsDVN.FieldByName( 'PROGRESS' ).AsString := aMark;
      cdsDVN.Post;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.SwitchingCAS(ACasText: String);
begin
  if SameText( ACasText, 'DVN' ) then
  begin
    FCasMode := 5;
    {}
    actNagraSetup.ShortCut := 0;
    actNagraClear.ShortCut := 0;
    actDvnSetup.ShortCut := VK_F6;
    actDvnClear.ShortCut := VK_F4;
    {}
    btnSetup.Action := actDvnSetup;
    btnClear.Action := actDvnClear;
    MainPage.ActivePageIndex := 0;
    Application.Title := 'DVN 配對燒機作業[SO1621L]';
  end else
  begin
    FCasMode := 2;
    {}
    actNagraSetup.ShortCut := VK_F6;
    actNagraClear.ShortCut := VK_F4;
    actDvnSetup.ShortCut := 0;
    actDvnClear.ShortCut := 0;
    {}
    btnSetup.Action := actNagraSetup;
    btnClear.Action := actNagraClear;
    MainPage.ActivePageIndex := 1;
    Application.Title := 'NAGRA 配對燒機作業[SO1621L]';
  end;
  Self.Caption := Application.Title;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.SwitchingCAS(ACasMode: Integer);
var
  aIndex: Integer;
  aFindText: String;
begin
  aFindText := 'NAGRA';
  if ( ACasMode = 5 ) then aFindText := 'DVN';
  for aIndex := 0 to cmbCAS.Properties.Items.Count - 1 do
  begin
    if SameText( cmbCAS.Properties.Items[aIndex], aFindText ) then
    begin
      cmbCAS.ItemIndex := aIndex;
      Break;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.SwitchNagraOperation(AOperationText: String);
begin
  if SameText( AOperationText, '新品開機配對' ) then
  begin
    FNagraOperation := 1;
    FNagraStep := 1;
    lbOpeartion.Caption := '開機 -> 開試看頻道';
  end else
  if SameText( AOperationText, '舊品開機配對' ) then
  begin
    FNagraOperation := 2;
    FNagraStep := 1;
    lbOpeartion.Caption := '開機 -> 開試看頻道 -> 取消所有頻道 -> 開試看頻道'
  end else if SameText( AOperationText, '舊品回收測試' ) then
  begin
    FNagraOperation := 3;
    FNagraStep := 1;
    lbOpeartion.Caption := '開機 -> 開試看頻道 -> 取消所有頻道'
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.SwitchNagraStep;
var
  aLimit: Integer; 
begin
  Inc( FNagraStep );
  if ( FNagraStep in [2,4] ) then
  begin
    if ( FNagraQueryParam.Param1 = EmptyStr ) then
      Inc( FNagraStep );
  end;
  case FNagraOperation of
    1: aLimit := 2;
    2: aLimit := 4;
    3: aLimit := 3;
  else
    aLimit := 2;
  end;
  if ( FNagraStep > aLimit ) then FNagraStep := 1;
  actNagraSetup.Caption := Format( 'F6.%s', [GetNagraStepText( FNagraStep ) ] );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.RunDvnSet;
begin
  BegindDvnSet;
  SetChannelBufferDays;
  ResetProgressFlag;
  FDvnThread := TDvnThread.Create( True );
  FDvnThread.OnTerminate := DvnThreadTerminate;
  TDvnThread( FDvnThread ).SourceDataSet := cdsDVN;
  TDvnThread( FDvnThread ).ChannelDataSet := cdsDvnChannel;
  TDvnThread( FDvnThread ).SourceConnection := DBController.DataConnection;
  TDvnThread( FDvnThread ).CompCode := DBController.LoginInfo.CompCode;
  TDvnThread( FDvnThread ).SmsOperator := DBController.LoginInfo.UserName;
  TDvnThread( FDvnThread ).CmdA1Text := GetCmdA1Text;
  TDvnThread( FDvnThread ).ReSendErrCmd := chkDvnReSendErr.Checked;

  FDvnThread.Resume;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.RunNagraSet;
begin
  BeginNagraSet;
  SetChannelBufferDays;
  ResetProgressFlag;
  SetNagraSkipFlag;
  FNagraThread := TNagraThread.Create( True );
  FNagraThread.OnTerminate := NagraThreadTerminate;
  TNagraThread( FNagraThread ).SourceDataSet := cdsNagra;
  TNagraThread( FNagraThread ).ChannelDataSet := cdsNagraChannel;
  TNagraThread( FNagraThread ).SourceConnection := DBController.DataConnection;
  TNagraThread( FNagraThread ).CompCode := DBController.LoginInfo.CompCode;
  TNagraThread( FNagraThread ).SmsOperator := DBController.LoginInfo.UserName;
  TNagraThread( FNagraThread ).ZipCode := txtZipCode.Text;
  TNagraThread( FNagraThread ).BatId := txtBatId.Text;
  TNagraThread( FNagraThread ).ResendErrCmd := chkNagraReSendErr.Checked;
  TNagraThread( FNagraThread ).NagraOperation := FNagraOperation;
  TNagraThread( FNagraThread ).NagraStep := FNagraStep;
  //#5216 如果CD039.NagraType=1 則不要進行Upd SOAC0201C By Kin 2009/07/30
  TNagraThread( FNagraThread ).AutoFill:=FAutoFill;
  FNagraThread.Resume;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.DoDvnProgress(const ACmdType, ARecNo: Integer;
  const ACmdStatus, AErrMsg: String);
var
  aCmdA1, aCmdB1, aLogText: String;
begin
  cdsDVN.DisableControls;
  try
    if ( aCmdType = CmdA1 ) then
    begin
      if ( cdsDVN.Locate( 'RECNO', aRecNo, [] ) ) then
      begin
        cdsDVN.Edit;
        cdsDVN.FieldByName( 'CMDA1' ).AsString := aCmdStatus;
        cdsDVN.FieldByName( 'ERRMSG' ).AsString := aErrMsg;
        cdsDVN.Post;
      end;
    end else
    if ( aCmdType = CmdB1 ) then
    begin
      if ( cdsDVN.Locate( 'RECNO', aRecNo, [] ) ) then
      begin
        cdsDVN.Edit;
        cdsDVN.FieldByName( 'CMDB1' ).AsString := aCmdStatus;
        cdsDVN.FieldByName( 'ERRMSG' ).AsString := aErrMsg;
        cdsDVN.Post;
      end;
    end;
    {}
    aCmdA1 := cdsDVN.FieldByName( 'CMDA1' ).AsString;
    aCmdB1 := cdsDVN.FieldByName( 'CMDB1' ).AsString;
    {}
    if ( aCmdA1 = 'E' ) then
    begin
      FCmdHasErr := True;
      aLogText := Format(
        '開機/配對: 第%d筆指令有誤, 機上盒序號=%s, 智慧卡序號=%s, 訊息:%S',
        [cdsDVN.RecNo,
         cdsDVN.FieldByName( 'STBNO' ).AsString,
         cdsDVN.FieldByName( 'ICCNO' ).AsString,
         aErrMsg] );
      FLogList.Add( aLogText );
    end;
    {}
    if ( aCmdB1 = 'E' ) then
    begin
      FCmdHasErr := True;
      aLogText := Format(
        '開試看頻道: 第%d筆指令有誤, 機上盒序號=%s, 智慧卡序號=%s, 訊息:%S',
        [cdsDVN.RecNo,
         cdsDVN.FieldByName( 'STBNO' ).AsString,
         cdsDVN.FieldByName( 'ICCNO' ).AsString,
         aErrMsg] );
      FLogList.Add( aLogText );
    end;
    {}
    DisplayDvnProgressBar;
    {}
  finally
    cdsDVN.EnableControls;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.DoNagraProgress(const ACmdType, ARecNo: Integer;
  const ACmdStatus, AErrMsg: String);
var
  aFieldName, aStepText, aLogText: String;
begin
  cdsNagra.DisableControls;
  try
    aFieldName := EmptyStr;
    {}
    if ( aCmdType = CmdA1 ) then
      aFieldName := 'CMDA1'
    else if ( aCmdType = CmdB1 ) then
      aFieldName := 'CMDB1'
    else if ( aCmdType = CmdB3 ) then
      aFieldName := 'CMDB3'
    else if ( aCmdType = CmdB1_1 ) then
      aFieldName := 'CMDB1_1';
    {}
    if ( aFieldName <> EmptyStr ) then
    begin
      if ( cdsNagra.Locate( 'RECNO', aRecNo, [] ) ) then
      begin
        cdsNagra.Edit;
        cdsNagra.FieldByName( aFieldName ).AsString := aCmdStatus;
        cdsNagra.FieldByName( 'ERRMSG' ).AsString := aErrMsg;
        cdsNagra.Post;
      end;
      if ( cdsNagra.FieldByName( aFieldName ).AsString = 'E' ) then
      begin
        FCmdHasErr := True;
        aStepText := GetNagraStepText( FNagraStep );
        aLogText := Format(
          '%s: 第%d筆指令有誤, 機上盒序號=%s, 智慧卡序號=%s, 訊息:%S',
          [aStepText,
           cdsNagra.RecNo,
           cdsNagra.FieldByName( 'STBNO' ).AsString,
           cdsNagra.FieldByName( 'ICCNO' ).AsString,
           aErrMsg] );
        FLogList.Add( aLogText );
      end;
    end;
    DisplayNagraProgressBar;
  finally
    cdsNagra.EnableControls;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.cdsDVNAfterScroll(DataSet: TDataSet);
var
  aRecNo: Integer;
  aText: String;
begin
  aRecNo := DataSet.RecNo;
  if ( aRecNo < 0 ) then aRecNo := 0;
  aText := Format( ' %d / %d ', [aRecNo, DataSet.RecordCount] );
  if ( FCasMode = 2 ) then
    lbNagraRecords.Caption := aText
  else
    lbDvnRecords.Caption := aText;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.txtDvnStbKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  aMsg1, aMsg2: String;
  aNextControl: TcxTextEdit;
  aNextCtl : Boolean;
begin

  if ( Key = VK_RETURN ) then
  begin
    aNextCtl := True;
    if ( TcxTextEdit( Sender ).Text <> EmptyStr ) then
    begin
      aMsg2 := EmptyStr;
      {}
      if not VdStbNo( TcxTextEdit( Sender ).Text, aMsg2 ) then
      begin
        aMsg1 := '該機上盒序號未入庫或入庫檔無此機上盒序號。';
        if ( aMsg2 <> EmptyStr ) then aMsg1 := ( aMsg1 + #13#10 + aMsg2 );
        WarningMsg( aMsg1 );
        TcxTextEdit( Sender ).SelectAll;
        Exit;
      end;
      {}
      if not VdStbNoLength( TcxTextEdit( Sender ).Text, aMsg2 ) then
      begin
        aMsg1 := '該機上盒序號檢核有誤。';
        if ( aMsg2 <> EmptyStr ) then aMsg1 := ( aMsg1 + #13#10 + aMsg2 );
        WarningMsg( aMsg1 );
        TcxTextEdit( Sender ).SelectAll;
        Exit;
      end;
      if FAutoFill then
      begin
        qryAutoFill.Close;
        qryAutoFill.SQL.Text := Format('Select SMARTCARDNO From SOAC0201C ' +
                                       'Where STBSNO=''%s'' And CompCode=%s ',
                                       [ txtNagraStb.Text ,DBController.LoginInfo.CompCode ]);
        qryAutoFill.Open;
        if not qryAutoFill.IsEmpty then
        begin
          txtNagraIcc.Text := qryAutoFill.FieldByName('SMARTCARDNO').AsString;
          //#5216 要自動輸入 By Kin 2009/07/30
          txtDvnIccKeyDown(txtNagraIcc ,Key,[]);
          aNextCtl:=False;
        end;
      end;
      aNextControl := txtDvnIcc;
      if ( FCasMode = 2 ) then
        aNextControl := txtNagraIcc;
      if aNextCtl then
      begin
        if ( aNextControl.CanFocusEx ) then aNextControl.SetFocus;
      end;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.txtDvnIccKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  aMsg1, aMsg2: String;
  aNextControl: TcxTextEdit;
  aGridResult: Boolean;
begin
  if ( Key = VK_RETURN ) then
  begin
    if ( TcxTextEdit( Sender ).Text <> EmptyStr ) then
    begin
      aMsg2 := EmptyStr;
      if not VdIccNo( TcxTextEdit( Sender ).Text, aMsg2 ) then
      begin
        aMsg1 := '該智慧卡序號未入庫或入庫檔無此智慧卡序號。';
        if ( aMsg2 <> EmptyStr ) then aMsg1 := ( aMsg1 + #13#10 + aMsg2 );
        WarningMsg( aMsg1 );
        TcxTextEdit( Sender ).SelectAll;
        Exit;
      end;
      {}
      if not VdIccNoLength( TcxTextEdit( Sender ).Text, aMsg2 ) then
      begin
        aMsg1 := '該智慧卡序號檢核有誤。';
        if ( aMsg2 <> EmptyStr ) then aMsg1 := ( aMsg1 + #13#10 + aMsg2 );
        WarningMsg( aMsg1 );
        TcxTextEdit( Sender ).SelectAll;
        Exit;
      end;
      {}
      aNextControl := txtDvnStb;
      if ( FCasMode = 2 ) then aNextControl := txtNagraStb;
      {}
      if ( aNextControl.Text = EmptyStr ) then
      begin
        if ( aNextControl.CanFocusEx ) then aNextControl.SetFocus;
        Exit;
      end;
      {}
      if not VdStbNo( aNextControl.Text, aMsg2 ) then
      begin
        aMsg1 := '該機上盒序號未入庫或入庫檔無此機上盒序號。';
        if ( aMsg2 <> EmptyStr ) then aMsg1 := ( aMsg1 + #13#10 + aMsg2 );
        WarningMsg( aMsg1 );
        if ( aNextControl.CanFocusEx ) then aNextControl.SetFocus;
        Exit;
      end;
      {}
      if not VdStbNoLength( aNextControl.Text, aMsg2 ) then
      begin
        aMsg1 := '該機上盒序號檢核有誤。';
        if ( aMsg2 <> EmptyStr ) then aMsg1 := ( aMsg1 + #13#10 + aMsg2 );
        WarningMsg( aMsg1 );
        if ( aNextControl.CanFocusEx ) then aNextControl.SetFocus;
        Exit;
      end;
      {}
      aMsg1 := EmptyStr;
      if ( FCasMode = 2 ) then
        aGridResult := AddToNagraGrid( aNextControl.Text, TcxTextEdit( Sender ).Text, aMsg1 )
      else
        aGridResult := AddToDvnGrid( aNextControl.Text, TcxTextEdit( Sender ).Text, aMsg1 );
      if ( not aGridResult ) then
      begin
        WarningMsg( '無法輸入此組機上盒/智慧卡資訊。'+ #13#10 +
          Format( '訊息:%s', [aMsg1] ) );
        TcxTextEdit( Sender ).SelectAll;
        Exit;
      end;
      {}
      TcxTextEdit( Sender ).Clear;
      aNextControl.Clear;
      if ( aNextControl.CanFocusEx ) then aNextControl.SetFocus;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.txtDaysKeyPress(Sender: TObject; var Key: Char);
begin
  if not ( Key in [ '0'..'9', #8 ] ) then Abort;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.txtZipCodeExit(Sender: TObject);
begin
  if ( TcxTextEdit( Sender ).Text = EmptyStr ) then
    LoadDefaultZipCode;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.txtBatIdExit(Sender: TObject);
begin
  if ( TcxTextEdit( Sender ).Text = EmptyStr ) then
    LoadDefaultBatId;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.txtNagraDaysExit(Sender: TObject);
begin
  if ( TcxTextEdit( Sender ).Text = EmptyStr ) then
    TcxTextEdit( Sender ).Text := '0';
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.RecvParamString: Boolean;
var
  aParamText, aDbText, aValue: String;
  aIndex: Integer;
  aCrypt32: _Crypt32;
begin
  Result := ( ParamCount >= 9 );
  if not Result then
  begin
    WarningMsg( '傳入之程式啟動參數有誤。' );
    Exit;
  end;
  { 傳入之參數 有 2 種,
    1種是掛在客服系統下,
    1種是掛在 powerget 寫的動態 menu 下 }
  if ( ParamCount > 9 ) then
  begin
    aParamText := EmptyStr;
    for aIndex := 1 to ParamCount do
      aParamText := ( aParamText + ParamStr( aIndex ) + #32 );
    {}
    aCrypt32 := CoCrypt32.Create;
    try
      aParamText := aCrypt32.DValue( aParamText, EmptyParam, 'CsMisk' );

      DBController.LoginInfo.IsSuperVisor := False;
      DBController.LoginInfo.UserId := ExtractValue( aParamText, #9 );
      DBController.LoginInfo.UserName := ExtractValue( aParamText, #9 );
      {}
      aDbText := ExtractValue( aParamText, #9 );
      ExtractValue( aDbText, ';' );
      {}
      aValue := ExtractValue( aDbText, ';' );
      ExtractValue( aValue, '=' );
      DBController.LoginInfo.DbAliase := aValue;
      ExtractValue( aDbText, ';' );
      aValue := ExtractValue( aDbText, ';' );
      ExtractValue( aValue, '=' );
      DBController.LoginInfo.DbAccount := aValue;
      ExtractValue( aDbText, '=' );
      DBController.LoginInfo.DbPassword := aDbText;
      {}
      ExtractValue( aParamText, #9 );
      DBController.LoginInfo.CompCode := ExtractValue( aParamText, #9 );
      DBController.LoginInfo.CompName := EmptyStr;
      DBController.LoginInfo.ServiceType := EmptyStr;
    finally
      aCrypt32 := nil;
    end;
  end else
  begin
    { 動態Menu下}
    { 客服系統下 N 1 Nick 12 大安文山 emcdw emcdw mis c }
    DBController.LoginInfo.IsSuperVisor := ( ParamStr( 1 ) = 'Y' );
    DBController.LoginInfo.UserId := ParamStr( 2 );
    DBController.LoginInfo.UserName := ParamStr( 3 );
    DBController.LoginInfo.CompCode := ParamStr( 4 );
    DBController.LoginInfo.CompName := ParamStr( 5 );
    DBController.LoginInfo.DbAccount := WideUpperCase( ParamStr( 6 ) );
    DBController.LoginInfo.DbPassword := ParamStr( 7 );
    DBController.LoginInfo.DbAliase := ParamStr( 8 );
    DBController.LoginInfo.ServiceType := ParamStr( 9 );
  end; 
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.AddToDvnGrid(const AStbNo, AIccNo: String; var AMsg: String): Boolean;
begin
  AMsg := EmptyStr;
  cdsDVN.DisableControls;
  try
    if ( not cdsDVN.IsEmpty ) then
    begin
      if cdsDVN.Locate( 'STBNO', AStbNo, [] ) then
        AMsg := '機上盒序號重覆輸入。';
      {}
      if cdsDVN.Locate( 'ICCNO', AIccNo, [] ) then
        AMsg := ( AMsg + '智慧卡卡號重覆輸入。' );
    end;
    {}
    if ( AMsg = EmptyStr ) then
    begin
      cdsDVN.Append;
      cdsDVN.FieldByName( 'RECNO' ).AsInteger := GetRecNo;
      cdsDVN.FieldByName( 'PROGRESS' ).AsString := '00';
      cdsDVN.FieldByName( 'SEQNO1' ).AsString := EmptyStr;
      cdsDVN.FieldByName( 'SEQNO2' ).AsString := EmptyStr;
      cdsDVN.FieldByName( 'STBNO' ).AsString := Copy( AStbNo, 1, FDvnCD022Info.StbSNoLength );
      cdsDVN.FieldByName( 'ICCNO' ).AsString := Copy( AIccNo, 1, FDvnCD022Info.IccSNoLength );
      cdsDVN.FieldByName( 'CMDA1' ).AsString := EmptyStr;
      cdsDVN.FieldByName( 'CMDB1' ).AsString := EmptyStr;
      cdsDVN.FieldByName( 'ERRMSG' ).AsString := EmptyStr;
      cdsDVN.FieldByName( 'HASLOG1' ).AsString := EmptyStr;
      cdsDVN.FieldByName( 'HASLOG2' ).AsString := EmptyStr;
      cdsDVN.Post;
    end;
  finally
    cdsDVN.EnableControls;
  end;
  Result := ( AMsg = EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.AddToNagraGrid(const AStbNo, AIccNo: String;
  var AMsg: String): Boolean;
begin
  AMsg := EmptyStr;
  cdsNagra.DisableControls;
  try
    if ( not cdsNagra.IsEmpty ) then
    begin
      if cdsNagra.Locate( 'STBNO', AStbNo, [] ) then
        AMsg := '機上盒序號重覆輸入。';
      {}
      if cdsNagra.Locate( 'ICCNO', AIccNo, [] ) then
        AMsg := ( AMsg + '智慧卡卡號重覆輸入。' );
    end;
    {}
    if ( AMsg = EmptyStr ) then
    begin
      cdsNagra.Append;
      cdsNagra.FieldByName( 'RECNO' ).AsInteger := GetRecNo;
      cdsNagra.FieldByName( 'PROGRESS' ).AsString := '0000';
      cdsNagra.FieldByName( 'SEQNO1' ).AsString := EmptyStr;
      cdsNagra.FieldByName( 'SEQNO2' ).AsString := EmptyStr;
      cdsNagra.FieldByName( 'SEQNO3' ).AsString := EmptyStr;
      cdsNagra.FieldByName( 'SEQNO4' ).AsString := EmptyStr;
      cdsNagra.FieldByName( 'STBNO' ).AsString := Copy( AStbNo, 1, FNagraCD022Info.StbSNoLength );
      cdsNagra.FieldByName( 'ICCNO' ).AsString := Copy( AIccNo, 1, FNagraCD022Info.IccSNoLength );
      cdsNagra.FieldByName( 'CMDA1' ).AsString := EmptyStr;
      cdsNagra.FieldByName( 'CMDB1' ).AsString := EmptyStr;
      cdsNagra.FieldByName( 'CMDB3' ).AsString := EmptyStr;
      cdsNagra.FieldByName( 'CMDB1_1' ).AsString := EmptyStr;
      cdsNagra.FieldByName( 'ERRMSG' ).AsString := EmptyStr;
      cdsNagra.FieldByName( 'HASLOG1' ).AsString := EmptyStr;
      cdsNagra.FieldByName( 'HASLOG2' ).AsString := EmptyStr;
      cdsNagra.FieldByName( 'HASLOG3' ).AsString := EmptyStr;
      cdsNagra.FieldByName( 'HASLOG4' ).AsString := EmptyStr;
      cdsNagra.Post;
    end;
  finally
    cdsNagra.EnableControls;
  end;
  Result := ( AMsg = EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.LoadFromFile(const AFileName: String): Boolean;
var
  aIndex: Integer;
  aText, aStb, aIcc, aLogText, aVdMsg: String;
  aCheckOk: Boolean;
begin
  FLogList.Clear;
  try
    FImportList.Clear;
    FImportList.LoadFromFile( AFileName );
    for aIndex := 0 to FImportList.Count - 1 do
    begin
      Application.ProcessMessages;
      aText := FImportList.Strings[aIndex];
      aStb := Trim( ExtractValue( aText ) );
      aIcc := Trim( ExtractValue( aText ) );
      aCheckOk := True;
      if ( aStb = EmptyStr ) or ( aIcc = EmptyStr ) then
      begin
        aLogText := Format(
          '匯檔: 第%d行有誤 --> %s, 機上盒序號=%s, 智慧卡序號=%s',
          [aIndex+1, aStb, Nvl( aStb, '空值' ), Nvl( aIcc, '空值' )] );
        FLogList.Add( aLogText );
        aCheckOk := False;
      end;
      {}
      if ( aCheckOk ) then
      begin
        if not VdStbNo( aStb, aVdMsg ) then
        begin
          aLogText := Format(
           '匯檔: 第%d行有誤, 機上盒序號 (%s) 未入庫或入庫檔無此機上盒序號。',
           [aIndex+1, aStb] );
          FLogList.Add( aLogText );
          aCheckOk := False;
        end;
        {}
        if not VdStbNoLength( aStb, aVdMsg ) then
        begin
          aLogText := Format(
           '匯檔: 第%d行有誤, 機上盒序號 (%s) 長度檢核不足。',
           [aIndex+1, aStb] );
          FLogList.Add( aLogText );
          aCheckOk := False;
        end;
        {}
        if not VdIccNo( aIcc, aVdMsg ) then
        begin
          aLogText := Format(
           '匯檔: 第%d行有誤, 智慧卡序號 (%s) 未入庫或入庫檔無此智慧卡序號。',
           [aIndex+1, aIcc] );
          FLogList.Add( aLogText );
          aCheckOk := False;
        end;
        {}
        if not VdIccNoLength( aIcc, aVdMsg ) then
        begin
          aLogText := Format(
           '匯檔: 第%d行有誤, 智慧卡序號 (%s) 長度檢核不足。',
           [aIndex+1, aIcc] );
          FLogList.Add( aLogText );
          aCheckOk := False;
        end;
        {}
      end;
      {}
      if ( aCheckOk ) then
      begin
        if FCasMode = 2 then
        begin
          if not AddToNagraGrid( aStb, aIcc, aVdMsg ) then
          begin
            aLogText := Format(
             '匯檔: 第%d行有誤 --> %s,%s, %s 。',
             [aIndex+1, aStb, aIcc, aVdMsg] );
            FLogList.Add( aLogText );
          end;
        end else
        begin
          if not AddToDvnGrid( aStb, aIcc, aVdMsg ) then
          begin
            aLogText := Format(
             '匯檔: 第%d行有誤 --> %s,%s, %s 。',
             [aIndex+1, aStb, aIcc, aVdMsg] );
            FLogList.Add( aLogText );
          end;
        end;
      end;
      Application.ProcessMessages;
    end;
    if ( FLogList.Count > 0 ) then
    begin
      FLogList.Insert( 0, Format(
        '***************************** %s *****************************',
        [FormatDateTime('yyyy/mm/dd hh:nn:ss', Now)] ) );
      FLogList.Add(
        '*******************************************************************************' );
      FLogList.SaveToFile( FImportLogFileName );
    end;
  finally
    Result := ( FLogList.Count <= 0 );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.VdIccNo(const AIccNo: String; var AMsg: String): Boolean;
var
  aSql: string;
begin
  AMsg := EmptyStr;
  if ( FCasMode = 2 ) then
  begin
    aSql := Format(
      ' SELECT COUNT(1) FROM SOAC0201B ' +
      '  WHERE COMPCODE = ''%S''       ' +
      '    AND FACISNO LIKE ''%S''        ',
      [DBController.LoginInfo.CompCode, Copy( AIccNo, 1, 10 ) + '%'] );
  end else
  begin
    aSql := Format(
      ' SELECT COUNT(1) FROM SOAC0501B ' +
      '  WHERE COMPCODE = ''%S''       ' +
      '    AND FACISNO = ''%S''        ',
      [DBController.LoginInfo.CompCode, AIccNo] );
  end;
  DBController.DataReader.Close;
  DBController.DataReader.SQL.Text := aSql;
  try
    DBController.DataReader.Open;
  except
    on E: Exception do AMsg := E.Message;
  end;
  Result := ( AMsg = EmptyStr );
  if ( Result ) then
    Result := ( DBController.DataReader.Fields[0].AsInteger > 0 );
  DBController.DataReader.Close;
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.VdStbNo(const AStbNo: String; var AMsg: String): Boolean;
var
  aSql: string;
begin
  AMsg := EmptyStr;
  if ( FCasMode = 2 ) then
  begin
    { Nagra }
    aSql := Format(
      ' SELECT COUNT(1) FROM SOAC0201A ' +
      '  WHERE COMPCODE = ''%S''       ' +
      '    AND FACISNO LIKE ''%S''        ',
      [DBController.LoginInfo.CompCode, Copy( AStbNo, 1, 10 ) + '%' ] );
  end else
  begin
    { DVN }
    aSql := Format(
      ' SELECT COUNT(1) FROM SOAC0501A ' +
      '  WHERE COMPCODE = ''%S''       ' +
      '    AND FACISNO = ''%S''        ',
      [DBController.LoginInfo.CompCode, AStbNo] );
  end;
  DBController.DataReader.Close;
  DBController.DataReader.SQL.Text := aSql;
  try
    DBController.DataReader.Open;
  except
    on E: Exception do AMsg := E.Message;
  end;
  Result := ( AMsg = EmptyStr );
  if ( Result ) then
    Result := ( DBController.DataReader.Fields[0].AsInteger > 0 );
  DBController.DataReader.Close;
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.VdIccNoLength(const AIccNo: String; var AMsg: String): Boolean;
begin
  AMsg := EmptyStr;
  if ( FCasMode = 2 ) then
  begin
    { Nagra }
    if ( Length( AIccNo ) < FNagraCD022Info.IccSNoLength ) then
      AMsg := Format( '智慧卡卡號長度不足, 卡號總長度應為%d碼。',
        [FNagraCD022Info.IccSNoLength] );
  end else
  begin
    { DVN }
    if ( Length( AIccNo ) < FDvnCD022Info.IccSNoLength ) then
      AMsg := Format( '智慧卡卡號長度不足, 卡號總長度應為%d碼。',
        [FDvnCD022Info.IccSNoLength] );
  end;
  Result := ( AMsg = EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.VdStbNoLength(const AStbNo: String; var AMsg: String): Boolean;
begin
  AMsg := EmptyStr;
  if ( FCasMode = 2 ) then
  begin
    { Nagra }
    if ( Length( AStbNo ) < FNagraCD022Info.StbSNoLength ) then
      AMsg := Format( '機上盒序號長度不足, 序號總長度應為%d碼。',
        [FNagraCD022Info.StbSNoLength] );
  end else
  begin
    { DVN }
    if ( Length( AStbNo ) < FDvnCD022Info.StbSNoLength ) then
      AMsg := Format( '機上盒序號長度不足, 序號總長度應為%d碼。',
        [FDvnCD022Info.StbSNoLength] );
  end;
  Result := ( AMsg = EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.VdMustInput: Boolean;
var
  aDays: Integer;
begin
  Result := True;
  if ( txtZipCode.Text = EmptyStr ) then
  begin
    Result := False;
    WarningMsg( '請輸入 Zip Code 。' );
    if ( txtZipCode.CanFocusEx ) then txtZipCode.SetFocus;
    Exit;
  end;
  {}
  if ( txtBatId.Text = EmptyStr ) then
  begin
    Result := False;
    WarningMsg( '請輸入BAT 區碼。' );
    if ( txtBatId.CanFocusEx ) then txtBatId.SetFocus;
    Exit;
  end;
  {}
  aDays := StrToIntDef( txtNagraDays.Text, 0 );
  if ( aDays > 7 ) then
  begin
    Result := False;
    WarningMsg( '工作緩衝日，必須小於等於7天。' );
    if ( txtNagraDays.CanFocusEx ) then txtNagraDays.SetFocus;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.GetImportLogFileName: String;
var
  aPath: String;
begin
  aPath := ExtractFilePath( ParamStr( 0 ) );
  StringReplace( aPath, 'Bin', 'Errlog', [rfReplaceAll,rfIgnoreCase] );
  Result := IncludeTrailingPathDelimiter( aPath ) + 'SO1621L_Importlog.txt';
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.GetCmdLogFileName: String;
var
  aPath: String;
begin
  aPath := ExtractFilePath( ParamStr( 0 ) );
  StringReplace( aPath, 'Bin', 'Errlog', [rfReplaceAll,rfIgnoreCase] );
  Result := IncludeTrailingPathDelimiter( aPath ) + 'SO1621L_Cmdlog.txt';
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.GetRecNo: Integer;
begin
  if ( FCasMode = 2 ) then
  begin
    Inc( FNagraRecNo );
    Result := FNagraRecNo;
  end else
  begin
    Inc( FDvnRecNo );
    Result := FDvnRecNo;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.GetCmdA1Text: String;
begin
  Result := 'A1';
  if ( rdoA7.Checked ) then Result := 'A7';
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.GetNagraStepText(const AStep: Integer): String;
begin
  case AStep of
    1: Result := '開機';
    2: Result := '開試看頻道';
    3: Result := '取消所有頻道';
    4: Result := '開試看頻道';
  else
    Result := EmptyStr;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.btnDvnCannelClick(Sender: TObject);
var
  aRefNo: Integer;
  aParam: TQueryParam;
  aEditor: TcxTextEdit;
  aSQL: string;
begin
  aRefNo := 10;
  aParam := FDvnQueryParam;
  aEditor := txtDvnChannel;
  if ( FCasMode = 2 ) then
  begin
    aRefNo := 11;
    aParam := FNagraQueryParam;
    aEditor := txtNagraChannel;
  end;
  {}
  fmMultiSelect := TfmMultiSelect.Create( Application );
  try
    fmMultiSelect.Connection := DBController.DataConnection;
    fmMultiSelect.KeyFields := 'CODENO';
    fmMultiSelect.KeyValues := aParam.Param1;
    fmMultiSelect.DisplayFields := 'CODENO,代碼,DESCRIPTION,名稱';
    fmMultiSelect.ResultFields := 'DESCRIPTION';
    {}
    aSQL :=Format(
      ' SELECT CODENO, DESCRIPTION,         ' +
      '        CHANNELID, CHANCEDAYS        ' +
      '   FROM CD024                        ' +
      '  WHERE STOPFLAG = 0                 ' +
      '    AND PAYFLAG IN ( 0, 1 )          ' +
      '    AND REFNO = %d                   ' +
      ' AND CHANNELID IS NOT NULL           '
      , [aRefNo] );
    //#5241 如果選項不勾,不過濾Chancedays By Kin 2009/08/18
    if  (not chkCHANCEDAYS.Checked) or ( FCasMode <> 2 )  then
      aSQL := aSQL +  ' AND CHANCEDAYS > 0';

    fmMultiSelect.SQL.Text := aSQL;

    {}
    if ( fmMultiSelect.ShowModal = mrOk ) then
    begin
      aParam.Param1 := fmMultiSelect.SelectedValue;
      aEditor.Text := fmMultiSelect.SelectedText;
      LoadChannelDataSet( Nvl( TrimChar( aParam.Param1, [#39] ), 'X' ) );
      SetChannelBufferDays;
    end;
    {}
  finally
    fmMultiSelect.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.btnChannelClearClick(Sender: TObject);
begin
  LoadChannelDataSet( 'X' );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.btnImportClick(Sender: TObject);
begin
  if ( OpenDialog.Execute ) then
  begin
    Screen.Cursor := crSQLWait;
    try
      cmbCAS.Enabled := False;
      if not LoadFromFile( OpenDialog.FileName ) then
      begin
        Delay( 500 );
        ShellExecute( 0, PChar( 'open' ), PChar( FImportLogFileName ), nil,
          PChar( ExtractFilePath( FImportLogFileName ) ), SW_SHOWNORMAL );
      end;
      cmbCAS.Enabled := True;
      ResetNagraCommandData;
    finally
      Screen.Cursor := crDefault;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.ProgressReport(const ACmdType, ARecNo: Integer;
  const ACmdStatus, AErrMsg: String);
begin
  if ( FCasMode = 2 ) then
  begin
    DoNagraProgress( ACmdType, ARecNo, ACmdStatus, AErrMsg );
  end else
  begin
    DoDvnProgress( ACmdType, ARecNo, ACmdStatus, AErrMsg );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.actDvnSetupExecute(Sender: TObject);
begin
  if ( cdsDVN.IsEmpty ) then Exit;
  RunDvnSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.actDvnClearExecute(Sender: TObject);
begin
  UnPrepareDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.actCancelExecute(Sender: TObject);
begin
  Self.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.actNagraSetupExecute(Sender: TObject);
begin
  if ( cdsNagra.IsEmpty ) then Exit;
  if not VdMustInput then Exit;
  RunNagraSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.actNagraClearExecute(Sender: TObject);
var
  aWanttoClear: Boolean;
begin
  aWanttoClear := True;
  if ( FNagraStep > 1 ) then
  begin
    aWanttoClear := ConfirmMsg( Format( '機上盒【%s】作業尚未完成,是否清除?',
      [cmbNagraOpeartion.Text] ) );
  end;
  if ( aWanttoClear ) then
  begin
    FNagraStep := 4;
    SwitchNagraStep;
    UnPrepareDataSet;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.DvnThreadTerminate(Sender: TObject);
begin
  EndDvnSet;
  if ( DvnThreadErrMsg <> EmptyStr ) then
    ErrorMsg( DvnThreadErrMsg )
  else begin
    if ( FCmdHasErr ) then
      ShellExecute( 0, PChar( 'open' ), PChar( FCmdLogFileName ), nil,
        PChar( ExtractFilePath( FCmdLogFileName ) ), SW_SHOWNORMAL );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.NagraThreadTerminate(Sender: TObject);
begin
  EndNagraSet;
  if ( NagraThreadErrMsg <> EmptyStr ) then
    ErrorMsg( NagraThreadErrMsg )
  else begin
    if ( FCmdHasErr ) then
      ShellExecute( 0, PChar( 'open' ), PChar( FCmdLogFileName ), nil,
        PChar( ExtractFilePath( FCmdLogFileName ) ), SW_SHOWNORMAL );
  end;
  SwitchNagraStep;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.gvDVNCustomDrawCell(Sender: TcxCustomGridTableView;
  ACanvas: TcxCanvas; AViewInfo: TcxGridTableDataCellViewInfo;
  var ADone: Boolean);

  { -------------------------------------------- }

  function GetIconIndex(const aValue: String): Integer;
  begin
    Result := -1;
    if ( ( aValue = 'W' ) or ( aValue = 'P' ) ) then
      Result := 4
    else if ( aValue = 'C' ) then
      Result := 5
    else if ( aValue = 'E' ) then
      Result := 6;
  end;

  { -------------------------------------------- }

var
  aItem: TcxGridDBBandedColumn;
  aValue: String;
  aIconIdx, aX, aY: Integer;
begin
  aItem := TcxGridDBBandedColumn( AViewInfo.Item );
  if ( aItem.DataBinding.FieldName = 'CMDA1' ) or
     ( aItem.DataBinding.FieldName = 'CMDB1' ) or
     ( aItem.DataBinding.FieldName = 'CMDB3' ) or
     ( aItem.DataBinding.FieldName = 'CMDB1_1' ) then
  begin
    aValue := VarToStrDef( AViewInfo.GridRecord.Values[aItem.Index], EmptyStr );
    aIconIdx := GetIconIndex( aValue );
    if ( aIconIdx >= 0 ) then
    begin
      ACanvas.FillRect( AViewInfo.ContentBounds );
      aX := AViewInfo.ContentBounds.Left +
        ( ( AViewInfo.ContentWidth - CmdImageList.Width ) div 2 );
      aY := AViewInfo.ContentBounds.Top + 2;
      ACanvas.DrawImage( CmdImageList, aX, aY, aIconIdx );
      ADone := True;
    end else
    if ( aValue = 'S' ) then
    begin
      ACanvas.Brush.Color := clBtnFace;
      ACanvas.FillRect( AViewInfo.ContentBounds );
      ADone := True;
    end;
  end else
  if ( aItem.DataBinding.FieldName = 'ERRMSG' ) then
  begin
    aValue := VarToStrDef( AViewInfo.GridRecord.Values[aItem.Index], EmptyStr );
    if ( aValue <> EmptyStr ) and ( not AViewInfo.GridRecord.Focused ) then
      ACanvas.Font.Color := clRed;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.Panel3Resize(Sender: TObject);
const
  aSpace = 20;
var
  aAllWidth: Integer;
begin
  aAllWidth := ( btnImport.Width + btnSetup.Width + btnClear.Width +
    btnCancel.Width );
  btnImport.Left := ( Panel3.Width - aAllWidth ) div 2;
  btnSetup.Left := ( btnImport.Left + btnImport.Width + aSpace );
  btnClear.Left := ( btnSetup.Left + btnSetup.Width + aSpace );
  btnCancel.Left := ( btnClear.Left + btnClear.Width + aSpace );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.cmbCASPropertiesChange(Sender: TObject);
begin
  SwitchingCAS( cmbCAS.Text );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.cmbOpeartionPropertiesChange(Sender: TObject);
begin
  SwitchNagraOperation( cmbNagraOpeartion.Text );
  actNagraSetup.Caption := Format( 'F6.%s', [GetNagraStepText( FNagraStep ) ] );
  Screen.Cursor := crSQLWait;
  try
    ResetNagraCommandData;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.GetAutoFill;
begin
  qryAutoFill.Close;
  qryAutoFill.SQL.Text:=Format('Select NagraType From CD039 Where CodeNo=%s',
                                [DBController.LoginInfo.CompCode]);
  qryAutoFill.Open;
  if qryAutoFill.FieldByName('NagraType').AsInteger = 1 then
    FAutoFill := True
  else
    FAutoFill := False;
end;

procedure TfmMain.chkCHANCEDAYSClick(Sender: TObject);
  var aParam: TQueryParam;
begin
  aParam := FNagraQueryParam;
  {
  if chkCHANCEDAYS.Checked then
    chkCHANCEDAYS.Caption := '試看緩衝日 , 為試看截止日'
  else
    chkCHANCEDAYS.Caption := '預設天數 + 工作緩衝日, 為試看截止日';
  }
  if  aParam.Param1 <> EmptyStr then
  begin
    LoadChannelDataSet( aParam.Param1);
  end;

end;
end.
