unit cbMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, dxBar, ExtCtrls, dxNavBar, cxPC, cxControls, StdCtrls, IniFiles, 
  dxNavBarCollns, dxNavBarBase, ImgList, cbStyleModule, cbSoModule, cbAppClass,
  cxGraphics, cxContainer, cxEdit, cxTextEdit, cxMaskEdit, cxDropDownEdit,
  Menus, cxLookAndFeelPainters, cxButtons, cxStyles, cxCustomData,
  cxFilter, cxData, cxDataStorage, DB, cxDBData, cxGridCustomView,
  cxGridCustomTableView, cxGridTableView, cxGridBandedTableView,
  cxGridDBBandedTableView, cxClasses, cxGridLevel, cxGrid, DBClient,
  dxBarExtItems, cxGridDBTableView, ComCtrls, cxListView, cxGridStrs,
  cxCheckBox, CodeSiteLogging, Grids, DBGrids, cxGroupBox,
  cxRadioGroup, cxCalendar, cxImageComboBox;

type
  TfmMain = class(TForm)
    dxBarManager: TdxBarManager;
    dxBarDockControl1: TdxBarDockControl;
    StartTimer: TTimer;
    dsInput: TDataSource;
    dxPg1: TdxBarLargeButton;
    Panel1: TPanel;
    dxPg2: TdxBarLargeButton;
    dxCompName: TdxBarStatic;
    dxComputer: TdxBarStatic;
    dxLoginName: TdxBarStatic;
    Panel7: TPanel;
    Bevel4: TBevel;
    dxNavBar: TdxNavBar;
    dxNavBarGroup1: TdxNavBarGroup;
    dxInput: TdxNavBarItem;
    dxList: TdxNavBarItem;
    dxOption: TdxNavBarItem;
    Panel8: TPanel;
    cxMainPage: TcxPageControl;
    cxTabSheet1: TcxTabSheet;
    Panel2: TPanel;
    Bevel3: TBevel;
    Bevel2: TBevel;
    cxInputGrid: TcxGrid;
    gvInput: TcxGridDBTableView;
    gvInputICCNO: TcxGridDBColumn;
    gvInputSTBNO: TcxGridDBColumn;
    gvInputSTBAUTOCB: TcxGridDBColumn;
    gvInputRCSTARTDATE: TcxGridDBColumn;
    gvInputOPERATOR: TcxGridDBColumn;
    glInput: TcxGridLevel;
    Panel6: TScrollBox;
    Label1: TLabel;
    Label2: TLabel;
    txtStbNo: TcxMaskEdit;
    txtIccNo: TcxMaskEdit;
    btnAdd: TcxButton;
    btnRemove: TcxButton;
    btnSave: TcxButton;
    Panel5: TPanel;
    Shape1: TShape;
    lbInputMsg: TLabel;
    cxTabSheet2: TcxTabSheet;
    Panel3: TPanel;
    Panel4: TPanel;
    lbTitle: TLabel;
    cmbSo: TcxComboBox;
    Bevel1: TBevel;
    Panel9: TScrollBox;
    btnSave2: TcxButton;
    Bevel5: TBevel;
    cxGroupBox1: TcxGroupBox;
    rdoCondition1: TcxRadioButton;
    rdoCondition2: TcxRadioButton;
    rdoCondition3: TcxRadioButton;
    rdoCondition4: TcxRadioButton;
    txtIccNo2: TcxMaskEdit;
    rdoCondition5: TcxRadioButton;
    txtRcDateSt: TcxDateEdit;
    txtRcDateEd: TcxDateEdit;
    btnQuery: TcxButton;
    cxListPage: TcxPageControl;
    cxTabSheet3: TcxTabSheet;
    cxTabSheet4: TcxTabSheet;
    glNoCall: TcxGridLevel;
    cxNoCallGrid: TcxGrid;
    Bevel6: TBevel;
    dsNoCall: TDataSource;
    gvNoCall: TcxGridDBBandedTableView;
    gvNoCallICC_NO: TcxGridDBBandedColumn;
    gvNoCallSTB_NO: TcxGridDBBandedColumn;
    gvNoCallRCSTARTDATE: TcxGridDBBandedColumn;
    gvNoCallOPERATOR: TcxGridDBBandedColumn;
    gvNoCallCONFIRM: TcxGridDBBandedColumn;
    gvNoCallERRMSG: TcxGridDBBandedColumn;
    gvNoCallCMD52_1: TcxGridDBBandedColumn;
    gvNoCallCMD8: TcxGridDBBandedColumn;
    gvNoCallCMD97_96: TcxGridDBBandedColumn;
    gvNoCallCMD7: TcxGridDBBandedColumn;
    gvNoCallCMD48: TcxGridDBBandedColumn;
    gvNoCallCMD52_2: TcxGridDBBandedColumn;
    Bevel7: TBevel;
    cxCallGrid: TcxGrid;
    gvCall: TcxGridDBBandedTableView;
    gvCallICC_NO: TcxGridDBBandedColumn;
    gvCallSTBNO: TcxGridDBBandedColumn;
    gvCallRCSTARTDATE: TcxGridDBBandedColumn;
    gvCallOPERATOR: TcxGridDBBandedColumn;
    gvCallCONFIRM: TcxGridDBBandedColumn;
    gvCallERRMSG: TcxGridDBBandedColumn;
    gvCallCMD52_1: TcxGridDBBandedColumn;
    gvCallCMD8: TcxGridDBBandedColumn;
    gvCallCMD60: TcxGridDBBandedColumn;
    gvCallCMD7: TcxGridDBBandedColumn;
    gvCallCMD48: TcxGridDBBandedColumn;
    gvCallCMD52_2: TcxGridDBBandedColumn;
    glCall: TcxGridLevel;
    gvCallCMD62: TcxGridDBBandedColumn;
    gvCallCMD99: TcxGridDBBandedColumn;
    dsCall: TDataSource;
    gvNoCallCMDFLAG: TcxGridDBBandedColumn;
    gvCallCMDFLAG: TcxGridDBBandedColumn;
    gvNoCallRECYCLETEXT: TcxGridDBBandedColumn;
    gvCallRECYCLETEXT: TcxGridDBBandedColumn;
    chkCallback: TcxCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnAddClick(Sender: TObject);
    procedure cxMainPageChange(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure txtStbNoPropertiesValidate(Sender: TObject;
      var DisplayValue: Variant; var ErrorText: TCaption;
      var Error: Boolean);
    procedure txtIccNoPropertiesValidate(Sender: TObject;
      var DisplayValue: Variant; var ErrorText: TCaption;
      var Error: Boolean);
    procedure dxNavBarLinkClick(Sender: TObject; ALink: TdxNavBarItemLink);
    procedure StartTimerTimer(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure cmbSoPropertiesChange(Sender: TObject);
    procedure btnRemoveClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure rdoCondition1Click(Sender: TObject);
    procedure btnQueryClick(Sender: TObject);
    procedure gvNoCallCustomDrawCell(Sender: TcxCustomGridTableView;
      ACanvas: TcxCanvas; AViewInfo: TcxGridTableDataCellViewInfo;
      var ADone: Boolean);
    procedure gvNoCallCellClick(Sender: TcxCustomGridTableView;
      ACellViewInfo: TcxGridTableDataCellViewInfo; AButton: TMouseButton;
      AShift: TShiftState; var AHandled: Boolean);
    procedure btnSave2Click(Sender: TObject);
  private
    { Private declarations }
    FAppIni: TIniFile;
    FConfigFileName: String;
    FOldSoItemIndex: Integer;
    procedure InitMainPage;
    procedure InitDevexComponentLang;
    procedure InitDevexComponentColor;
    procedure LoadAppIni;
    procedure SaveAppIni;
    procedure InitObj(aObj: TCARecycle);
    procedure BuildSoChangeItem;
  public
    { Public declarations }
    property ConfigFileName: String read FConfigFileName;
  end;

var
  fmMain: TfmMain;

implementation

uses cbUtilis, cbLogin, cbOption;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormCreate(Sender: TObject);
var
  aMsg, aFileName: String;
begin
  aFileName := IncludeTrailingPathDelimiter( ExtractFilePath( ParamStr( 0 ) ) ) +
    'AppInfo.ini';
  FAppIni := TIniFile.Create( aFileName );
  LoadAppIni;
  SoDataModule := TSoDataModule.Create( nil );
  if not SoDataModule.ConfigLoadFromFile( aMsg ) then
  begin
    ErrorMsg( aMsg );
    Application.Terminate;
    Exit;
  end;
  InitMainPage;
  InitDevexComponentLang;
  InitDevexComponentColor;
  dxOption.Visible := False;
  rdoCondition1.Checked := True;
  chkCallback.Visible := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormShow(Sender: TObject);
begin
  dxNavBarLinkClick( dxNavBar, dxNavBar.ActiveGroup.Links[0] );
  StartTimer.Enabled := True;
  {}
  Panel5.Height := 0;
  Panel5.Visible := False;
  Bevel2.Visible := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  {}
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormDestroy(Sender: TObject);
begin
  SaveAppIni;
  FAppIni.Free;
  SoDataModule.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.InitMainPage;
begin
  cxMainPage.HideTabs := True;
  cxMainPage.ActivePageIndex := 0;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.btnAddClick(Sender: TObject);
var
  aObj: TCARecycle;
  aMsg: String;
begin
  if ( Trim( txtStbNo.Text  ) = EmptyStr ) then
  begin
    WarningMsg( '請輸入STB序號。' );
    if ( txtStbNo.CanFocus ) then txtStbNo.SetFocus;
    Exit;
  end;
  {}
  if ( Trim( txtIccNo.Text ) = EmptyStr ) then
  begin
    WarningMsg( '請輸入ICC卡號。' );
    if ( txtIccNo.CanFocus ) then txtIccNo.SetFocus;
    Exit;
  end;
  aObj := TCARecycle.Create;
  try
     Screen.Cursor := crSQLWait;
     try
       {}
       InitObj( aObj );
       {}
       if not SoDataModule.VdInputData( aObj, aMsg ) then
       begin
         WarningMsg( aMsg );
         if Pos( 'STB', aMsg ) > 0 then
         begin
           if txtStbNo.CanFocus then txtStbNo.SetFocus;
         end else
         begin
           if txtIccNo.CanFocus then txtIccNo.SetFocus;
         end;
         Exit;
       end;
       if not SoDataModule.VdDuplicateIcc( aObj, SoDataModule.cdsInput, aMsg ) then
       begin
         WarningMsg( aMsg );
         if Pos( 'STB', aMsg ) > 0 then
         begin
           if txtStbNo.CanFocus then txtStbNo.SetFocus;
         end else
         begin
           if txtIccNo.CanFocus then txtIccNo.SetFocus;
         end;
         Exit;
       end;
       if not SoDataModule.GetOtherInfo( aObj, aMsg ) then
       begin
         WarningMsg( aMsg );
         Exit;
       end;
       {}
       if ( chkCallback.Visible ) and ( chkCallback.Checked ) then
       begin
         if ConfirmMsg( '此張ICC卡號回收作業, 請確認須要回撥?' ) then
           aObj.StbAutoFlag := 1
         else
           Exit;
       end;
       {}
       if aObj.StbAutoFlag = 1 then
       begin
         aObj.ResetRecycleText;
         aObj.ResetTransNum;
       end;
       {}
       if not SoDataModule.ObjectToInsertRecord( SoDataModule.cdsInput, aObj, aMsg ) then
       begin
         WarningMsg( aMsg );
         Exit;
       end;
       {}
       txtIccNo.Clear;
       txtStbNo.Clear;
       chkCallback.Checked := False;
       if ( txtStbNo.CanFocus ) then txtStbNo.SetFocus;
       {}
     finally
       Screen.Cursor := crDefault;
     end;
  finally
    aObj.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.InitDevexComponentLang;
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

procedure TfmMain.InitDevexComponentColor;
begin
  Exit;
  Panel3.Color := TcxOffice11LookAndFeelPainter.DefaultContentColor;
  Panel6.Color := TcxOffice11LookAndFeelPainter.DefaultGroupColor;
  Panel9.Color := TcxOffice11LookAndFeelPainter.DefaultGroupColor;
  Panel8.Color := TcxOffice11LookAndFeelPainter.DefaultGroupColor;
  dxNavBar.DefaultStyles.GroupBackground.BackColor :=
    TcxOffice11LookAndFeelPainter.DefaultGroupColor;
  dxNavBar.DefaultStyles.GroupBackground.BackColor2 :=
    TcxOffice11LookAndFeelPainter.DefaultHeaderBackgroundColor;
  dxNavBar.DefaultStyles.ItemHotTracked.BackColor :=
    TcxOffice11LookAndFeelPainter.DefaultTabColor;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.cxMainPageChange(Sender: TObject);
begin
  if ( cxMainPage.ActivePageIndex = 0 ) then
  begin
    if ( txtStbNo.CanFocusEx ) then txtStbNo.SetFocus;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.txtStbNoPropertiesValidate(Sender: TObject;
  var DisplayValue: Variant; var ErrorText: TCaption; var Error: Boolean);
begin
  if Error then
  begin
    WarningMsg( '輸入的STB序號有誤,請重新輸入。' );
    ErrorText := EmptyStr;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.txtIccNoPropertiesValidate(Sender: TObject;
  var DisplayValue: Variant; var ErrorText: TCaption; var Error: Boolean);
begin
  if Error then
  begin
    WarningMsg( '輸入的ICC卡號有誤,請重新輸入。' );
    ErrorText := EmptyStr;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxNavBarLinkClick(Sender: TObject;
  ALink: TdxNavBarItemLink);
begin
  if ( ALink.Item = dxInput ) then
  begin
    cxMainPage.ActivePageIndex := 0;
    lbTitle.Caption := ALink.Item.Caption;
  end else
  if ( ALink.Item = dxList ) then
  begin
    cxMainPage.ActivePageIndex := 1;
    lbTitle.Caption := ALink.Item.Caption;
  end else
  if ( ALink.Item = dxOption ) then
  begin
    fmOption := TfmOption.Create( nil );
    try
      fmOption.ShowModal;
    finally
      FreeAndNil( fmOption );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.StartTimerTimer(Sender: TObject);
begin
  StartTimer.Enabled := False;
  fmLogin := TfmLogin.Create( nil );
  try
    if ( fmLogin.ShowModal <> mrOk ) then
    begin
      Application.Terminate;
      Exit;
    end;
    dxCompName.Caption := Format( '系統台:%s', [SoDataModule.User.CompName] );
    dxComputer.Caption := Format( '電腦名稱:%s', [SoDataModule.User.ComputerName] );
    dxLoginName.Caption := Format( '作業人員:%s', [SoDataModule.User.UserName] );
  finally
    fmLogin.Free;
  end;
  BuildSoChangeItem;
  dxOption.Visible := SoDataModule.User.Admin;
  SoDataModule.PrepareInputDataSet;
  SoDataModule.PrepareNoCallDataSet;
  SoDataModule.PrepareCallDataSet;
  chkCallback.Visible := SoDataModule.User.IsWareHouse;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.LoadAppIni;
begin
  Self.WindowState := TWindowState( FAppIni.ReadInteger(
    'WindowPositon', 'State', Ord( wsNormal ) ) );
  if ( Self.WindowState = wsNormal ) then
  begin
    Self.Width := FAppIni.ReadInteger( 'WindowPositon', 'Width', Self.Width );
    Self.Height := FAppIni.ReadInteger( 'WindowPositon', 'Height', Self.Height );
    Self.Left := FAppIni.ReadInteger( 'WindowPositon', 'Left', Self.Left );
    Self.Top := FAppIni.ReadInteger( 'WindowPositon', 'Top', Self.Top );
  end;
  FConfigFileName := FAppIni.ReadString( 'ConfigFile', 'ActiveFile', EmptyStr );
  if ( FConfigFileName = EmptyStr ) then
    FConfigFileName := IncludeTrailingPathDelimiter( ExtractFilePath( ParamStr( 0 ) ) ) + 'Config.cfg';
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.SaveAppIni;
begin
  FAppIni.WriteInteger( 'WindowPositon', 'State', Ord( Self.WindowState ) );
  if ( Self.WindowState = wsNormal ) then
  begin
    FAppIni.WriteInteger( 'WindowPositon', 'Width', Self.Width );
    FAppIni.WriteInteger( 'WindowPositon', 'Height', Self.Height );
    FAppIni.WriteInteger( 'WindowPositon', 'Left', Self.Left );
    FAppIni.WriteInteger( 'WindowPositon', 'Top', Self.Top );
  end;
  FAppIni.WriteString( 'ConfigFile', 'ActiveFile', FConfigFileName );
  FAppIni.UpdateFile;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.InitObj(aObj: TCARecycle);
begin
 if ( aObj.CompCode < 0 ) then
   aObj.CompCode := StrToInt( Nvl( SoDataModule.User.CompCode, '-1' ) );
 if ( aObj.IccNo = EmptyStr ) then aObj.IccNo := txtIccNo.Text;
 if ( aObj.StbNo = EmptyStr ) then aObj.StbNo := txtStbNo.Text;
 if ( aObj.RCStattDate <= 0 ) then aObj.RCStattDate := Now;
 if ( aObj.Operator = EmptyStr ) then aObj.Operator := SoDataModule.User.UserName;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.BuildSoChangeItem;
var
  aIndex, aLocateItem: Integer;
  aText: String;
begin
  cmbSo.Properties.OnChange := nil;
  try
    cmbSo.Properties.Items.Clear;
    aLocateItem := -1;
    for aIndex := 0 to SoDataModule.SoList.Count - 1 do
    begin
      if ( TSo( SoDataModule.SoList[aIndex] ).CanSwitch ) then
      begin
        aText := Format( '%s-%s', [
          TSo( SoDataModule.SoList[aIndex] ).CompCode,
          TSo( SoDataModule.SoList[aIndex] ).CompName] );
        cmbSo.Properties.Items.Add( aText );
        if ( TSo( SoDataModule.SoList[aIndex] ).Active ) then
          aLocateItem := cmbSo.Properties.Items.Count - 1;
      end;
    end;
    if ( aLocateItem >= 0 ) then cmbSo.ItemIndex := aLocateItem;
    FOldSoItemIndex := cmbSo.ItemIndex;
  finally
    cmbSo.Properties.OnChange := cmbSoPropertiesChange;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.cmbSoPropertiesChange(Sender: TObject);
var
  aNewCompCode, aOldCompCode, aNewCompName, aMsg: String;

  { ------------------------------------ }

  procedure AbortChange(const aIdx: Integer);
  begin
    cmbSo.Properties.OnChange := nil;
    try
      cmbSo.ItemIndex := aIdx;
    finally
      cmbSo.Properties.OnChange := cmbSoPropertiesChange;
    end;
  end;

  { ------------------------------------ }

begin
  if not ConfirmMsg( Format( '切換至系統台:%s?', [cmbSo.Text] ) ) then
  begin
    AbortChange( FOldSoItemIndex );
    Exit;
  end;
  if ( cmbSo.ItemIndex < 0  ) then
  begin
    WarningMsg( '無系統台可切換。' );
    Exit;
  end;
  Application.ProcessMessages;
  aNewCompCode := Copy( cmbSo.Properties.Items[cmbSo.ItemIndex],
    1, Pos( '-', cmbSo.Properties.Items[cmbSo.ItemIndex] ) - 1 );
  aOldCompCode := Copy( cmbSo.Properties.Items[FOldSoItemIndex],
    1, Pos( '-', cmbSo.Properties.Items[FOldSoItemIndex] ) - 1 );
  aNewCompName := Copy( cmbSo.Properties.Items[cmbSo.ItemIndex],
    Pos( '-', cmbSo.Properties.Items[cmbSo.ItemIndex] ) + 1,
    Length( cmbSo.Properties.Items[cmbSo.ItemIndex] ) );
  {}
  Screen.Cursor := crSQLWait;
  try
    SoDataModule.DbLogout( aOldCompCode );
    {}
    if SoDataModule.DbLogin( aNewCompCode, aMsg ) then
    begin
      FOldSoItemIndex := cmbSo.ItemIndex;
    end else
    begin
      WarningMsg( Format( '無法切換系統台, 訊息:%s。', [aMsg] ) );
      if not SoDataModule.DbLogin( aOldCompCode, aMsg ) then
        WarningMsg( Format( '無法切換回原本系統台, 訊息:%s。', [aMsg] ) );
      AbortChange( FOldSoItemIndex );
    end;
    SoDataModule.UnPrepareInputDataSet;
    SoDataModule.UnPrepareNoCallDataSet;
    SoDataModule.UnPrepareCallDataSet;
    {}
    SoDataModule.PrepareInputDataSet;
    SoDataModule.PrepareNoCallDataSet;
    SoDataModule.PrepareCallDataSet;
    {}
    SoDataModule.User.CompCode := aNewCompCode;
    SoDataModule.User.CompName := aNewCompName;
    SoDataModule.User.IsWareHouse := SoDataModule.IsWareHouse( [aNewCompCode] );
    chkCallback.Visible := SoDataModule.User.IsWareHouse;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.btnRemoveClick(Sender: TObject);
begin
  if ( gvInput.DataController.GetSelectedCount > 0 ) then
  begin
    if not ConfirmMsg( '確認刪除選取的資料?' ) then Exit;
    Screen.Cursor := crSQLWait;
    try
      gvInput.DataController.DeleteSelection;
    finally
      Screen.Cursor := crDefault;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.btnSaveClick(Sender: TObject);
var
  aMsg: String;
begin
  if ( SoDataModule.cdsInput.IsEmpty ) then
  begin
    WarningMsg( '尚未輸入任何資料, 請先填入回收作業的ICC卡號及STB序號。' );
    if ( txtStbNo.CanFocusEx ) then txtStbNo.SetFocus;
    Exit;
  end;
  Screen.Cursor := crSQLWait;
  try
    if not SoDataModule.AddToRecycleTable( SoDataModule.cdsInput, aMsg ) then
    begin
      WarningMsg( aMsg );
      Exit;
    end;
    InfoMsg( '回收作業待處理資料存檔完成。' );
    SoDataModule.UnPrepareInputDataSet;
    SoDataModule.PrepareInputDataSet;
    if ( txtStbNo.CanFocusEx ) then txtStbNo.SetFocus;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.rdoCondition1Click(Sender: TObject);
begin
  txtIccNo2.Enabled := ( Sender = rdoCondition4 );
  if not txtIccNo2.Enabled then txtIccNo2.Clear;
  {}
  txtRcDateSt.Enabled := ( Sender = rdoCondition5 );
  txtRcDateEd.Enabled := ( Sender = rdoCondition5 );
  if ( not txtRcDateSt.Enabled ) or ( not txtRcDateEd.Enabled ) then
  begin
    txtRcDateSt.Text := EmptyStr;
    txtRcDateEd.Text := EmptyStr;
  end;
  {}
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.btnQueryClick(Sender: TObject);
var
  aCondition: Integer;
  aMsg, aText1, aText2: String;
begin
  if ( rdoCondition4.Checked ) and ( txtIccNo2.Text = EmptyStr ) then
  begin
    WarningMsg( '請輸入ICC卡號。' );
    if ( txtIccNo2.CanFocusEx ) then txtIccNo2.SetFocus;
    Exit;
  end;
  if ( rdoCondition5.Checked ) then
  begin
    if ( txtRcDateSt.Text = EmptyStr ) and ( txtRcDateEd.Text = EmptyStr ) then
    begin
      WarningMsg( '請輸入處理日期。' );
      if ( txtRcDateSt.CanFocusEx ) then
        txtRcDateSt.SetFocus
      else if ( txtRcDateEd.CanFocusEx ) then
        txtRcDateEd.SetFocus;
      Exit;
    end;
  end;
  aCondition := 1;
  aText1 := EmptyStr;
  aText2 := EmptyStr;
  if rdoCondition1.Checked then
    aCondition := 1
  else if rdoCondition2.Checked then
    aCondition := 2
  else if rdoCondition3.Checked then
    aCondition := 3
  else if rdoCondition4.Checked then
  begin
    aCondition := 4;
    aText1 := txtIccNo2.Text;
  end else
  if rdoCondition5.Checked then
  begin
    aCondition := 5;
    if ( txtRcDateSt.Text <> EmptyStr ) then
      aText1 := FormatDateTime( 'yyyymmdd', txtRcDateSt.Date );
    if ( txtRcDateEd.Text <> EmptyStr ) then
      aText2 := FormatDateTime( 'yyyymmdd', txtRcDateEd.Date );
  end;  
  {}
  Screen.Cursor := crSQLWait;
  try
    {}
    SoDataModule.UnPrepareNoCallDataSet;
    SoDataModule.UnPrepareCallDataSet;
    {}
    SoDataModule.PrepareNoCallDataSet;
    SoDataModule.PrepareCallDataSet;
    {}
    SoDataModule.cdsNoCall.DisableControls;
    SoDataModule.cdsCall.DisableControls;
    try
      if not SoDataModule.QueryProcessRecord( aCondition, aMsg, aText1, aText2 ) then
      begin
        WarningMsg( aMsg );
        Exit;
      end;
      if ( SoDataModule.cdsNoCall.IsEmpty ) and
         ( SoDataModule.cdsCall.IsEmpty ) then
      begin
        WarningMsg( '本次查詢無符合資料' );
        Exit;
      end;
      if not SoDataModule.cdsNoCall.IsEmpty then
        cxListPage.ActivePageIndex := 0
      else
        cxListPage.ActivePageIndex := 1;
    finally
      SoDataModule.cdsNoCall.EnableControls;
      SoDataModule.cdsCall.EnableControls;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.gvNoCallCustomDrawCell(Sender: TcxCustomGridTableView;
  ACanvas: TcxCanvas; AViewInfo: TcxGridTableDataCellViewInfo;
  var ADone: Boolean);

  { ------------------------------------------ }

  function GetDrawImageIndex(const aText: String): Integer;
  begin
    if ( Pos( '2', aText ) > 0 ) then //失敗
      Result := 2
    else if ( Pos( '0', aText ) > 0 ) then //尚未處理
      Result :=  0
    else if ( Pos( '3', aText ) > 0 ) then //處理中
      Result := 5
    else
      Result := 1;  
  end;

  { ------------------------------------------ }

var
  aValue, aCmdFlag: String;
  aX, aY, aChkIdx: Integer;
  aCmdFlagItem: TcxGridDBBandedColumn;
  aColor: TColor;
begin
  if ( AViewInfo.Item.Tag = 1 ) then
  begin
    aValue := VarToStrDef( AViewInfo.GridRecord.Values[AViewInfo.Item.Index], '0' );
    ACanvas.FillRect( AViewInfo.Bounds );
    aX := ( ( AViewInfo.Bounds.Right - AViewInfo.Bounds.Left ) -
      StyleModule.MsgImageList.Width ) div 2;
    aY := ( ( AViewInfo.Bounds.Bottom - AViewInfo.Bounds.Top ) -
      StyleModule.MsgImageList.Height ) div 2;
    ACanvas.DrawImage( StyleModule.MsgImageList, AViewInfo.Bounds.Left + aX ,
      AViewInfo.Bounds.Top + aY, GetDrawImageIndex( aValue ) );
    ADone := True;
  end else
  if ( AViewInfo.Item.Tag = 2 ) then
  begin
    aCmdFlagItem := gvCallCMDFLAG;
    if Sender = gvNoCall then aCmdFlagItem := gvNoCallCMDFLAG;
    aCmdFlag := VarToStrDef( AViewInfo.GridRecord.Values[aCmdFlagItem.Index], '0' );
    aValue := VarToStrDef( AViewInfo.GridRecord.Values[AViewInfo.Item.Index], 'N' );
    { 成功/失敗 }
    aColor := clBtnFace;
    if ( aCmdFlag = '2' ) or ( aCmdFlag = '4' ) then
      aColor := clWindow;
    ACanvas.Brush.Color := aColor;
    ACanvas.FillRect( AViewInfo.Bounds );
    if ( aValue = 'Y' ) then
    begin
      aX := ( ( AViewInfo.Bounds.Right - AViewInfo.Bounds.Left ) -
        StyleModule.MsgImageList.Width ) div 2;
      aY := ( ( AViewInfo.Bounds.Bottom - AViewInfo.Bounds.Top ) -
        StyleModule.MsgImageList.Height ) div 2;
      { 成功的 Icon }
      if ( aCmdFlag = '2' ) or ( aCmdFlag = '3' ) then
        aChkIdx := 0
      else
        aChkIdx := 1;
      ACanvas.DrawImage( StyleModule.CheckImageList, AViewInfo.Bounds.Left + aX ,
        AViewInfo.Bounds.Top + aY, aChkIdx );
    end;
    ADone := True;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.gvNoCallCellClick(Sender: TcxCustomGridTableView;
  ACellViewInfo: TcxGridTableDataCellViewInfo; AButton: TMouseButton;
  AShift: TShiftState; var AHandled: Boolean);
var
  aValue, aCmdFlag: String;
  aCmdFlagItem: TcxGridDBBandedColumn;
begin
  if ( AButton <> mbLeft ) then Exit;
  if ( ACellViewInfo.Item.Tag = 2 ) then
  begin
    aCmdFlagItem := gvCallCMDFLAG;
    if Sender = gvNoCall then aCmdFlagItem := gvNoCallCMDFLAG;
    aCmdFlag := VarToStrDef( ACellViewInfo.GridRecord.Values[aCmdFlagItem.Index], '0' );
    if ( aCmdFlag <> '2' ) and ( aCmdFlag <> '4' ) then Exit;
    aValue := IIF( ( VarToStrDef( ACellViewInfo.Value, 'N' ) = 'Y' ), 'N', 'Y' );
    TcxGridDBBandedColumn( ACellViewInfo.Item ).DataBinding.DataController.Edit;
    TcxGridDBBandedColumn( ACellViewInfo.Item ).DataBinding.DataController.DataSet.FieldByName(
      TcxGridDBBandedColumn( ACellViewInfo.Item ).DataBinding.FieldName ).AsString := aValue;
    TcxGridDBBandedColumn( ACellViewInfo.Item ).DataBinding.DataController.Post;  
    AHandled := True;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.btnSave2Click(Sender: TObject);
var
  aMsg: String;
  aCount1, aCount2: Integer;
begin
  if ( SoDataModule.cdsNoCall.IsEmpty ) and
     ( SoDataModule.cdsCall.IsEmpty ) then
  begin
    WarningMsg( '無任何資料可存檔, 請先使用查詢功能。' );
    Exit;
  end;
  Screen.Cursor := crSQLWait;
  try
    { 不須回撥的資料 }
    if not SoDataModule.ConfirmData( SoDataModule.cdsNoCall, aCount1, aMsg ) then
    begin
      cxListPage.ActivePageIndex := 0;
      Self.ActiveControl := cxNoCallGrid;
      WarningMsg( aMsg );
      Exit;
    end;
    { 須要回撥的資料 }
    if not SoDataModule.ConfirmData( SoDataModule.cdsCall, aCount2, aMsg ) then
    begin
      cxListPage.ActivePageIndex := 1;
      Self.ActiveControl := cxCallGrid;
      WarningMsg( aMsg );
      Exit;
    end;
    if ( aCount1 + aCount2 ) > 0 then
      InfoMsg( '回收作業確認/取消資料存檔完成。' )
    else
      InfoMsg( '無確認/取消資料可存檔。' )
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
