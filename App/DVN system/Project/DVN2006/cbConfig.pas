unit cbConfig;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, ImgList, DB, ADODB, 
  {$IFDEF APPDEBUG} CodeSiteLogging, {$ENDIF}
  { Developer Express Suite }
  cxLookAndFeelPainters, cxPC, cxButtons, cxControls, cxContainer, cxEdit,
  cxTextEdit, cxMaskEdit, cxCheckBox, cxSpinEdit, cxTimeEdit, cxRadioGroup,
  cxGraphics, cxCustomData, cxStyles, cxTL, cxImageComboBox,
  cxInplaceContainer, cxDropDownEdit, Menus;



type
  TfmConfig = class(TForm)
    ConfigPage: TcxPageControl;
    Panel1: TPanel;
    Bevel1: TBevel;
    btnConfirm: TcxButton;
    btnCancel: TcxButton;
    cxCAS: TcxTabSheet;
    lblCASIP: TLabel;
    txtCASIP: TcxMaskEdit;
    lblSendPort: TLabel;
    txtSendPort: TcxMaskEdit;
    lblRecvPort: TLabel;
    txtRecvPort: TcxMaskEdit;
    chkEnableSend: TcxCheckBox;
    chkEnableRecv: TcxCheckBox;
    cxCommon: TcxTabSheet;
    lbDbRetryFreq: TLabel;
    txtDbRetryFreq: TcxSpinEdit;
    lblBusyTimeStart: TLabel;
    txtBusyTimeStart: TcxTimeEdit;
    lblBusyTimeEnd: TLabel;
    txtBusyTimeEnd: TcxTimeEdit;
    lblBusyTimeReadFrquence: TLabel;
    txtBusyTimeReadFreq: TcxSpinEdit;
    lblNormalTime: TLabel;
    txtNorTimeFreq: TcxSpinEdit;
    lblDbProcRecords: TLabel;
    txtDbProcRecords: TcxSpinEdit;
    lblBusyTimeReadFreq: TLabel;
    lblDbWriteErrorWhenSocketFail: TLabel;
    Panel2: TPanel;
    chkWriteNo: TcxRadioButton;
    chkWriteYes: TcxRadioButton;
    lblProcessIPPV: TLabel;
    Panel3: TPanel;
    chkProcessIPPV_A: TcxRadioButton;
    chkProcessIPPV_N: TcxRadioButton;
    chkProcessIPPV_O: TcxRadioButton;
    lblProcessBatch: TLabel;
    Panel4: TPanel;
    ChkProcessBatch_A: TcxRadioButton;
    ChkProcessBatch_N: TcxRadioButton;
    ChkProcessBatch_O: TcxRadioButton;
    lblCARertyFeq: TLabel;
    txtCARertyFeq: TcxSpinEdit;
    lblCasCommCheck: TLabel;
    txtCACheckFreq: TcxSpinEdit;
    HeaderPanel: TPanel;
    HeaderImage: TImage;
    cxSoDb: TcxTabSheet;
    ConfigSoTree: TcxTreeList;
    ConfigSoTreeCol1: TcxTreeListColumn;
    ConfigSoTreeCol2: TcxTreeListColumn;
    ConfigSoTreeCol3: TcxTreeListColumn;
    ConfigSoTreeCol4: TcxTreeListColumn;
    ConfigSoTreeCol5: TcxTreeListColumn;
    ConfigSoTreeCol6: TcxTreeListColumn;
    ConfigSoTreeCol7: TcxTreeListColumn;
    ConfigSoTreeCol9: TcxTreeListColumn;
    btnDeleteSo: TcxButton;
    btnAddSo: TcxButton;
    cxHighCmd: TcxTabSheet;
    HighCmdTree: TcxTreeList;
    btnTest: TcxButton;
    lblTest: TLabel;
    HighCmdTreeCol1: TcxTreeListColumn;
    HighCmdTreeCol2: TcxTreeListColumn;
    HighCmdTreeCol3: TcxTreeListColumn;
    HighCmdTreeCol5: TcxTreeListColumn;
    HighCmdTreeCol6: TcxTreeListColumn;
    HighCmdTreeCol7: TcxTreeListColumn;
    HighCmdTreeCol4: TcxTreeListColumn;
    btnAddHighCmd: TcxButton;
    btnDeleteHighCmd: TcxButton;
    cxLowCmd: TcxTabSheet;
    btnAddLowCmd: TcxButton;
    btnDeleteLowCmd: TcxButton;
    LowCmdTree: TcxTreeList;
    LowCmdTreeCol1: TcxTreeListColumn;
    LowCmdTreeCol2: TcxTreeListColumn;
    cxCmdError: TcxTabSheet;
    btnAddError: TcxButton;
    btnDeleteError: TcxButton;
    ErrorTree: TcxTreeList;
    ErrorTreeCol1: TcxTreeListColumn;
    ErrorTreeCol2: TcxTreeListColumn;
    ErrorTreeCol3: TcxTreeListColumn;
    lblCAProductDefine: TLabel;
    txtCAProductDefine: TcxMaskEdit;
    ConfigSoTreeCol10: TcxTreeListColumn;
    ConfigSoTreeCol8: TcxTreeListColumn;
    lblCAProtocol: TLabel;
    txtCAProtocol: TcxMaskEdit;
    lblCASendDelay: TLabel;
    txtCASendDelay: TcxSpinEdit;
    lblCAMaxError: TLabel;
    txtCAMaxError: TcxSpinEdit;
    lblIdleTime: TLabel;
    txtCAIdleTime: TcxSpinEdit;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure ConfigSoTreeFocusedNodeChanged(Sender: TObject; APrevFocusedNode,
      AFocusedNode: TcxTreeListNode);
    procedure ConfigSoTreeEdited(Sender: TObject; AColumn: TcxTreeListColumn);
    procedure btnAddSoClick(Sender: TObject);
    procedure btnDeleteSoClick(Sender: TObject);
    procedure btnTestClick(Sender: TObject);
    procedure btnAddHighCmdClick(Sender: TObject);
    procedure btnDeleteHighCmdClick(Sender: TObject);
    procedure btnAddLowCmdClick(Sender: TObject);
    procedure btnDeleteLowCmdClick(Sender: TObject);
    procedure btnAddErrorClick(Sender: TObject);
    procedure btnDeleteErrorClick(Sender: TObject);
    procedure ConfigPageChange(Sender: TObject);
  private
    { Private declarations }
    procedure GetCommon;
    procedure GetHighCmdEnv;
    procedure GetLowCmdEnv;
    procedure GetCASEnv;
    procedure GetCmdError;
    procedure GetSoList;
    procedure EnableConfigControls;
    {$HINTS OFF}
       procedure EnableConfigControl(aControl: TControl);
       procedure DisableConfigControl(aControl: TControl);
    {$HINTS ON}
    procedure DisableConfigControls;
    procedure InitControlLanguage;
  public
    { Public declarations }
  end;

var
  fmConfig: TfmConfig;

implementation

uses cbMain, cbAppClass, cbStyleModule, cbConfigModule;

{$R *.dfm}

var aGlobeKeyValue: String;


{ ---------------------------------------------------------------------------- }

procedure TfmConfig.FormCreate(Sender: TObject);
begin
  fmConfig := Self;
  {}
  lblCasCommCheck.Enabled := False;
  txtCACheckFreq.Enabled := False;
  {}
  cxHighCmd.TabVisible := False;
  cxLowCmd.TabVisible := False;
  cxCmdError.TabVisible := False;
  {}
  GetCASEnv;
  GetHighCmdEnv;
  GetLowCmdEnv;
  GetCmdError;
  GetCommon;
  GetSoList;
  if ConfigPage.ActivePageIndex <> 0 then
    ConfigPage.ActivePageIndex := 0;
  ConfigPageChange( ConfigPage );
  lblTest.Caption := EmptyStr;
  DisableConfigControls;
  {}
  InitControlLanguage;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.FormDestroy(Sender: TObject);
begin
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.GetCASEnv;
begin
  txtCASIP.Text := ConfigModule.CASEnv.IP;
  txtSendPort.Text := ConfigModule.CASEnv.SendPort;
  txtRecvPort.Text := ConfigModule.CASEnv.RecvPort;
  txtCAProtocol.Text := ConfigModule.CasEnv.Protocol;
  txtCAProductDefine.Text := ConfigModule.CommEnv.CAProductDefine;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.GetCmdError;
var
  aNode: TcxTreeListNode;
begin
  ErrorTree.Clear;
  ConfigModule.CmdErrorEnv.First;
  while not ConfigModule.CmdErrorEnv.Eof do
  begin
    aNode := ErrorTree.Add;
    aNode.AssignValues( [
      ConfigModule.CmdErrorEnv.FieldByName( 'ERRORFLAG' ).AsString,
      ConfigModule.CmdErrorEnv.FieldByName( 'ERRORCODE' ).AsString,
      ConfigModule.CmdErrorEnv.FieldByName( 'ERRORDESC' ).AsString ] );
    aNode.ImageIndex := 6;
    aNode.SelectedIndex := aNode.ImageIndex;
    ConfigModule.CmdErrorEnv.Next;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.GetCommon;
begin
  txtDbRetryFreq.Value := ConfigModule.CommEnv.DbRetryFreq;
  txtBusyTimeStart.Text := ConfigModule.CommEnv.BusyTimeStart;
  txtBusyTimeEnd.Text := ConfigModule.CommEnv.BusyTimeEnd;
  txtBusyTimeReadFreq.Value := ConfigModule.CommEnv.BusyTimeReadFreq;
  txtNorTimeFreq.Value := ConfigModule.CommEnv.NorTimeReadFreq;
  txtDbProcRecords.Value := ConfigModule.CommEnv.DbProcRecords;
  chkWriteYes.Checked := ConfigModule.CommEnv.DbWriteError;
  if ConfigModule.CommEnv.DbProcIPPV = 'A' then
    chkProcessIPPV_A.Checked := True
  else if ConfigModule.CommEnv.DbProcIPPV = 'O' then
    chkProcessIPPV_O.Checked := True
  else
    chkProcessIPPV_N.Checked := True;
  if ConfigModule.CommEnv.DbProcBatch = 'A' then
    ChkProcessBatch_A.Checked := True
  else if ConfigModule.CommEnv.DbProcBatch = 'O' then
    ChkProcessBatch_O.Checked := True
  else
    ChkProcessBatch_N.Checked := True;
  {}
  txtCARertyFeq.Text := IntToStr( ConfigModule.CommEnv.CARetryFreq );
  txtCAMaxError.Value := ConfigModule.CommEnv.CAMaxError;
  txtCACheckFreq.Value := ConfigModule.CommEnv.CACheckFreq;
  txtCAIdleTime.Value := ConfigModule.CommEnv.CAIdleTime;
  {}
  chkEnableSend.Checked := ConfigModule.CommEnv.CAEnableSend;
  chkEnableRecv.Checked := ConfigModule.CommEnv.CAEnableRecv;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.GetHighCmdEnv;
var
  aNode: TcxTreeListNode;
begin
  HighCmdTree.Clear;
  ConfigModule.HighCmdEnv.First;
  while not ConfigModule.HighCmdEnv.Eof do
  begin
    aNode := HighCmdTree.Add;
    aNode.AssignValues( [
      ConfigModule.HighCmdEnv.FieldByName( 'HIGHLEVELCMD' ).AsString,
      ConfigModule.HighCmdEnv.FieldByName( 'DESCRIPTION' ).AsString,
      ConfigModule.HighCmdEnv.FieldByName( 'CMDTYPE' ).AsString,
      ConfigModule.HighCmdEnv.FieldByName( 'LOWLEVELCMD' ).AsString,
      ConfigModule.HighCmdEnv.FieldByName( 'CMDBYNOTE' ).AsString,
      ConfigModule.HighCmdEnv.FieldByName( 'CMDTYPEPRIORITY' ).AsString,
      ConfigModule.HighCmdEnv.FieldByName( 'CMDPRIORITY' ).AsString] );
    aNode.ImageIndex := 8;
    aNode.SelectedIndex := aNode.ImageIndex;
    ConfigModule.HighCmdEnv.Next;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.GetLowCmdEnv;
var
  aNode: TcxTreeListNode;
begin
  LowCmdTree.Clear;
  ConfigModule.LowCmdEnv.First;
  while not ConfigModule.LowCmdEnv.Eof do
  begin
    aNode := LowCmdTree.Add;
    aNode.AssignValues( [
      ConfigModule.LowCmdEnv.FieldByName( 'LOWLEVELCMD' ).AsString,
      ConfigModule.LowCmdEnv.FieldByName( 'DESCRIPTION' ).AsString] );
    aNode.ImageIndex := 1;
    aNode.SelectedIndex := aNode.ImageIndex;
    ConfigModule.LowCmdEnv.Next;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.EnableConfigControls;
var
  aIndex: Integer;
begin
  cxHighCmd.TabVisible := True;
  cxLowCmd.TabVisible := True;
  cxCmdError.TabVisible := True;
  {}
  txtCASIP.Enabled := True;
  txtSendPort.Enabled := True;
  txtRecvPort.Enabled := True;
  txtCAProtocol.Enabled := True;
  txtCAProductDefine.Enabled := True;
  {}
  ConfigSoTreeCol5.Visible := True;
  ConfigSoTreeCol6.Visible := True;
  ConfigSoTreeCol7.Visible := True;
  {}
  for aIndex := 0 to ConfigSoTree.ColumnCount - 1 do
  begin
    if not ConfigSoTree.Columns[aIndex].Visible then Continue;
    ConfigSoTree.Columns[aIndex].Options.Focusing := True;
  end;
  btnAddSo.Enabled := True;
  btnDeleteSo.Enabled := True;
  {}
  HighCmdTree.OptionsSelection.CellSelect := True;
  btnAddHighCmd.Enabled := True;
  btnDeleteHighCmd.Enabled := True;
  {}
  LowCmdTree.OptionsSelection.CellSelect := True;
  btnAddLowCmd.Enabled := True;
  btnDeleteLowCmd.Enabled := True;
  {}
  ErrorTree.OptionsSelection.CellSelect := True;
  btnAddError.Enabled := True;
  btnDeleteError.Enabled := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.EnableConfigControl(aControl: TControl);
var
  aIndex: Integer;
begin
  if aControl = ConfigSoTree then
  begin
    ConfigSoTree.Enabled := True;
    for aIndex := 0 to ConfigSoTree.ColumnCount - 1 do
    begin
      if not ConfigSoTree.Columns[aIndex].Visible then Continue;
      ConfigSoTree.Columns[aIndex].Options.Focusing := True;
    end;
    btnAddSo.Enabled := True;
    btnDeleteSo.Enabled := True;
  end else
  if aControl = HighCmdTree then
  begin
    HighCmdTree.OptionsSelection.CellSelect := True;
    btnAddHighCmd.Enabled := True;
    btnDeleteHighCmd.Enabled := True;
  end else
  if aControl = LowCmdTree then
  begin
    LowCmdTree.OptionsSelection.CellSelect := True;
    btnAddLowCmd.Enabled := True;
    btnDeleteLowCmd.Enabled := True;
  end else
  if aControl = ErrorTree then
  begin
    ErrorTree.OptionsSelection.CellSelect := True;
    btnAddError.Enabled := True;
    btnDeleteError.Enabled := True;
  end else
  begin
    aControl.Enabled := True;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.DisableConfigControls;
var
  aIndex: Integer;
begin
  for aIndex := 0 to ConfigSoTree.ColumnCount - 1 do
  begin
    if not ConfigSoTree.Columns[aIndex].Visible then Continue;
    if ( aIndex > 0 ) and ( aIndex <> 9 ) then
      ConfigSoTree.Columns[aIndex].Options.Focusing := False;
  end;
  btnAddSo.Enabled := False;
  btnDeleteSo.Enabled := False;
  {}
  HighCmdTree.OptionsSelection.CellSelect := False;
  btnAddHighCmd.Enabled := False;
  btnDeleteHighCmd.Enabled := False;
  {}
  LowCmdTree.OptionsSelection.CellSelect := False;
  btnAddLowCmd.Enabled := False;
  btnDeleteLowCmd.Enabled := False;
  {}
  ErrorTree.OptionsSelection.CellSelect := False;
  btnAddError.Enabled := False;
  btnDeleteError.Enabled := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.DisableConfigControl(aControl: TControl);
var
  aIndex: Integer;
begin
  if aControl = ConfigSoTree then
  begin
    for aIndex := 0 to ConfigSoTree.ColumnCount - 1 do
    begin
      if not ConfigSoTree.Columns[aIndex].Visible then Continue;
      if ( aIndex > 0 ) and ( aIndex <> 9 ) then
        ConfigSoTree.Columns[aIndex].Options.Focusing := False;
    end;
    btnAddSo.Enabled := False;
    btnDeleteSo.Enabled := False;
  end else
  if aControl = HighCmdTree then
  begin
    HighCmdTree.OptionsSelection.CellSelect := False;
    btnAddHighCmd.Enabled := False;
    btnDeleteHighCmd.Enabled := False;
  end else
  if aControl = LowCmdTree then
  begin
    LowCmdTree.OptionsSelection.CellSelect := False;
    btnAddLowCmd.Enabled := False;
    btnDeleteLowCmd.Enabled := False;
  end else
  if aControl = ErrorTree then
  begin
    ErrorTree.OptionsSelection.CellSelect := False;
    btnAddError.Enabled := False;
    btnDeleteError.Enabled := False;
  end else
  begin
    aControl.Enabled := False;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if not ( ssCtrl in Shift ) then aGlobeKeyValue := EmptyStr;
  if not ( Char( Key ) in ['a'..'z', 'A'..'Z'] ) then
    aGlobeKeyValue := EmptyStr
  else
    aGlobeKeyValue := ( aGlobeKeyValue + Char( Key ) );
  if UpperCase( aGlobeKeyValue ) = 'CABLESOFT' then
  begin
    EnableConfigControls;
    aGlobeKeyValue := EmptyStr;
  end;
end;

{ ---------------------------------------------------------------------------- }

function GetPosNode(const aPos, aPosType, aPosName: String; const aCanAdd: Boolean = True):
  TcxTreeListNode;
begin
  Result := nil;
  if ( aPosName = EmptyStr ) then Exit;
  Result := fmConfig.ConfigSoTree.FindNodeByText( aPosName, fmConfig.ConfigSoTreeCol3 );
  if not Assigned( Result ) and aCanAdd then
  begin
    Result := fmConfig.ConfigSoTree.Add( nil );
    Result.AssignValues( [EmptyStr, EmptyStr, aPosName, EmptyStr,
      EmptyStr, EmptyStr, EmptyStr, aPosType, aPos] );
    Result.ImageIndex := 0;
    Result.SelectedIndex := 0;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.GetSoList;

  { -------------------------------------------------- }

  function GetParnetNode(const aPosName: String):
    TcxTreeListNode;
  begin
    Result := nil;
    if ( aPosName = EmptyStr ) then Exit;
    Result := ConfigSoTree.FindNodeByText( aPosName, ConfigSoTreeCol3 );
    if not Assigned( Result ) then
    begin
      Result := ConfigSoTree.Add( nil );
      Result.AssignValues( [EmptyStr, EmptyStr, aPosName] );
      Result.ImageIndex := 0;
      Result.SelectedIndex := 0;
    end;
  end;

  { -------------------------------------------------- }

var
  aIndex: Integer;
  aSo: TSo;
  aParnetNode, aNode: TcxTreeListNode;
begin
  for aIndex := 0 to ConfigModule.SoList.Count - 1 do
  begin
    aSo := ConfigModule.SoList[aIndex];
    { 取平台節點, 此系統台要掛在那一個平台下 }
    aParnetNode := GetParnetNode( aSo.PosName );
    aNode := ConfigSoTree.AddChild( aParnetNode );
    aNode.AssignValues( [aSo.Selected, aSo.CompCode, aSo.CompName,
      aSo.NetworkId, aSo.DbUserId, aSo.DbUserPass, aSo.DbAliase, aSo.SoType,
      aSo.PosName, aSo.AreaCode] );
    aNode.ImageIndex := StyleModule.GetSoImageIndex( aSo );
    aNode.SelectedIndex := aNode.ImageIndex;
    aParnetNode.Expanded := True;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.ConfigSoTreeFocusedNodeChanged(Sender: TObject;
  APrevFocusedNode, AFocusedNode: TcxTreeListNode);
begin
  if Assigned( AFocusedNode ) then
    ConfigSoTree.OptionsSelection.CellSelect := ( AFocusedNode.Level > 0 );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.ConfigSoTreeEdited(Sender: TObject;
  AColumn: TcxTreeListColumn);
begin
 if AColumn = ConfigSoTreeCol1 then
 begin
   if VarAsType( AColumn.Value, varBoolean ) then
     ConfigSoTree.FocusedNode.ImageIndex := 2
   else
     ConfigSoTree.FocusedNode.ImageIndex := 1;
   ConfigSoTree.FocusedNode.SelectedIndex := ConfigSoTree.FocusedNode.ImageIndex;
 end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.btnAddSoClick(Sender: TObject);
var
  aParentNode, aNode: TcxTreeListNode;
begin
  if not Assigned( ConfigSoTree.FocusedNode ) then Exit;
  aParentNode := nil;
  if ConfigSoTree.FocusedNode.Level = 0 then
    aParentNode := ConfigSoTree.FocusedNode
  else begin
    if Assigned( ConfigSoTree.FocusedNode.Parent ) then
      aParentNode := ConfigSoTree.FocusedNode.Parent;
  end;
  if Assigned( aParentNode ) then
  begin
    aNode := aParentNode.AddChild;
    aNode.ImageIndex := 1;
    aNode.SelectedIndex := aNode.ImageIndex;
    aNode.AssignValues( [False, EmptyStr, EmptyStr, EmptyStr, EmptyStr,
      EmptyStr, EmptyStr, aParentNode.Values[ConfigSoTreeCol8.ItemIndex],
      aParentNode.Values[ConfigSoTreeCol9.ItemIndex] ] );
    aNode.Focused := True;
    ConfigSoTree.SetFocus;
    ConfigSoTreeCol2.Editing := True;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.btnDeleteSoClick(Sender: TObject);
var
  aParent: TcxTreeListNode;
begin
  if Assigned( ConfigSoTree.FocusedNode ) then
    if ConfigSoTree.FocusedNode.Level > 0 then
    begin
      aParent := ConfigSoTree.FocusedNode.Parent;
      ConfigSoTree.FocusedNode.Delete;
      if Assigned( aParent ) then aParent.Focused := True;
    end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.btnTestClick(Sender: TObject);
var
  aCompName, aUser, aPass, aAlias, aMessage: String;
  aConnection: TADOConnection;
begin
  if not Assigned( ConfigSoTree.FocusedNode ) then Exit;
  if ConfigSoTree.FocusedNode.Level <= 0 then Exit;
  aCompName := VarToStrDef( ConfigSoTree.FocusedNode.Values[ConfigSoTreeCol3.ItemIndex], EmptyStr );
  aUser := VarToStrDef( ConfigSoTree.FocusedNode.Values[ConfigSoTreeCol5.ItemIndex], EmptyStr );
  aPass := VarToStrDef( ConfigSoTree.FocusedNode.Values[ConfigSoTreeCol6.ItemIndex], EmptyStr );
  aAlias := VarToStrDef( ConfigSoTree.FocusedNode.Values[ConfigSoTreeCol7.ItemIndex], EmptyStr );
  if ( aUser = EmptyStr ) or ( aPass = EmptyStr ) or ( aAlias = EmptyStr ) then
  begin
    lblTest.Caption := Format( LanguageManager.Get( 'SConfigDbTestParamIsEmpty' ),
      [aCompName] );
    Exit;
  end;
  lblTest.Caption := Format( LanguageManager.Get( 'SConfigDbTestNow' ), [aCompName] );
  Screen.Cursor := crHourGlass;
  Application.ProcessMessages;
  try
    aConnection := TADOConnection.Create( nil );
    aConnection.LoginPrompt := False;
    try
      aConnection.ConnectionString := Format(
        'Provider=MSDAORA.1;Persist Security Info=True;' +
        'Password=%s;User ID=%s;Data Source=%s;',
        [aPass, aUser, aAlias] );
      aMessage := EmptyStr;
      try
        aConnection.Connected := True;
        aMessage := Format( LanguageManager.Get( 'SConfigDbTestOk' ),
          [aCompName] );
      except
        on E: Exception do
          aMessage := Format( LanguageManager.Get( 'SConfigDbTestError' ),
            [aCompName, E.Message] );
      end;
      if aConnection.Connected then aConnection.Connected := False;
      lblTest.Caption := aMessage;
    finally
      aConnection.Free;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.btnAddHighCmdClick(Sender: TObject);
var
  aNode: TcxTreeListNode;
begin
  HighCmdTree.SetFocus;
  aNode := HighCmdTree.Add;
  aNode.ImageIndex := 8;
  aNode.SelectedIndex := aNode.ImageIndex;
  HighCmdTree.FocusedNode := aNode;
  HighCmdTreeCol1.Editing := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.btnDeleteHighCmdClick(Sender: TObject);
begin
  if Assigned( HighCmdTree.FocusedNode ) then
    HighCmdTree.FocusedNode.Delete;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.btnAddLowCmdClick(Sender: TObject);
var
  aNode: TcxTreeListNode;
begin
  LowCmdTree.SetFocus;
  aNode := LowCmdTree.Add;
  aNode.ImageIndex := 1;
  aNode.SelectedIndex := aNode.ImageIndex;
  LowCmdTree.FocusedNode := aNode;
  LowCmdTreeCol1.Editing := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.btnDeleteLowCmdClick(Sender: TObject);
begin
  if Assigned( LowCmdTree.FocusedNode ) then
    LowCmdTree.FocusedNode.Delete;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.btnAddErrorClick(Sender: TObject);
var
  aNode: TcxTreeListNode;
begin
  ErrorTree.SetFocus;
  aNode := ErrorTree.Add;
  aNode.ImageIndex := 6;
  aNode.SelectedIndex := aNode.ImageIndex;
  ErrorTree.FocusedNode := aNode;
  ErrorTreeCol1.Editing := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.btnDeleteErrorClick(Sender: TObject);
begin
  if Assigned( ErrorTree.FocusedNode ) then
    ErrorTree.FocusedNode.Delete;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.ConfigPageChange(Sender: TObject);
begin
  HeaderPanel.Parent := ConfigPage.ActivePage;
  StyleModule.ConfigImageList.GetIcon( ConfigPage.ActivePageIndex,
    HeaderImage.Picture.Icon );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.InitControlLanguage;
var
  aIndex: Integer;
begin
  Self.Caption := LanguageManager.Get( 'dxMenuItem1' );
  {}
  cxCAS.Caption := LanguageManager.Get( 'ConfigPage1' );
  cxCommon.Caption := LanguageManager.Get( 'ConfigPage2' );
  cxSoDb.Caption := LanguageManager.Get( 'ConfigPage3' );
  cxHighCmd.Caption := LanguageManager.Get( 'ConfigPage4' );
  cxLowCmd.Caption := LanguageManager.Get( 'ConfigPage5' );
  cxCmdError.Caption := LanguageManager.Get( 'ConfigPage6' );
  {}
  btnConfirm.Caption := LanguageManager.Get( 'SConfirm' );
  btnCancel.Caption := LanguageManager.Get( 'SCancel' );
  {}
  lblCASIP.Caption := LanguageManager.Get( lblCASIP.Name );
  lblSendPort.Caption := LanguageManager.Get( lblSendPort.Name );
  lblRecvPort.Caption := LanguageManager.Get( lblRecvPort.Name );
  lblCAProtocol.Caption := LanguageManager.Get( lblCAProtocol.Name );
  lblCAProductDefine.Caption := LanguageManager.Get( lblCAProductDefine.Name );
  lblCARertyFeq.Caption := LanguageManager.Get( lblCARertyFeq.Name );
  lblCasCommCheck.Caption := LanguageManager.Get( lblCasCommCheck.Name );
  lblIdleTime.Caption := LanguageManager.Get( lblIdleTime.Name );
  chkEnableSend.Caption := LanguageManager.Get( chkEnableSend.Name );
  chkEnableRecv.Caption := LanguageManager.Get( chkEnableRecv.Name );
  {}
  lbDbRetryFreq.Caption := LanguageManager.Get( lbDbRetryFreq.Name );
  lblBusyTimeStart.Caption := LanguageManager.Get( lblBusyTimeStart.Name );
  lblBusyTimeEnd.Caption := LanguageManager.Get( lblBusyTimeEnd.Name );
  lblBusyTimeReadFreq.Caption := LanguageManager.Get( lblBusyTimeReadFreq.Name );
  lblNormalTime.Caption := LanguageManager.Get( lblNormalTime.Name );
  lblDbProcRecords.Caption := LanguageManager.Get( lblDbProcRecords.Name );
  lblCASendDelay.Caption := LanguageManager.Get( lblCASendDelay.Name );
  lblCAMaxError.Caption := LanguageManager.Get( lblCAMaxError.Name );
  lblDbWriteErrorWhenSocketFail.Caption := LanguageManager.Get( lblDbWriteErrorWhenSocketFail.Name );
  chkWriteNo.Caption := LanguageManager.Get( chkWriteNo.Name );
  chkWriteYes.Caption := LanguageManager.Get( chkWriteYes.Name );
  lblProcessIPPV.Caption := LanguageManager.Get( lblProcessIPPV.Name );
  chkProcessIPPV_A.Caption := LanguageManager.Get( chkProcessIPPV_A.Name );
  chkProcessIPPV_N.Caption := LanguageManager.Get( chkProcessIPPV_N.Name );
  chkProcessIPPV_O.Caption := LanguageManager.Get( chkProcessIPPV_O.Name );
  lblProcessBatch.Caption := LanguageManager.Get( lblProcessBatch.Name );
  ChkProcessBatch_A.Caption := LanguageManager.Get( ChkProcessBatch_A.Name );
  ChkProcessBatch_N.Caption := LanguageManager.Get( ChkProcessBatch_N.Name );
  ChkProcessBatch_O.Caption := LanguageManager.Get( ChkProcessBatch_O.Name );
  {}
  btnAddSo.Caption := LanguageManager.Get( 'SAdd' );
  btnDeleteSo.Caption := LanguageManager.Get( 'SDel' );
  btnTest.Caption := LanguageManager.Get( 'STest' );
  lblTest.Caption := EmptyStr;
  for aIndex := 0 to ConfigSoTree.ColumnCount - 1 do
    ConfigSoTree.Columns[aIndex].Caption.Text:=
      LanguageManager.Get( ConfigSoTree.Columns[aIndex].Name );
  {}
  btnAddHighCmd.Caption := LanguageManager.Get( 'SAdd' );
  btnDeleteHighCmd.Caption := LanguageManager.Get( 'SDel' );
  for aIndex := 0 to HighCmdTree.ColumnCount - 1 do
    HighCmdTree.Columns[aIndex].Caption.Text :=
      LanguageManager.Get( HighCmdTree.Columns[aIndex].Name );
  {}
  btnAddLowCmd.Caption := LanguageManager.Get( 'SAdd' );
  btnDeleteLowCmd.Caption := LanguageManager.Get( 'SDel' );
  for aIndex := 0 to LowCmdTree.ColumnCount - 1 do
    LowCmdTree.Columns[aIndex].Caption.Text :=
      LanguageManager.Get( LowCmdTree.Columns[aIndex].Name );
  {}
  btnAddError.Caption := LanguageManager.Get( 'SAdd' );
  btnDeleteError.Caption := LanguageManager.Get( 'SDel' );
  for aIndex := 0 to ErrorTree.ColumnCount - 1 do
    ErrorTree.Columns[aIndex].Caption.Text :=
      LanguageManager.Get( ErrorTree.Columns[aIndex].Name );
end;

{ ---------------------------------------------------------------------------- }

end.
