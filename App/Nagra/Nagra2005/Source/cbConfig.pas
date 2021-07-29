unit cbConfig;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Menus, ImgList, DB, ADODB, 
  {$IFDEF DEBUG} CodeSiteLogging, {$ENDIF}
  { Developer Express Suite }
  dxSkinsCore, cxLookAndFeelPainters, dxSkinsDefaultPainters, dxSkinscxPCPainter,
  cxPC, cxButtons, cxControls, cxContainer, cxEdit, cxTextEdit, cxMaskEdit,
  cxCheckBox, cxSpinEdit, cxTimeEdit, cxRadioGroup, cxGraphics, cxCustomData,
  cxStyles, cxTL, cxImageComboBox, cxInplaceContainer, cxDropDownEdit,
  cxProgressBar, dxSkinBlack, dxSkinBlue, dxSkinCaramel, dxSkinCoffee;



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
    lblControlPort: TLabel;
    txtControlPort: TcxMaskEdit;
    lblFeedbackPort: TLabel;
    txtFeedbackPort: TcxMaskEdit;
    lblSourceId: TLabel;
    txtControlSourceId: TcxMaskEdit;
    lblDestId: TLabel;
    txtControlDestId: TcxMaskEdit;
    lblMopPpid: TLabel;
    txtControlMopPpid: TcxMaskEdit;
    chkEnableControlSend: TcxCheckBox;
    chkEnableFeedbackRecv: TcxCheckBox;
    lblSendCommandDelay: TLabel;
    txtSendCommandDelay: TcxSpinEdit;
    chkUseSimulator: TcxCheckBox;
    cxCommon: TcxTabSheet;
    chkDbMultiThread: TcxCheckBox;
    lbDbRetryFrequence: TLabel;
    txtDbRetryFrequence: TcxSpinEdit;
    lblBusyTimeStart: TLabel;
    txtBusyTimeStart: TcxTimeEdit;
    lblBusyTimeEnd: TLabel;
    txtBusyTimeEnd: TcxTimeEdit;
    lblBusyTimeReadFrquence: TLabel;
    txtlBusyTimeReadFrquence: TcxSpinEdit;
    lblNormalTime: TLabel;
    txtNormalTime: TcxSpinEdit;
    lblProcessRecordCount: TLabel;
    txtProcessRecordCount: TcxSpinEdit;
    Label2: TLabel;
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
    lblCasRetryFrquence: TLabel;
    txtCasRetryFrquence: TcxSpinEdit;
    lblBatchOperator: TLabel;
    txtBatchOperator: TcxTextEdit;
    lblCasSendErrMax: TLabel;
    txtCasSendErrMax: TcxSpinEdit;
    lblCasRecvErrMax: TLabel;
    txtCasRecvErrMax: TcxSpinEdit;
    lblCasCommCheck: TLabel;
    txtCasCommCheck: TcxSpinEdit;
    HeaderPanel: TPanel;
    HeaderImage: TImage;
    cxSoDb: TcxTabSheet;
    SoTree: TcxTreeList;
    SoTreeColumn1: TcxTreeListColumn;
    SoTreeColumn2: TcxTreeListColumn;
    SoTreeColumn3: TcxTreeListColumn;
    SoTreeColumn4: TcxTreeListColumn;
    SoTreeColumn5: TcxTreeListColumn;
    SoTreeColumn6: TcxTreeListColumn;
    SoTreeColumn7: TcxTreeListColumn;
    SoTreeColumn8: TcxTreeListColumn;
    SoTreeColumn9: TcxTreeListColumn;
    btnDeleteSo: TcxButton;
    btnAddSo: TcxButton;
    cxHighCmd: TcxTabSheet;
    HighCmdTree: TcxTreeList;
    btnTest: TcxButton;
    lblTest: TLabel;
    HighCmdTreeColumn1: TcxTreeListColumn;
    HighCmdTreeColumn2: TcxTreeListColumn;
    HighCmdTreeColumn3: TcxTreeListColumn;
    HighCmdTreeColumn5: TcxTreeListColumn;
    HighCmdTreeColumn6: TcxTreeListColumn;
    HighCmdTreeColumn7: TcxTreeListColumn;
    HighCmdTreeColumn4: TcxTreeListColumn;
    btnAddHighCmd: TcxButton;
    btnDeleteHighCmd: TcxButton;
    cxLowCmd: TcxTabSheet;
    btnAddLowCmd: TcxButton;
    btnDeleteLowCmd: TcxButton;
    LowCmdTree: TcxTreeList;
    LowCmdTreeColumn1: TcxTreeListColumn;
    LowCmdTreeColumn2: TcxTreeListColumn;
    cxCmdError: TcxTabSheet;
    btnAddError: TcxButton;
    btnDeleteError: TcxButton;
    ErrorTree: TcxTreeList;
    ErrorTreeColumn1: TcxTreeListColumn;
    ErrorTreeColumn2: TcxTreeListColumn;
    ErrorTreeColumn3: TcxTreeListColumn;
    Label1: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    txtFeedbackSourceId: TcxMaskEdit;
    txtFeedbackDestId: TcxMaskEdit;
    txtFeedbackMopPpid: TcxMaskEdit;
    Label6: TLabel;
    Label7: TLabel;
    txtControlTransId: TcxMaskEdit;
    Label8: TLabel;
    txtFeedbackTransId: TcxMaskEdit;
    HighCmdTreeColumn8: TcxTreeListColumn;
    Label9: TLabel;
    txtCasCmdMaxCounter: TcxSpinEdit;
    HighCmdTreeColumn9: TcxTreeListColumn;
    HighCmdTreeColumn10: TcxTreeListColumn;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure SoTreeFocusedNodeChanged(Sender: TObject; APrevFocusedNode,
      AFocusedNode: TcxTreeListNode);
    procedure SoTreeEdited(Sender: TObject; AColumn: TcxTreeListColumn);
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
    procedure GetCmdTransEnv;
    procedure GetCASEnv;
    procedure GetCmdError;
    procedure GetSoList;
    procedure GetFeedbackSo;
    procedure EnableConfigControls;
    {$HINTS OFF}
       procedure EnableConfigControl(aControl: TControl);
       procedure DisableConfigControl(aControl: TControl);
    {$HINTS ON}
    procedure DisableConfigControls;

  public
    { Public declarations }
  end;

var
  fmConfig: TfmConfig;

implementation

uses cbMain, cbClass, cbStyleModule, cbConfigModule, cbResStr;

{$R *.dfm}

var aGlobeKeyValue: String;


{ ---------------------------------------------------------------------------- }

procedure TfmConfig.FormCreate(Sender: TObject);
begin
  Self.Caption := SMenuItem1;
  btnConfirm.Caption := SButtonSave;
  btnCancel.Caption := SButtonCancel;
  btnAddSo.Caption := SButtonAdd;
  btnDeleteSo.Caption := SButtonDelete;
  btnTest.Caption := SButtonDbTest;
  btnAddHighCmd.Caption := SButtonAdd;
  btnDeleteHighCmd.Caption := SButtonDelete;
  btnAddLowCmd.Caption := SButtonAdd;
  btnDeleteLowCmd.Caption := SButtonDelete;
  btnAddError.Caption := SButtonAdd;
  btnDeleteError.Caption := SButtonDelete;
  fmConfig := Self;
  cxHighCmd.TabVisible := False;
  cxLowCmd.TabVisible := False;
  cxCmdError.TabVisible := False;
  GetCASEnv;
  GetHighCmdEnv;
  GetLowCmdEnv;
  GetCmdError;
  GetCmdTransEnv;
  GetCommon;
  GetSoList;
  GetFeedbackSo;
  if ConfigPage.ActivePageIndex <> 0 then
    ConfigPage.ActivePageIndex := 0;
  ConfigPageChange( ConfigPage );
  lblTest.Caption := EmptyStr;
  DisableConfigControls;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.FormDestroy(Sender: TObject);
begin
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.GetCASEnv;
begin
  txtCASIP.Text := ConfigModule.CASEnv.IP;
  txtControlPort.Text := IntToStr( ConfigModule.CASEnv.ControlPort );
  txtFeedbackPort.Text := IntToStr( ConfigModule.CASEnv.FeedbackPort );
  txtControlSourceId.Text := ConfigModule.CASEnv.ControlSourceId;
  txtControlDestId.Text := ConfigModule.CASEnv.ControlDestId;
  txtControlMopPpid.Text := ConfigModule.CASEnv.ControlMopPPId;
  txtFeedbackSourceId.Text := ConfigModule.CASEnv.FeedbackSourceId;
  txtFeedbackDestId.Text := ConfigModule.CASEnv.FeedbackDestId;
  txtFeedbackMopPpid.Text := ConfigModule.CASEnv.FeedbackMopPPId;
  txtControlTransId.Text := ConfigModule.CASEnv.ControlTransId;
  txtFeedbackTransId.Text := ConfigModule.CASEnv.FeedbackTransId;
  txtCasCmdMaxCounter.Text := IntToStr( ConfigModule.CASEnv.CmdMaxCounter );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.GetCmdError;
var
  aNode: TcxTreeListNode;
begin
  ErrorTree.Clear;
  ConfigModule.CmdError.First;
  while not ConfigModule.CmdError.Eof do
  begin
    aNode := ErrorTree.Add;
    aNode.AssignValues( [
      ConfigModule.CmdError.FieldByName( 'ERROR_FLAG' ).AsString,
      ConfigModule.CmdError.FieldByName( 'ERROR_CODE' ).AsString,
      ConfigModule.CmdError.FieldByName( 'ERROR_DESC' ).AsString ] );
    aNode.ImageIndex := 6;
    aNode.SelectedIndex := aNode.ImageIndex;
    ConfigModule.CmdError.Next;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.GetCmdTransEnv;
begin
  chkEnableControlSend.Checked := ConfigModule.CmdTransEnv.EnableControlSend;
  chkEnableFeedbackRecv.Checked := ConfigModule.CmdTransEnv.EnableFeedbackRecv;
  txtSendCommandDelay.Value := ConfigModule.CmdTransEnv.SendCommandDelay;
  chkUseSimulator.Checked := ConfigModule.CmdTransEnv.UseSimulator;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.GetCommon;
begin
  chkDbMultiThread.Checked := ConfigModule.CommEnv.DbMultiThread;
  txtDbRetryFrequence.Value := ConfigModule.CommEnv.DbRetryFrequence;
  txtBusyTimeStart.Text := ConfigModule.CommEnv.BusyTimeStart;
  txtBusyTimeEnd.Text := ConfigModule.CommEnv.BusyTimeEnd;
  txtlBusyTimeReadFrquence.Value := ConfigModule.CommEnv.BusyTimeDbReadFrequence;
  txtNormalTime.Value := ConfigModule.CommEnv.NormalTimeDbReadFrequence;
  txtProcessRecordCount.Value := ConfigModule.CommEnv.ProcessRecordCount;
  chkWriteYes.Checked := ConfigModule.CommEnv.DbWriteErrorWhenSocketFail;
  if ConfigModule.CommEnv.ProcessIPPV = 'A' then
    chkProcessIPPV_A.Checked := True
  else if ConfigModule.CommEnv.ProcessIPPV = 'O' then
    chkProcessIPPV_O.Checked := True
  else
    chkProcessIPPV_N.Checked := True;
  if ConfigModule.CommEnv.ProcessBatch = 'A' then
    ChkProcessBatch_A.Checked := True
  else if ConfigModule.CommEnv.ProcessBatch = 'O' then
    ChkProcessBatch_O.Checked := True
  else
    ChkProcessBatch_N.Checked := True;
  txtBatchOperator.Text := ConfigModule.CommEnv.BatchOperator;
  {}
  txtCasRetryFrquence.Value := ConfigModule.CommEnv.CASRetryFrequence;
  txtCasSendErrMax.Value := ConfigModule.CommEnv.CASSendErrMax;
  txtCasRecvErrMax.Value := ConfigModule.CommEnv.CASRecvErrMax;
  txtCasCommCheck.Value := ConfigModule.CommEnv.CASCommCheck;
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
      ConfigModule.HighCmdEnv.FieldByName( 'HIGH_LEVEL_CMD' ).AsString,
      ConfigModule.HighCmdEnv.FieldByName( 'DESCRIPTION' ).AsString,
      ConfigModule.HighCmdEnv.FieldByName( 'CMD_TYPE' ).AsString,
      ConfigModule.HighCmdEnv.FieldByName( 'LOW_LEVEL_CMD' ).AsString,
      ConfigModule.HighCmdEnv.FieldByName( 'NOTES_USE' ).AsString,
      ConfigModule.HighCmdEnv.FieldByName( 'CMD_TYPE_PRIORITY' ).AsString,
      ConfigModule.HighCmdEnv.FieldByName( 'CMD_PRIORITY' ).AsString,
      ConfigModule.HighCmdEnv.FieldByName( 'IRD_CMD' ).AsString,
      ConfigModule.HighCmdEnv.FieldByName( 'NEW_CMD_COUNT' ).AsString,
      ConfigModule.HighCmdEnv.FieldByName( 'OLD_CMD_COUNT' ).AsString] );
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
      ConfigModule.LowCmdEnv.FieldByName( 'LOW_LEVEL_CMD' ).AsString,
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
  txtControlPort.Enabled := True;
  txtFeedbackPort.Enabled := True;
  txtControlSourceId.Enabled := True;
  txtControlDestId.Enabled := True;
  txtControlMopPpid.Enabled := True;
  txtFeedbackSourceId.Enabled := True;
  txtFeedbackDestId.Enabled := True;
  txtFeedbackMopPpid.Enabled := True;
  {}
  txtControlTransId.Enabled := True;
  txtFeedbackTransId.Enabled := True;
  {}
  chkDbMultiThread.Enabled := True;
  chkUseSimulator.Enabled := True;
  {}
  SoTreeColumn5.Visible := True;
  SoTreeColumn6.Visible := True;
  SoTreeColumn7.Visible := True;
  {}
  for aIndex := 0 to SoTree.ColumnCount - 1 do
  begin
    if not SoTree.Columns[aIndex].Visible then Continue;
    SoTree.Columns[aIndex].Options.Focusing := True;
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
  if aControl = SoTree then
  begin
    SoTree.Enabled := True;
    for aIndex := 0 to SoTree.ColumnCount - 1 do
    begin
      if not SoTree.Columns[aIndex].Visible then Continue;
      SoTree.Columns[aIndex].Options.Focusing := True;
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
  chkDbMultiThread.Enabled := False;
  chkUseSimulator.Enabled := False;
  for aIndex := 0 to SoTree.ColumnCount - 1 do
  begin
    if not SoTree.Columns[aIndex].Visible then Continue;
    if ( aIndex > 0 ) and ( aIndex <> 9 ) then
      SoTree.Columns[aIndex].Options.Focusing := False;
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
  if aControl = SoTree then
  begin
    for aIndex := 0 to SoTree.ColumnCount - 1 do
    begin
      if not SoTree.Columns[aIndex].Visible then Continue;
      if ( aIndex > 0 ) and ( aIndex <> 9 ) then
        SoTree.Columns[aIndex].Options.Focusing := False;
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
  if UpperCase( aGlobeKeyValue ) = 'DEVELOPER' then
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
  Result := fmConfig.SoTree.FindNodeByText( aPosName, fmConfig.SoTreeColumn3 );
  if not Assigned( Result ) and aCanAdd then
  begin
    Result := fmConfig.SoTree.Add( nil );
    Result.AssignValues( [EmptyStr, EmptyStr, aPosName, EmptyStr,
      EmptyStr, EmptyStr, EmptyStr, aPosType, aPos] );
    Result.ImageIndex := 0;
    Result.SelectedIndex := 0;
  end;
end;

{ ---------------------------------------------------------------------------- }

function GetSoImageIndex(const aStatus: TDbConnectStatus): Integer;
begin
  case aStatus of
    dbNoSelect: Result := 1;
  else
    Result := 2;
  end;
end;

{ ---------------------------------------------------------------------------- }

function IsSelected(const aStatus: TDbConnectStatus): String;
begin
  Result := 'Y';
  if ( aStatus in [dbNoSelect] ) then Result := 'N';
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.GetSoList;

var
  aIndex: Integer;
  aList: TList;
  aSo: PSoInfo;
  aNodePos, aNodeComp: TcxTreeListNode;
  aSelected: String;
begin
  aList := ConfigModule.SoList.LockList;
  try
    for aIndex := 0 to aList.Count - 1 do
    begin
      aSo := PSoInfo( aList.Items[aIndex] );
      { 取平台節點, 此系統台要掛在那一個平台下 }
      aNodePos := GetPosNode( IntToStr( aSo.Pos ), aSo.SoType, aSo.PosName );
      aNodeComp := SoTree.AddChild( aNodePos );
      aSelected := IsSelected( aSo.DbConnectStatus );
      aNodeComp.AssignValues( [aSelected, aSo.CompCode, aSo.CompName,
        aSo.NetworkId, aSo.LoginUser, aSo.LoginPass, aSo.DbAliase, aSo.SoType, aSo.Pos] );
      aNodeComp.ImageIndex := GetSoImageIndex( aSo.DbConnectStatus );
      aNodeComp.SelectedIndex := aNodeComp.ImageIndex;
      aNodePos.Expanded := True;
    end;
  finally
    ConfigModule.SoList.UnlockList;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.GetFeedbackSo;
var
  aSo: PSoInfo;
  aNodePos, aNodeComp: TcxTreeListNode;
  aSelected: String;
begin
  if ( ConfigModule.FeedbackSo.Pos < 0 ) then Exit;
  aSo := ConfigModule.FeedbackSo;
  aNodePos := GetPosNode( IntToStr( aSo.Pos ), aSo.SoType, aSo.PosName );
  aNodeComp := SoTree.AddChild( aNodePos );
  aSelected := IsSelected( aSo.DbConnectStatus );
  aNodeComp.AssignValues( [aSelected, aSo.CompCode, aSo.CompName,
    aSo.NetworkId, aSo.LoginUser, aSo.LoginPass, aSo.DbAliase, aSo.SoType, aSo.Pos] );
  aNodeComp.ImageIndex := GetSoImageIndex( ConfigModule.FeedbackSo.DbConnectStatus );
  aNodeComp.SelectedIndex := aNodeComp.ImageIndex;
  aNodePos.Expanded := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.SoTreeFocusedNodeChanged(Sender: TObject;
  APrevFocusedNode, AFocusedNode: TcxTreeListNode);
begin
  if Assigned( AFocusedNode ) then
    SoTree.OptionsSelection.CellSelect := ( AFocusedNode.Level > 0 );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.SoTreeEdited(Sender: TObject;
  AColumn: TcxTreeListColumn);
begin
 if AColumn = SoTreeColumn1 then
 begin
   if AColumn.Value = 'Y' then
     SoTree.FocusedNode.ImageIndex := 2
   else
     SoTree.FocusedNode.ImageIndex := 1;
   SoTree.FocusedNode.SelectedIndex := SoTree.FocusedNode.ImageIndex;
 end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.btnAddSoClick(Sender: TObject);
var
  aParentNode, aNode: TcxTreeListNode;
begin
  if not Assigned( SoTree.FocusedNode ) then Exit;
  aParentNode := nil;
  if SoTree.FocusedNode.Level = 0 then
    aParentNode := SoTree.FocusedNode
  else begin
    if Assigned( SoTree.FocusedNode.Parent ) then
      aParentNode := SoTree.FocusedNode.Parent;
  end;
  if Assigned( aParentNode ) then
  begin
    aNode := aParentNode.AddChild;
    aNode.ImageIndex := 1;
    aNode.SelectedIndex := aNode.ImageIndex;
    aNode.AssignValues( ['N', EmptyStr, EmptyStr, EmptyStr, EmptyStr,
      EmptyStr, EmptyStr, aParentNode.Values[SoTreeColumn8.ItemIndex],
      aParentNode.Values[SoTreeColumn9.ItemIndex] ] );
    aNode.Focused := True;
    SoTree.SetFocus;
    SoTreeColumn2.Editing := True;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.btnDeleteSoClick(Sender: TObject);
var
  aParent: TcxTreeListNode;
begin
  if Assigned( SoTree.FocusedNode ) then
    if SoTree.FocusedNode.Level > 0 then
    begin
      aParent := SoTree.FocusedNode.Parent;
      SoTree.FocusedNode.Delete;
      if Assigned( aParent ) then aParent.Focused := True;
    end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.btnTestClick(Sender: TObject);
var
  aCompName, aUser, aPass, aAlias, aMessage: String;
  aConnection: TADOConnection;
begin
  if not Assigned( SoTree.FocusedNode ) then Exit;
  if SoTree.FocusedNode.Level <= 0 then Exit;
  aCompName := VarToStrDef( SoTree.FocusedNode.Values[SoTreeColumn3.ItemIndex], EmptyStr );
  aUser := VarToStrDef( SoTree.FocusedNode.Values[SoTreeColumn5.ItemIndex], EmptyStr );
  aPass := VarToStrDef( SoTree.FocusedNode.Values[SoTreeColumn6.ItemIndex], EmptyStr );
  aAlias := VarToStrDef( SoTree.FocusedNode.Values[SoTreeColumn7.ItemIndex], EmptyStr );
  if ( aUser = EmptyStr ) or ( aPass = EmptyStr ) or ( aAlias = EmptyStr ) then
  begin
    lblTest.Caption := Format( SDbTestParamIsEmpty, [aCompName] );
    Exit;
  end;
  lblTest.Caption := Format( SDbTestNow, [aCompName] );
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
        aMessage := Format( SDbTestOk, [aCompName] );
      except
        on E: Exception do
          aMessage := Format( SDbTestError, [aCompName, E.Message] );
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
  HighCmdTreeColumn1.Editing := True;
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
  LowCmdTreeColumn1.Editing := True;
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
  ErrorTreeColumn1.Editing := True;
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

end.
