unit cbMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ImgList, ExtCtrls, StdCtrls, Contnrs, ComCtrls,
  cbSoInfo,
  cxPC, cxControls, dxBar, cxInplaceContainer, cxTL, dxBarExtItems, cxEdit,
  cxContainer, cxListView, cxSplitter, cxLookAndFeels, cxImageComboBox,
  cxBlobEdit, dxSkinsCore, dxSkinsDefaultPainters, dxSkinscxPCPainter,
  dxSkinsdxBarPainter, cxClasses;

const
  DBINIFILE = 'CSIS.INI';

type
  TMain = class(TForm)
    dxBarManager: TdxBarManager;
    BarImageList: TImageList;
    SoPageControl: TcxPageControl;
    Bevel1: TBevel;
    PageImageList: TImageList;
    dxPlay: TdxBarButton;
    dxStop: TdxBarButton;
    dxConfig: TdxBarButton;
    dxWebServer: TdxBarStatic;
    cxMainSplitter: TcxSplitter;
    cxDefaultEditStyleController: TcxDefaultEditStyleController;
    Panel2: TPanel;
    MsgListView: TcxListView;
    dxBarDockControl1: TdxBarDockControl;
    TreeImageList: TImageList;
    dxStopGroup: TdxBarGroup;
    dxPlayGroup: TdxBarGroup;
    StartupTimer: TTimer;
    MsgImageList: TImageList;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormDestroy(Sender: TObject);
    procedure StartupTimerTimer(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure dxPlayClick(Sender: TObject);
    procedure dxStopClick(Sender: TObject);
  private
    { Private declarations }
    FSoList: TObjectList;
    FConfig: TConfigInfo;
    FMsg: PMsg;
    procedure ReleaseSoList;
    procedure CreateSoList;
    procedure CreateTabeSheet;
    procedure CreateTreeList;
    procedure ReleaseConfig;
    procedure CreateConfig;
    procedure CreateSoThread;
    procedure ReleaseSoThread;
  public
    { Public declarations }
    procedure MsgNotify(aMsg: Pointer);
    procedure DbNotify(aSoInfo: TSoInfo);
    procedure RecordNotify(aSoInfo: TSoInfo; aSO311: PSO311);
  end;

var
  Main: TMain;

implementation

uses cbUtils, cbCMCP;


{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TMain.FormCreate(Sender: TObject);
begin
  New( FMsg );
  dxPlayGroup.Enabled := False;
  dxStopGroup.Enabled := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TMain.FormShow(Sender: TObject);
begin
  FormResize( Self );
  StartupTimer.Enabled := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TMain.FormResize(Sender: TObject);
begin
  MsgListView.Columns[0].Width := ( Self.Width - 10 );
end;

{ ---------------------------------------------------------------------------- }

procedure TMain.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  if not dxPlayGroup.Enabled then
  begin
    CanClose := ConfirmMsg( '確認結束程式?' );
    if CanClose then
    begin
      try
        dxStop.Click;
      except
        { ... }
      end;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TMain.FormDestroy(Sender: TObject);
begin
  ReleaseSoThread;
  ReleaseConfig;
  ReleaseSoList;
  Dispose( FMsg );
end;

{ ---------------------------------------------------------------------------- }

procedure TMain.ReleaseSoList;
begin
  if not Assigned( FSoList ) then Exit;
  while FSoList.Count > 0 do
    FSoList.Delete( 0 );
  FreeAndNil( FSoList );
end;

{ ---------------------------------------------------------------------------- }

procedure TMain.CreateSoList;
var
  aLoader: TSoLoader;
  aFileName: String;
begin
  ReleaseSoList;
  aLoader := TSoLoader.Create;
  try
    aFileName := IncludeTrailingPathDelimiter( ExtractFilePath(
      ParamStr( 0 ) ) ) + DBINIFILE;
    aLoader.LoadFromFile( aFileName );
    FSoList := aLoader.LoaderResult;
  finally
    aLoader.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TMain.CreateTabeSheet;
var
  aIndex: Integer;
  aSoInfo: TSoInfo;
  aTabSheet: TcxTabSheet; 
begin
  for aIndex := 0 to FSoList.Count - 1 do
  begin
    if not Assigned( FSoList[aIndex] ) then Continue;
    aSoInfo := TSoInfo( FSoList[aIndex] );
    if not aSoInfo.Enable then Continue;
    aTabSheet := TcxTabSheet.Create( Self );
    aTabSheet.Name := Format( 'TabSheet%s', [aSoInfo.CompCode] );
    aTabSheet.ImageIndex := 0;
    aTabSheet.Caption := Format( '   %s   ', [aSoInfo.CompName] );
    aTabSheet.BorderWidth := 2;
    aTabSheet.PageControl := SoPageControl;
    aSoInfo.DisplaySheet := aTabSheet;
    SoPageControl.Update;
    Delay( 300 );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TMain.CreateTreeList;
var
  aIndex: Integer;
  aSoInfo: TSoInfo;
  aTreeList: TcxTreeList;
  aTreeColumn: TcxTreeListColumn;
  aItem: TcxImageComboBoxItem;
begin
  for aIndex := 0 to FSoList.Count - 1 do
  begin
    if not Assigned( FSoList[aIndex] ) then Continue;
    aSoInfo := TSoInfo( FSoList[aIndex] );
    if not aSoInfo.Enable then Continue;
    aTreeList := TcxTreeList.Create( Main );
    aTreeList.Name := Format( 'TreeList%s', [aSoInfo.CompCode] );
    aTreeList.Images := TreeImageList;
    aTreeList.LookAndFeel.Kind := lfOffice11;
    aTreeList.OptionsBehavior.Sorting := False;
    aTreeList.OptionsData.Deleting := False;
    aTreeList.OptionsData.Editing := True;
    aTreeList.OptionsData.Inserting := False;
    aTreeList.OptionsSelection.CellSelect := True;
    aTreeList.OptionsSelection.HideFocusRect := False;
    aTreeList.OptionsView.Bands := True;
    aTreeList.OptionsView.CellEndEllipsis := True;
    aTreeList.OptionsView.GridLines := tlglBoth;
    aTreeList.OptionsView.Indicator := True;
    {}
    with aTreeList.Bands.Add do
    begin
      FixedKind := tlbfNone;
      Caption.Text := 'CM/CP';
      Caption.AlignHorz := taLeftJustify;
    end;
    with aTreeList.Bands.Add do
    begin
      FixedKind := tlbfRight;
      Caption.Text := 'WEB SERVICE';
      Caption.AlignHorz := taCenter;
    end;
    {}
    aTreeColumn := aTreeList.CreateColumn( aTreeList.Bands.Items[0] );
    aTreeColumn.Name := Format( '%s_CmdSeqNo', [aTreeList.Name] );
    aTreeColumn.Caption.Text := '指令序號';
    aTreeColumn.Options.Editing := False;
    aTreeList.Update;
    {}
    aTreeColumn := aTreeList.CreateColumn( aTreeList.Bands.Items[0] );
    aTreeColumn.Name := Format( '%s_MsgCode', [aTreeList.Name] );
    aTreeColumn.Caption.Text := '動作別';
    aTreeColumn.Options.Editing := False;
    aTreeList.Update;
    {}
    aTreeColumn := aTreeList.CreateColumn( aTreeList.Bands.Items[0] );
    aTreeColumn.Name := Format( '%s_WorkId', [aTreeList.Name] );
    aTreeColumn.Caption.Text := '操作人員';
    aTreeColumn.Options.Editing := False;
    aTreeList.Update;
    {}
    aTreeColumn := aTreeList.CreateColumn( aTreeList.Bands.Items[0] );
    aTreeColumn.Name := Format( '%s_CusId', [aTreeList.Name] );
    aTreeColumn.Caption.Text := '客編';
    aTreeColumn.Options.Editing := False;
    aTreeList.Update;
    {}
    aTreeColumn := aTreeList.CreateColumn( aTreeList.Bands.Items[0] );
    aTreeColumn.Name := Format( '%s_CMMAC', [aTreeList.Name] );
    aTreeColumn.Caption.Text := 'CMMAC';
    aTreeColumn.Options.Editing := False;
    aTreeList.Update;
    {}
    aTreeColumn := aTreeList.CreateColumn( aTreeList.Bands.Items[0] );
    aTreeColumn.Name := Format( '%s_EMTAMAC', [aTreeList.Name] );
    aTreeColumn.Caption.Text := 'EMTAMAC';
    aTreeColumn.Options.Editing := False;
    aTreeList.Update;
    {}
    aTreeColumn := aTreeList.CreateColumn( aTreeList.Bands.Items[0] );
    aTreeColumn.Name := Format( '%s_OperType', [aTreeList.Name] );
    aTreeColumn.Caption.Text := '功能指令';
    aTreeColumn.Options.Editing := False;
    aTreeList.Update;
    {}
    aTreeColumn := aTreeList.CreateColumn( aTreeList.Bands.Items[0] );
    aTreeColumn.Name := Format( '%s_SendDateTime', [aTreeList.Name] );
    aTreeColumn.Caption.Text := '傳送/回應時間';
    aTreeColumn.Options.Editing := False;
    aTreeList.Update;
    {}
    aTreeColumn := aTreeList.CreateColumn( aTreeList.Bands.Items[1] );
    aTreeColumn.Name := Format( '%s_Call', [aTreeList.Name] );
    aTreeColumn.Caption.Text := '呼叫';
    aTreeColumn.Options.Editing := False;
    aTreeColumn.Caption.AlignHorz := taCenter;
    aTreeColumn.Width := 40;
    aTreeColumn.PropertiesClass := TcxImageComboBoxProperties;
    TcxImageComboBoxProperties( aTreeColumn.Properties ).Images := TreeImageList;
    aItem := TcxImageComboBoxProperties( aTreeColumn.Properties ).Items.Add;
    aItem.Value := '1';
    aItem.ImageIndex := 1;
    aItem := TcxImageComboBoxProperties( aTreeColumn.Properties ).Items.Add;
    aItem.Value := '2';
    aItem.ImageIndex := 2;
    aTreeList.Update;
    {}
    aTreeColumn := aTreeList.CreateColumn( aTreeList.Bands.Items[1] );
    aTreeColumn.Name := Format( '%s_SOAPRequest', [aTreeList.Name] );
    aTreeColumn.Caption.Text := 'HTTP(Request)';
    aTreeColumn.PropertiesClass := TcxBlobEditProperties;
    TcxBlobEditProperties( aTreeColumn.Properties ).BlobEditKind := bekMemo;
    TcxBlobEditProperties( aTreeColumn.Properties ).ReadOnly := True;
    aTreeList.Update;
    {}
    aTreeColumn := aTreeList.CreateColumn( aTreeList.Bands.Items[1] );
    aTreeColumn.Name := Format( '%s_Resp', [aTreeList.Name] );
    aTreeColumn.Caption.Text := '回應';
    aTreeColumn.Options.Editing := False;
    aTreeColumn.Caption.AlignHorz := taCenter;
    aTreeColumn.Width := 40;
    aTreeColumn.PropertiesClass := TcxImageComboBoxProperties;
    TcxImageComboBoxProperties( aTreeColumn.Properties ).Images := TreeImageList;
    aItem := TcxImageComboBoxProperties( aTreeColumn.Properties ).Items.Add;
    aItem.Value := '1';
    aItem.ImageIndex := 1;
    aItem := TcxImageComboBoxProperties( aTreeColumn.Properties ).Items.Add;
    aItem.Value := '2';
    aItem.ImageIndex := 2;
    aTreeList.Update;
    {}
    aTreeColumn := aTreeList.CreateColumn( aTreeList.Bands.Items[1] );
    aTreeColumn.Name := Format( '%s_SOAPResponse', [aTreeList.Name] );
    aTreeColumn.Caption.Text := 'HTTP(Response)';
    aTreeColumn.PropertiesClass := TcxBlobEditProperties;
    TcxBlobEditProperties( aTreeColumn.Properties ).BlobEditKind := bekMemo;
    TcxBlobEditProperties( aTreeColumn.Properties ).ReadOnly := True;
    aTreeList.Update;
    {}
    aTreeList.Parent := aSoInfo.DisplaySheet;
    aTreeList.Align := alClient;
    aSoInfo.DisplayTreeList := aTreeList;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TMain.ReleaseConfig;
begin
  if Assigned( FConfig ) then FreeAndNil( FConfig );
end;

{ ---------------------------------------------------------------------------- }

procedure TMain.CreateConfig;
var
  aLoader: TConfigLoader;
  aFileName: String;
begin
  ReleaseConfig;
  aLoader := TConfigLoader.Create;
  try
    aFileName := IncludeTrailingPathDelimiter( ExtractFilePath( ParamStr( 0 )
      ) ) + DBINIFILE;
    aLoader.LoadFromFile( aFileName );
    FConfig := aLoader.LoaderResult;
  finally
    aLoader.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TMain.ReleaseSoThread;
var
  aIndex: Integer;
  aSoInfo: TSoInfo;
begin
  for aIndex := 0 to FSoList.Count - 1 do
  begin
    aSoInfo := TSoInfo( FSoList[aIndex] );
    if not Assigned( aSoInfo.ExecThread ) then Continue;
    if aSoInfo.Enable then
    begin
      try
        aSoInfo.ExecThread.Terminate;
        aSoInfo.ExecThread.WaitFor;
        aSoInfo.ExecThread.Free;
        FMsg.Kind := MB_ICONINFORMATION;
        FMsg.Msg := Format( '系統台: %s, 執行緒已停止。', [aSoInfo.CompName] );
      except
        on E: Exception do
        begin
          FMsg.Kind := MB_ICONERROR;
          FMsg.Msg := Format( '系統台: %s, 停止執行緒失敗, 原因: %s。',
           [aSoInfo.CompName, E.Message] );
        end;
      end;
    end;  
    aSoInfo.ExecThread := nil;
    MsgNotify( FMsg );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TMain.CreateSoThread;
var
  aIndex: Integer;
  aSoInfo: TSoInfo;
begin
  for aIndex := 0 to FSoList.Count - 1 do
  begin
    aSoInfo := TSoInfo( FSoList[aIndex] );
    if aSoInfo.Enable then
      aSoInfo.ExecThread := TCMCPThread.Create( aSoInfo, FConfig );
  end;
  for aIndex := 0 to FSoList.Count - 1 do
  begin
    aSoInfo := TSoInfo( FSoList[aIndex] );
    if aSoInfo.Enable then
      aSoInfo.ExecThread.Resume;
  end;
  FMsg.Kind := MB_OK;
  FMsg.Msg := 'CM/CP WebService Call Gateway開始執行。';
  MsgNotify( FMsg );
end;

{ ---------------------------------------------------------------------------- }

procedure TMain.StartupTimerTimer(Sender: TObject);
begin
  StartupTimer.Enabled := False;
  Screen.Cursor := crHourGlass;
  try
   FMsg.Kind := MB_ICONINFORMATION;
   FMsg.Msg := '啟動環境配置中, 請稍候。';
   MsgNotify( FMsg );
   Delay( 300 );
   CreateSoList;
   Delay( 300 );
   CreateTabeSheet;
   Delay( 300 );
   CreateTreeList;
   Delay( 300 );
   CreateConfig;
   Delay( 300 );
   dxWebServer.Caption := Format( 'WebService: %s', [FConfig.SOAPUrl] );
   FMsg.Kind := MB_ICONINFORMATION;
   FMsg.Msg := '環境配置完成。';
   MsgNotify( FMsg );
   Delay( 1000 );
  finally
    Screen.Cursor := crDefault;
  end;
  dxPlayGroup.Enabled := ( SoPageControl.PageCount > 0 );
end;

{ ---------------------------------------------------------------------------- }

procedure TMain.MsgNotify(aMsg: Pointer);

    { ------------------------------------------------ }

    function GetImageIndex(aIndex: Integer): Integer;
    begin
      case aIndex of
        MB_OK: Result := 1;
        MB_ICONERROR: Result := 2;
      else
        Result := 0;
      end;
    end;

    { ------------------------------------------------ }

var
  aItem: TListItem;
  aDelCount: Integer;
begin
  if MsgListView.Items.Count > 1000 then
  begin
    aDelCount := 500;
    MsgListView.Items.BeginUpdate;
    try
      while ( aDelCount > 0 ) do
      begin
        if MsgListView.Items.Count <= 0 then Break;          
        MsgListView.Items.Delete( 0 );
        Dec( aDelCount );
      end;
    finally
      MsgListView.Items.EndUpdate;
    end;
  end;    
  if PMsg( aMsg ).Msg = EmptyStr then Exit;
  aItem := MsgListView.Items.Add;
  aItem.Caption := Format( '%s> %s',
    [FormatDateTime( 'yyyy-mm-dd hh:nn:ss', Now ), PMsg( aMsg ).Msg] );
  aItem.ImageIndex := GetImageIndex( PMsg( aMsg ).Kind );
  aItem.Selected := True;
  aItem.MakeVisible( True );
  MsgListView.Update;
  Application.ProcessMessages;
end;

{ ---------------------------------------------------------------------------- }

procedure TMain.DbNotify(aSoInfo: TSoInfo);
begin
  case aSoInfo.DbStatus of
    dbNone: aSoInfo.DisplaySheet.ImageIndex := 0;
    dbOK: aSoInfo.DisplaySheet.ImageIndex := 2;
    dbActive: aSoInfo.DisplaySheet.ImageIndex := 1;
  end;
  aSoInfo.DisplaySheet.Update;
  Application.ProcessMessages;
end;

{ ---------------------------------------------------------------------------- }

procedure TMain.dxPlayClick(Sender: TObject);
begin
  CreateSoThread;
  dxPlayGroup.Enabled := False;
  dxStopGroup.Enabled := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TMain.dxStopClick(Sender: TObject);
begin
  ReleaseSoThread;
  dxPlayGroup.Enabled := True;
  dxStopGroup.Enabled := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TMain.RecordNotify(aSoInfo: TSoInfo; aSO311: PSO311);
var
  aIsNewRecord: Boolean;
  aNode: TcxTreeListNode;
  aColumn1: TcxTreeListColumn;
  aColumn2, aColumn3, aColumn4, aColumn5, aColumn6, aColumn7: TcxTreeListColumn;
  aColumn8, aColumn9, aColumn10, aColumn11, aColumn12: TcxTreeListColumn;
begin
  aColumn1 := aSoInfo.DisplayTreeList.ColumnByName( Format( '%s_CmdSeqNo', [aSoInfo.DisplayTreeList.Name] ) );
  aColumn2 := aSoInfo.DisplayTreeList.ColumnByName( Format( '%s_MsgCode', [aSoInfo.DisplayTreeList.Name] ) );
  aColumn3 := aSoInfo.DisplayTreeList.ColumnByName( Format( '%s_WorkId', [aSoInfo.DisplayTreeList.Name] ) );
  aColumn4 := aSoInfo.DisplayTreeList.ColumnByName( Format( '%s_CusId', [aSoInfo.DisplayTreeList.Name] ) );
  aColumn5 := aSoInfo.DisplayTreeList.ColumnByName( Format( '%s_CMMAC', [aSoInfo.DisplayTreeList.Name] ) );
  aColumn6 := aSoInfo.DisplayTreeList.ColumnByName( Format( '%s_EMTAMAC', [aSoInfo.DisplayTreeList.Name] ) );
  aColumn7 := aSoInfo.DisplayTreeList.ColumnByName( Format( '%s_OperType', [aSoInfo.DisplayTreeList.Name] ) );
  aColumn8 := aSoInfo.DisplayTreeList.ColumnByName( Format( '%s_Call', [aSoInfo.DisplayTreeList.Name] ) );
  aColumn9 := aSoInfo.DisplayTreeList.ColumnByName( Format( '%s_SOAPRequest', [aSoInfo.DisplayTreeList.Name] ) );
  aColumn10 := aSoInfo.DisplayTreeList.ColumnByName( Format( '%s_Resp', [aSoInfo.DisplayTreeList.Name] ) );
  aColumn11 := aSoInfo.DisplayTreeList.ColumnByName( Format( '%s_SOAPResponse', [aSoInfo.DisplayTreeList.Name] ) );
  aColumn12 := aSoInfo.DisplayTreeList.ColumnByName( Format( '%s_SendDateTime', [aSoInfo.DisplayTreeList.Name] ) );
  aIsNewRecord := not Assigned( aSO311.DisplayNode );
  if aIsNewRecord then
    aNode := aSoInfo.DisplayTreeList.Add
  else
    aNode := TcxTreeListNode( aSO311.DisplayNode );
  if aIsNewRecord then
  begin
    aNode.Texts[aColumn1.ItemIndex] := aSO311.CmdSeqNO;
    aNode.Texts[aColumn2.ItemIndex] := Format( '%s-%s', [aSO311.MsgCode, aSO311.MsgName] );
    aNode.Texts[aColumn3.ItemIndex] := aSO311.WorkId;
    aNode.Texts[aColumn4.ItemIndex] := aSO311.CusId;
    aNode.Texts[aColumn5.ItemIndex] := aSO311.CMMac;
    aNode.Texts[aColumn6.ItemIndex] := aSO311.EMTAMac;
    aNode.Texts[aColumn7.ItemIndex] := aSO311.OperType;
    {}
    aNode.Values[aColumn8.ItemIndex] := Null;
    aNode.Texts[aColumn9.ItemIndex] := EmptyStr;
    {}
    aNode.Values[aColumn10.ItemIndex] := Null;
    aNode.Texts[aColumn11.ItemIndex] := EmptyStr;
    {}
    aNode.Texts[aColumn12.ItemIndex] := EmptyStr;
    {}
    if ( aSo311.IsBatch ) then
    begin
      aNode.ImageIndex := 3;
      aNode.SelectedIndex := 3;
    end;
    aSO311.DisplayNode := aNode;
  end else
  begin
    if ( aSO311.ErrMsg <> EmptyStr ) then
    begin
      aNode.Values[aColumn8.ItemIndex] := '2';
      aNode.Values[aColumn8.ItemIndex] := '2';
      aNode.Texts[aColumn9.ItemIndex] := aSO311.SOAPRequest;
      aNode.Texts[aColumn11.ItemIndex] := aSO311.SOAPResponse;
      aNode.Texts[aColumn12.ItemIndex] := FormatDateTime( 'yyyy-mm-dd hh:nn:ss', Now );
      aNode.MakeVisible;
      Exit;
    end else
    begin
      if ( aSO311.SOAPRequest <> EmptyStr ) then
      begin
        aNode.Values[aColumn8.ItemIndex] := '1';
        aNode.Texts[aColumn9.ItemIndex] := aSO311.SOAPRequest;
        aNode.Texts[aColumn12.ItemIndex] := FormatDateTime( 'yyyy-mm-dd hh:nn:ss', Now );
        aNode.MakeVisible;
      end;
      if ( aSO311.SOAPResponse <> EmptyStr ) then
      begin
        aNode.Values[aColumn10.ItemIndex] := '1';
        aNode.Texts[aColumn11.ItemIndex] := aSO311.SOAPResponse;
        aNode.Texts[aColumn12.ItemIndex] := Format( '%s - %s',
          [aNode.Texts[aColumn12.ItemIndex], FormatDateTime( 'yyyy-mm-dd hh:nn:ss', Now )] );
        aNode.MakeVisible;   
      end;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
