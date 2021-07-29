unit frmCD078B3_1U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, DBClient, ADODB, Menus, StdCtrls, ExtCtrls, ActnList,
  frmCD078B1U, 
  cbDBController, cbStyleController, cbDataLookup,
  dxSkinsCore, dxSkinsDefaultPainters, cxControls,
  cxContainer, cxEdit, cxTextEdit, cxButtonEdit, dxSkinscxPCPainter,
  cxGraphics, cxDataStorage,
  cxDBData, cxGridLevel, cxGridCustomTableView, cxGridTableView,
  cxGridBandedTableView, cxGridDBBandedTableView, cxClasses, cxGridCustomView,
  cxGrid, cxPC, Buttons, cxCheckBox, cxDropDownEdit, cxCurrencyEdit,
  cxImageComboBox, cxMaskEdit, cxDataUtils, cxStyles, cxCustomData,
  cxFilter, cxData, dxSkinBlack, dxSkinBlue, dxSkinCaramel, dxSkinCoffee;

const
  COMMON_DELAY = 100;

type

  TRankInfo = record
    StepNo: String;
    Description: String;
    MasterSale: Integer;
    LinkKey: String;
    LinkDescription: String;
    Period: String;
  end;

  PRankInfo = ^TRankInfo;

  TRankAmount = class(TObject)
  private
    FSign: String;
    FDefaultPeriod: Integer;
    FDefaultDiscountAmount: Double;
    FDefaultMonthAmount: Double;
    FDefaultDayAmount: Double;
    {}
    FPeriod: Integer;
    FDiscountAmount: Double;
    FDayAmount: Double;
    FMonthAmount: Double;
    {}
    procedure SetDefaultPeriod(const AValue: Integer);
    procedure SetDefaultDiscountAmount(const AValue: Double);
    procedure SetPeriod(const AValue: Integer);
    procedure SetDiscountAmount(const AValue: Double);
    procedure SetSign(const AValue: String);
  public
    procedure ComputeMonthAndDayAmount;
    {}
    property Sign: String read FSign write SetSign;
    property DefaultPeriod: Integer read FDefaultPeriod write SetDefaultPeriod;
    property DefaultDiscountAmount: Double read FDefaultDiscountAmount write SetDefaultDiscountAmount;
    property DefaultMonthAmount: Double read FDefaultMonthAmount;
    property DefaultDayAmount: Double read FDefaultDayAmount;
    {}
    property Period: Integer read FPeriod write SetPeriod;
    property DiscountAmount: Double read FDiscountAmount write SetDiscountAmount;
    property MonthAmount: Double read FMonthAmount;
    property DayAmount: Double read FDayAmount;
  end;


  TfrmCD078B3_1 = class(TForm)
    ActionList1: TActionList;
    actSave: TAction;
    actCancel: TAction;
    Panel1: TPanel;
    edtDML: TcxTextEdit;
    btnSave: TButton;
    Button1: TButton;
    Panel2: TPanel;
    Bevel1: TBevel;
    RankPage: TcxPageControl;
    RankDataSource: TDataSource;
    RankDataSet: TClientDataSet;
    Panel3: TPanel;
    btnAddRank: TBitBtn;
    btnDelRnak: TBitBtn;
    Panel4: TPanel;
    Label30: TLabel;
    edtStepNo: TcxTextEdit;
    edtDescription: TcxTextEdit;
    chkMasterSale: TcxCheckBox;
    RankGrid: TcxGrid;
    gvRank: TcxGridDBBandedTableView;
    gvRankColumn1: TcxGridDBBandedColumn;
    gvRankColumn2: TcxGridDBBandedColumn;
    gvRankColumn3: TcxGridDBBandedColumn;
    gvRankColumn4: TcxGridDBBandedColumn;
    gvRankColumn5: TcxGridDBBandedColumn;
    gvRankColumn6: TcxGridDBBandedColumn;
    gvRankColumn7: TcxGridDBBandedColumn;
    gvRankColumn8: TcxGridDBBandedColumn;
    glRank: TcxGridLevel;
    gvRankColumn10: TcxGridDBBandedColumn;
    gvRankColumn11: TcxGridDBBandedColumn;
    LinkKey: TDataLookup;
    Label18: TLabel;
    gvRankColumn12: TcxGridDBBandedColumn;
    lbl1: TLabel;
    edtPeriod: TcxCurrencyEdit;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure actSaveExecute(Sender: TObject);
    procedure actCancelExecute(Sender: TObject);
    procedure RankDataSetAfterCancel(DataSet: TDataSet);
    procedure RankPageChange(Sender: TObject);
    procedure gvRankCustomDrawCell(Sender: TcxCustomGridTableView;
      ACanvas: TcxCanvas; AViewInfo: TcxGridTableDataCellViewInfo;
      var ADone: Boolean);
    procedure RankPageDrawTabEx(AControl: TcxCustomTabControl; ATab: TcxTab;
      Font: TFont);
    procedure chkMasterSalePropertiesChange(Sender: TObject);
    procedure gvRankEditing(Sender: TcxCustomGridTableView;
      AItem: TcxCustomGridTableItem; var AAllow: Boolean);
    procedure gvRankColumn2PropertiesChange(Sender: TObject);
    procedure gvRankColumn2PropertiesEditValueChanged(Sender: TObject);
    procedure gvRankColumn4PropertiesEditValueChanged(Sender: TObject);
    procedure gvRankColumn6PropertiesValidate(Sender: TObject;
      var DisplayValue: Variant; var ErrorText: TCaption; var Error: Boolean);
    procedure btnDelRnakClick(Sender: TObject);
    procedure edtDescriptionPropertiesChange(Sender: TObject);
    procedure btnAddRankClick(Sender: TObject);
    procedure gvRankColumn11GetCellHint(Sender: TcxCustomGridTableItem;
      ARecord: TcxCustomGridRecord; ACellViewInfo: TcxGridTableDataCellViewInfo;
      const AMousePos: TPoint; var AHintText: TCaption;
      var AIsHintMultiLine: Boolean; var AHintTextRect: TRect);
    procedure gvRankColumn11PropertiesButtonClick(Sender: TObject;
      AButtonIndex: Integer);
    procedure LinkKeyCodeNoPropertiesChange(Sender: TObject);
    procedure LinkKeyCodeNamePropertiesChange(Sender: TObject);
    procedure LinkKeyCodeNamePropertiesInitPopup(Sender: TObject);
    procedure edtPeriodPropertiesChange(Sender: TObject);
  private
    { Private declarations }
    FMode: TDML;
    FSourceDataSet: TClientDataSet;
    FMonMap: array [1..12] of Boolean;
    FLockBindArray: Boolean;
    FKeyArray: array [1..4] of String;
    FRankAmount: TRankAmount;
    FLinkKeyDataSet: TClientDataSet;
    FEnableLinkKey: Boolean;
    FLockCheck: Boolean;
    procedure ClearEditor;
    procedure DataToEditor;
    procedure DMLModeChange(aMode: TDML);
    procedure ButtonStateChange;
    procedure EditorStateChange;
    function DataSetStateChange(const aMode: TDML): Boolean;
    function DataValidate: Boolean;
    function EditorToData: Boolean; overload;
    function EditorToData(const aIndex: Integer): Boolean; overload;
    {}
    procedure InternalCalcMonth;
    procedure PrepareRankDataSet;
    procedure UnPrepareRankDataSet;
    {}
    function CompoundNullRankData(const aState: TDataSetState): String;
    procedure DeCompoundRankData;
    {}
    procedure AllocateEmptyRankInfo;
    procedure AllocateFullRankInfo;
    procedure AllocateNewRankInfo;
    {}
    procedure CleanupRankPage; overload;
    procedure CleanupRankPage(const aIndex: Integer); overload;
    procedure FilterRankData(const aStepNo: String);
    procedure RankDataBindToArray;
    procedure MakeColumnEditMode;
    procedure MoveToLastRecord;
    {}
    function VdMustInput: Boolean;
    function VdMonthBreak: Boolean;
    function VdRateAndAmt: Boolean;
    {}
    function SaveData: Boolean;
    procedure DeleteRankData(const aStepNo: String);
    procedure ReCaptionTabSheet;
    procedure InitControls;
    procedure SetEnableLinkKey(const AEnabled: Boolean);
  public
    { Public declarations }
    constructor Create(aMode: TDML; aDataSet: TClientDataSet;  aKeys: array of String); reintroduce;
    property RankAmount: TRankAmount read FRankAmount;
    property LinkKeyDataSet: TClientDataSet read FLinkKeyDataSet write FLinkKeyDataSet;
    property EnableLinkKey: Boolean read FEnableLinkKey write SetEnableLinkKey;
  end;

var
  frmCD078B3_1: TfrmCD078B3_1;

implementation

{$R *.dfm}

uses cbUtilis, frmCD078B3_2U;


{ ---------------------------------------------------------------------------- }

{ TRankAmount }

procedure TRankAmount.SetSign(const AValue: String);
begin
  if ( AValue <> EmptyStr ) and
     ( ( AValue = '+' ) or ( AValue = '-' ) ) then
  begin
    FSign := AValue;
    SetDefaultDiscountAmount( FDefaultDiscountAmount );
    ComputeMonthAndDayAmount;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TRankAmount.SetDefaultPeriod(const AValue: Integer);
begin
  FDefaultPeriod := Abs( AValue );
  ComputeMonthAndDayAmount;
end;

{ ---------------------------------------------------------------------------- }

procedure TRankAmount.SetDefaultDiscountAmount(const AValue: Double);
begin
  FDefaultDiscountAmount := Abs( AValue );
  if ( FSign = '-' ) then
    FDefaultDiscountAmount := ( 0 - FDefaultDiscountAmount );
  ComputeMonthAndDayAmount;
end;

{ ---------------------------------------------------------------------------- }

procedure TRankAmount.SetPeriod(const AValue: Integer);
begin
  FPeriod := AValue;
  ComputeMonthAndDayAmount;
end;

{ ---------------------------------------------------------------------------- }

procedure TRankAmount.SetDiscountAmount(const AValue: Double);
begin
  FDiscountAmount := AValue;
  ComputeMonthAndDayAmount;
end;

{ ---------------------------------------------------------------------------- }

procedure TRankAmount.ComputeMonthAndDayAmount;
begin
  FDefaultMonthAmount := 0;
  FDefaultDayAmount := 0;
  if ( FDefaultPeriod > 0 ) then
  begin
    FDefaultMonthAmount := CbRoundTo( ( Abs( FDefaultDiscountAmount ) / FDefaultPeriod ), 3  );
    FDefaultDayAmount := CbRoundTo( ( ( Abs( FDefaultDiscountAmount ) / FDefaultPeriod ) / 30  ), 3  );
    if ( FSign = '-' ) then
    begin
      FDefaultMonthAmount := ( 0 - FDefaultMonthAmount );
      FDefaultDayAmount := ( 0 - FDefaultDayAmount );
    end;
  end;
  {}
  FMonthAmount := 0;
  FDayAmount := 0;
  if ( FPeriod > 0 ) then
  begin
    {****************************************************************************}
    { #4329 計算方式改變 By Kin 2009/01/21}
    //FMonthAmount := CbRoundTo( ( Abs( FDiscountAmount ) / FPeriod ), 3  );
    //FDayAmount := CbRoundTo( ( ( Abs( FDiscountAmount ) / FPeriod ) / 30  ), 3  );
    FMonthAmount := CbRoundTo( ( Abs( FDiscountAmount ) / FPeriod ), 3  );
    FDayAmount := CbRoundTo( ( ( Abs( FMonthAmount ))  / 30  ), 3  );
    {****************************************************************************}
    if ( FSign = '-' ) then
    begin
      FMonthAmount := ( 0 - FMonthAmount );
      FDayAmount := ( 0 - FDayAmount );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

const
  aNumText: array [1..12] of String =
    ( '一', '二', '三', '四', '五', '六', '七', '八', '九', '十', '十一', '十二' );
    
{ TfrmCD078B3_1 }

{ ---------------------------------------------------------------------------- }

constructor TfrmCD078B3_1.Create(aMode: TDML; aDataSet: TClientDataSet; aKeys: array of String);
begin
  inherited Create( Application );
  FMode := aMode;
  FSourceDataSet := aDataSet;
  FKeyArray[1] := aKeys[0];
  FKeyArray[2] := aKeys[1];
  FKeyArray[3] := aKeys[2];
  FKeyArray[4] := aKeys[3];
  FRankAmount := TRankAmount.Create;
  FRankAmount.Period := 0;
  FRankAmount.DiscountAmount := 0;
  FLockCheck := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.FormCreate(Sender: TObject);
begin
  InitControls;
  FLockBindArray := False;
  RankGrid.Visible := False;
  FillChar( FMonMap, SizeOf( FMonMap ), Ord( False ) );
  PrepareRankDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.FormShow(Sender: TObject);
begin
  Screen.Cursor := crSQLWait;
  try
    if ( FSourceDataSet.IsEmpty ) then
    begin
      CompoundNullRankData( dsInsert );
      AllocateEmptyRankInfo;
    end else
    begin
      DeCompoundRankData;
      AllocateFullRankInfo;
    end;
    LinkKey.DataSet := FLinkKeyDataSet;
    DataToEditor;
    DMLModeChange( FMode );
    MoveToLastRecord;
    MakeColumnEditMode;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if ( FMode <> dmSave ) then
    DMLModeChange( dmCancel );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.FormDestroy(Sender: TObject);
begin
  FSourceDataSet := nil;
  UnPrepareRankDataSet;
  CleanupRankPage;
  LinkKey.Finaliza;
  FRankAmount.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.ButtonStateChange;
begin
  if ( FMode = dmInsert ) then
  begin
    actSave.Enabled := True;
    btnAddRank.Enabled := True;
    btnDelRnak.Enabled := True;
  end else
  if ( FMode = dmUpdate ) then
  begin
    actSave.Enabled := True;
    btnAddRank.Enabled := True;
    btnDelRnak.Enabled := True;
  end else
  begin
    actSave.Enabled := False;
    actCancel.Caption := '結束(&C)';
    actCancel.ShortCut := TextToShortCut( 'Alt+C' );
    btnAddRank.Enabled := False;
    btnDelRnak.Enabled := False;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.ClearEditor;
begin
  edtDescription.LockChangeEvents( True, False );
  chkMasterSale.LockChangeEvents( True, False );
  edtPeriod.LockChangeEvents(True,False);
  try
    edtDescription.Clear;
    chkMasterSale.Checked := False;
    {#6865 增加Period By Kin 2014/09/10}
    edtPeriod.Text := '';
  finally
    edtDescription.LockChangeEvents( False, False );
    chkMasterSale.LockChangeEvents( False, False );
    edtPeriod.LockChangeEvents(False,False);
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B3_1.DataSetStateChange(const aMode: TDML): Boolean;
begin
  Result := True;
  case aMode of
    dmInsert:
      begin
      end;
    dmUpdate:
      begin
      end;
    dmCancel:
      begin
        RankDataSet.Cancel;
        FSourceDataSet.Cancel;
      end;
    dmBrowser:
      begin
        ClearEditor;
        DataToEditor;
      end;
    dmSave:
      begin
        Result := DataValidate;
        if Result then
        begin
          try
            Result := SaveData;
          except
            on E: Exception do
            begin
              ErrorMsg( Format( '存檔有誤, 原因:%s。', [E.Message] ) );
              Result := False;
            end;
          end;
        end;
      end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.DataToEditor;
var
  aRankInfo: PRankInfo;
begin
  aRankInfo := PRankInfo( RankPage.Tabs.Objects[RankPage.ActivePageIndex] );
  if not Assigned( aRankInfo ) then
  begin
    New( aRankInfo );
    aRankInfo.StepNo := IntToStr( MaxInt );
    aRankInfo.Description := EmptyStr;
    aRankInfo.MasterSale := Ord( False );
    aRankInfo.LinkKey := EmptyStr;
    aRankInfo.LinkDescription := EmptySTr;
    {#6865 增加Period By Kin 2014/09/10}
    aRankInfo.Period := EmptyStr;
    RankPage.Tabs.Objects[RankPage.ActivePageIndex] := TObject( aRankInfo );
  end;
  edtStepNo.Text := EmptyStr;
  if ( aRankInfo.StepNo <> IntToStr( MaxInt ) ) then
    edtStepNo.Text := aRankInfo.StepNo;
  edtDescription.LockChangeEvents( True, False );
  chkMasterSale.LockChangeEvents( True, False );
  edtPeriod.LockChangeEvents(True,False);
  try
    edtDescription.Text := aRankInfo.Description;
    chkMasterSale.Checked := ( aRankInfo.MasterSale = 1 );
    {#6865 增加Period By Kin 2014/09/10}
    edtPeriod.Text := aRankinfo.Period;
  finally
    edtDescription.LockChangeEvents( False, False );
    chkMasterSale.LockChangeEvents( False, False );
    edtPeriod.LockChangeEvents(False,False);
  end;
  if Assigned( LinkKey.DataSet ) then
  begin
    FLockCheck := True;
    try
      LinkKey.Value := aRankInfo.LinkKey;
    finally
      FLockCheck := False;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B3_1.DataValidate: Boolean;
begin
  Screen.Cursor := crSQLWait;
  try
    Result := VdMustInput;
    if not Result then Exit;
    {}
    Result := VdMonthBreak;
    if not Result then Exit;
    {}
    Result := VdRateAndAmt;
    if not Result then Exit;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.DMLModeChange(aMode: TDML);
begin
  if not DataSetStateChange( aMode ) then Exit;
  if ( aMode in [dmSave, dmCancel] ) then
  begin
    FMode := aMode;
    Self.Close;
    Exit;
  end;
  ButtonStateChange;
  EditorStateChange;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.EditorStateChange;
begin
  case FMode of
    dmBrowser:
      begin
        edtStepNo.Enabled := False;
        edtDescription.Enabled := False;
        chkMasterSale.Enabled := False;
      end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B3_1.EditorToData: Boolean;
begin
  Result := EditorToData( RankPage.ActivePageIndex );
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B3_1.EditorToData(const aIndex: Integer): Boolean;
var
  aRankInfo: PRankInfo;
  aBookMark: TBookmark;
begin
  Result := False;
  aRankInfo := PRankInfo( RankPage.Tabs.Objects[aIndex] );
  if Assigned( aRankInfo ) then
  begin
    aBookMark := RankDataSet.GetBookmark;
    try
      RankDataSet.DisableControls;
      try
        RankDataSet.First;
        while not RankDataSet.Eof do
        begin
          RankDataSet.Edit;
          RankDataSet.FieldByName( 'StepNo' ).AsString := aRankInfo.StepNo;
          RankDataSet.FieldByName( 'Description' ).AsString := aRankInfo.Description;
          RankDataSet.FieldByName( 'MasterSale' ).AsInteger := aRankInfo.MasterSale;
          RankDataSet.FieldByName( 'MasterPeriod' ).AsString := aRankInfo.Period;
          RankDataSet.FieldByName( 'LinkKey' ).AsString := EmptyStr;
          RankDataSet.FieldByName( 'LinkKeyName' ).AsString := EmptyStr;
          if ( FEnableLinkKey ) then
          begin
            RankDataSet.FieldByName( 'LinkKey' ).AsString := aRankInfo.LinkKey;
            RankDataSet.FieldByName( 'LinkKeyName' ).AsString := aRankInfo.LinkDescription;
          end;
          RankDataSet.Post;
          RankDataSet.Next;
        end;
        RankDataSet.GotoBookmark( aBookMark );
      finally
        RankDataSet.EnableControls;
      end;
    finally
      RankDataSet.FreeBookmark( aBookMark );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.InternalCalcMonth;
var
  aSValue, aEvalue, aFdVal: Integer;
  aText: String;
  aBMark: TBookmark;
begin
  RankDataSet.DisableControls;
  try
    aBMark := RankDataSet.GetBookmark;
    try
      RankDataSet.First;
      while not RankDataSet.Eof do
      begin
        if ( RankDataSet.FieldByName( 'RankLv' ).AsInteger = 1 ) then
        begin
          aSValue := 0;
          aEValue := 0;
        end;
        aFdVal := RankDataSet.FieldByName( 'Mon' ).AsInteger;
        if ( aFdVal > 0 ) then
        begin
          aSValue := aSValue + 1;
          aEValue := aSValue + aFdVal - 1;
          aText := Format( '%d~%d', [aSValue, aEValue] );
          aSValue := aEValue;
        end else
          aText := EmptyStr;
        RankDataSet.Edit;
        RankDataSet.FieldByName( 'LblPeriod' ).AsString := aText;
        RankDataSet.Post;
        RankDataSet.Next;
      end;
      RankDataSet.GotoBookmark( aBMark );
    finally
      RankDataSet.FreeBookmark( aBMark );
    end;
  finally
    RankDataSet.EnableControls;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.LinkKeyCodeNoPropertiesChange(Sender: TObject);
var
  aRankInfo: PRankInfo;
begin
  LinkKey.CodeNoPropertiesChange( Sender );
  aRankInfo := PRankInfo( RankPage.Tabs.Objects[RankPage.ActivePageIndex] );
  if Assigned( aRankInfo ) then
  begin
    if not FLockCheck then
    begin
      aRankInfo.LinkKey := LinkKey.Value;
      aRankInfo.LinkDescription := LinkKey.ValueName;
      EditorToData( RankPage.ActivePageIndex );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.LinkKeyCodeNamePropertiesChange(Sender: TObject);
var
  aRankInfo: PRankInfo;
begin
  LinkKey.CodeNamePropertiesChange( Sender );
  aRankInfo := PRankInfo( RankPage.Tabs.Objects[RankPage.ActivePageIndex] );
  if Assigned( aRankInfo ) then
  begin
    if not FLockCheck then
    begin
      aRankInfo.LinkKey := LinkKey.Value;
      aRankInfo.LinkDescription := LinkKey.ValueName;
      EditorToData( RankPage.ActivePageIndex );
    end;
  end;                                       
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.LinkKeyCodeNamePropertiesInitPopup(Sender: TObject);
begin
  LinkKey.CodeNamePropertiesInitPopup( Sender );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.PrepareRankDataSet;
begin
  if ( RankDataSet.FieldDefs.Count <= 0 ) then
  begin
    RankDataSet.FieldDefs.Add( 'StepNo', ftInteger );
    RankDataSet.FieldDefs.Add( 'RankLv', ftInteger );
    RankDataSet.FieldDefs.Add( 'RankLvName', ftString, 10 );
    RankDataSet.FieldDefs.Add( 'Description', ftString, 50 );
    RankDataSet.FieldDefs.Add( 'MasterSale', ftInteger );
    RankDataSet.FieldDefs.Add( 'Mon', ftInteger );
    RankDataSet.FieldDefs.Add( 'LblPeriod', ftString, 20 );
    RankDataSet.FieldDefs.Add( 'Period', ftInteger );
    RankDataSet.FieldDefs.Add( 'RateType', ftInteger );
    RankDataSet.FieldDefs.Add( 'DiscountRate', ftInteger );
    RankDataSet.FieldDefs.Add( 'DiscountAmt', ftFloat );
    RankDataSet.FieldDefs.Add( 'MonthAmt', ftFloat );
    RankDataSet.FieldDefs.Add( 'DayAmt', ftFloat );
    RankDataSet.FieldDefs.Add( 'Comment', ftString, 50 );
    RankDataSet.FieldDefs.Add( 'LinkKey', ftString, 10 );
    RankDataSet.FieldDefs.Add( 'LinkKeyName', ftString, 50 );
    RankDataSet.FieldDefs.Add( 'MasterPeriod',ftString,5);
    RankDataSet.CreateDataSet;
    RankDataSet.AddIndex( 'IDX1', 'stepno;ranklv', [ixUnique] );
    RankDataSet.IndexName := 'IDX1';
  end;
  RankDataSet.EmptyDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.UnPrepareRankDataSet;
begin
  RankDataSet.EmptyDataSet;
  if ( RankDataSet.IndexName <> EmptyStr ) then
    RankDataSet.DeleteIndex( RankDataSet.IndexName );
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B3_1.VdMustInput: Boolean;
var
  aIndex: Integer;
  aHasMasterSale: Boolean;
begin
  Result := False;
  aHasMasterSale := False;
  for aIndex := 0 to RankPage.PageCount - 1 do
  begin
    RankPage.ActivePageIndex := aIndex;
    Delay( COMMON_DELAY );
    if ( edtDescription.Text = EmptyStr ) then
    begin
      WarningMsg( '請輸入【多階方案說明】。' );
      if edtDescription.CanFocusEx then edtDescription.SetFocus;
      Exit;
    end;
    {#6893 增加繳別期數的判斷 By Kin 2014/10/20}
    if ( edtPeriod.EditValue = Null ) then
    begin
      WarningMsg( '請輸入【繳別期數】。' );
      if edtPeriod.CanFocusEx then edtPeriod.SetFocus;
      Exit;
    end;
    if ( StrToInt( VarToStr( edtPeriod.EditValue ) ) <= 0  ) then
    begin
      WarningMsg( '請輸入【繳別期數需 >0】。' );
      if edtPeriod.CanFocusEx then edtPeriod.SetFocus;
      Exit;
    end;
    {}
    if ( LinkKey.Value = EmptyStr ) and ( FEnableLinkKey ) then
    begin
      WarningMsg( '請輸入【對應正項多階方案】。' );
      if LinkKey.CodeNo.CanFocusEx then LinkKey.CodeNo.SetFocus;
      Exit;
    end;
    if ( not aHasMasterSale ) and chkMasterSale.Checked then
      aHasMasterSale := chkMasterSale.Checked;
  end;
  { 若是未設定主推多階, 而且只有設定一筆 多階方案的話,
    自動把主推多階勾選 }
  if ( not aHasMasterSale ) then
  begin
    if ( RankPage.PageCount = 1  ) then
       chkMasterSale.Checked := True
    else begin
      WarningMsg( '請設定【主推多階方案】。' );
      if chkMasterSale.CanFocusEx then chkMasterSale.SetFocus;
      Exit;
    end;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B3_1.VdMonthBreak: Boolean;
var
  aIndex, aLastIdx: Integer;
  aBookMark: TBookmark;
  aMonIsEmpty: Boolean;
begin
  Result := True;
  for aIndex := 0 to RankPage.PageCount - 1 do
  begin
    FLockBindArray := True;
    try
      RankPage.ActivePageIndex := aIndex;
      Delay( COMMON_DELAY );
    finally
      FLockBindArray := False;
    end;
    RankDataSet.DisableControls;
    try
      aLastIdx := 0;
      aMonIsEmpty := True;
      RankDataSet.Last;
      while not RankDataSet.Bof do
      begin
        if ( aMonIsEmpty ) and
           ( RankDataSet.FieldByName( 'Mon' ).AsString <> EmptyStr ) then
          aMonIsEmpty := False;
        {}
        if ( aLastIdx = 0 ) and
           ( RankDataSet.FieldByName( 'Mon' ).AsString <> EmptyStr ) then
        begin
          aLastIdx := RankDataSet.FieldByName( 'RankLv' ).AsInteger;
        end else
        if ( aLastIdx > 0 ) and
           ( RankDataSet.FieldByName( 'Mon' ).AsString = EmptyStr ) then
        begin
          Result := False;
          Break;
        end;
        RankDataSet.Prior;
      end;
    finally
      RankDataSet.EnableControls;
    end;
    if not Result then Break;
  end;
  {
    因為在 PageControl.OnPageChange 裏, 把 BindArray 關掉, 必須在開起來
    不然 Grid 畫面重繪會不正確
  }
  aBookMark := RankDataSet.GetBookmark;
  try
    RankDataBindToArray;
    RankDataSet.GotoBookmark( aBookMark );
  finally
    RankDataSet.FreeBookmark( aBookMark );
  end;
  {}
  if ( not Result ) then
  begin
    WarningMsg( '多階優惠參數【使用月數】必須連續輸入, 不可中斷。' );
    if RankGrid.CanFocusEx then RankGrid.SetFocus;
    gvRank.Controller.EditingItem := gvRankColumn2;
    Exit;
  end;
  {}
  if ( aMonIsEmpty ) then
  begin
    WarningMsg( '多階優惠參數必須輸入資料, 不可有空白內容。' );
    if RankGrid.CanFocusEx then RankGrid.SetFocus;
    gvRank.Controller.EditingItem := gvRankColumn2;
    Result := False;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B3_1.VdRateAndAmt: Boolean;
var
  aIndex: Integer;
  aItem: TcxGridDBBandedColumn;
  aBookMark: TBookmark;
  aErrMsg: String;
begin
  Result := True;
  for aIndex := 0 to RankPage.PageCount - 1 do
  begin
    FLockBindArray := True;
    try
      RankPage.ActivePageIndex := aIndex;
      Delay( COMMON_DELAY );
    finally
      FLockBindArray := False;
    end;
    RankDataSet.DisableControls;
    try
      aItem := nil;
      aErrMsg := EmptyStr;
      RankDataSet.First;
      while not RankDataSet.Eof do
      begin
        if ( RankDataSet.FieldByName( 'Mon' ).AsString <> EmptyStr ) then
        begin
          {}
          if ( RankDataSet.FieldByName( 'Period' ).AsString = EmptyStr ) then
          begin
            aItem := gvRankColumn4;
            aErrMsg := '請輸入【繳費期數】。';
            Result := False;
            Break;
          end;
          {}
          if ( RankDataSet.FieldByName( 'RateType' ).AsString = EmptyStr ) then
          begin
            aItem := gvRankColumn5;
            aErrMsg := '請輸入【費率依據】。';
            Result := False;
            Break;
          end;
          {}
        end;
        RankDataSet.Next;
      end;
    finally
      RankDataSet.EnableControls;
    end;
    if not Result then Break;
  end;
  {
    因為在 PageControl.OnPageChange 裏, 把 BindArray 關掉, 必須再開起來
    不然 Grid 畫面重繪會不正確
  }
  aBookMark := RankDataSet.GetBookmark;
  try
    RankDataBindToArray;
    RankDataSet.GotoBookmark( aBookMark );
  finally
    RankDataSet.FreeBookmark( aBookMark );
  end;
  if ( not Result ) then
  begin
    WarningMsg( aErrMsg );
    if RankGrid.CanFocusEx then RankGrid.SetFocus;
    gvRank.Controller.EditingItem := aItem;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.DeCompoundRankData;
var
  aIndex: Integer;
begin
  FSourceDataSet.First;
  while not FSourceDataSet.Eof do
  begin
    for aIndex := 1 to 12 do
    begin
      RankDataSet.Append;
      if ( FSourceDataSet.FieldByName( Format( 'mon%d', [aIndex] ) ).AsString <> EmptyStr ) then
      begin
        RankDataSet.FieldByName( 'StepNo' ).AsString := FSourceDataSet.FieldByName( 'StepNo' ).AsString;
        RankDataSet.FieldByName( 'RankLv' ).AsInteger := aIndex;
        RankDataSet.FieldByName( 'RankLvName' ).AsString := aNumText[aIndex];
        RankDataSet.FieldByName( 'Description' ).AsString := FSourceDataSet.FieldByName( 'Description' ).AsString;
        RankDataSet.FieldByName( 'MasterSale' ).AsString := FSourceDataSet.FieldByName( 'MasterSale' ).AsString;
        RankDataSet.FieldByName( 'MasterPeriod' ).AsString := FSourceDataSet.FieldByName( 'Period' ).AsString;
        RankDataSet.FieldByName( 'LinkKey' ).AsString := FSourceDataSet.FieldByName( 'LinkKey' ).AsString;
        RankDataSet.FieldByName( 'LinkKeyName' ).AsString := FSourceDataSet.FieldByName( 'LinkKeyName' ).AsString;
        RankDataSet.FieldByName( 'Mon' ).AsString := FSourceDataSet.FieldByName( Format( 'Mon%d', [aIndex] ) ).AsString;
        RankDataSet.FieldByName( 'LblPeriod' ).AsString := EmptyStr;
        RankDataSet.FieldByName( 'Period' ).Value := FSourceDataSet.FieldByName( Format( 'Period%d', [aIndex] ) ).Value;
        RankDataSet.FieldByName( 'RateType' ).Value := FSourceDataSet.FieldByName( Format( 'RateType%d', [aIndex] ) ).Value;
        if ( FEnableLinkKey ) then
          RankDataSet.FieldByName( 'DiscountRate' ).AsInteger := FSourceDataSet.FieldByName( Format( 'DiscountRate%d', [aIndex] ) ).AsInteger
        else
          RankDataSet.FieldByName( 'DiscountRate' ).Value := Null;
        RankDataSet.FieldByName( 'DiscountAmt' ).Value := FSourceDataSet.FieldByName( Format( 'DiscountAmt%d', [aIndex] ) ).Value;
        RankDataSet.FieldByName( 'MonthAmt' ).Value := FSourceDataSet.FieldByName( Format( 'MonthAmt%d', [aIndex] ) ).Value;
        RankDataSet.FieldByName( 'DayAmt' ).Value := FSourceDataSet.FieldByName( Format( 'DayAmt%d', [aIndex] ) ).Value;
        RankDataSet.FieldByName( 'Comment' ).AsString := FSourceDataSet.FieldByName( Format( 'Comment%d', [aIndex] ) ).AsString;
      end else
      begin
        RankDataSet.FieldByName( 'StepNo' ).AsString := FSourceDataSet.FieldByName( 'StepNo' ).AsString;
        RankDataSet.FieldByName( 'RankLv' ).AsInteger := aIndex;
        RankDataSet.FieldByName( 'RankLvName' ).AsString := aNumText[aIndex];
        RankDataSet.FieldByName( 'Description' ).AsString := FSourceDataSet.FieldByName( 'Description' ).AsString;
        RankDataSet.FieldByName( 'MasterSale' ).AsInteger := FSourceDataSet.FieldByName( 'MasterSale' ).AsInteger;
        RankDataSet.FieldByName( 'MasterPeriod' ).AsString := FSourceDataSet.FieldByName( 'Period' ).AsString;
        RankDataSet.FieldByName( 'LinkKey' ).AsString := FSourceDataSet.FieldByName( 'LinkKey' ).AsString;
        RankDataSet.FieldByName( 'LinkKeyName' ).AsString := FSourceDataSet.FieldByName( 'LinkKeyName' ).AsString;
        RankDataSet.FieldByName( 'LblPeriod' ).AsString := EmptyStr;
        RankDataSet.FieldByName( 'Period' ).Value := Null;
        RankDataSet.FieldByName( 'RateType' ).Value := Null;
        if ( FEnableLinkKey ) then
          RankDataSet.FieldByName( 'DiscountRate' ).AsInteger := 0
        else
          RankDataSet.FieldByName( 'DiscountRate' ).Value := Null;
        RankDataSet.FieldByName( 'DiscountAmt' ).Value := 0;
        RankDataSet.FieldByName( 'MonthAmt' ).Value := 0;
        RankDataSet.FieldByName( 'DayAmt' ).Value := 0;
        RankDataSet.FieldByName( 'Comment' ).AsString := EmptyStr;
      end;
      RankDataSet.Post;
    end;
    FSourceDataSet.Next;
  end;
  InternalCalcMonth;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B3_1.CompoundNullRankData(const aState: TDataSetState): String;
var
  aIndex: Integer;
  aStepNo: String;
begin
  aStepNo := DBController.GetStepNo;
  for aIndex := 1 to 12 do
  begin
    if ( dsInsert = aState ) then
      RankDataSet.Append
    else
      RankDataSet.Edit;
    RankDataSet.FieldByName( 'stepno' ).AsString := aStepNo;
    RankDataSet.FieldByName( 'ranklv' ).AsInteger := aIndex;
    RankDataSet.FieldByName( 'ranklvname' ).AsString := aNumText[aIndex];
    RankDataSet.FieldByName( 'description' ).AsString := EmptyStr;
    RankDataSet.FieldByName( 'MasterPeriod' ).AsString := EmptyStr;
    RankDataSet.FieldByName( 'MasterSale' ).AsInteger := 0;
    RankDataSet.FieldByName( 'Mon' ).Value := Null;
    RankDataSet.FieldByName( 'Period' ).Value := Null;
    RankDataSet.FieldByName( 'LblPeriod' ).AsString := EmptyStr;
    RankDataSet.FieldByName( 'RateType' ).Value := Null;
    RankDataSet.FieldByName( 'DiscountRate' ).Value := Null;
    RankDataSet.FieldByName( 'DiscountAmt' ).Value := 0;
    RankDataSet.FieldByName( 'MonthAmt' ).Value := 0;
    RankDataSet.FieldByName( 'DayAmt' ).Value := 0;
    RankDataSet.FieldByName( 'Comment' ).AsString := EmptyStr;
    RankDataSet.Post;
  end;
  Result := aStepNo;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.CleanupRankPage;
var
  aIndex: Integer;
begin
  RankPage.OnChange := nil;
  try
    for aIndex := RankPage.TabCount - 1 downto 0 do
      CleanupRankPage( aIndex );
  finally
    RankPage.OnChange := RankPageChange;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.CleanupRankPage(const aIndex: Integer);
var
  aTab: TcxTabSheet;
begin
  aTab := RankPage.Pages[aIndex];
  if Assigned( RankPage.Tabs.Objects[aIndex] ) then
  begin
    Dispose( PRankInfo( RankPage.Tabs.Objects[aIndex] ) );
    RankPage.Tabs.Objects[aIndex] := nil;
  end;
  aTab.PageControl := nil;
  RankGrid.Parent := RankPage.ActivePage;
  aTab.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.AllocateEmptyRankInfo;
var
  aRankInfo: PRankInfo;
  aCaption: String;
  aTab: TcxTabSheet;
begin
  RankDataSet.DisableControls;
  try
    aCaption := Format( '方案:%d', [1] );
    {}
    if ( RankPage.ActivePageIndex >= 0 ) then
    begin
      aRankInfo := PRankInfo( RankPage.Tabs.Objects[RankPage.ActivePageIndex] );
      Dispose( aRankInfo );
    end;
    New( aRankInfo );
    aRankInfo.StepNo := RankDataSet.FieldByName( 'stepno' ).AsString;
    aRankInfo.Description := EmptyStr;
    aRankInfo.MasterSale := 0;
    aRankInfo.LinkKey := EmptyStr;
    aRankInfo.LinkDescription := EmptyStr;
    {#6865 增加Period By Kin 2014/09/10}
    aRankInfo.Period := EmptyStr;
    {}
    if ( RankPage.PageCount <= 0) then
    begin
      aTab := TcxTabSheet.Create( RankPage );
      aTab.Name := Format( 'Rank%d', [RankDataSet.FieldByName( 'stepno' ).AsInteger] );
      aTab.Caption := aCaption;
      RankPage.OnChange := nil;
      try
        aTab.PageControl := RankPage;
      finally
        RankPage.OnChange := RankPageChange;
      end;
      RankPage.Tabs.Objects[aTab.PageIndex] := TObject( aRankInfo );
      RankPage.ActivePageIndex := 0;
      RankGrid.Parent := RankPage.ActivePage;
      RankGrid.Align := alClient;
      RankGrid.Visible := True;
    end else
    begin
      RankPage.Tabs.Objects[RankPage.ActivePageIndex] := TObject( aRankInfo );
      RankPage.ActivePage.Caption := aCaption;
      RankPageChange( RankPage );
    end;
  finally
    RankDataSet.EnableControls;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.AllocateFullRankInfo;
var
  aFindIdx, aMasterSalePageIdx, aPageNum: Integer;
  aTab: TcxTabSheet;
  aCaption: String;
  aRankInfo: PRankInfo;
begin
  aMasterSalePageIdx := 0;
  aPageNum := 0;
  FSourceDataSet.DisableControls;
  try
    FSourceDataSet.First;
    while not FSourceDataSet.Eof do
    begin
      Inc( aPageNum );
      aCaption := Format( '方案:%d', [aPageNum] );
      aFindIdx := RankPage.Tabs.IndexOf( aCaption );
      if ( aFindIdx < 0 ) then
      begin
        {}
        New( aRankInfo );
        aRankInfo.StepNo := FSourceDataSet.FieldByName( 'StepNo' ).AsString;
        aRankInfo.Description := FSourceDataSet.FieldByName( 'Description' ).AsString;
        aRankInfo.MasterSale := FSourceDataSet.FieldByName( 'MasterSale' ).AsInteger;
        aRankInfo.LinkKey := FSourceDataSet.FieldByName( 'LinkKey' ).AsString;
        aRankInfo.LinkDescription:= FSourceDataSet.FieldByName( 'LinKkeyName' ).AsString;
        {#6865 增加Period By Kin 2014/09/10}
        aRankInfo.Period := FSourceDataSet.FieldByName( 'Period' ).AsString;
        {}
        aTab := TcxTabSheet.Create( RankPage );
        aTab.Name := Format( 'Rank%d', [FSourceDataSet.FieldByName( 'stepno' ).AsInteger] );
        aTab.Caption := aCaption;
        RankPage.OnChange := nil;
        try
          aTab.PageControl := RankPage;
        finally
          RankPage.OnChange := RankPageChange;
        end;
        RankPage.Tabs.Objects[aTab.PageIndex] := TObject( aRankInfo );
        {}
        if ( aRankInfo.MasterSale = 1 ) then
          aMasterSalePageIdx := aTab.PageIndex;
      end;
      FSourceDataSet.Next;
    end;
    if ( RankPage.ActivePageIndex <> aMasterSalePageIdx ) then
      RankPage.ActivePageIndex := aMasterSalePageIdx
    else
      RankPageChange( RankPage );
    RankGrid.Parent := RankPage.ActivePage;
    RankGrid.Align := alClient;
    RankGrid.Visible := True;
  finally
    FSourceDataSet.EnableControls;
  end;

end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.AllocateNewRankInfo;
var
  aRankInfo : PRankInfo;
  aCaption: String;
  aTab: TcxTabSheet;
begin
  RankDataSet.DisableControls;
  try
    RankDataSet.First;
    {}
    aCaption := Format( '方案:%d', [RankPage.PageCount+1] );
    New( aRankInfo );
    aRankInfo.StepNo := RankDataSet.FieldByName( 'stepno' ).AsString;
    aRankInfo.Description := EmptyStr;
    aRankInfo.MasterSale := 0;
    aRankInfo.LinkKey := EmptyStr;
    aRankInfo.LinkDescription := EmptyStr;
    {#6865 增加Period By Kin 2014/09/10}
    aRankInfo.Period := EmptyStr;
    {}
    aTab := TcxTabSheet.Create( RankPage );
    aTab.Name := Format( 'Rank%d', [RankDataSet.FieldByName( 'stepno' ).AsInteger] );
    aTab.Caption := aCaption;
    RankPage.OnChange := nil;
    try
      aTab.PageControl := RankPage;
    finally
      RankPage.OnChange := RankPageChange;
    end;
    RankPage.Tabs.Objects[aTab.PageIndex] := TObject( aRankInfo );
    RankPage.ActivePage := aTab;
    RankGrid.Parent := RankPage.ActivePage;
    RankGrid.Align := alClient;
    RankGrid.Visible := True;
  finally
    RankDataSet.EnableControls;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.FilterRankData(const aStepNo: String);
begin
  RankDataSet.DisableControls;
  try
    RankDataSet.Filtered := False;
    RankDataSet.Filter := Format( 'stepno=''%s''', [aStepNo] );
    RankDataSet.Filtered := True;
  finally
    RankDataSet.EnableControls;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.RankDataBindToArray;
begin
  FillChar( FMonMap, SizeOf( FMonMap ), Ord( False ) );
  RankDataSet.DisableControls;
  try
    RankDataSet.First;
    while not RankDataSet.Eof do
    begin
      FMonMap[RankDataSet.FieldByName( 'RankLv' ).AsInteger] :=
        ( RankDataSet.FieldByName( 'Mon' ).AsString <> EmptyStr );
      RankDataSet.Next;
    end;
  finally
    RankDataSet.EnableControls;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.MakeColumnEditMode;
begin
  if ( RankGrid.CanFocusEx ) then RankGrid.SetFocus;
  gvRank.Controller.EditingItem := gvRankColumn2;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.MoveToLastRecord;
var
  aIndex: Integer;
begin
  for aIndex := Low( FMonMap ) to High( FMonMap ) - 1 do
  begin
    if ( not FMonMap[aIndex] ) then
    begin
      RankDataSet.First;
      RankDataSet.MoveBy( aIndex - 1 );
      Break;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.RankDataSetAfterCancel(DataSet: TDataSet);
begin
  FMonMap[RankDataSet.FieldByName( 'RankLv' ).AsInteger] :=
    ( RankDataSet.FieldByName( 'Mon' ).AsString <> EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.RankPageDrawTabEx(AControl: TcxCustomTabControl;
  ATab: TcxTab; Font: TFont);
var
  aRankInfo: PRankInfo;
begin
  aRankInfo := PRankInfo( RankPage.Tabs.Objects[ATab.Index] );
  if Assigned( aRankInfo ) then
  begin
    if ( aRankInfo.MasterSale = 1 ) then
      Font.Style := ( Font.Style + [fsBold] );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.RankPageChange(Sender: TObject);
var
  aRankInfo: PRankInfo;
begin
  Screen.Cursor := crSQLWait;
  try
    aRankInfo := PRankInfo( RankPage.Tabs.Objects[RankPage.ActivePageIndex] );
    if Assigned( aRankInfo ) then
    begin
      FilterRankData( aRankInfo.StepNo );
      RankGrid.Parent := RankPage.ActivePage;
      DataToEditor;
    end;
    if not FLockBindArray then RankDataBindToArray;
    MoveToLastRecord;
    MakeColumnEditMode;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.gvRankCustomDrawCell(Sender: TcxCustomGridTableView;
  ACanvas: TcxCanvas; AViewInfo: TcxGridTableDataCellViewInfo;
  var ADone: Boolean);
var
  aItem: TcxGridDBBandedColumn;
  aGridRecord: TcxGridDataRow;
  aDrawDisable: Boolean;
  aIndex, aDiscountRate: Integer;
begin

  aItem := TcxGridDBBandedColumn( AViewInfo.Item );
  aGridRecord := TcxGridDataRow( AViewInfo.GridRecord );
  aIndex := VarAsType( aGridRecord.Values[gvRankColumn10.Index], varInteger );
  aDiscountRate := StrToInt( VarToStrDef( aGridRecord.Values[gvRankColumn12.Index], '0' ) );
  {}
  aDrawDisable :=
    ( aItem = gvRankColumn1 ) or ( aItem = gvRankColumn3 ) or
    ( ( aItem <> gvRankColumn2 ) and ( FMonMap[aIndex] = False ) ) or
    ( ( aItem = gvRankColumn12 ) and ( not FEnableLinkKey ) ) or
    ( ( aItem = gvRankColumn6 ) and ( ( FEnableLinkKey and ( aDiscountRate > 0 ) ) ) );
  {}
  if ( aDrawDisable ) then
  begin
    ACanvas.Brush.Color := clBtnFace;
    ACanvas.Font.Color := clDefault;
  end else
  begin
    ACanvas.Brush.Color := clWindow;
    ACanvas.Font.Color := clDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.gvRankColumn2PropertiesChange(Sender: TObject);
var
  aInnerEdit: TcxMaskEdit;
  aIndex: Integer;
begin
  aInnerEdit := TcxMaskEdit( gvRank.Controller.EditingController.Edit );
  aIndex := RankDataSet.FieldByName( 'RankLv' ).AsInteger;
  FMonMap[aIndex] := ( aInnerEdit.EditingText <> EmptyStr );
  if ( not FMonMap[aIndex] ) then aInnerEdit.PostEditValue;
  gvRank.Invalidate( True );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.gvRankColumn2PropertiesEditValueChanged(
  Sender: TObject);
var
  aInnerEdit: TcxMaskEdit;
  aIndex: Integer;
begin
  aInnerEdit := TcxMaskEdit( Sender );
  aInnerEdit.PostEditValue;
  aIndex := RankDataSet.FieldByName( 'RankLv' ).AsInteger;
  FMonMap[aIndex] := ( aInnerEdit.EditingText <> EmptyStr );
  { 若該多階方案代碼, 非負項, 則不使用折扣率, 用FEnab\leLinkKey 來判斷 }
  if ( FEnableLinkKey ) then
    gvRank.DataController.SetEditValue( gvRankColumn12.Index, '0', evsValue )
  else
    gvRank.DataController.SetEditValue( gvRankColumn12.Index, Null, evsValue );
  { 該筆資料是清除 }
  if ( not FMonMap[aIndex] ) then
  begin

    gvRank.DataController.SetEditValue( gvRankColumn3.Index, Null, evsValue );
    gvRank.DataController.SetEditValue( gvRankColumn4.Index, Null, evsValue );
    gvRank.DataController.SetEditValue( gvRankColumn5.Index, Null, evsValue );
    gvRank.DataController.SetEditValue( gvRankColumn6.Index, 0, evsValue );
    gvRank.DataController.SetEditValue( gvRankColumn7.Index, 0, evsValue );
    gvRank.DataController.SetEditValue( gvRankColumn8.Index, 0, evsValue );
  end else
  begin
    if ( RankDataSet.FieldByName( 'Period' ).AsInteger = 0 ) then
      RankDataSet.FieldByName( 'Period' ).AsInteger := FRankAmount.DefaultPeriod;
    if ( RankDataSet.FieldByName( 'DiscountAmt' ).AsFloat = 0 ) then
      RankDataSet.FieldByName( 'DiscountAmt' ).AsFloat := FRankAmount.DefaultDiscountAmount;
    if ( RankDataSet.FieldByName( 'MonthAmt' ).AsFloat = 0 ) then
      RankDataSet.FieldByName( 'MonthAmt' ).AsFloat := FRankAmount.DefaultMonthAmount;
    if ( RankDataSet.FieldByName( 'DayAmt' ).AsFloat = 0 ) then
      RankDataSet.FieldByName( 'DayAmt' ).AsFloat := FRankAmount.DefaultDayAmount;
    if ( RankDataSet.FieldByName( 'RateType' ).AsString = EmptyStr ) then
      RankDataSet.FieldByName( 'RateType' ).AsString := '2';
    InternalCalcMonth;
  end;
  gvRank.Invalidate( True );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.gvRankColumn4PropertiesEditValueChanged(Sender: TObject);
var
  aInnerEdit: TcxMaskEdit;
  aCInnerEdit: TcxCurrencyEdit;
  aPeriod, aDiscountRate: Integer;
  aDiscountAmount: Double;

begin
  if ( gvRank.Controller.EditingItem = gvRankColumn4 ) then
  begin
    aInnerEdit := TcxMaskEdit( Sender );
    aPeriod := 0;
    if VarToStrDef( aInnerEdit.EditValue, EmptyStr ) <> EmptyStr then
      aPeriod := VarAsType( aInnerEdit.EditValue , varInteger );
    aInnerEdit.LockChangeEvents( True, False );
    try
      aInnerEdit.EditValue := aPeriod;
      aInnerEdit.ModifiedAfterEnter := True;
      aInnerEdit.PostEditValue;
    finally
      aInnerEdit.LockChangeEvents( False, False );
    end;
    FRankAmount.Period := aPeriod;
    if gvRankColumn12.EditValue = Null then
      aDiscountRate := 0
    else
      aDiscountRate := VarAsType( gvRankColumn12.EditValue,varInteger );
    if ( aDiscountRate )>0 then
    begin
      FRankAmount.DiscountAmount := CbRoundTo(( FRankAmount.DefaultDiscountAmount * ( aDiscountRate / 100 ) )
                                              * (aPeriod) / (FRankAmount.FDefaultPeriod) ,3) ;

    end else
    begin
      if (FRankAmount.DefaultDiscountAmount=0) or (FRankAmount.DefaultPeriod=0) then
      begin
        FRankAmount.DiscountAmount :=0;
      end else
      begin
        FRankAmount.DiscountAmount :=CbRoundTo(( FRankAmount.DefaultDiscountAmount ) *( aPeriod )
                                                /( FRankAmount.DefaultPeriod ) ,3)  ;
      end;
    end;
    gvRank.DataController.SetEditValue( gvRankColumn6.Index, FRankAmount.DiscountAmount, evsValue );

//    FRankAmount.DiscountAmount := VarAsType( gvRankColumn6.EditValue, varDouble );
  end else
  { 當為負項時, 用折扣率連動計算優惠金額, 單月/日金額 }
  if ( gvRank.Controller.EditingItem = gvRankColumn12 )    then
  begin
    aCInnerEdit := TcxCurrencyEdit( Sender );
    aDiscountRate := VarAsType( aCInnerEdit.EditValue, varInteger );
    {}
    aCInnerEdit.LockChangeEvents( True, False );
    try
      aCInnerEdit.ModifiedAfterEnter := True;
      aCInnerEdit.PostEditValue;
    finally
      aCInnerEdit.LockChangeEvents( False, False );
    end;
    {}
    if ( aDiscountRate > 0 ) then
    begin
      FRankAmount.Period := VarAsType( gvRankColumn4.EditValue, varInteger );
      //************************************************************************************
      //#4329 計算方式改變 By Kin 2009/01/21
      //FRankAmount.DiscountAmount := ( FRankAmount.DefaultDiscountAmount * ( aDiscountRate / 100 ) );
      FRankAmount.DiscountAmount := CbRoundTo(( FRankAmount.DefaultDiscountAmount * ( aDiscountRate / 100 ) )
                                              * (FRankAmount.Period) / (FRankAmount.FDefaultPeriod) ,3) ;

      //************************************************************************************
      gvRank.DataController.SetEditValue( gvRankColumn6.Index, FRankAmount.DiscountAmount, evsValue );
    end else
    begin
      FRankAmount.Period := VarAsType( gvRankColumn4.EditValue, varInteger );
      //**************************************************************************
      //#4329 計算方式改變 By Kin 2009/01/21
      //FRankAmount.DiscountAmount := FRankAmount.DefaultDiscountAmount;
      FRankAmount.DiscountAmount :=CbRoundTo(( FRankAmount.DefaultDiscountAmount ) *( FRankAmount.Period )
                                              /( FRankAmount.DefaultPeriod ) ,3)  ;
      //***************************************************************************
      gvRank.DataController.SetEditValue( gvRankColumn6.Index, FRankAmount.DiscountAmount, evsValue );
    end;
  end else
  { 修改優惠金額, 連動計算單月/日金額 }
  if ( gvRank.Controller.EditingItem = gvRankColumn6 ) then
  begin
    aInnerEdit := TcxMaskEdit( Sender );
    aDiscountAmount := 0;
    if VarToStrDef( aInnerEdit.EditValue, EmptyStr ) <> EmptyStr then
      aDiscountAmount := VarAsType( aInnerEdit.EditValue, varDouble );
    if ( FRankAmount.Sign = '-' ) then
    begin
      if aDiscountAmount > 0 then
        aDiscountAmount := ( 0 - aDiscountAmount )
    end;

    { 避免再次觸發 OnEditValueChanged }
    aInnerEdit.LockChangeEvents( True, False );
    try
      aInnerEdit.EditValue := aDiscountAmount;
      aInnerEdit.ModifiedAfterEnter := True;
      aInnerEdit.PostEditValue;
    finally
      aInnerEdit.LockChangeEvents( False, False );
    end;
    FRankAmount.Period := VarAsType( gvRankColumn4.EditValue, varInteger );
    FRankAmount.DiscountAmount := aDiscountAmount;
  end;
  gvRank.DataController.SetEditValue( gvRankColumn7.Index, FRankAmount.MonthAmount, evsValue );
  gvRank.DataController.SetEditValue( gvRankColumn8.Index, FRankAmount.DayAmount, evsValue );

end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.gvRankColumn6PropertiesValidate(Sender: TObject;
  var DisplayValue: Variant; var ErrorText: TCaption; var Error: Boolean);
begin
  if not Error then Exit;
  if VarToStrDef( DisplayValue, EmptyStr ) = EmptyStr then Exit;
  if ( TcxCurrencyEdit( Sender ).Properties.DecimalPlaces = 0 ) then
  begin
     ErrorText := Format( '輸入值有誤, 值必須介於%d~%d',
      [Trunc( TcxCurrencyEdit( Sender ).Properties.MinValue ),
       Trunc( TcxCurrencyEdit( Sender ).Properties.MaxValue )] );
  end else
  begin
     ErrorText := Format( '輸入值有誤, 值必須介於%10.3f~%10.3f',
      [TcxCurrencyEdit( Sender ).Properties.MinValue,
       TcxCurrencyEdit( Sender ).Properties.MaxValue] );
  end;
  WarningMsg( ErrorText );
  Error := False;
  ErrorText := EmptyStr;
  TcxCurrencyEdit( Sender ).SelectAll;
  Abort;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.gvRankEditing(Sender: TcxCustomGridTableView;
  AItem: TcxCustomGridTableItem; var AAllow: Boolean);
var
  aDiscountRate: Integer;
begin
  if ( AItem = gvRankColumn12 ) then
  begin
    AAllow := FEnableLinkKey;
  end else
  if ( AItem = gvRankColumn6 ) then
  begin
    aDiscountRate := StrToInt( VarToStrDef( gvRank.Controller.FocusedRecord.Values[gvRankColumn12.Index], '0' ) );
    AAllow := not FEnableLinkKey;
    if ( not AAllow ) then
      AAllow := ( aDiscountRate <= 0 );
  end else
  if ( AItem <> gvRankColumn2  ) and
     ( AItem <> gvRankColumn10 ) then
  begin
    AAllow := FMonMap[RankDataSet.FieldByName( 'RankLv' ).AsInteger];
  end;

end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.edtDescriptionPropertiesChange(Sender: TObject);
var
  aRankInfo: PRankInfo;
begin
  aRankInfo := PRankInfo( RankPage.Tabs.Objects[RankPage.ActivePageIndex] );
  if Assigned( aRankInfo ) then
  begin
    aRankInfo.Description := edtDescription.Text;
    EditorToData( RankPage.ActivePageIndex );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.chkMasterSalePropertiesChange(Sender: TObject);
var
  aRankInfo: PRankInfo;
  aIndex: Integer;
  aOldStepNo: String;
begin
  aOldStepNo := EmptyStr;
  for aIndex := 0 to RankPage.PageCount - 1 do
  begin
    aRankInfo := PRankInfo( RankPage.Tabs.Objects[aIndex] );
    if Assigned( aRankInfo ) then
      aRankInfo.MasterSale := 0;
    if ( RankPage.Pages[aIndex] = RankPage.ActivePage ) then
    begin
      aRankInfo.MasterSale := Ord( chkMasterSale.Checked );
      aOldStepNo := aRankInfo.StepNo;
    end;
    Screen.Cursor := crSQLWait;
    try
      FilterRankData( aRankInfo.StepNo );
      EditorToData( aIndex );
    finally
      Screen.Cursor := crDefault;
    end;
  end;
  FilterRankData( aOldStepNo );
  RankPage.Invalidate;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.actCancelExecute(Sender: TObject);
begin
  DMLModeChange( dmCancel );
  Self.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.actSaveExecute(Sender: TObject);
begin
  edtDescription.PostEditValue;
  chkMasterSale.PostEditValue;
  edtPeriod.PostEditValue;
  gvRank.DataController.PostEditingData;
  DMLModeChange( dmSave );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.DeleteRankData(const aStepNo: String);
var
  aIsDeleteData: Boolean;
  aNewStepNo: String;
begin
  aIsDeleteData := ( RankPage.PageCount > 1 );
  RankDataSet.DisableControls;
  try
    RankDataSet.First;
    if ( aIsDeleteData ) then
    begin
      while not RankDataSet.Eof do
      begin
        if ( RankDataSet.FieldByName( 'StepNo' ).AsString = aStepNo ) then
          RankDataSet.Delete
        else
          RankDataSet.Next;
      end;
    end else
    begin
      aNewStepNo := CompoundNullRankData( dsEdit );
      FilterRankData( aNewStepNo );
    end;
  finally
    RankDataSet.EnableControls;
  end;
  if aIsDeleteData then
  begin
    CleanupRankPage( RankPage.ActivePageIndex );
    ReCaptionTabSheet;
  end else
    AllocateEmptyRankInfo;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.btnDelRnakClick(Sender: TObject);
var
  aRankInfo: PRankInfo;
begin
  aRankInfo := PRankInfo( RankPage.Tabs.Objects[RankPage.ActivePageIndex] );
  if Assigned( aRankInfo ) then
  begin
    Screen.Cursor := crSQLWait;
    try
      DeleteRankData( aRankInfo.StepNo );
    finally
      Screen.Cursor := crDefault;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.btnAddRankClick(Sender: TObject);
var
  aNewStepNo: String;
begin
  Screen.Cursor := crSQLWait;
  try
    RankDataSet.DisableControls;
    try
      aNewStepNo := CompoundNullRankData( dsInsert );
      FilterRankData( aNewStepNo );
      AllocateNewRankInfo;
    finally
      RankDataSet.EnableControls;
    end;
    MakeColumnEditMode;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B3_1.SaveData: Boolean;


    { --------------------------------------------------- }

    function WantToDelete(const aStepNo: String): Boolean;
    var
      aRankInfo: PRankInfo;
      aIndex: Integer;
    begin
      Result := True;
      for aIndex := 0 to RankPage.PageCount - 1 do
      begin
        aRankInfo := PRankInfo( RankPage.Tabs.Objects[aIndex] );
        if not Assigned( aRankInfo ) then Continue;
        if ( aRankInfo.StepNo = aStepNo ) then
        begin
          Result := False;
          Break;
        end;
      end;
    end;

    { --------------------------------------------------- }

var
  aSaveStepNo: String;
  aIndex: Integer;
  aAlreadyExists: Boolean;
begin
  { 先做刪除的資料 }
  RankPage.OnChange := nil;
  try
    FSourceDataSet.First;
    while not FSourceDataSet.Eof do
    begin
      if WantToDelete( FSourceDataSet.FieldByName( 'StepNo' ).AsString ) then
        FSourceDataSet.Delete
      else
        FSourceDataSet.Next;
    end;
  finally
    RankPage.OnChange := RankPageChange;
  end;
  RankDataSet.DisableControls;
  try
    aSaveStepNo := RankDataSet.FieldByName( 'StepNo' ).AsString;
    RankDataSet.Filtered := False;
    RankDataSet.First;
    while not RankDataSet.Eof do
    begin
      
      aAlreadyExists :=
        ( FSourceDataSet.Locate( 'StepNo', RankDataSet.FieldByName( 'StepNo' ).AsString, [] ) );
      if aAlreadyExists then
        FSourceDataSet.Edit
      else
        FSourceDataSet.Append;
      FSourceDataSet.FieldByName( 'BPCode' ).AsString := FKeyArray[1];
      FSourceDataSet.FieldByName( 'CitemCode' ).AsString := FKeyArray[2];
      FSourceDataSet.FieldByName( 'ServiceType' ).AsString := FKeyArray[3];
      FSourceDataSet.FieldByName( 'FaciItem' ).AsString := FKeyArray[4];

      {}
      for aIndex := 1 to 12 do
      begin
        FSourceDataSet.FieldByName( 'StepNo' ).AsString := RankDataSet.FieldByName( 'StepNo' ).AsString;
        FSourceDataSet.FieldByName( 'Description' ).AsString := RankDataSet.FieldByName( 'Description' ).AsString;
        FSourceDataSet.FieldByName( 'MasterSale' ).AsInteger := RankDataSet.FieldByName( 'MasterSale' ).AsInteger;
        FSourceDataSet.FieldByName( 'LinkKey' ).AsString := EmptyStr;
        FSourceDataSet.FieldByName( 'LinkKeyName' ).AsString := EmptyStr;
        FSourceDataSet.FieldByName( 'Period' ).AsString := RankDataSet.FieldByName( 'MasterPeriod' ).AsString;
        if ( FEnableLinkKey ) then
        begin
          FSourceDataSet.FieldByName( 'LinkKey' ).AsString := RankDataSet.FieldByName( 'LinkKey' ).AsString;
          FSourceDataSet.FieldByName( 'LinkKeyName' ).AsString := RankDataSet.FieldByName( 'LinkKeyName' ).AsString;
          FSourceDataSet.FieldByName( Format( 'DiscountRate%d', [aIndex] ) ).AsInteger := RankDataSet.FieldByName( 'DiscountRate' ).AsInteger;
        end;
        FSourceDataSet.FieldByName( Format( 'Mon%d', [aIndex] ) ).Value := RankDataSet.FieldByName( 'Mon' ).Value;
        FSourceDataSet.FieldByName( Format( 'Period%d', [aIndex] ) ).Value := RankDataSet.FieldByName( 'Period' ).Value;
        FSourceDataSet.FieldByName( Format( 'RateType%d', [aIndex] ) ).Value := RankDataSet.FieldByName( 'RateType' ).Value;
        FSourceDataSet.FieldByName( Format( 'DiscountAmt%d', [aIndex] ) ).Value := RankDataSet.FieldByName( 'DiscountAmt' ).Value;
        FSourceDataSet.FieldByName( Format( 'MonthAmt%d', [aIndex] ) ).Value := RankDataSet.FieldByName( 'MonthAmt' ).Value;
        FSourceDataSet.FieldByName( Format( 'DayAmt%d', [aIndex] ) ).Value := RankDataSet.FieldByName( 'DayAmt' ).Value;
        FSourceDataSet.FieldByName( Format( 'Comment%d', [aIndex] ) ).AsString := RankDataSet.FieldByName( 'Comment' ).AsString;
        RankDataSet.Next;
      end;
      FSourceDataSet.Post;
      {}
    end;
  finally
    RankDataSet.EnableControls;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.ReCaptionTabSheet;
var
  aIndex: Integer;
  aCaption: String;
begin
  for aIndex := 0 to RankPage.PageCount - 1 do
  begin
    aCaption := Format( '方案:%d', [aIndex+1] );
    RankPage.Pages[aIndex].Caption := aCaption;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.InitControls;
begin
  TcxButtonEditProperties( gvRankColumn11.Properties ).Buttons[0].Width :=
    gvRankColumn11.Width - 3;
  LinkKey.CodeName.Properties.ListFieldNames := 'STEPNO;DESCRIPTION;PERIOD1;DISCOUNTAMT1;MON1';
  LinkKey.CodeName.Properties.KeyFieldNames := 'STEPNO';
  LinkKey.CodeName.Properties.ListFieldIndex := 1;
  LinkKey.Initializa;
  LinkKey.CodeName.Properties.ListColumns[0].Caption := '代碼';
  LinkKey.CodeName.Properties.ListColumns[1].Caption := '說明';
  LinkKey.CodeName.Properties.ListColumns[2].Caption := '期數';
  LinkKey.CodeName.Properties.ListColumns[2].Width := 100;
  LinkKey.CodeName.Properties.ListColumns[3].Caption := '優惠金額';
  LinkKey.CodeName.Properties.ListColumns[3].Width := 100;
  LinkKey.CodeName.Properties.ListColumns[4].Caption := '月數';
  LinkKey.CodeName.Properties.ListColumns[4].Width := 100;
  LinkKey.CodeName.Properties.ListOptions.ShowHeader := True;
  LinkKey.CodeName.Properties.DropDownWidth := 500;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.SetEnableLinkKey(const AEnabled: Boolean);
begin
  FEnableLinkKey := AEnabled;
  LinkKey.Enabled := FEnableLinkKey;
  Label18.Enabled := FEnableLinkKey;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.gvRankColumn11GetCellHint(
  Sender: TcxCustomGridTableItem; ARecord: TcxCustomGridRecord;
  ACellViewInfo: TcxGridTableDataCellViewInfo; const AMousePos: TPoint;
  var AHintText: TCaption; var AIsHintMultiLine: Boolean;
  var AHintTextRect: TRect);
begin
  AHintText := VarToStrDef( ARecord.Values[gvRankColumn11.Index], EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.gvRankColumn11PropertiesButtonClick(Sender: TObject;
  AButtonIndex: Integer);
var
  aData: String;
  aRow: TcxGridDataRow;
begin
  aData := EmptyStr;
  if ( gvRank.Controller.FocusedRowIndex >= 0 ) then
  begin
    aRow := TcxGridDataRow( gvRank.ViewData.Records[gvRank.Controller.FocusedRowIndex] );
    aData := VarToStrDef( aRow.Values[gvRankColumn11.Index], EmptyStr );
  end;
  {}
  frmCD078B3_2 := TfrmCD078B3_2.Create( nil );
  try
    frmCD078B3_2.CommentText := aData;
    if ( frmCD078B3_2.ShowModal = mrOk ) then
    begin
      RankDataSet.Edit;
      RankDataSet.FieldByName( 'Comment' ).AsString := frmCD078B3_2.CommentText;
      RankDataSet.Post;
    end;
  finally
    frmCD078B3_2.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_1.edtPeriodPropertiesChange(Sender: TObject);
var
  aRankInfo: PRankInfo;
begin
  aRankInfo := PRankInfo( RankPage.Tabs.Objects[RankPage.ActivePageIndex] );
  if Assigned( aRankInfo ) then
  begin
    aRankInfo.Period := edtPeriod.Text;
    EditorToData( RankPage.ActivePageIndex );
  end;
end;

end.

