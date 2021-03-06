unit frmCD078G0U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, cxStyles, dxSkinsCore, dxSkinBlack, dxSkinBlue,
  dxSkinCaramel, dxSkinCoffee, dxSkinsDefaultPainters, dxSkinscxPCPainter,
  cxCustomData, cxGraphics, cxFilter, cxData, cxDataStorage, cxEdit, DB,
  cxDBData, cxContainer, cxCheckBox, cxGridLevel, cxGridCustomTableView,
  cxGridTableView, cxGridDBTableView, cxClasses, cxControls,
  cxGridCustomView, cxGrid, cbDataLookup, StdCtrls,ADODB,DBClient,
  frmCD078B1U, ActnList,Menus, Buttons, cxLookAndFeelPainters, cxTextEdit,
  cxButtons,frmCD078B0U ;

type
  TfrmCD078GU = class(TForm)
    Panel1: TPanel;
    Label2: TLabel;
    ServiceType: TDataLookup;
    Label1: TLabel;
    CItem: TDataLookup;
    Panel2: TPanel;
    PenalGrid: TcxGrid;
    gvPenal: TcxGridDBTableView;
    gvPenalCol1: TcxGridDBColumn;
    glPenal: TcxGridLevel;
    chkMasterSale: TcxCheckBox;
    dsCD078G1: TDataSource;
    Panel3: TPanel;
    FaciItem: TDataLookup;
    Label3: TLabel;
    cdsCD078G1: TClientDataSet;
    ActionList1: TActionList;
    actSave: TAction;
    actCancel: TAction;
    actUpdate: TAction;
    MSalePoint: TDataLookup;
    Label4: TLabel;
    btnUpdate: TButton;
    btnSave: TButton;
    Button1: TButton;
    btnRankList: TBitBtn;
    btnInstCodeStr: TButton;
    cxButton1: TcxButton;
    edtInstCodeStr: TcxTextEdit;
    edtDML: TcxTextEdit;
    btnDelRnak: TBitBtn;
    gvPenalCol2: TcxGridDBColumn;
    gvPenalCol3: TcxGridDBColumn;
    gvPenalCol4: TcxGridDBColumn;
    procedure FormCreate(Sender: TObject);
    procedure cdsCD078G1AfterScroll(DataSet: TDataSet);
    procedure ServiceTypeCodeNamePropertiesChange(Sender: TObject);
    procedure ServiceTypeCodeNamePropertiesInitPopup(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure ServiceTypeCodeNoPropertiesChange(Sender: TObject);
    procedure CItemCodeNamePropertiesChange(Sender: TObject);
    procedure CItemCodeNamePropertiesInitPopup(Sender: TObject);
    procedure CItemCodeNoPropertiesChange(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FaciItemCodeNamePropertiesChange(Sender: TObject);
    procedure FaciItemCodeNamePropertiesInitPopup(Sender: TObject);
    procedure FaciItemCodeNoPropertiesChange(Sender: TObject);
    procedure MSalePointCodeNamePropertiesChange(Sender: TObject);
    procedure MSalePointCodeNamePropertiesInitPopup(Sender: TObject);
    procedure MSalePointCodeNoPropertiesChange(Sender: TObject);
    procedure actSaveExecute(Sender: TObject);
    procedure actCancelExecute(Sender: TObject);
    procedure actUpdateExecute(Sender: TObject);
    procedure EditorStateChange;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnRankListClick(Sender: TObject);
    procedure btnInstCodeStrClick(Sender: TObject);
    procedure cxButton1Click(Sender: TObject);
    procedure btnDelRnakClick(Sender: TObject);
  private
    { Private declarations }
    FMode: TDML;
    FKeyValue: String;
    FDataSet: TClientDataSet;
    FAlreadySetCItemValue: String;
    FCD078G1DataSet: TClientDataSet;
    FaciDataSet: TClientDataSet;
    FInstCodeDataSet : TClientDataSet;
    FQueryParam: TQueryParam;
    FCanEnable : Boolean;
    procedure SetCD078G1DataSet( ADataSet: TClientDataSet);
    procedure ButtonStateChange;
    procedure FilterCD078G1DataWhenNew;
    procedure UpdateCD078G1LinkKey;
    function DataValidate(): Boolean;
    function VdMustInput: Boolean;
    procedure ClearEditor;
    procedure DataToEditor;
    function EditorToData: Boolean;
    function DataSetStateChange(const aMode: TDML): Boolean;
    procedure DMLModeChange(aMode: TDML);

  public
    constructor Create(aMode: TDML; aKeyValue: String; aDataSet:TClientDataSet );reintroduce;
    property AlreadySetCItemValue: String read FAlreadySetCItemValue write FAlreadySetCItemValue;
    property FaciItemDataSet: TClientDataSet read FaciDataSet write FaciDataSet;
    property CD078G1DataSet:  TClientDataSet read FCD078G1DataSet write SetCD078G1DataSet;
    property InstCodeDataSet: TClientDataSet read FInstCodeDataSet write FInstCodeDataSet;
    { Public declarations }
  end;

var
  frmCD078GU: TfrmCD078GU;

implementation
  uses cbUtilis, frmMultiSelectU, cbDBController,frmCD078G1U;
{$R *.dfm}

{ TfrmCD078GU }

constructor TfrmCD078GU.Create(aMode: TDML; aKeyValue: String;
  aDataSet: TClientDataSet);
begin
  inherited Create( Application );
  FMode := aMode;
  FKeyValue := aKeyValue;
  FDataSet := aDataSet;
  {
  FReader := DBController.CodeReader;
  FLockCheck := False;
  FLockCItemChange := False;
  FOldValue.CItemCode := EmptyStr;
  FOldValue.ServiceType := EmptyStr;
  FOldValue.PenalCode := EmptyStr;
  InitCD019Data;
  }
end;



procedure TfrmCD078GU.SetCD078G1DataSet(  ADataSet: TClientDataSet);
begin
  cdsCD078G1 := ADataSet;
  dsCD078G1.DataSet := ADataSet;
  ADataSet.AfterScroll := cdsCD078G1AfterScroll;
end;

procedure TfrmCD078GU.FormCreate(Sender: TObject);
begin
  CItem.Initializa;
  ServiceType.Initializa;
  FaciItem.Initializa;
  MSalePoint.Initializa;
end;

procedure TfrmCD078GU.cdsCD078G1AfterScroll(DataSet: TDataSet);
begin
  if cdsCD078G1.FieldByName('MasterSale').AsInteger > 0 then
    chkMasterSale.Checked := True
  else
    chkMasterSale.Checked := False;
end;

procedure TfrmCD078GU.ServiceTypeCodeNamePropertiesChange(Sender: TObject);
begin
  ServiceType.CodeNamePropertiesChange( Sender );
  CItem.Clear;
  FaciItem.Clear;

end;

procedure TfrmCD078GU.ServiceTypeCodeNamePropertiesInitPopup(
  Sender: TObject);
begin
{ ?A?????O }
  Screen.Cursor := crSQLWait;
  try
    ServiceType.SQL.Text := Format(
    ' SELECT CODENO, DESCRIPTION    ' +
    '  FROM %s.CD046                ' +
    '  WHERE CODENO=''%s''          ' ,
    [DBController.LoginInfo.DbAccount,'D'] );
    ServiceType.CodeNamePropertiesInitPopup( Sender );
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TfrmCD078GU.ButtonStateChange;
begin
  if ( FMode = dmInsert ) then
  begin
    actSave.Enabled := True;
  end else
  if ( FMode = dmUpdate ) then
  begin
    actSave.Enabled := True;
  end else
  begin
    actSave.Enabled := False;
    actCancel.Caption := '????(&C)';
    actCancel.ShortCut := TextToShortCut( 'Alt+C' );
  end;
  
end;

procedure TfrmCD078GU.FilterCD078G1DataWhenNew;
begin
  cdsCD078G1.Filtered := False;
  cdsCD078G1.Filtered := False;
  cdsCD078G1.Filter := Format(
    'BPCODE=''%s'' AND CITEMCODE=''-1'' AND SERVICETYPE=''-1''', [FKeyValue] );
  cdsCD078G1.Filtered := True;

end;

procedure TfrmCD078GU.UpdateCD078G1LinkKey;
var
  aRecordCount: Integer;
begin
  aRecordCount := cdsCD078G1.RecordCount;
  cdsCD078G1.First;
  while not cdsCD078G1.Eof do
  begin
    cdsCD078G1.Edit;
    cdsCD078G1.FieldByName( 'BPCode' ).AsString := FKeyValue;
    cdsCD078G1.FieldByName( 'CitemCode' ).AsString := CItem.Value;
    cdsCD078G1.FieldByName( 'ServiceType' ).AsString := ServiceType.Value;
    cdsCD078G1.FieldByName( 'FaciItem' ).AsString := FaciItem.Value;
    cdsCD078G1.Post;
    if ( aRecordCount = cdsCD078G1.RecordCount ) then
      cdsCD078G1.Next;
  end;
    cdsCD078G1.First;
end;

function TfrmCD078GU.DataValidate(): Boolean;
begin
  Result := VdMustInput;
end;

function TfrmCD078GU.VdMustInput: Boolean;
  var aBookMark : Pointer;
      aExit : Boolean;
begin
  Result := False;
  aExit := False;
  if ( CItem.Value = EmptyStr ) then
  begin
    WarningMsg( '?????J?i???O?????j?C' );
    if CItem.CodeNo.CanFocusEx then CItem.CodeNo.SetFocus;
    Exit;
  end;
  {}
  if ( ServiceType.Value = EmptyStr ) then
  begin
    WarningMsg( '?????J?i?A?????O?j?C' );
    if ServiceType.CodeNo.CanFocusEx then ServiceType.CodeNo.SetFocus;
    Exit;
  end;
  if ( FQueryParam.Param1 = EmptyStr ) then
  begin
    WarningMsg( '?????w???u???O?I');
    if btnInstCodeStr.CanFocus then btnInstCodeStr.Click;
    Exit;
  end;
  if cdsCD078G1.IsEmpty then
  begin
    WarningMsg( '???]?w?I???????I');
    if btnRankList.CanFocus then btnRankList.SetFocus;
    Exit;
  end;
  aBookMark := cdsCD078G1.GetBookmark;
  try
    cdsCD078G1.DisableControls;
    cdsCD078G1.First;
    while not cdsCD078G1.Eof do
    begin
      if cdsCD078G1.FieldByName( 'MasterSale' ).AsString = '1' then
      begin
        aExit := True;
        Break;
      end;
      cdsCD078G1.Next;
    end;

  finally
    cdsCD078G1.GotoBookmark( aBookMark );
    cdsCD078G1.FreeBookmark( aBookMark );
    cdsCD078G1.EnableControls;
  end;
  if not aExit then
  begin
    MessageDlg( '???????]?w?@???D???I',mtWarning,[mbOK],0);
    Result := False;
    Exit;
  end;
  {
  if ( FaciItem.Value = EmptyStr ) then
  begin
    WarningMsg( '?????J?i???w?]???j?C' );
    if FaciItem.CodeNo.CanFocusEx then FaciItem.CodeNo.SetFocus;
    Exit;
  end;
  }
  Result := True;
end;

procedure TfrmCD078GU.ClearEditor;
begin
  ServiceType.Clear;
  CItem.Clear;
  FaciItem.Clear;
  MSalePoint.Clear;
  chkMasterSale.Checked := False;
  FQueryParam.Param1 := EmptyStr;
  edtInstCodeStr.Text := EmptyStr;
end;

procedure TfrmCD078GU.DataToEditor;
begin
  ServiceType.Value := FDataSet.FieldByName( 'SERVICETYPE' ).AsString;
  CItem.Value := FDataSet.FieldByName( 'CITEMCODE' ).AsString;
  CItem.ValueName := FDataSet.FieldByName( 'CITEMNAME' ).AsString;
  FaciItem.Value := FDataSet.FieldByName( 'FACIITEM' ).AsString;
  MSalePoint.Value := FDataSet.FieldByName( 'SalePointcode' ).AsString;
  FQueryParam.Param1 :=  FDataSet.FieldByName( 'INSTCODESTR' ).AsString;
  edtInstCodeStr.Text := DBController.GetInstName( FQueryParam.Param1 );
  cdsCD078G1AfterScroll( cdsCD078G1 );

end;

procedure TfrmCD078GU.FormDestroy(Sender: TObject);
begin
  ServiceType.Finaliza;
  CItem.Finaliza;
  FaciItem.Finaliza;
  MSalePoint.Finaliza;
end;

function TfrmCD078GU.EditorToData: Boolean;
begin
  Result := False;
  try
    if ( FMode in [dmInsert] ) then
    begin
      FDataSet.FieldByName( 'BPCODE' ).AsString := FKeyValue;
    end;
    FDataSet.FieldByName( 'CITEMCODE' ).AsString := CItem.Value;
    FDataSet.FieldByName( 'CITEMNAME' ).AsString := CItem.ValueName;
    FDataSet.FieldByName( 'SERVICETYPE' ).AsString := ServiceType.Value;
    FDataSet.FieldByName( 'FACIITEM' ).AsString := FaciItem.Value;
    FDataSet.FieldByName( 'SalePointcode' ).AsString := MSalePoint.Value;
    FDataSet.FieldByName( 'SalePointName' ).AsString := MSalePoint.ValueName;
    FDataSet.FieldByName( 'INSTCODESTR' ).AsString := TrimChar( FQueryParam.Param1, [''''] );
   except
    on E: Exception do
    begin
      ErrorMsg( Format( '?s?????~, ???]:%s', [E.Message] ) );
      Exit;
    end;
  end;
  Result := True;
end;

procedure TfrmCD078GU.ServiceTypeCodeNoPropertiesChange(Sender: TObject);
begin
  ServiceType.CodeNoPropertiesChange( Sender );
  CItem.Clear;
  FaciItem.Clear;
end;

procedure TfrmCD078GU.CItemCodeNamePropertiesChange(Sender: TObject);
begin
  CItem.CodeNamePropertiesChange( Sender );
end;

procedure TfrmCD078GU.CItemCodeNamePropertiesInitPopup(Sender: TObject);
var
  aCItems: String;
begin

  Screen.Cursor := crSQLWait;
  try
    CItem.SQL.Text := Format(
      ' SELECT A.CODENO, A.DESCRIPTION      ' +
      '  FROM %s.CD019 A                    ' +
      ' WHERE A.REFNO = 22                  ' +
      '   AND Nvl(A.STOPFLAG,0) = 0         ',
      [DBController.LoginInfo.DbAccount] );
    {}
    if ( ServiceType.value <> EmptyStr )then
      CItem.SQL.Text := CItem.SQL.Text +
        ' AND A.SERVICETYPE=''' + ServiceType.Value+''' ';
    {}
    if ( FMode in [dmInsert] ) and ( FAlreadySetCItemValue <> EmptyStr ) then
    begin

      aCItems := QuotedValue( FAlreadySetCItemValue );
      CItem.SQL.Text := CItem.SQL.Text +
        ' AND A.CODENO NOT IN ( ' + aCItems + ' ) ';
    end;
    {}
    CItem.SQL.Text := CItem.SQL.Text +
      '  ORDER BY A.CODENO  ';
    CItem.CodeNamePropertiesInitPopup( Sender );
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TfrmCD078GU.CItemCodeNoPropertiesChange(Sender: TObject);
begin
  CItem.CodeNoPropertiesChange( Sender );
end;

procedure TfrmCD078GU.FormShow(Sender: TObject);
begin
  Screen.Cursor := crSQLWait;
  FCanEnable := True;
  try
    FaciItem.DataSet := FaciDataSet;
    //#7238 add globle variable to check whether the outside buttons can be use By Kin 2016/06/08
    if FMode = dmCannotModify then FCanEnable := False;
    DMLModeChange( FMode );
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TfrmCD078GU.FaciItemCodeNamePropertiesChange(Sender: TObject);
begin
  FaciItem.CodeNamePropertiesChange( Sender );
end;

procedure TfrmCD078GU.FaciItemCodeNamePropertiesInitPopup(Sender: TObject);
begin
  FaciDataSet.Filtered := False;
  FaciDataSet.Filter := Format( 'SERVICETYPE=''%s''', [ServiceType.Value] );
  FaciDataSet.Filtered := True;
  FaciItem.CodeNamePropertiesInitPopup( Sender );
end;

procedure TfrmCD078GU.FaciItemCodeNoPropertiesChange(Sender: TObject);
begin
  FaciDataSet.Filtered := False;
  FaciDataSet.Filter := Format( 'SERVICETYPE=''%s''', [ServiceType.Value] );
  FaciDataSet.Filtered := True;
  FaciItem.CodeNoPropertiesChange( Sender );
end;

procedure TfrmCD078GU.MSalePointCodeNamePropertiesChange(Sender: TObject);
begin
  MSalePoint.CodeNamePropertiesChange( Sender );
end;

procedure TfrmCD078GU.MSalePointCodeNamePropertiesInitPopup(
  Sender: TObject);
begin
  Screen.Cursor := crSQLWait;
  try
    MSalePoint.SQL.Text := Format(
    ' SELECT CODENO, DESCRIPTION    ' +
    '  FROM %s.CD107                ' +
    '  WHERE STOPFLAG <> 1          ' +
    '  AND POINTDOU >= 1            ' ,
    [DBController.LoginInfo.DbAccount] );
    MSalePoint.CodeNamePropertiesInitPopup( Sender );
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TfrmCD078GU.MSalePointCodeNoPropertiesChange(Sender: TObject);
begin
  MSalePoint.CodeNoPropertiesChange( Sender );
end;

function TfrmCD078GU.DataSetStateChange(const aMode: TDML): Boolean;

  procedure DoSubDetailSave;
  begin
    if ( not cdsCD078G1.IsEmpty ) then
      UpdateCD078G1LinkKey;
    frmCD078B1.CloneCD078G1Data;
  end;

begin
  Result := True;
  case aMode of
    dmInsert:
      begin
        FilterCD078G1DataWhenNew;
        ClearEditor;
        FDataSet.Append;
      end;
    dmUpdate:
      begin
        ClearEditor;
        DataToEditor;
        FDataSet.Edit;
      end;
    dmCancel:
      begin
        FDataSet.Cancel;
      end;
    dmBrowser,dmCannotModify:
      begin
        ClearEditor;
        DataToEditor;
      end;  
    dmSave:
      begin
        Result := DataValidate;
        if Result then Result := EditorToData;
        if Result then
        begin
          try
            FDataSet.Post;
            DoSubDetailSave;
//            UpdateMCItemDataSet( EmptyStr );
  //          frmCD078B1.RestoreLinkKeyData;
            FAlreadySetCItemValue := frmCD078B1.GetCD078GCItemValue;
            Self.Close;
          except
            on E: Exception do
            begin
              ErrorMsg( Format( '?s?????~, ???]:%s?C', [E.Message] ) );
              Result := False;
            end;  
          end;
        end;
      end;
  end;
end;

procedure TfrmCD078GU.DMLModeChange(aMode: TDML);
begin

  if not DataSetStateChange( aMode ) then Exit;
  if ( aMode = dmSave  ) and ( FMode = dmInsert ) then
  begin
    DMLModeChange( dmInsert );
    Exit;
  end else
  if ( aMode in [dmSave, dmCancel] ) then
    FMode := dmBrowser;
  ButtonStateChange;
  EditorStateChange;
  edtDML.Text := TDMLString[FMode];
  if ( not FDataSet.IsEmpty ) and ( FCanEnable ) then
  begin
    frmCD078B1.btnUpdate.Enabled := True;
    frmCD078B1.actUpdate.Enabled := True;
    frmCD078B1.btnDelete.Enabled := True;
    frmCD078B1.actDelete.Enabled := True;
  end;
end;

procedure TfrmCD078GU.actSaveExecute(Sender: TObject);
begin
  DMLModeChange( dmSave );
//  Self.Close;
end;

procedure TfrmCD078GU.actCancelExecute(Sender: TObject);
begin
  DMLModeChange( dmCancel );
  Self.Close;
end;

procedure TfrmCD078GU.actUpdateExecute(Sender: TObject);
begin
  FMode := dmUpdate;
  DMLModeChange(dmUpdate);
end;

procedure TfrmCD078GU.EditorStateChange;
  var aIndex : Integer;
begin
  for aIndex := 0 to ControlCount - 1 do
  begin
    if ( Controls[aIndex].Name <> 'btnSave' ) and
       ( Controls[aIndex].Name <> 'Panel3' )  then
      Controls[aIndex].Enabled := not ( FMode in [dmBrowser,dmCannotModify] );
  end;
  btnUpdate.Enabled := ( FMode in [dmBrowser] );
  case FMode of
    dmInsert:
      begin
        if ( ServiceType.CodeNo.CanFocusEx ) then ServiceType.CodeNo.SetFocus;
        CItem.Enabled := True;
        ServiceType.Enabled := True;
      end;
    dmUpdate:
      begin
        if ( ServiceType.CodeNo.CanFocusEx ) then ServiceType.CodeNo.SetFocus;
        CItem.Enabled := True;
        ServiceType.Enabled := False;
      end;
    dmBrowser,dmCannotModify:
      begin
        if ( ServiceType.CodeNo.CanFocusEx ) then ServiceType.CodeNo.SetFocus;
      end;
  end;
end;

procedure TfrmCD078GU.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if ( FMode <> dmCancel ) then
    DMLModeChange( dmCancel );
end;

procedure TfrmCD078GU.btnRankListClick(Sender: TObject);
 var aMode : TDML;
begin

  if ServiceType.Value = EmptyStr then
  begin
    WarningMsg('?????]?w?A?????O');
    if ServiceType.CanFocus then ServiceType.SetFocus;
    Exit;
  end;
  if CItem.Value = EmptyStr then
  begin
    WarningMsg('?????]?w???O????');
    if CItem.CanFocus then CItem.SetFocus;
    Exit;
  end;

  if cdsCD078G1.IsEmpty then
    aMode := dmInsert
  else
    aMode := dmBrowser;

  frmCD078G1 := TfrmCD078G1.Create(aMode,FKeyValue,cdsCD078G1 );
  try
    frmCD078G1.Caption := '?I???????]?w [ ' + frmCD078G1.Caption + ']';
    cdsCD078G1.AfterScroll := nil;
    frmCD078G1.MSalePointCode := MSalePoint.Value;
    frmCD078G1.MSalePointName := MSalePoint.ValueName;
    frmCD078G1.MCitemCode := CItem.Value;
    frmCD078G1.MServiceType := ServiceType.Value;
    frmCD078G1.MFaciItem := FaciItem.Value;
    frmCD078G1.ShowModal;
  finally
    cdsCD078G1.AfterScroll := cdsCD078G1AfterScroll;
    cdsCD078G1.First;
    frmCD078G1.Free;
  end;

end;

procedure TfrmCD078GU.btnInstCodeStrClick(Sender: TObject);
begin
  if ( FMode in [dmInsert, dmUpdate] ) then
  begin
    frmMultiSelect := TfrmMultiSelect.Create( Application );
    try
      frmMultiSelect.Connection := DBController.DataConnection;
      frmMultiSelect.KeyFields := 'INSTCODE';
      frmMultiSelect.KeyValues := FQueryParam.Param1;
      frmMultiSelect.DisplayFields := 'INSTCODE,?????N?X,INSTNAME,???????O?W??';
      frmMultiSelect.ResultFields := 'INSTNAME';
      FInstCodeDataSet.Filtered := False;
      FInstCodeDataSet.Filter := Format( 'SERVICETYPE=''%s''', [ServiceType.Value] );
      FInstCodeDataSet.Filtered := True;
      try
        frmMultiSelect.DataSet := FInstCodeDataSet;
        if frmMultiSelect.ShowModal = mrOk then
        begin
          if ( FMode in [dmInsert, dmUpdate] ) then
          begin
            FQueryParam.Param1 := frmMultiSelect.SelectedValue;
            edtInstCodeStr.Text := frmMultiSelect.SelectedDisplay;
          end;
        end;
      finally
        FInstCodeDataSet.Filtered := False;
        FInstCodeDataSet.Filter := EmptyStr;
      end;
    finally
      frmMultiSelect.Free;
    end;
  end;
end;

procedure TfrmCD078GU.cxButton1Click(Sender: TObject);
begin
  if ( FMode in [dmInsert, dmUpdate] ) then
  begin
    FQueryParam.Param1 := EmptyStr;
    edtInstCodeStr.Text := EmptyStr;
  end;
end;

procedure TfrmCD078GU.btnDelRnakClick(Sender: TObject);
begin
  if not cdsCD078G1.IsEmpty then
  begin
    cdsCD078G1.Edit;
    cdsCD078G1.Delete;
    if cdsCD078G1.RecordCount = 1 then
    begin
      if cdsCD078G1.FieldByName('MasterSale').AsString = '0' then
      begin
        MessageDlg('???????]?w?@???D???I',mtInformation,[mbOK],0);
      end;
    end;
  end;
end;

end.
