unit frmCD078G1U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, dxSkinsCore, dxSkinBlack, dxSkinBlue, dxSkinCaramel,
  dxSkinCoffee, dxSkinsDefaultPainters, cxStyles, dxSkinscxPCPainter,
  cxCustomData, cxGraphics, cxFilter, cxData, cxDataStorage, cxEdit, DB,
  cxDBData, DBClient, cxGridLevel, cxGridCustomTableView, cxGridTableView,
  cxGridDBTableView, cxClasses, cxControls, cxGridCustomView, cxGrid,
  cxCheckBox, cxTextEdit, cbDataLookup, cxContainer, cxCurrencyEdit,
  StdCtrls, ExtCtrls,frmCD078B1U, ActnList,ADODB,Menus;

type
  TfrmCD078G1 = class(TForm)
    Panel2: TPanel;
    Label30: TLabel;
    Label1: TLabel;
    Label8: TLabel;
    edtDiscountAmt: TcxCurrencyEdit;
    MSalePoint: TDataLookup;
    edtDescription: TcxTextEdit;
    chkMasterSale: TcxCheckBox;
    Panel1: TPanel;
    gvCD078G1: TcxGrid;
    gvCD078G1View: TcxGridDBTableView;
    gvCD078G1ViewCol1: TcxGridDBColumn;
    glCD078G1: TcxGridLevel;
    cdsCD078G1: TClientDataSet;
    dsCD078G1: TDataSource;
    Panel3: TPanel;
    btnUpdate: TButton;
    btnSave: TButton;
    Button1: TButton;
    ActionList1: TActionList;
    actSave: TAction;
    actCancel: TAction;
    actUpdate: TAction;
    btnInsert: TButton;
    actInsert: TAction;
    edtDML: TcxTextEdit;
    gvCD078G1ViewCol2: TcxGridDBColumn;
    gvCD078G1ViewCol3: TcxGridDBColumn;
    gvCD078G1ViewCol4: TcxGridDBColumn;
    Label2: TLabel;
    edtGiftAmt: TcxCurrencyEdit;
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure MSalePointCodeNamePropertiesChange(Sender: TObject);
    procedure MSalePointCodeNamePropertiesInitPopup(Sender: TObject);
    procedure MSalePointCodeNoPropertiesChange(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure cdsCD078G1AfterScroll(DataSet: TDataSet);
    procedure ClearEditor;
    procedure ButtonStateChange;
    procedure chkMasterSalePropertiesChange(Sender: TObject);
    procedure actSaveExecute(Sender: TObject);
    procedure actCancelExecute(Sender: TObject);
    procedure actUpdateExecute(Sender: TObject);
    procedure actInsertExecute(Sender: TObject);
    procedure btnUpdateClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnInsertClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure edtDiscountAmtPropertiesChange(Sender: TObject);

  private
    { Private declarations }
    FSalePointCode : string;
    FSalePointName : String;
    FCitemCode : string;
    FFaciItem : string;
    FServiceType : string;
    FKeyValue : String;
    FMode:TDML;
    FMainCD078G1: TClientDataSet;
    FWriter : TADOQuery;
    FChkIndex:Integer;
    procedure PrepareCD078G1DataSet;
    procedure DeCompoundCD078G1DataSet;
    procedure EditorStateChange;
    procedure SetSourceDataSet(ADataSet: TClientDataSet);
    procedure DataToEditor;
    function EditorToData : Boolean ;
    function GetStepNo : string;
    function DataValidate: Boolean;
    function DataSetStateChange(const aMode: TDML): Boolean;
    procedure DMLModeChange(aMode: TDML);
    procedure FilterCD078G1;
    procedure UpdMasterSale(aBookMark : Pointer ) ;
    procedure CaculatGift;
  public
    { Public declarations }
    constructor Create(aMode: TDML; aKeyValue: String;aDataSet: TClientDataSet );reintroduce;
    property MSalePointCode : String read FSalePointCode write FSalePointCode;
    property MSalePointName : String read FSalePointName write FSalePointName;
    property MCitemCode : string read FCitemCode write FCitemCode;
    property MServiceType : string read FServiceType write FServiceType;
    property MFaciItem : string read FFaciItem write FFaciItem;
    property SetCD078G1DataSet:  TClientDataSet read FMainCD078G1 write SetSourceDataSet;

  end;

var
  frmCD078G1: TfrmCD078G1;

implementation

uses cbDBController,cbUtilis;

{$R *.dfm}

{ TForm1 }

constructor TfrmCD078G1.Create(aMode: TDML; aKeyValue: String; aDataSet: TClientDataSet);
begin
  inherited Create( Application );
  FMode := aMode;
  FKeyValue := aKeyValue;
  FMainCD078G1 := aDataSet;

//  FMainCD078G1 := aDataSet;
//  cdsCD078G1 := aDataSet;
//  aDataSet.AfterScroll := cdsCD078G1AfterScroll;
  //cdsCD078G1 := aDataSet;
//  cdsCD078G1.FieldByName( 'DiscountAmt' ).AsString
end;

procedure TfrmCD078G1.PrepareCD078G1DataSet;
begin
  if cdsCD078G1.FieldDefs.Count <= 0  then
  begin
    cdsCD078G1.FieldDefs.Add( 'BPCode', ftString, 10);
    cdsCD078G1.FieldDefs.Add( 'CitemCode',ftInteger );
    cdsCD078G1.FieldDefs.Add( 'CitemName', ftString, 40);
    cdsCD078G1.FieldDefs.Add( 'ServiceType',ftString,1);
    cdsCD078G1.FieldDefs.Add( 'FaciItem',ftString,10 );
    cdsCD078G1.FieldDefs.Add( 'StepNo',ftInteger );
    cdsCD078G1.FieldDefs.Add( 'Description',ftString,50);
    cdsCD078G1.FieldDefs.Add('MasterSale',ftInteger);
    cdsCD078G1.FieldDefs.Add('DiscountAmt',ftInteger);
    cdsCD078G1.FieldDefs.Add('SalePointCode',ftString,10);
    cdsCD078G1.FieldDefs.Add( 'SalePointName',ftString,100);
    cdsCD078G1.CreateDataSet;
  end;
  cdsCD078G1.EmptyDataSet;
end;

procedure TfrmCD078G1.FormShow(Sender: TObject);
begin
  Screen.Cursor := crSQLWait;
  try
    dsCD078G1.DataSet := FMainCD078G1;
    cdsCD078G1 := FMainCD078G1;
    FilterCD078G1;
    FMainCD078G1.AfterScroll := cdsCD078G1AfterScroll;
    FChkIndex := -1;
    DMLModeChange( FMode );
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TfrmCD078G1.FormCreate(Sender: TObject);
begin
  MSalePoint.Initializa;
  FWriter := DBController.DataWriter;
end;

procedure TfrmCD078G1.DeCompoundCD078G1DataSet;
begin
  cdsCD078G1.EmptyDataSet;
  cdsCD078G1.First;
  FMainCD078G1.First;
  while not FMainCD078G1.Eof do
  begin
    cdsCD078G1.Append;
    cdsCD078G1.FieldByName( 'BPCode' ).AsString := FKeyValue;
    cdsCD078G1.FieldByName( 'CitemCode' ).AsString := FCitemCode;
    cdsCD078G1.FieldByName( 'FaciItem' ).AsString := FFaciItem;
    cdsCD078G1.FieldByName( 'StepNo' ).AsString := FMainCD078G1.FieldByName( 'StepNo').AsString;
    cdsCD078G1.Post;
    FMainCD078G1.Next;

  end;
end;

procedure TfrmCD078G1.EditorStateChange;
var aIndex : Integer;
begin

  for aIndex := 0 to ControlCount - 1 do
  begin
    if ( Controls[aIndex].Name <> 'btnSave' ) and
       ( Controls[aIndex].Name <> 'Panel3' )  then
      Controls[aIndex].Enabled := not ( FMode in [dmBrowser] );
  end;
  btnUpdate.Enabled := ( FMode in [dmBrowser] );
  btnInsert.Enabled := ( FMode in [dmBrowser] );
  btnSave.Enabled := not (FMode in [dmBrowser]);
  Panel1.Enabled := ( FMode in [ dmBrowser ] );
  case FMode of
    dmInsert:
      begin
        edtDiscountAmt.Enabled := True;
        MSalePoint.Enabled := True;
        edtDescription.Enabled := True;
        edtDescription.SetFocus;
      end;
    dmUpdate:
      begin
        edtDiscountAmt.Enabled := True;
        MSalePoint.Enabled := True;
        edtDescription.Enabled := True;
        edtDescription.SetFocus;
      end;
    dmBrowser:
      begin
        edtDiscountAmt.Enabled := False;
        MSalePoint.Enabled := False;
        edtDescription.Enabled := False;
      end;
  end;
end;

procedure TfrmCD078G1.MSalePointCodeNamePropertiesChange(Sender: TObject);
begin
  MSalePoint.CodeNamePropertiesChange( Sender );
  CaculatGift;
end;

procedure TfrmCD078G1.MSalePointCodeNamePropertiesInitPopup(
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

procedure TfrmCD078G1.MSalePointCodeNoPropertiesChange(Sender: TObject);
begin
  MSalePoint.CodeNoPropertiesChange( Sender );
  CaculatGift;
end;

procedure TfrmCD078G1.FormDestroy(Sender: TObject);
begin
  MSalePoint.Finaliza;
end;

procedure TfrmCD078G1.DataToEditor;
begin
  cdsCD078G1AfterScroll( cdsCD078G1 );
end;

procedure TfrmCD078G1.cdsCD078G1AfterScroll(DataSet: TDataSet);
begin
  edtDiscountAmt.Value := cdsCD078G1.FieldByName( 'DiscountAmt').AsFloat;
  MSalePoint.Value := cdsCD078G1.FieldByName( 'SalePointCode' ).AsString;
  edtDescription.Text := cdsCD078G1.FieldByName( 'Description' ).AsString;
  chkMasterSale.Checked := (cdsCD078G1.FieldByName( 'MasterSale' ).AsInteger = 1);
  CaculatGift;
end;

procedure TfrmCD078G1.SetSourceDataSet(ADataSet: TClientDataSet);
begin
  //cdsCD078G1 := ADataSet;
  //dsCD078G1.DataSet := ADataSet;
  //ADataSet.AfterScroll := cdsCD078G1AfterScroll;

end;

procedure TfrmCD078G1.ClearEditor;
begin
  edtDiscountAmt.Value := 0;
  MSalePoint.Clear;
  edtDescription.Text := EmptyStr;
  chkMasterSale.Checked := False;
  if FMode = dmInsert then
  begin
    MSalePoint.Value := FSalePointCode;
    MSalePoint.ValueName := FSalePointName;
  end;
end;



function TfrmCD078G1.EditorToData: Boolean;
  var aBookmark : Pointer;
begin
  Result := False;

  try
    if (FMode in [dmInsert] ) then
    begin
      cdsCD078G1.FieldByName( 'BPCODE' ).AsString := FKeyValue;
      cdsCD078G1.FieldByName( 'CITEMCODE' ).AsString := FCitemCode;
      cdsCD078G1.FieldByName( 'SERVICETYPE' ).AsString := FServiceType;
      cdsCD078G1.FieldByName( 'FACIITEM' ).AsString := FFaciItem;
      cdsCD078G1.FieldByName( 'STEPNO' ).AsString := GetStepNo;
    end;
    cdsCD078G1.FieldByName( 'DESCRIPTION' ).AsString := edtDescription.Text;
    cdsCD078G1.FieldByName( 'MASTERSALE' ).AsInteger := 0;
    edtDiscountAmt.PostEditValue;
    cdsCD078G1.FieldByName( 'DISCOUNTAMT' ).AsString := VarToStrDef(edtDiscountAmt.EditValue,'0');
    cdsCD078G1.FieldByName( 'SALEPOINTCODE' ).AsString := MSalePoint.Value;
    cdsCD078G1.FieldByName( 'SALEPOINTNAME' ).AsString := MSalePoint.ValueName;

  except
    on E: Exception do
    begin
      ErrorMsg( Format( '存檔有誤, 原因:%s', [E.Message] ) );
      Exit;
    end;
  end;

  Result := True;

end;
function TfrmCD078G1.GetStepNo: string;
begin
  FWriter.Close;
  FWriter.SQL.Text := Format(
    ' SELECT %s.S_CD078G1.NEXTVAL FROM DUAL ', [DBController.LoginInfo.DbAccount] );
  FWriter.Open;
  Result := FWriter.Fields[0].AsString;
  FWriter.Close;
end;

function TfrmCD078G1.DataValidate: Boolean;
begin
  Result := False;
  if edtDiscountAmt.Value <= 0 then
  begin
    WarningMsg( '金額不可小於0！');
    if edtDiscountAmt.CanFocus then edtDiscountAmt.SetFocus;
    Exit;
  end;
  if MSalePoint.Value = EmptyStr then
  begin
    WarningMsg( '請設定點數辦法！');
    Exit;
  end;
  Result := True;
end;

function TfrmCD078G1.DataSetStateChange(const aMode: TDML): Boolean;
begin
  Result := True;
  case aMode of
    dmInsert:
      begin
        FMainCD078G1.AfterScroll := nil;
        FMainCD078G1.Append;
        ClearEditor;
      end;
    dmUpdate:
      begin
        FMainCD078G1.AfterScroll := nil;
        ClearEditor;
        DataToEditor;
        FMainCD078G1.Edit;
      end;
    dmCancel:
      begin
        FMainCD078G1.Cancel;
      end;
    dmBrowser:
      begin
        ClearEditor;
        DataToEditor;
      end;
    dmSave:
      begin
        Screen.Cursor := crSQLWait;
        try
          Result := DataValidate;
          if Result then Result := EditorToData;
          if Result then
          begin
            try
              FMainCD078G1.Post;
              UpdMasterSale( FMainCD078G1.GetBookmark );

            except
              on E: Exception do
              begin
                ErrorMsg( Format( '存檔有誤, 原因:%s。', [E.Message] ) );
                Result := False;
              end;
            end;

          end;
        finally
          FMainCD078G1.AfterScroll := cdsCD078G1AfterScroll;
          Screen.Cursor := crDefault;
        end;
      end;



  end;
end;

procedure TfrmCD078G1.ButtonStateChange;
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
    actCancel.Caption := '結束(&C)';
    actCancel.ShortCut := TextToShortCut( 'Alt+C' );
  end;
end;

procedure TfrmCD078G1.chkMasterSalePropertiesChange(Sender: TObject);

begin
  if chkMasterSale.Checked then
  begin
    if  FMode <> dmBrowser then
    begin
      FChkIndex := cdsCD078G1.RecNo;
    end;
  end;
//  FChkIndex := cdsCD078G1.RecNo;
  {
  if FMode in [dmInsert,dmUpdate ] then
  begin
    aBookMark := cdsCD078G1.GetBookmark;
    try
      FMainCD078G1.Edit;
//      cdsCD078G1.DisableControls;
      cdsCD078G1.AfterScroll := nil;
      cdsCD078G1.First;
      while not cdsCD078G1.Eof do
      begin
        cdsCD078G1.FieldByName( 'MasterSale' ).AsInteger := 0;
        cdsCD078G1.Next;
      end;
    finally
      cdsCD078G1.GotoBookmark( aBookMark );
      cdsCD078G1.FieldByName( 'MasterSale' ).AsInteger := 1;
      cdsCD078G1.Post;
      cdsCD078G1.AfterScroll := cdsCD078G1AfterScroll;
//      cdsCD078G1.EnableControls;
    end;
  end;
  }

end;
procedure TfrmCD078G1.DMLModeChange(aMode: TDML);
begin
  if not DataSetStateChange( aMode ) then Exit;
  if ( aMode = dmSave  ) and ( FMode = dmInsert ) then
  begin
    FMode := dmBrowser;
    DMLModeChange( dmBrowser );
    Exit;
  end else
  if ( aMode in [dmSave, dmCancel] ) then
    FMode := dmBrowser;
  ButtonStateChange;
  EditorStateChange;
  edtDML.Text := TDMLString[FMode];
end;

procedure TfrmCD078G1.actSaveExecute(Sender: TObject);
begin
  DMLModeChange( dmSave );
end;

procedure TfrmCD078G1.actCancelExecute(Sender: TObject);
  var aBookMark : Pointer;
      aExit : Boolean;
begin
  DMLModeChange( dmCancel );
  aBookMark := nil;
  aExit := False;
  if cdsCD078G1.IsEmpty then
  begin
    MessageDlg('請至少設定一筆資料！',mtWarning,[mbOK],0);
    Exit;
  end;
  aBookMark := cdsCD078G1.GetBookmark;
  try
    cdsCD078G1.DisableControls;
    cdsCD078G1.First;
    while not cdsCD078G1.Eof do
    begin
      if cdsCD078G1.FieldByName('MasterSale').AsString = '1' then
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
  if aExit then
    Self.Close
  else begin
    MessageDlg('請至少設定一筆主推！',mtWarning,[mbOK],0);
  end;

end;

procedure TfrmCD078G1.actUpdateExecute(Sender: TObject);
begin
  FMode := dmUpdate;
  DMLModeChange(dmUpdate);
end;

procedure TfrmCD078G1.actInsertExecute(Sender: TObject);
begin
  FMode := dmInsert;
  DMLModeChange( dmInsert );
end;

procedure TfrmCD078G1.btnUpdateClick(Sender: TObject);
begin
  actUpdate.Execute;
end;

procedure TfrmCD078G1.btnSaveClick(Sender: TObject);
begin
  actSave.Execute;
end;

procedure TfrmCD078G1.btnInsertClick(Sender: TObject);
begin
  actInsert.Execute;
end;

procedure TfrmCD078G1.Button1Click(Sender: TObject);
begin
  actCancel.Execute;
end;



procedure TfrmCD078G1.FilterCD078G1;
  var aBookMark : Pointer;
begin
  if cdsCD078G1.IsEmpty then
  begin
    try
      aBookMark := cdsCD078G1.GetBookmark;
      cdsCD078G1.AfterScroll := nil;
      cdsCD078G1.Filter := EmptyStr;
      cdsCD078G1.Filter := Format(
          'BPCODE=''%s'' AND CITEMCODE=''%s''             ',
          [ FKeyValue,FCitemCode ]);
      cdsCD078G1.Filtered := True;
    finally
      cdsCD078G1.AfterScroll := cdsCD078G1AfterScroll;
      cdsCD078G1.GotoBookmark( aBookMark );
      cdsCD078G1.FreeBookmark( aBookMark );

    end;
  end;

end;
{ 主推多階只能有一個,所以先把所有的資料改成0,在到當筆回填 }
procedure TfrmCD078G1.UpdMasterSale(aBookMark : Pointer);
  var aBook : Pointer;
begin
  try
    aBook := nil;
    cdsCD078G1.DisableControls;
    cdsCD078G1.First;
    //先將所有的主推都清成0,但先記錄那一筆是1的
    while not cdsCD078G1.Eof do
    begin
      if cdsCD078G1.FieldByName( 'MasterSale' ).AsInteger = 1 then
      begin
        aBook := cdsCD078G1.GetBookmark;
      end;
      cdsCD078G1.Edit;
      cdsCD078G1.FieldByName( 'MasterSale' ).AsInteger := 0;
      cdsCD078G1.Post;
      cdsCD078G1.Next;
    end;
  finally
    cdsCD078G1.GotoBookmark( aBookMark );
    //如果畫面上的主推不是True,還要檢查剛剛的BookMark是不是有值,如果有值,要Upd成1
    if chkMasterSale.Checked then
    begin
      cdsCD078G1.Edit;
      cdsCD078G1.FieldByName( 'MasterSale' ).AsInteger := 1;
      cdsCD078G1.Post;
    end else
    begin
      if Assigned(aBook) then  //檢查aBook是否有值
      begin
        cdsCD078G1.GotoBookmark( aBook );
        cdsCD078G1.Edit;
        cdsCD078G1.FieldByName( 'MasterSale' ).AsInteger := 1;
        cdsCD078G1.Post;
        cdsCD078G1.GotoBookmark( aBookMark ); //再跳回剛剛那筆資料
      end;
    end;
    cdsCD078G1.FreeBookmark( aBookMark );
    cdsCD078G1.FreeBookmark( aBook );
    cdsCD078G1.EnableControls;
  end;

end;

procedure TfrmCD078G1.edtDiscountAmtPropertiesChange(Sender: TObject);
begin
  CaculatGift;
end;

// #5318 測試後調整---
//要Show出贈送金額(CD107.Pointdue*該筆銷售金額+ CD107.Addvalue)- 銷售金額)  By Kin 2009/10/30
procedure TfrmCD078G1.CaculatGift;
  var aPointdue,aAddvalue :Double;
begin
  aPointdue := 0;
  aAddvalue := 0;
  if MSalePoint.Value <> EmptyStr then
  begin
    FWriter.Close;
    FWriter.SQL.Text :=Format('SELECT Pointdou,Addvalue FROM %s.CD107  '+
                              ' WHERE CodeNo = ''%s''                  ',
                              [DBController.LoginInfo.DbAccount,MSalePoint.Value]);
    FWriter.Open;
    aPointdue := FWriter.FieldByName( 'Pointdou' ).AsFloat ;
    aAddvalue := FWriter.FieldByName( 'Addvalue' ).AsFloat ;
    FWriter.Close;
  end;
  if (aPointdue > 0) and (edtDiscountAmt.Value > 0 ) then
  begin
    edtGiftAmt.Value := ((edtDiscountAmt.Value * aPointdue)+ aAddvalue)-edtDiscountAmt.Value;
  end else
  begin
    edtGiftAmt.Value := 0;
  end;
end;

end.
