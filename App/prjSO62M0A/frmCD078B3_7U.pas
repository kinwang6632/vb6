unit frmCD078B3_7U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,DB, DBClient, ADODB,StdCtrls,frmCD078B1U, cxGraphics,
  dxSkinsCore, dxSkinBlack, dxSkinBlue, dxSkinCaramel, dxSkinCoffee,
  dxSkinsDefaultPainters, cxControls, cxContainer, cxEdit, cxTextEdit,
  cxMaskEdit, cxDropDownEdit, cxCurrencyEdit,cbDBController, ActnList;

type
  TfrmCD078B3_7 = class(TForm)
    cmbUseMonth: TcxComboBox;
    Label5: TLabel;
    Label1: TLabel;
    edtBillPeriod: TcxCurrencyEdit;
    Label3: TLabel;
    edtMasterPeriod: TcxCurrencyEdit;
    Label6: TLabel;
    cmbBillType: TcxComboBox;
    Label7: TLabel;
    edtMonthAmt1: TcxCurrencyEdit;
    btnSave: TButton;
    Label2: TLabel;
    edtUseMonth: TcxCurrencyEdit;
    btnCancel: TButton;
    ActionList1: TActionList;
    actSave: TAction;
    actCancel: TAction;
    edtMonthAmt: TcxCurrencyEdit;
    procedure btnSaveClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure actSaveExecute(Sender: TObject);
    procedure actCancelExecute(Sender: TObject);
    procedure cmbUseMonthPropertiesChange(Sender: TObject);
  private
    { Private declarations }
    FMode: TDML;
    FSourceDataSet: TClientDataSet;
    FKeyArray: array [1..4] of String;
    FEnableLinkKey : Boolean;
    FDefAmt: string;
    FDefMonth: String;
    FReader: TADOQuery;
    FLinkKeyDataSet: TClientDataSet;
    FLinkCitem : String;
    procedure SetEnableLinkKey(const AEnabled: Boolean);
    procedure FindLinkKey(aPeriod : Integer);
    procedure SaveData;
    function DataValidate: Boolean;

    function VdMustInput: Boolean;
    function VdChkData:Boolean;
  public
    { Public declarations }
    constructor Create(AMode: TDML; ADataSet: TClientDataSet; AKeys: array of String); reintroduce;
    property EnableLinkKey: Boolean read FEnableLinkKey write SetEnableLinkKey;
    property DefaultAmt : string read FDefAmt write FDefAmt;
    property DefaultMonth: String read FDefMonth write FDefMonth;
    property LinkKeyDataSet: TClientDataSet read FLinkKeyDataSet write FLinkKeyDataSet;
    property LinkCitemCode: String read FLinkCitem write FLinkCitem;

  end;

var
  frmCD078B3_7: TfrmCD078B3_7;

implementation
uses
  cbUtilis;
{$R *.dfm}

{ TfrmCD078B3_7 }

constructor TfrmCD078B3_7.Create(AMode: TDML; ADataSet: TClientDataSet;
  AKeys: array of String);
begin
  inherited Create( Application );
  FMode := aMode;
  FKeyArray[1] := aKeys[0]; { BPCode }
  FKeyArray[2] := aKeys[1]; { CItemCode }
  FKeyArray[3] := aKeys[2]; { ServiceType }
  FSourceDataSet := ADataSet;
end;

procedure TfrmCD078B3_7.SetEnableLinkKey(const AEnabled: Boolean);
begin
  FEnableLinkKey := AEnabled;
end;

procedure TfrmCD078B3_7.btnSaveClick(Sender: TObject);

begin
  {
  for aIndex := 1  to StrToInt(VarToStrDef(edtBillPeriod.EditValue,'0'))   do
  begin
    FSourceDataSet.Append;
    FSourceDataSet.FieldByName( 'BPCode' ).AsString := FKeyArray[1];
    FSourceDataSet.FieldByName( 'CitemCode' ).AsString := FKeyArray[2];
    FSourceDataSet.FieldByName( 'ServiceType' ).AsString := FKeyArray[3];
    FSourceDataSet.FieldByName( 'FaciItem' ).AsString := FKeyArray[4];
    FSourceDataSet.FieldByName( 'StepNo' ).AsString :=DBController.GetStepNo;
    FSourceDataSet.FieldByName( 'Period1' ).AsInteger := aIndex;
    FSourceDataSet.FieldByName( 'Description' ).AsString := IntToStr(aIndex);
    if cmbUseMonth.ItemIndex = 0 then
    begin
      FSourceDataSet.FieldByName( 'Mon1' ).AsVariant := edtUseMonth.EditValue;
    end else
    begin
      FSourceDataSet.FieldByName( 'Mon1' ).AsInteger := aIndex;
    end;
    if FEnableLinkKey then
    begin
        FindLinkKey( aIndex );
        if FLinkKeyDataSet.RecordCount > 0 then
        begin
          FSourceDataSet.FieldByName( 'LINKKEY' ).AsString :=
            FLinkKeyDataSet.FieldByName('StepNo').AsString ;
          FSourceDataSet.FieldByName( 'LINKKEYNAME' ).AsString :=
            FLinkKeyDataSet.FieldByName('LINKKEYNAME').AsString ;
        end;
    end;
    FSourceDataSet.FieldByName( 'MonthAmt1' ).AsVariant := edtMonthAmt.EditValue;
    FSourceDataSet.FieldByName( 'DayAmt1' ).AsFloat := CbRoundTo(Abs( StrToFloat( edtMonthAmt.Text ) / 30),3);
    FSourceDataSet.FieldByName( 'DiscountAmt1' ).AsFloat := StrToFloat( edtMonthAmt.Text ) * aIndex;
    if FEnableLinkKey then
    begin
      FSourceDataSet.FieldByName( 'MonthAmt1' ).AsFloat := 0-FSourceDataSet.FieldByName( 'MonthAmt1' ).AsFloat;
      FSourceDataSet.FieldByName( 'DayAmt1' ).AsFloat := 0-FSourceDataSet.FieldByName( 'DayAmt1' ).AsFloat;
      FSourceDataSet.FieldByName( 'DiscountAmt1' ).AsFloat := 0-FSourceDataSet.FieldByName( 'DiscountAmt1' ).AsFloat;
    end;
    if VarToStr( edtMasterPeriod.EditValue ) = IntToStr( aIndex ) then
      FSourceDataSet.FieldByName( 'MasterSale' ).AsString := '1';
    FSourceDataSet.FieldByName( 'RateType1' ).AsInteger := cmbBillType.ItemIndex + 1;
    FSourceDataSet.Post;
  end;
  }
end;

procedure TfrmCD078B3_7.FormShow(Sender: TObject);
begin
  if FDefAmt <> EmptyStr then
  begin
    edtMonthAmt.Text := FDefAmt;
  end else
  begin
    edtMonthAmt.Text := EmptyStr;
  end;
  if FDefMonth <> EmptyStr then
    edtUseMonth.Text := FDefMonth
  else
    edtUseMonth.Text := EmptyStr;
end;

procedure TfrmCD078B3_7.FormCreate(Sender: TObject);
begin
  FReader := DBController.CodeReader;
  try
    FReader.Close;
    FReader.SQL.Text := Format('SELECT OrderPeriod FROM %s.SO044 ' +
                      ' WHERE SERVICETYPE=''%s''',
                      [DBController.LoginInfo.DbAccount,FKeyArray[3]]);
    FReader.Open;
    edtMasterPeriod.Text := FReader.Fields[0].AsString;
  finally
    FReader.Close;
  end;
end;

procedure TfrmCD078B3_7.FindLinkKey(aPeriod : Integer);
begin
  FLinkKeyDataSet.Filtered := False;
  FLinkKeyDataSet.Filter := Format(
  'BPCODE=''%s'' AND CITEMCODE=''%s'' AND SERVICETYPE=''%s''' +
  ' AND PERIOD1=%s',
  [FKeyArray[1],FLinkCitem,FKeyArray[3],IntToStr(aPeriod)]);
  FLinkKeyDataSet.Filtered := True;
end;

function TfrmCD078B3_7.DataValidate: Boolean;
begin
  Result := False;
  Screen.Cursor := crSQLWait;
  try
    Result := VdMustInput;
    if not Result then Exit;
    Result := VdChkData;
    if not Result then Exit;
  finally
    Screen.Cursor := crDefault;
  end;
end;

function TfrmCD078B3_7.VdMustInput: Boolean;
begin
  Result := False;
  edtUseMonth.PostEditValue;
  edtMonthAmt.PostEditValue;
  edtMasterPeriod.PostEditValue;
  edtBillPeriod.PostEditValue;
  if StrToInt(VarToStrDef(edtUseMonth.EditValue,'0')) <= 0 then
  begin
    WarningMsg( '請設定使用月數' );
    Exit;
  end;
  if StrToInt(VarToStrDef(edtBillPeriod.EditValue,'0')) <=0 then
  begin
    WarningMsg('請設定繳費期數');
    Exit;
  end;
  if StrToInt(VarToStrDef(edtMasterPeriod.EditValue,'0')) <=0 then
  begin
    WarningMsg('請設定主推期數');
    Exit;
  end;
  if StrToFloat(VarToStrDef(edtMonthAmt.EditValue,'0')) = 0 then
  begin
    WarningMsg('請設定單月金額');
    Exit;
  end;
  Result := True;

end;

function TfrmCD078B3_7.VdChkData: Boolean;
begin
  Result := False;
  if StrToInt(edtUseMonth.Text) < StrToInt(edtBillPeriod.Text ) then
  begin
    WarningMsg( '使用月數不得小於繳費期數' );
    Exit;
  end;
  if StrToInt(edtMasterPeriod.Text ) > StrToInt( edtBillPeriod.Text ) then
  begin
    WarningMsg('主推期數不得大於繳費期數' );
    Exit;
  end;
  Result := True;
end;

procedure TfrmCD078B3_7.actSaveExecute(Sender: TObject);
begin
  if not DataValidate then Exit;
  SaveData;
  Self.Close;
end;

procedure TfrmCD078B3_7.SaveData;
  var aIndex : Integer;
begin

  try
    Screen.Cursor := crSQLWait;
    for aIndex := 1  to StrToInt(VarToStrDef(edtBillPeriod.EditValue,'0'))   do
    begin
      FSourceDataSet.Append;
      FSourceDataSet.FieldByName( 'BPCode' ).AsString := FKeyArray[1];
      FSourceDataSet.FieldByName( 'CitemCode' ).AsString := FKeyArray[2];
      FSourceDataSet.FieldByName( 'ServiceType' ).AsString := FKeyArray[3];
      FSourceDataSet.FieldByName( 'FaciItem' ).AsString := FKeyArray[4];
      FSourceDataSet.FieldByName( 'StepNo' ).AsString :=DBController.GetStepNo;
      FSourceDataSet.FieldByName( 'Period1' ).AsInteger := aIndex;
      {#6893 增加Period欄位 By Kin 2014/10/20}
      FSourceDataSet.FieldByName( 'Period' ).AsInteger := aIndex;
      FSourceDataSet.FieldByName( 'Description' ).AsString := IntToStr(aIndex);
      if cmbUseMonth.ItemIndex = 0 then
      begin
        FSourceDataSet.FieldByName( 'Mon1' ).AsVariant := edtUseMonth.EditValue;
      end else
      begin
        FSourceDataSet.FieldByName( 'Mon1' ).AsInteger := aIndex;
      end;
      if FEnableLinkKey then
      begin
          FindLinkKey( aIndex );
          if FLinkKeyDataSet.RecordCount > 0 then
          begin
            FSourceDataSet.FieldByName( 'LINKKEY' ).AsString :=
              FLinkKeyDataSet.FieldByName('StepNo').AsString ;
            FSourceDataSet.FieldByName( 'LINKKEYNAME' ).AsString :=
              FLinkKeyDataSet.FieldByName('LINKKEYNAME').AsString ;
          end;
      end;
      FSourceDataSet.FieldByName( 'MonthAmt1' ).AsVariant := edtMonthAmt.Value;
      FSourceDataSet.FieldByName( 'DayAmt1' ).AsFloat := CbRoundTo(Abs(  edtMonthAmt.Value  / 30),3);
      FSourceDataSet.FieldByName( 'DiscountAmt1' ).AsFloat :=  edtMonthAmt.Value * aIndex;
      if FEnableLinkKey then
      begin
        FSourceDataSet.FieldByName( 'MonthAmt1' ).AsFloat := 0-Abs(FSourceDataSet.FieldByName( 'MonthAmt1' ).AsFloat);
        FSourceDataSet.FieldByName( 'DayAmt1' ).AsFloat := 0-Abs( FSourceDataSet.FieldByName( 'DayAmt1' ).AsFloat );
        FSourceDataSet.FieldByName( 'DiscountAmt1' ).AsFloat := 0-Abs(FSourceDataSet.FieldByName( 'DiscountAmt1' ).AsFloat);
      end;
      //#5815 如果不是主推,其它要UPD為0 By Kin 2010/10/21
      if VarToStr( edtMasterPeriod.EditValue ) = IntToStr( aIndex ) then
        FSourceDataSet.FieldByName( 'MasterSale' ).AsString := '1'
      else
        FSourceDataSet.FieldByName( 'MasterSale' ).AsString := '0';
      FSourceDataSet.FieldByName( 'RateType1' ).AsInteger := cmbBillType.ItemIndex + 1;
      FSourceDataSet.Post;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TfrmCD078B3_7.actCancelExecute(Sender: TObject);
begin
  Self.Close;
end;

procedure TfrmCD078B3_7.cmbUseMonthPropertiesChange(Sender: TObject);
begin
  if cmbUseMonth.ItemIndex = 1 then
  begin
    edtBillPeriod.Text := FDefMonth;
  end;
end;

end.
