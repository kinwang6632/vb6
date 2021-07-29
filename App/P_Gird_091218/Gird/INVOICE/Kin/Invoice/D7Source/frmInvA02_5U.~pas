unit frmInvA02_5U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Buttons, ADODB, DB, DBClient,
  cxControls, cxContainer, cxEdit, cxLabel, cxStyles, cxCustomData, cxGraphics,
  cxFilter, cxData,  cxDataStorage, cxDBData, cxGridLevel, cxClasses, cxGridCustomView,
  cxGridCustomTableView, cxGridTableView, cxGridDBTableView, cxGrid,
  cxGridBandedTableView, cxGridDBBandedTableView, dxSkinsCore,
  dxSkinsDefaultPainters, dxSkinscxPCPainter;

type
  TfrmInvA02_5 = class(TForm)
    Panel1: TPanel;
    btnCancel: TBitBtn;
    StaticText1: TStaticText;
    StaticText2: TStaticText;
    btnOk: TBitBtn;
    lblCustID: TcxLabel;
    lblCustSName: TcxLabel;
    cxGridLevel1: TcxGridLevel;
    cxGrid: TcxGrid;
    cxGridView: TcxGridDBBandedTableView;
    DataSource1: TDataSource;
    cxGridViewTITLESNAME: TcxGridDBBandedColumn;
    cxGridViewTITLENAME: TcxGridDBBandedColumn;
    cxGridViewBUSINESSID: TcxGridDBBandedColumn;
    cxGridViewMZIPCODE: TcxGridDBBandedColumn;
    cxGridViewMAILADDR: TcxGridDBBandedColumn;
    cxGridViewINVADDR: TcxGridDBBandedColumn;
    cxGridViewMEMO: TcxGridDBBandedColumn;
    procedure FormShow(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure btnOkClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure cxGridViewDblClick(Sender: TObject);
  private
    { Private declarations }
    FCustId: String;
    FTitleDataSet: TClientDataSet;
    function GetCustomerName: String;
    procedure OpenInv019DataSet;
    procedure CreateFieldDefs;
    procedure UpdateSelectedRecord;
  public
    { Public declarations }
    constructor Create(const aCustId: String); overload;
    property TitleDataSet: TClientDataSet read FTitleDataSet;
  end;

var
  frmInvA02_5: TfrmInvA02_5;

implementation

uses cbUtilis, dtmMainHU, dtmMainU, frmMainU;

{$R *.dfm}

{ frmInvA02_5 }

{ ---------------------------------------------------------------------------- }

constructor TfrmInvA02_5.Create(const aCustId: String);
begin
  inherited Create( Application );
  FCustId := aCustId;
  FTitleDataSet := TClientDataSet.Create( nil );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_5.FormShow(Sender: TObject);
begin
  Caption := frmMain.GetFormTitleString( 'A02_5', '客戶抬頭' );
  OpenInv019DataSet;
  lblCustID.Caption := FCustId;
  lblCustSName.Caption := GetCustomerName;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_5.FormDestroy(Sender: TObject);
begin
  FTitleDataSet.Data := Null;
  FTitleDataSet.FieldDefs.Clear;
  FTitleDataSet.Free;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA02_5.GetCustomerName: String;
begin
  Result := '';
  if DataSource1.DataSet.Active then
    Result := DataSource1.DataSet.FieldByName( 'TITLESNAME' ).AsString;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_5.OpenInv019DataSet;
var
  aDataSet: TADOQuery;
begin
  aDataSet := dtmMainH.adoComm;
  aDataSet.Close;
  { 不須檢查 CompId, 只要是 CustId 符合的就帶出取第一筆資料 }
  aDataSet.SQL.Text := Format(
    '  SELECT A.TITLESNAME,                   ' +
    '         A.TITLENAME,                    ' +
    '         A.BUSINESSID,                   ' +
    '         A.MZIPCODE,                     ' +
    '         A.MAILADDR,                     ' +
    '         A.INVADDR,                      ' +
    '         A.MEMO                          ' +
    '    FROM INV019 A, INV002 B              ' +
    '   WHERE A.IDENTIFYID1 = B.IDENTIFYID1   ' +
    '     AND A.IDENTIFYID2 = B.IDENTIFYID2   ' +
    '     AND A.COMPID = B.COMPID             ' +
    '     AND A.CUSTID = B.CUSTID             ' +
    '     AND A.CUSTID = ''%s''               ' +
    '     AND A.IDENTIFYID1 = ''%s''          ' +
    '     AND A.IDENTIFYID2 = %s              ' +
    '   ORDER BY A.TITLEID                    ',
    [ FCustId, IDENTIFYID1, IDENTIFYID2 ] );
  aDataSet.Open;      
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_5.btnCancelClick(Sender: TObject);
begin
  Self.ModalResult := mrCancel;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_5.btnOkClick(Sender: TObject);
begin
  CreateFieldDefs;
  UpdateSelectedRecord;
  Self.ModalResult := mrOk;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_5.CreateFieldDefs;
var
  aIndex: Integer;
  aDataSet: TADOQuery;
begin
  FTitleDataSet.Data := Null;
  FTitleDataSet.FieldDefs.Clear;
  aDataSet := TADOQuery( DataSource1.DataSet );
  for aIndex := 0 to aDataSet.FieldCount - 1 do
  begin
    if FTitleDataSet.FieldDefs.IndexOf( aDataSet.Fields[aIndex].FieldName ) <= -1 then
    begin
      FTitleDataSet.FieldDefs.Add( aDataSet.Fields[aIndex].FieldName,
        aDataSet.Fields[aIndex].DataType, aDataSet.Fields[aIndex].Size );
    end;
  end;
  FTitleDataSet.CreateDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_5.UpdateSelectedRecord;
var
  aIndex: Integer;
  aDataSet: TADOQuery;
  aField: TField;
begin
  aDataSet := TADOQuery( DataSource1.DataSet );
  FTitleDataSet.Append;
  for aIndex := 0 to aDataSet.FieldCount - 1 do
  begin
    aField := FTitleDataSet.FindField( aDataSet.Fields[aIndex].FieldName );
    if Assigned( aField ) then
      aField.Value := aDataSet.Fields[aIndex].Value;
  end;
  FTitleDataSet.Post;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_5.cxGridViewDblClick(Sender: TObject);
begin
  btnOkClick( btnOk );
end;

{ ---------------------------------------------------------------------------- }

end.
