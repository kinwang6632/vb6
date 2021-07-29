unit frmInvA02U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, ExtCtrls, Buttons, DB, ADODB, Menus,
  cxControls, cxContainer, cxEdit, cxTextEdit,
  cxMaskEdit, cxDropDownEdit, cxCalendar, cxStyles, cxCustomData,
  cxGraphics, cxFilter, cxData, cxDataStorage, cxDBData, cxGridLevel,
  cxClasses, cxGridCustomView, cxGridCustomTableView, cxGridTableView,
  cxGridDBTableView, cxGrid, cxGridBandedTableView, cxGridDBBandedTableView,
  cxCurrencyEdit, cxPropertiesStore, cxSpinEdit, cxButtonEdit,
  cxDBExtLookupComboBox, cxGroupBox, cxRadioGroup,
  cxLookAndFeelPainters, cxButtons, dxSkinsCore, dxSkinsDefaultPainters,
  dxSkinscxPCPainter;

type
  TfrmInvA02 = class(TForm)
    dsrInv007: TDataSource;
    StatusBar1: TStatusBar;
    Panel1: TPanel;
    pnlFun1: TPanel;
    btnQuery: TcxButton;
    btnExit: TcxButton;
    btnCreate: TcxButton;
    btnView: TcxButton;
    BtnClear: TcxButton;
    pnlMaster: TPanel;
    cxGrid1Level1: TcxGridLevel;
    cxGrid1: TcxGrid;
    Bevel1: TBevel;
    cxGrid1DBBandedTableView1: TcxGridDBBandedTableView;
    cxGrid1DBBandedTableView1INVID: TcxGridDBBandedColumn;
    cxGrid1DBBandedTableView1INVDATE: TcxGridDBBandedColumn;
    cxGrid1DBBandedTableView1CUSTID: TcxGridDBBandedColumn;
    cxGrid1DBBandedTableView1CUSTSNAME: TcxGridDBBandedColumn;
    cxGrid1DBBandedTableView1CHARGEDATE: TcxGridDBBandedColumn;
    cxGrid1DBBandedTableView1INVTITLE: TcxGridDBBandedColumn;
    cxGrid1DBBandedTableView1ZIPCODE: TcxGridDBBandedColumn;
    cxGrid1DBBandedTableView1INVADDR: TcxGridDBBandedColumn;
    cxGrid1DBBandedTableView1MAILADDR: TcxGridDBBandedColumn;
    cxGrid1DBBandedTableView1BUSINESSID: TcxGridDBBandedColumn;
    cxGrid1DBBandedTableView1INVFORMAT: TcxGridDBBandedColumn;
    cxGrid1DBBandedTableView1TAXTYPE: TcxGridDBBandedColumn;
    cxGrid1DBBandedTableView1TAXRATE: TcxGridDBBandedColumn;
    cxGrid1DBBandedTableView1SALEAMOUNT: TcxGridDBBandedColumn;
    cxGrid1DBBandedTableView1TAXAMOUNT: TcxGridDBBandedColumn;
    cxGrid1DBBandedTableView1INVAMOUNT: TcxGridDBBandedColumn;
    cxGrid1DBBandedTableView1ISOBSOLETE: TcxGridDBBandedColumn;
    cxGrid1DBBandedTableView1OBSOLETEREASON: TcxGridDBBandedColumn;
    cxGrid1DBBandedTableView1MEMO1: TcxGridDBBandedColumn;
    cxGrid1DBBandedTableView1MEMO2: TcxGridDBBandedColumn;
    cxGrid1DBBandedTableView1UPTTIME: TcxGridDBBandedColumn;
    cxGrid1DBBandedTableView1UPTEN: TcxGridDBBandedColumn;
    btnColumnChooce: TSpeedButton;
    cxPropertiesStore: TcxPropertiesStore;
    cxGrid1DBBandedTableView1MAININVID: TcxGridDBBandedColumn;
    rgpQueryCond: TcxRadioGroup;
    edtCustID: TcxTextEdit;
    edtBusinessID: TcxTextEdit;
    txtInvDate: TcxDateEdit;
    txtChargeDate: TcxDateEdit;
    edtBeginInvID: TcxTextEdit;
    procedure FormShow(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnQueryClick(Sender: TObject);
    procedure btnCreateClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure rgpQueryCondClick(Sender: TObject);
    procedure BtnClearClick(Sender: TObject);
    procedure InputEditEnter(Sender: TObject);
    procedure N11Click(Sender: TObject);
    procedure btnColumnChooceClick(Sender: TObject);
    procedure cxGrid1DBBandedTableView1DblClick(Sender: TObject);
    procedure btnViewClick(Sender: TObject);
    procedure txtInvDatePropertiesValidate(Sender: TObject;
      var DisplayValue: Variant; var ErrorText: TCaption;
      var Error: Boolean);
    procedure txtInvDateExit(Sender: TObject);
  private
    { Private declarations }
    FMasterDataSet: TADOQuery;
    FCallFromProgram: TClass;
    procedure SetGridDisplayProperty;
  public
    { Public declarations }
    property CallFromProgram: TClass read FCallFromProgram write FCallFromProgram;
  end;

var
  frmInvA02: TfrmInvA02;

implementation

uses cbUtilis, dtmMainHU, frmMainU, dtmMainU,
     frmInvA02_1U, frmInvA02_4U, Ustru; 

{$R *.dfm}

{ TfrmInvA02 }

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02.FormShow(Sender: TObject);
begin
  if ( FCallFromProgram = TfrmMain ) then
    Caption := frmMain.GetFormTitleString( 'A0C', '發票查詢' )
  else
    Caption := frmMain.GetFormTitleString( 'A02', '開立發票' );
  txtInvDate.Date := Date;
  txtChargeDate.Date := Date;
  FMasterDataSet := dtmMainH.adoInv007;
  cxPropertiesStore.StorageName := GetGridStoragePath( Self.Name );
  try
    cxPropertiesStore.RestoreFrom;
  except
    DeleteFile( cxPropertiesStore.StorageName );
  end;
  if FCallFromProgram = TfrmMain then
  begin
    btnCreate.Enabled := False;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  FMasterDataSet.Close;
  cxPropertiesStore.StorageName := GetGridStoragePath( Self.Name );
  cxPropertiesStore.StoreTo;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02.btnExitClick(Sender: TObject);
begin
  Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02.btnQueryClick(Sender: TObject);
var
  aCustId, aBusinessId, aWhereCond, aBeginInvId: String;
begin
    case rgpQueryCond.ItemIndex of
      2:  //發票日期
       aWhereCond := ' AND INVDATE = ' + TUstr.getOracleSQLDateStr( txtInvDate.Date );
      3:  //收費日期
       aWhereCond := ' AND CHARGEDATE = ' + TUstr.getOracleSQLDateStr( txtChargeDate.Date);
      0:  //客編
       begin
         aCustId := Trim( edtCustID.Text);
         if ( aCustId = '' ) then
         begin
           WarningMsg( '請輸入客編!' );
           edtCustID.SetFocus;
           Exit;
         end;
         aWhereCond := ' AND CUSTID = ' + QuotedStr( aCustId );
       end;
      1:  //統編
       begin
          aBusinessId := Trim( edtBusinessID.Text );
        if ( aBusinessId = '' ) then
        begin
          WarningMsg( '請輸入統編!' );
          edtBusinessID.SetFocus;
          Exit;
        end;
        aWhereCond := ' AND BUSINESSID = ' + QuotedStr( aBusinessID );
       end;
      4://發票號碼
       begin
        aBeginInvId := Trim( edtBeginInvID.Text);
        if ( aBeginInvId = '' ) then
        begin
          WarningMsg( '請輸入發票/主發票號碼!' );
          edtBeginInvID.SetFocus;
          Exit;
        end;
        aWhereCond := ' AND ( INVID = ' + QuotedStr( aBeginInvId ) +
          ' OR MAININVID = ' +  QuotedStr( aBeginInvId ) + ' ) ';
       end;
    end;
    Screen.Cursor := crSQLWait;
    try
      btnView.Enabled := False;
      BtnClear.Enabled := False;
      if ( dtmMainH.getInv007Data( aWhereCond ) <= 0 ) then
      begin
         WarningMsg( '本次查詢無發票資料!' );
         Exit;
      end;
      btnView.Enabled := True;
      BtnClear.Enabled := True;
      cxGrid1.SetFocus;
      SetGridDisplayProperty;
    finally
      Screen.Cursor := crDefault;
    end;

end;

{ ---------------------------------------------------------------------------- }


procedure TfrmInvA02.btnCreateClick(Sender: TObject);
var
  aInvId: String;
  aForm: TfrmInvA02_4;
begin
  aInvId := EmptyStr;
  aForm := TfrmInvA02_4.Create( aInvId );
  try
    aForm.ShowModal;
  finally
    aForm.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02.btnViewClick(Sender: TObject);
begin
  cxGrid1DBBandedTableView1DblClick( cxGrid1DBBandedTableView1 );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02.rgpQueryCondClick(Sender: TObject);
begin
  case rgpQueryCond.ItemIndex of
    0: edtCustID.SetFocus;
    1: edtBusinessID.SetFocus;
    2: txtInvDate.SetFocus;
    3: txtChargeDate.SetFocus;
    4: edtBeginInvID.SetFocus;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02.BtnClearClick(Sender: TObject);
begin
  FMasterDataSet.Close;
  rgpQueryCond.ItemIndex := 0;
  edtCustID.Clear;
  edtBusinessID.Clear;
  edtBeginInvID.Clear;
  txtInvDate.Date := Date;
  txtChargeDate.Date := Date;
  btnView.Enabled := False;
  BtnClear.Enabled := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02.InputEditEnter(Sender: TObject);
begin
  rgpQueryCond.ItemIndex := TControl( Sender ).Tag
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02.N11Click(Sender: TObject);
begin
  cxGrid1DBBandedTableView1.Controller.Customization := True;
end;

procedure TfrmInvA02.btnColumnChooceClick(Sender: TObject);
begin
  cxGrid1DBBandedTableView1.Controller.Customization := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02.SetGridDisplayProperty;
begin
  FMasterDataSet.DisableControls;
  try
    FMasterDataSet.FieldByName( 'TAXTYPE' ).Alignment := taCenter;
    FMasterDataSet.FieldByName( 'ISOBSOLETE' ).Alignment := taCenter;
    FMasterDataSet.FieldByName( 'INVFORMAT' ).Alignment := taCenter;
  finally
    FMasterDataSet.EnableControls;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02.cxGrid1DBBandedTableView1DblClick(Sender: TObject);
var
  aForm: TfrmInvA02_4;
  aInvId: String;
begin
  Screen.Cursor := crSQLWait;
  try
    if not FMasterDataSet.Active then Exit;
    aInvId := FMasterDataSet.FieldByName( 'INVID' ).AsString;
    aForm := TfrmInvA02_4.Create( aInvId );
    try
      aForm.CallFromProgram := FCallFromProgram;
      aForm.ShowModal;
    finally
      aForm.Free;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02.txtInvDatePropertiesValidate(Sender: TObject;
  var DisplayValue: Variant; var ErrorText: TCaption; var Error: Boolean);
begin
  if Error then
  begin
    WarningMsg( '輸入的日期不正確。' );
    ErrorText := '';
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02.txtInvDateExit(Sender: TObject);
begin
  if TcxDateEdit( Sender ).EditText = '' then
    TcxDateEdit( Sender ).Date := Date;
end;

{ ---------------------------------------------------------------------------- }

end.


