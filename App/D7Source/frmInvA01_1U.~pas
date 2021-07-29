unit frmInvA01_1U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Buttons, DB, ADODB,
  cxControls, cxSplitter, cxStyles,
  cxCustomData, cxGraphics, cxFilter, cxData, cxDataStorage, cxEdit,
  cxDBData, cxGridLevel, cxClasses, cxGridCustomView,
  cxGridCustomTableView, cxGridTableView, cxGridDBTableView, cxGrid,
  cxCurrencyEdit, cxPropertiesStore, cxGridBandedTableView,
  cxGridDBBandedTableView, dxSkinsCore, dxSkinsDefaultPainters,
  dxSkinscxPCPainter;

type
  TfrmInvA01_1 = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    cxSplitter1: TcxSplitter;
    GridMasterDBTableView1: TcxGridDBTableView;
    GridMasterLevel1: TcxGridLevel;
    GridMaster: TcxGrid;
    Panel5: TPanel;
    BitBtn1: TBitBtn;
    drInv016: TDataSource;
    drInv017: TDataSource;
    Bevel1: TBevel;
    GridMasterDBTableView1SEQ: TcxGridDBColumn;
    GridMasterDBTableView1CUSTID: TcxGridDBColumn;
    GridMasterDBTableView1TITLE: TcxGridDBColumn;
    GridMasterDBTableView1TEL: TcxGridDBColumn;
    GridMasterDBTableView1BUSINESSID: TcxGridDBColumn;
    GridMasterDBTableView1ZIPCODE: TcxGridDBColumn;
    GridMasterDBTableView1INVADDR: TcxGridDBColumn;
    GridMasterDBTableView1MAILADDR: TcxGridDBColumn;
    GridMasterDBTableView1CHARGEDATE: TcxGridDBColumn;
    GridMasterDBTableView1DESCRIPTION: TcxGridDBColumn;
    GridMasterDBTableView1TAXRATE: TcxGridDBColumn;
    GridMasterDBTableView1SALEAMOUNT: TcxGridDBColumn;
    GridMasterDBTableView1TAXAMOUNT: TcxGridDBColumn;
    GridMasterDBTableView1INVAMOUNT: TcxGridDBColumn;
    GridMasterDBTableView1HOWTOCREATE: TcxGridDBColumn;
    GridMasterDBTableView1CHARGETITLE: TcxGridDBColumn;
    GridMasterDBTableView1UPTTIME: TcxGridDBColumn;
    GridMasterDBTableView1UPTEN: TcxGridDBColumn;
    lblMasterTitle: TLabel;
    Timer1: TTimer;
    cxPropertiesStore: TcxPropertiesStore;
    Panel7: TPanel;
    Panel6: TPanel;
    lblDetailTitle: TLabel;
    Bevel2: TBevel;
    Panel4: TPanel;
    GridDetail: TcxGrid;
    GridDetailDBTableView1: TcxGridDBTableView;
    GridDetailDBTableView1BILLID: TcxGridDBColumn;
    GridDetailDBTableView1BILLIDITEMNO: TcxGridDBColumn;
    GridDetailDBTableView1ITEMDESCRIPTION: TcxGridDBColumn;
    GridDetailDBTableView1CHARGEDATE: TcxGridDBColumn;
    GridDetailDBTableView1TAXDESCRIPTION: TcxGridDBColumn;
    GridDetailDBTableView1TAXRATE: TcxGridDBColumn;
    GridDetailDBTableView1QUANTITY: TcxGridDBColumn;
    GridDetailDBTableView1UNITPRICE: TcxGridDBColumn;
    GridDetailDBTableView1TAXAMOUNT: TcxGridDBColumn;
    GridDetailDBTableView1TOTALAMOUNT: TcxGridDBColumn;
    GridDetailDBTableView1STARTDATE: TcxGridDBColumn;
    GridDetailDBTableView1ENDDATE: TcxGridDBColumn;
    GridDetailDBTableView1CHARGEEN: TcxGridDBColumn;
    GridDetailDBTableView1SERVICETYPE: TcxGridDBColumn;
    GridDetailLevel1: TcxGridLevel;
    procedure FormCreate(Sender: TObject);
    procedure MasterDataAfterScroll(DataSet: TDataSet);
    procedure FormShow(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BitBtn1Click(Sender: TObject);
  private
    { Private declarations }
    FMasterDataSet: TADOQuery;
    FDetailDataSet: TADOQuery;
    FSQL: String;
    // FDataType = 1 預計開立, FDataType = 2 不開立
    FDataType: Integer;
    procedure OpenData;
    procedure InitMasterFieldAlign;
    procedure InitDetailFieldAlign;
  public
    { Public declarations }
    property SQL: String read FSQL write FSQL;
    property DataType: Integer read FDataType write FDataType;
  end;

var
  frmInvA01_1: TfrmInvA01_1;

implementation

uses
  cbUtilis, dtmMainU;


{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA01_1.FormCreate(Sender: TObject);
begin
  FDataType := 1;
  FMasterDataSet := dtmMain.adoComm;
  FDetailDataSet := dtmMain.adoComm2;
  drInv016.DataSet := FMasterDataSet;
  drInv017.DataSet := FDetailDataSet;
  FMasterDataSet.AfterScroll := Self.MasterDataAfterScroll;
  cxPropertiesStore.StorageName := GetGridStoragePath( Self.Name );
  try
    cxPropertiesStore.RestoreFrom;
  except
    DeleteFile( cxPropertiesStore.StorageName );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA01_1.FormShow(Sender: TObject);
begin
  Timer1.Enabled := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA01_1.MasterDataAfterScroll(DataSet: TDataSet);
begin
  Screen.Cursor := crSQLWait;
  try
    FDetailDataSet.Close;
    if FDataType = 1 then
    begin
      FDetailDataSet.SQL.Text :=
         '   SELECT                                      ' +
         '         A.SEQ,                                ' +
         '         A.BILLID,                             ' +
         '         A.BILLIDITEMNO,                       ' +
         '         DECODE( A.TAXTYPE, ''1'', ''應稅'',    ' +
         '                            ''2'', ''零稅率'',  ' +
         '                            ''3'', ''免稅'',    ' +
         '                            A.TAXTYPE ) AS TAXDESCRIPTION, ' +
         '         A.CHARGEDATE,                       ' +
         '         A.DESCRIPTION AS ITEMDESCRIPTION,   ' +
         '         A.QUANTITY,                         ' +
         '         A.UNITPRICE,                        ' +
         '         A.TAXRATE,                          ' +
         '         A.TAXAMOUNT,                        ' +
         '         A.TOTALAMOUNT,                      ' +
         '         A.STARTDATE,                        ' +
         '         A.ENDDATE,                          ' +
         '         A.CHARGEEN,                         ' +
         '         A.SERVICETYPE                       ' +
         '   FROM INV017 A                             ' +
         '  WHERE A.SEQ = :SEQ                         ' +
         '    AND A.SHOULDBEASSIGNED = ''Y''           ' +
         '    AND A.TOTALAMOUNT <> 0                   ' +
         '    AND A.TAXTYPE <> ''0''                   ' +
         '  ORDER BY A.BILLIDITEMNO                    ';
    end else
    begin
      FDetailDataSet.SQL.Text :=
         '   SELECT                                      ' +
         '         A.SEQ,                                ' +
         '         A.BILLID,                             ' +
         '         A.BILLIDITEMNO,                       ' +
         '         DECODE( A.TAXTYPE, ''1'', ''應稅'',    ' +
         '                            ''2'', ''零稅率'',  ' +
         '                            ''3'', ''免稅'',    ' +
         '                            A.TAXTYPE ) AS TAXDESCRIPTION, ' +
         '         A.CHARGEDATE,                       ' +
         '         A.DESCRIPTION AS ITEMDESCRIPTION,   ' +
         '         A.QUANTITY,                         ' +
         '         A.UNITPRICE,                        ' +
         '         A.TAXRATE,                          ' +
         '         A.TAXAMOUNT,                        ' +
         '         A.TOTALAMOUNT,                      ' +
         '         A.STARTDATE,                        ' +
         '         A.ENDDATE,                          ' +
         '         A.CHARGEEN,                         ' +
         '         A.SERVICETYPE                       ' +
         '   FROM INV017 A, INV016 B                   ' +
         '  WHERE A.SEQ = B.SEQ                        ' +
         '    AND ( B.SHOULDBEASSIGNED = ''N'' OR      ' +
         '          ( B.SHOULDBEASSIGNED = ''Y'' AND   ' +
         '            ( A.SHOULDBEASSIGNED = ''N'' OR  ' +
         '              A.TOTALAMOUNT = 0 OR           ' +
         '              B.TAXTYPE = ''0'' ) ) )        ' +
         '    AND A.SEQ = :SEQ                         ' +
         '  ORDER BY A.BILLIDITEMNO                    ';
    end;
    FDetailDataSet.Parameters.ParamByName( 'SEQ' ).Value :=
      FMasterDataSet.FieldByName( 'SEQ' ).AsString;  
    FDetailDataSet.Open;
  finally
    Screen.Cursor := crDefault;
  end;

//  aShouldveAssigned := 'N';
//  if FDataType = 1 then aShouldveAssigned := 'Y';
//  Screen.Cursor := crSQLWait;
//  try
//    FDetailDataSet.Close;
//    FDetailDataSet.SQL.Text :=
//       '   SELECT                                      ' +
//       '         A.SEQ,                                ' +
//       '         A.BILLID,                             ' +
//       '         A.BILLIDITEMNO,                       ' +
//       '        DECODE( A.TAXTYPE, ''1'', ''應稅'',    ' +
//       '                           ''2'', ''零稅率'',  ' +
//       '                           ''3'', ''免稅'',    ' +
//       '                           A.TAXTYPE ) AS TAXDESCRIPTION, ' +
//       '         A.CHARGEDATE,                       ' +
//       '         A.DESCRIPTION AS ITEMDESCRIPTION,   ' +
//       '         A.QUANTITY,                         ' +
//       '         A.UNITPRICE,                        ' +
//       '         A.TAXRATE,                          ' +
//       '         A.TAXAMOUNT,                        ' +
//       '         A.TOTALAMOUNT,                      ' +
//       '         A.STARTDATE,                        ' +
//       '         A.ENDDATE,                          ' +
//       '         A.CHARGEEN,                         ' +
//       '         A.SERVICETYPE                       ' +
//       '   FROM INV017 A                             ' +
//       '  WHERE A.SEQ = :SEQ                         ' +
//       '    AND A.SHOULDBEASSIGNED = :FLAG           ' +
//       '  ORDER BY A.BILLIDITEMNO                    ';
//    FDetailDataSet.Parameters.ParamByName( 'SEQ' ).Value :=
//      FMasterDataSet.FieldByName( 'SEQ' ).Value;
//    FDetailDataSet.Parameters.ParamByName( 'FLAG' ).Value := aShouldveAssigned;
//    FDetailDataSet.Open;
//  finally
//    Screen.Cursor := crDefault;
//  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA01_1.OpenData;
begin
  FMasterDataSet.Close;
  FDetailDataSet.Close;
  if ( FSQL <> '' ) then
  begin
     Screen.Cursor := crSQLWait;
     try
       FMasterDataSet.SQL.Text := FSQL;
       FMasterDataSet.AfterScroll := nil;
       try
         GridMasterDBTableView1.BeginUpdate;
         try
           FMasterDataSet.Open;
           InitMasterFieldAlign;
         finally
           GridMasterDBTableView1.EndUpdate;
         end;
         MasterDataAfterScroll( FMasterDataSet );
         InitDetailFieldAlign;
       finally
         FMasterDataSet.AfterScroll := MasterDataAfterScroll;
       end;
     finally
       Screen.Cursor := crDefault;
     end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA01_1.Timer1Timer(Sender: TObject);
begin
  Timer1.Enabled := False;
  OpenData;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA01_1.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
   FMasterDataSet.AfterScroll := nil;
   FDetailDataSet.Close;
   FMasterDataSet.Close;
   cxPropertiesStore.StorageName := GetGridStoragePath( Self.Name );
   cxPropertiesStore.StoreTo;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA01_1.BitBtn1Click(Sender: TObject);
begin
  Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA01_1.InitDetailFieldAlign;
begin
  if not FDetailDataSet.Active then Exit;
  FDetailDataSet.FieldByName( 'TAXDESCRIPTION' ).Alignment := taCenter;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA01_1.InitMasterFieldAlign;
begin
  if not FMasterDataSet.Active then Exit;
  FMasterDataSet.FieldByName( 'DESCRIPTION' ).Alignment := taCenter;
  FMasterDataSet.FieldByName( 'TAXRATE' ).Alignment := taCenter;
  FMasterDataSet.FieldByName( 'HOWTOCREATE' ).Alignment := taCenter;
end;

{ ---------------------------------------------------------------------------- }

end.
