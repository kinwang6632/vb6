unit dtmReportModule;
  {
    ���n:
      1. frxA05_1.CloseDataSource ���i�]�w�� True, �|�v�T�o���M����,
         ��s INV007
      2. frxA07_1.CloseDataSource ���i�]�w�� True, �|�v�T������ƿ�J
  }

interface

uses
  SysUtils, Classes, frxClass, frxDBSet, DB, ADODB, DBClient,
  frxDesgn, frxRes, frxChBox, frxExportXLS, frxBarcode;

type
  TdtmReport = class(TDataModule)
    ADOMaster: TADOQuery;
    ADODetail: TADOQuery;
    cdsTempory: TClientDataSet;
    frxA01_1: TfrxDBDataset;
    frxA01_2: TfrxDBDataset;
    frxA05_1: TfrxDBDataset;
    frxA07_1: TfrxDBDataset;
    frxB04_1: TfrxDBDataset;
    frxC01_1: TfrxDBDataset;
    frxC03_1: TfrxDBDataset;
    frxC04_1: TfrxDBDataset;
    frxC05_1: TfrxDBDataset;
    frxC01_2: TfrxDBDataset;
    frxB06_1: TfrxDBDataset;
    frxDBDataset1: TfrxDBDataset;
    frxMainReport: TfrxReport;
    procedure DataModuleCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure AddA05DataField(aAutoCreateNum: Integer);
  end;

var
  dtmReport: TdtmReport;

implementation

uses cbUtilis, dtmMainU;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TdtmReport.DataModuleCreate(Sender: TObject);
begin
  if FileExists( IncludeTrailingPathDelimiter(
    ExtractFilePath( ParamStr( 0 ) ) ) + 'frxrcTaiwan.frc' ) then
  begin
    frxResources.LoadFromFile( IncludeTrailingPathDelimiter(
      ExtractFilePath( ParamStr( 0 ) ) ) + 'frxrcTaiwan.frc' );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmReport.AddA05DataField(aAutoCreateNum: Integer);
var
  aIndex: Integer;
begin
  frxA05_1.FieldAliases.Clear;
  frxA05_1.FieldAliases.Add( 'CHECKNO=�ˬd�X' );
  frxA05_1.FieldAliases.Add( 'INVID=�o�����X' );
  frxA05_1.FieldAliases.Add( 'BUSINESSID=�Τ@�s��' );
  frxA05_1.FieldAliases.Add( 'INVDATE=�o�����' );
  frxA05_1.FieldAliases.Add( 'CUSTID=�Ȥ�N�X' );
  frxA05_1.FieldAliases.Add( 'CUSTSNAME=�Ȥ�²��' );
  frxA05_1.FieldAliases.Add( 'INVTITLE=�o�����Y' );
  frxA05_1.FieldAliases.Add( 'ZIPCODE=�l���ϸ�' );
  frxA05_1.FieldAliases.Add( 'MAILADDR=�l�H�a�}' );
  frxA05_1.FieldAliases.Add( 'INVADDR=�o���a�}' );
  frxA05_1.FieldAliases.Add( 'INSTADDR=�˾��a�}' );
  frxA05_1.FieldAliases.Add( 'INVFORMAT=�o���榡' );
  frxA05_1.FieldAliases.Add( 'TAXTYPE=�|�O' );
  frxA05_1.FieldAliases.Add( 'SALEAMOUNT=�P���B' );
  frxA05_1.FieldAliases.Add( 'TAXAMOUNT=�|�B' );
  frxA05_1.FieldAliases.Add( 'INVAMOUNT=�o���`���B' );

  if aAutoCreateNum <= 0 then aAutoCreateNum := 5;

  for aIndex := 1 to aAutoCreateNum do
  begin
    if aIndex >= 100 then Break;
    frxA05_1.FieldAliases.Add( Format( 'BILLID%d=���O�渹(%s)', [aIndex, MappingChineseNumber( aIndex )] ) );
    frxA05_1.FieldAliases.Add( Format( 'DESCRIPTION%d=�~�W(%s)', [aIndex, MappingChineseNumber( aIndex )] ) );
    frxA05_1.FieldAliases.Add( Format( 'STARTDATE%d=���Ĥ�_(%s)', [aIndex, MappingChineseNumber( aIndex )] ) );
    frxA05_1.FieldAliases.Add( Format( 'ENDDATE%d=���Ĥ騴(%s)', [aIndex, MappingChineseNumber( aIndex )] ) );
    frxA05_1.FieldAliases.Add( Format( 'QUANTITY%d=�ƶq(%s)', [aIndex, MappingChineseNumber( aIndex )] ) );
    frxA05_1.FieldAliases.Add( Format( 'UNITPRICE%d=���(%s)', [aIndex, MappingChineseNumber( aIndex )] ) );
    frxA05_1.FieldAliases.Add( Format( 'SALEAMOUNT%d=�P���B(%s)', [aIndex, MappingChineseNumber( aIndex )] ) );
    frxA05_1.FieldAliases.Add( Format( 'TOTALAMOUNT%d=�`���B(%s)', [aIndex, MappingChineseNumber( aIndex )] ) );
    frxA05_1.FieldAliases.Add( Format( 'FACISNO%d=�]�Ƹ��X(%s)', [aIndex, MappingChineseNumber( aIndex )] ) );
  end;
  frxA05_1.FieldAliases.Add( 'MEMO1=�Ƶ�' );
  frxA05_1.FieldAliases.Add( 'ACCOUNTNO=�H�Υd��' );
  frxA05_1.FieldAliases.Add( 'MAININVID=�D�o�����X' );
  frxA05_1.FieldAliases.Add( 'MAINSALEAMOUNT=�D�o���`�P���B' );
  frxA05_1.FieldAliases.Add( 'MAINTAXAMOUNT=�D�o���`�|�B' );
  frxA05_1.FieldAliases.Add( 'MAININVAMOUNT=�D�o���`���B' );
  frxA05_1.FieldAliases.Add( 'HOWTOCREATE=������Ʀp�󲣥�' );
  frxA05_1.FieldAliases.Add( 'INVUSEDESC=�o���γ~' );
  {#6600 �W�[�o�������H���X By Kin 2013/10/02}
  frxA05_1.FieldAliases.Add( 'RANDOMNUM=�o�������H���X' );
  {#6629 �W�[�ӳ��榡 By Kin 2013/11/01}
  frxA05_1.FieldAliases.Add( 'APPLYFORMAT=�ӳ��榡' );
end;

{ ---------------------------------------------------------------------------- }

end.