unit dtmMainU;

interface

uses
  SysUtils, Classes, DBXpress, DB, SqlExpr, FMTBcd, Controls, Dialogs,
  ADODB;

type
  TdtmMain = class(TDataModule)
    ADOConnection1: TADOConnection;
    qryCommon: TADOQuery;
    procedure DataModuleCreate(Sender: TObject);
    procedure ADOConnection1Login(Sender: TObject; Username,
      Password: String);
  private
    { Private declarations }
  public
    { Public declarations }
    function getSequence(dI_Date:TDate):String;
  end;

var
  dtmMain: TdtmMain;

implementation

uses UdateTimeu;

{$R *.dfm}

procedure TdtmMain.DataModuleCreate(Sender: TObject);
var
    sL_DbUserID, sL_DbPassword, sL_DbAlias : String;
begin
    sL_DbUserID := 'v30';
    sL_DbPassword := 'v30';
    sL_DbAlias := 'mis';
    ADOConnection1.ConnectionString := 'Provider=OraOLEDB.Oracle.1;Password=' + sL_DbPassword + ';Persist Security Info=True;User ID=' + sL_DbUserID + ';Data Source=' + sL_DbAlias;
    ADOConnection1.Connected := true;
end;

function TdtmMain.getSequence(dI_Date: TDate): String;
var
    nL_MaxSeq : Integer;
    sL_DateStr, sL_MaxSeq, sL_UsefulSeq : String;
begin
    sL_DateStr := TUdateTime.GetPureDateStr(dI_Date,'');

    with qryCommon do
    begin

      Close;
      SQL.Clear;
      SQL.Add('select max(seq) SEQ from NDS001 ');
      SQL.Add('where substr(seq,1,8)=:DateStr');
      Parameters.ParamByName('DateStr').Value :=  sL_DateStr;
      Open;
      sL_MaxSeq := FieldByName('SEQ').AsString;

      nL_MaxSeq := StrToInt(Copy(sL_MaxSeq,9,8));
      Close;
    end;

    sL_UsefulSeq := Format('%.8d',[nL_MaxSeq+1]);
    result := sL_DateStr+sL_UsefulSeq;
//    select max(seq) SEQ from NDS001 where substr(seq,1,8)='20010708'
end;

procedure TdtmMain.ADOConnection1Login(Sender: TObject; Username,
  Password: String);
begin
    {
    Username := 'gicmis';
    Password := 'may';
    }
end;

end.
