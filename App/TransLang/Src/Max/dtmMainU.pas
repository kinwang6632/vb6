unit dtmMainU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  dtmBasicU, Db, ADODB, DBClient;

type
  TdtmMain = class(TdtmBasic)
    ADOPrjInfoQry: TADOQuery;
    CDSTmpQry: TClientDataSet;
    ADOExecSQL: TADOQuery;
    ADOPrjInfoQrynSeq: TIntegerField;
    ADOPrjInfoQrysProjName: TWideStringField;
    ADOPrjInfoQrysUnitName: TWideStringField;
  private
    { Private declarations }
  public
    Procedure ExecSQL(StyleFlag:integer);
//    Procedure InsertCmd;
    { Public declarations }
  end;

var
  dtmMain: TdtmMain;

implementation

{$R *.DFM}

procedure TdtmMain.ExecSQL(StyleFlag:integer);
begin
  ADOPrjInfoQry.close;
  case StyleFlag of
    1:
    begin
		  ADOExecSQL.SQL.Clear;
		  ADOExecSQL.SQL.Add('Delete from ProjInfo');
		  ADOExecSQL.ExecSQL;
    end;
  end;
  ADOPrjInfoQry.open;
end;

{procedure TdtmMain.InsertCmd;
begin
  CDSTmpQry.First;
  while not CDSTmpQry.Eof do
  begin
    if ADOPrjInfoQry.Lookup('sUnitName',CDSTmpQry.Fieldbyname('sUnitName').asstring,
      'sProjName') = Null then
      ADOPrjInfoQry.AppendRecord([nil,CDSTmpQry.fieldbyname('sProjName').asstring,
        CDSTmpQry.fieldbyname('sUnitName').asstring]);
      CDSTmpQry.Next;
  end;
end;}

end.
