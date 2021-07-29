unit dtmCommonU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  dtmBasicU, Db, ADODB;

type
  TdtmCommon = class(TdtmBasic)
    adoProjName: TADOQuery;
    adoProjInfo: TADOQuery;
    adoCommon: TADOQuery;
  private
    { Private declarations }
  public
    { Public declarations }
    procedure DeleteAllWordsInfo;
    function GetProjName : TDataSet;
    function GetProjInfo(sI_ProjName:String) : TDataSet;        
  end;

var
  dtmCommon: TdtmCommon;

implementation

{$R *.DFM}

{ TdtmCommon }

procedure TdtmCommon.DeleteAllWordsInfo;
begin
    with adoCommon do
    begin
      Close;
      SQL.Clear;
      SQL.Add('delete from WordsInfo ');
      ExecSQL;
      Close;
    end;
    
end;

function TdtmCommon.GetProjInfo(sI_ProjName: String): TDataSet;
begin
    with adoProjInfo do
    begin
      Close;
      SQL.Clear;
      SQL.Add('select sUnitName from ProjInfo ');
      SQL.Add('where sProjName=:sProjName ');
      Parameters.ParamByName('sProjName').Value := sI_ProjName;
      Open;
      Result := adoProjInfo;
    end;
end;

function TdtmCommon.GetProjName: TDataSet;
begin
    with adoProjName do
    begin
      if state = dsInactive then
        Active := True;
      Result := adoProjName;
    end;
end;

end.
