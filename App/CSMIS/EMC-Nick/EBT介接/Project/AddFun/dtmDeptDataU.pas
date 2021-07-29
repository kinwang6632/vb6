unit dtmDeptDataU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, dtmSTabMaintainU, DB, DBTables, DBClient, Provider, ADODB;

type
  TdtmDeptData = class(TdtmSTabMaintain)
    ClientDataSet1CODENO: TIntegerField;
    ClientDataSet1DESCRIPTION: TStringField;
    ClientDataSet1STOPFLAG: TIntegerField;
    ClientDataSet1UPDTIME: TDateTimeField;
    ClientDataSet1UPDEN: TStringField;
    ClientDataSet1STOPFLAGDESC: TStringField;
    procedure dtmSTabMaintainCreate(Sender: TObject);
    procedure ClientDataSet1AfterInsert(DataSet: TDataSet);
    procedure ClientDataSet1BeforePost(DataSet: TDataSet);
    procedure ClientDataSet1CalcFields(DataSet: TDataSet);
  private
    { Private declarations }
  public
    { Public declarations }
  protected
    procedure SetMasterData(var MstQry: TDataSet; var MasterQuery : TQuery) ; override ;

  end;

var
  dtmDeptData: TdtmDeptData;

implementation

uses dtmMainU;

{$R *.dfm}

{ TdtmCompData }

procedure TdtmDeptData.SetMasterData(var MstQry: TDataSet;
  var MasterQuery: TQuery);
begin
  inherited;
    MstQry := ClientDataSet1;

end;

procedure TdtmDeptData.dtmSTabMaintainCreate(Sender: TObject);
begin
  inherited;
//
end;

procedure TdtmDeptData.ClientDataSet1AfterInsert(DataSet: TDataSet);
begin
  inherited;
    DataSet.FieldByName('STOPFLAG').AsInteger := 0;
end;

procedure TdtmDeptData.ClientDataSet1BeforePost(DataSet: TDataSet);
begin
  inherited;
    DataSet.FieldByName('UpdTime').AsDateTime := now;
    DataSet.FieldByName('UpdEn').AsString := dtmMain.getUserName;    
end;

procedure TdtmDeptData.ClientDataSet1CalcFields(DataSet: TDataSet);
begin
  inherited;
    if DataSet.FieldByName('STOPFLAG').AsInteger = 1 then
      DataSet.FieldByName('STOPFLAGDESC').AsString := '¬O'
    else
      DataSet.FieldByName('STOPFLAGDESC').AsString := '§_';    
end;

end.
