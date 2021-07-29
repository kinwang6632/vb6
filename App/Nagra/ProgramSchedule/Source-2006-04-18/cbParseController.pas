unit cbParseController;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, DB, DBClient, ADODB,
  Provider, cbAppClass;

type
  TParseController = class(TDataModule)
    DataConnection: TADOConnection;
    DataReader: TADOQuery;
    DataWriter: TADOQuery;
    DataInsert: TADOQuery;
    DataUpdate: TADOQuery;
    DataDelete: TADOQuery;
    ParseOrderDataSet: TClientDataSet;
    DataSetProvider: TDataSetProvider;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  ParseController: TParseController;

implementation


{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TParseController.DataModuleCreate(Sender: TObject);
begin
  ParseOrderDataSet.CreateDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TParseController.DataModuleDestroy(Sender: TObject);
begin
  ParseOrderDataSet.EmptyDataSet;
  ParseOrderDataSet.Data := Null;
end;

{ ---------------------------------------------------------------------------- }

end.
