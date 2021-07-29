unit cbDataController;

interface

uses
  Windows, SysUtils, Classes, DB, ADODB, cbSoInfo, xmldom, XMLIntf,
  msxmldom, XMLDoc;

type
  TDataController = class(TDataModule)
    LogConnection: TADOConnection;
    LogWriter: TADOQuery;
    LogReader: TADOQuery;
    CMConnection: TADOConnection;
    CMDataReader: TADOQuery;
    CMDataWriter: TADOQuery;
    XMLDoc: TXMLDocument;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  DataController: TDataController;

implementation


{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TDataController.DataModuleCreate(Sender: TObject);
begin
  {...}
end;

{ ---------------------------------------------------------------------------- }

procedure TDataController.DataModuleDestroy(Sender: TObject);
begin
  {...}
end;

{ ---------------------------------------------------------------------------- }

end.
