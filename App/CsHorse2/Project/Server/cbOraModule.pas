unit cbOraModule;

interface

uses
  SysUtils, Classes, DB,
{ ODAC: }
  MemDS, DBAccess, Ora, OraSmart, DAScript, OraScript;

type
  TOraModule = class(TDataModule)
    OraSession: TOraSession;
    OraTable: TOraTable;
    OraSQL: TOraSQL;
    OraQuery: TOraQuery;
    OraScript: TOraScript;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;


implementation

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TOraModule.DataModuleCreate(Sender: TObject);
begin
  {}
end;

{ ---------------------------------------------------------------------------- }

procedure TOraModule.DataModuleDestroy(Sender: TObject);
begin
  {}
end;

{ ---------------------------------------------------------------------------- }

end.
