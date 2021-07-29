unit cbMainUIModule;

interface

uses
  SysUtils, Classes, Variants, DB, cbSyncObjs,
{App:}
  cbDesignPattern, cbHrHelper,
{ODAC:} MemDS, VirtualTable;

  
type
  TMainUIModule = class(TDataModule)
    dsUserBuffer: TDataSource;
    UserBuffer: TVirtualTable;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure UpdateUserBuffer(ASubject: TSubject);
  end;

var
  MainUIModule: TMainUIModule;

implementation


{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TMainUIModule.DataModuleCreate(Sender: TObject);
begin
  TBufferHelper.CreateFieldDefs( biCH005, UserBuffer );
end;

{ ---------------------------------------------------------------------------- }

procedure TMainUIModule.DataModuleDestroy(Sender: TObject);
begin
  UserBuffer.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TMainUIModule.UpdateUserBuffer(ASubject: TSubject);
var
  aSource: TVirtualTable;
begin
  aSource := TVirtualTable( TDataSubject( ASubject ).Data );
  UserBuffer.DisableControls;
  try
    UserBuffer.Assign( aSource );
  finally
    UserBuffer.EnableControls;
  end;
end;

{ ---------------------------------------------------------------------------- }


end.
