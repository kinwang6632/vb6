unit dtmBasicU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, ADODB;

type
  TdtmBasic = class(TDataModule)
    ADOConnection1: TADOConnection;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  dtmBasic: TdtmBasic;

implementation

{$R *.DFM}

end.
