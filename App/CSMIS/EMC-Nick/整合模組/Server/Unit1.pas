unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls;

type
  TfrmMain = class(TForm)
    Panel1: TPanel;
  private
    { Private declarations }
    FQueryCount:Integer;
    FClientCount:Integer;
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

{$R *.dfm}





end.
 