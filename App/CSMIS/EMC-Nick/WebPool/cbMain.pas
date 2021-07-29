unit cbMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, DBGrids, DB, ADODB;

type
  TfmMain = class(TForm)
    Button1: TButton;
    Edit1: TEdit;
    DBGrid1: TDBGrid;
    DataSource1: TDataSource;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmMain: TfmMain;

implementation

{$R *.dfm}

uses WebPool_TLB;

procedure TfmMain.Button1Click(Sender: TObject);
var
  aPool: PoolManager;
  aConn: _Connection;
  aEffect: OleVariant;
  aRecord: _Recordset;
  aR: TADODataSet;
begin
  aPool := CoPoolManager.Create;
  aConn := ( IInterface( aPool.GetConnection ) as _Connection );
  aRecord := aConn.Execute( Edit1.Text, aEffect, 0 );
  aConn := nil;
  aR := TADODataSet.Create( nil );
  aR.Recordset := aRecord;
  DataSource1.DataSet := aR;
end;

end.
