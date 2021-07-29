unit cbTestMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, ExtCtrls, DB, ADODB;

type
  TTestMain = class(TForm)
    Panel1: TPanel;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Panel2: TPanel;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    Memo1: TMemo;
    Memo2: TMemo;
    TabSheet3: TTabSheet;
    Memo3: TMemo;
    TabSheet4: TTabSheet;
    Memo4: TMemo;
    Button4: TButton;
    ADOQuery1: TADOQuery;
    ADOConnection1: TADOConnection;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
  private
    { Private declarations }
    FOleObj: OleVariant;
  public
    { Public declarations }
  end;

var
  TestMain: TTestMain;

implementation

uses ComObj;

{$R *.dfm}

procedure TTestMain.FormCreate(Sender: TObject);
begin
  PageControl1.ActivePageIndex := 0;
end;

procedure TTestMain.FormDestroy(Sender: TObject);
begin
  FOleObj := Unassigned;
end;

procedure TTestMain.Button1Click(Sender: TObject);
var
  aReturn: String;
begin
  FOleObj := CreateOleObject( 'prjE2C.E2C' );
  FOleObj.E2C_DLL1( Memo1.Lines.Text, 'EMPTY' );
  aReturn := FOleObj.E2C_DLL1_RESULT;
  if ( aReturn <> EmptyStr ) then Application.MessageBox( PChar( aReturn ),
    '警告', MB_OK + MB_ICONWARNING );
  FOleObj := Unassigned;
end;

procedure TTestMain.Button2Click(Sender: TObject);
var
  aReturn: String;
begin
  FOleObj := CreateOleObject( 'prjE2C.E2C' );
  FOleObj.E2C_DLL1( 'EMPTY', Memo2.Lines.Text );
  aReturn := FOleObj.E2C_DLL1_RESULT;
  if ( aReturn <> EmptyStr ) then  Application.MessageBox( PChar( aReturn ),
    '警告', MB_OK + MB_ICONWARNING );
  FOleObj := Unassigned;
end;

procedure TTestMain.Button3Click(Sender: TObject);
var
  aReturn: String;
begin
  FOleObj := CreateOleObject( 'prjE2C.E2C' );
  FOleObj.E2C_DLL2( Memo3.Lines.Text, EmptyStr, '0,6' );
  aReturn := FOleObj.E2C_DLL2_RESULT;
  if ( aReturn <> EmptyStr ) then  Application.MessageBox( PChar( aReturn ),
    '警告', MB_OK + MB_ICONWARNING );
  FOleObj := Unassigned;
end;

procedure TTestMain.Button4Click(Sender: TObject);
var
  aReturn: String;
begin
  FOleObj := CreateOleObject( 'prjE2C.E2C' );
  FOleObj.E2C_DLL3( Memo4.Lines.Text );
  aReturn := FOleObj.E2C_DLL3_RESULT;
  if ( aReturn <> EmptyStr ) then  Application.MessageBox( PChar( aReturn ),
    '警告', MB_OK + MB_ICONWARNING );
  FOleObj := Unassigned;
end;

end.

