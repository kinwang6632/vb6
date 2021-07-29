unit cbTestMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, ExtCtrls;

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
    TabSheet5: TTabSheet;
    Memo5: TMemo;
    Button5: TButton;
    TabSheet6: TTabSheet;
    Edit1: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    Edit2: TEdit;
    Button6: TButton;
    Label3: TLabel;
    Edit3: TEdit;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
  private
    { Private declarations }
    FOleObj: OleVariant;
  public
    { Public declarations }
  end;

var
  TestMain: TTestMain;

implementation

{ LbCipher, LbString --> Turbopower 的加解密 Library ( LockBox ) }

uses ComObj, LbCipher, LbString;

{$R *.dfm}

procedure TTestMain.FormCreate(Sender: TObject);
begin
  PageControl1.ActivePageIndex := 0;
end;

procedure TTestMain.FormDestroy(Sender: TObject);
begin
  FOleObj := Null;
end;

procedure TTestMain.Button1Click(Sender: TObject);
var
  aReturn: String;
begin
  FOleObj := CreateOleObject( 'LibCM2CT.CM2CT' );
  aReturn := FOleObj.CM2CT_DLL1( Memo1.Lines.Text, 'EMPTY' );
  if ( aReturn <> EmptyStr ) then Application.MessageBox( PChar( aReturn ),
    '警告', MB_OK + MB_ICONWARNING );
  FOleObj := Unassigned;
end;


procedure TTestMain.Button2Click(Sender: TObject);
var
  aReturn: String;
begin
  FOleObj := CreateOleObject( 'LibCM2CT.CM2CT' );
  aReturn := FOleObj.CM2CT_DLL1( 'EMPTY', Memo2.Lines.Text );
  if ( aReturn <> EmptyStr ) then  Application.MessageBox( PChar( aReturn ),
    '警告', MB_OK + MB_ICONWARNING );
  FOleObj := Unassigned;
end;


procedure TTestMain.Button3Click(Sender: TObject);
var
  aReturn: String;
begin
  FOleObj := CreateOleObject( 'LibCM2CT.CM2CT' );
  aReturn := FOleObj.CM2CT_DLL2( Memo3.Lines.Text );
  if ( aReturn <> EmptyStr ) then  Application.MessageBox( PChar( aReturn ),
    '警告', MB_OK + MB_ICONWARNING );
  FOleObj := Unassigned;
end;

procedure TTestMain.Button4Click(Sender: TObject);
var
  aReturn: String;
begin
  FOleObj := CreateOleObject( 'LibCM2CT.CM2CT' );
  aReturn := FOleObj.CM2CT_DLL3( Memo4.Lines.Text );
  if ( aReturn <> EmptyStr ) then  Application.MessageBox( PChar( aReturn ),
    '警告', MB_OK + MB_ICONWARNING );
  FOleObj := Unassigned;
end;

procedure TTestMain.Button5Click(Sender: TObject);
var
  aReturn: String;
begin
  FOleObj := CreateOleObject( 'LibCM2CT.CM2CT' );
  aReturn := FOleObj.CM2CT_DLL5( Memo5.Lines.Text );
  if ( aReturn <> EmptyStr ) then  Application.MessageBox( PChar( aReturn ),
    '警告', MB_OK + MB_ICONWARNING );
  FOleObj := Unassigned;
end;

procedure TTestMain.Button6Click(Sender: TObject);
var
  FPassPhrase: String;
  aKey: TKey64;
begin
  { DES 加解密用的 Key --> CS (大寫字母) }
  FPassPhrase := Chr( 67 ) + Chr( 83 );
  { 用加解密FPassPhrase值產生一組解密用的 array 結構 }
  GenerateLMDKey( aKey, SizeOf( aKey ), FPassPhrase );
  Edit2.Text := DESEncryptStringEx( Edit1.Text, aKey, True );
  Edit3.Text := DESEncryptStringEx( Edit2.Text, aKey, False );
end;

end.

