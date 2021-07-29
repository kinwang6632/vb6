unit cbNgwProfileTool;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, StdCtrls, ExtCtrls, ComCtrls, ImgList, DB, ADODB,
  LbCipher, LbClass,
  cxLookAndFeels, cxLookAndFeelPainters,
  dxSkinsCore, dxSkinsDefaultPainters, dxSkinsdxStatusBarPainter, dxSkinsForm,
  cxGraphics, cxControls, cxContainer,
  cxEdit, cxTextEdit, cxMaskEdit, cxButtonEdit, cxButtons, cxRadioGroup, cxGroupBox,
  cxListView, dxStatusBar;

type
  TfmNgwProfileTool = class(TForm)
    btnExecute: TcxButton;
    OpenDialog1: TOpenDialog;
    Profileonnection: TADOConnection;
    tbProfile: TADOTable;
    Label1: TLabel;
    dxSkinController1: TdxSkinController;
    cxGroupBox1: TcxGroupBox;
    txtFile: TcxButtonEdit;
    rbEncrypt: TcxRadioButton;
    rbDecrypt: TcxRadioButton;
    cxGroupBox2: TcxGroupBox;
    ListViewTable: TcxListView;
    dxStatusBar: TdxStatusBar;
    ImgSmall: TcxImageList;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure txtFilePropertiesButtonClick(Sender: TObject;
      AButtonIndex: Integer);
    procedure btnExecuteClick(Sender: TObject);
  private
    { Private declarations }
    FLbDES: TLbDES;
    function DetermineIsEncryption: Boolean;
    function EncryptOrDecryptData(const AMode: Integer; const ATableName: String): Boolean;
    function OpenProfile: Boolean;
    procedure CloseProfile;
    procedure LoadProfileTables;
    procedure Encrypt;
    procedure Decrypt;
  public
    { Public declarations }
  end;

var
  fmNgwProfileTool: TfmNgwProfileTool;

implementation

uses cbUtilis;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TfmNgwProfileTool.FormCreate(Sender: TObject);
begin
  FLbDES := TLbDES.Create( nil );
  FLbDES.GenerateKey( 'CS' );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmNgwProfileTool.FormDestroy(Sender: TObject);
begin
  FLbDES.Free;
end;

{ ---------------------------------------------------------------------------- }

function TfmNgwProfileTool.OpenProfile: Boolean;
const
  aDbPassword = 'cyc84177282';
var
  aFileName: String;
begin
  aFileName := txtFile.Text;
  Profileonnection.Connected := False;
  Profileonnection.ConnectionString := Format(
    'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=%s;Jet OLEDB:Database Password=%s;'+
    'Mode=Share Deny Read|Share Deny Write;', [aFileName, aDbPassword] );
  try
    Profileonnection.Connected := True;
  except
    on E: Exception do dxStatusBar.Panels[0].Text := E.Message;
  end;
  Result := Profileonnection.Connected;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmNgwProfileTool.CloseProfile;
begin
  if ( Profileonnection.Connected ) then
    Profileonnection.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmNgwProfileTool.LoadProfileTables;
var
  aList: TStringList;
  aIndex: Integer;
  aItem: TListItem;
begin
  aList := TStringList.Create;
  try
    Profileonnection.GetTableNames( aList, False );
    for aIndex := 0 to aList.Count - 1 do
    begin
      aItem := ListViewTable.Items.Add;
      aItem.Caption := aList[aIndex];
      aItem.ImageIndex := 0;
      aItem.Checked := True;
      Delay( 100 );
    end;
  finally
    aList.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmNgwProfileTool.Encrypt;
var
  aTableName: String;
  aIndex: Integer;
begin
  Screen.Cursor := crSQLWait;
  try
    for aIndex := 0 to ListViewTable.Items.Count - 1 do
    begin
      aTableName := ListViewTable.Items[aIndex].Caption;
      dxStatusBar.Panels[0].Text := Format( '資料表: %s, %s....', [aTableName, '加密中'] );
      Application.ProcessMessages;
      if not EncryptOrDecryptData( 0, aTableName ) then Exit;
    end;
    dxStatusBar.Panels[0].Text := EmptyStr;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmNgwProfileTool.Decrypt;
var
  aTableName: String;
  aIndex: Integer;
begin
  Screen.Cursor := crSQLWait;
  try
    for aIndex := 0 to ListViewTable.Items.Count - 1 do
    begin
      aTableName := ListViewTable.Items[aIndex].Caption;
      dxStatusBar.Panels[0].Text := Format( '資料表: %s, %s....', [aTableName, '解密中'] );
      Application.ProcessMessages;
      if not EncryptOrDecryptData( 1, aTableName ) then Exit;
    end;
    dxStatusBar.Panels[0].Text := EmptyStr;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfmNgwProfileTool.DetermineIsEncryption: Boolean;
begin
  tbProfile.Close;
  tbProfile.TableName := 'ENCRYPTIONTEST';
  tbProfile.Open;
  Result := ( tbProfile.FieldByName( 'COL1' ).AsString = 'QN+6JrN/3Z0=' );
  tbProfile.Close;
end;

{ ---------------------------------------------------------------------------- }

function TfmNgwProfileTool.EncryptOrDecryptData(const AMode: Integer; const ATableName: String): Boolean;
var
  aIndex: Integer;
begin
  tbProfile.Close;
  tbProfile.TableName := ATableName;
  tbProfile.Open;
  tbProfile.First;
  try
    while not tbProfile.Eof do
    begin
      tbProfile.Edit;
      for aIndex := 0 to tbProfile.FieldCount - 1 do
      begin
        if ( AMode = 0 ) then
          tbProfile.Fields[aIndex].AsString := FLbDES.EncryptString(
            tbProfile.Fields[aIndex].AsString )
        else
          tbProfile.Fields[aIndex].AsString := FLbDES.DecryptString(
            tbProfile.Fields[aIndex].AsString );
      end;
      tbProfile.Post;
      tbProfile.Next;
      Application.ProcessMessages;
    end;
    Result := True;
  except
    on E: Exception do
    begin
      dxStatusBar.Panels[0].Text := Format( 'Encrypt Error: %s --> %s',
        [ATableName, E.Message] );
      Result := False;
    end;
  end;
  tbProfile.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmNgwProfileTool.txtFilePropertiesButtonClick(Sender: TObject;
  AButtonIndex: Integer);
begin
  if OpenDialog1.Execute then
  begin
    txtFile.Text := OpenDialog1.FileName;
    if not OpenProfile then Exit;
    rbDecrypt.Checked := DetermineIsEncryption;
    LoadProfileTables;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmNgwProfileTool.btnExecuteClick(Sender: TObject);
begin
  if rbEncrypt.Checked then
    Encrypt
  else
    Decrypt;
  rbEncrypt.Checked := True;
  if ( DetermineIsEncryption ) then rbDecrypt.Checked := True;
  CloseProfile;
end;

{ ---------------------------------------------------------------------------- }

end.
