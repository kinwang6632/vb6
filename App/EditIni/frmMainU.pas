unit frmMainU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Buttons, Encryption_TLB, DevPower_TLB,
  Crypt32DLL_TLB, XPMan;

type
  TForm1 = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    Label1: TLabel;
    edtIniFileName: TEdit;
    SpeedButton1: TSpeedButton;
    memContent: TMemo;
    OpenDialog1: TOpenDialog;
    Panel3: TPanel;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    BitBtn3: TBitBtn;
    XPManifest1: TXPManifest;
    RadioButton1: TRadioButton;
    RadioButton2: TRadioButton;
    RadioButton3: TRadioButton;
    RadioButton4: TRadioButton;
    procedure SpeedButton1Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
  private
    { Private declarations }
    FIntf: OleVariant;
    procedure DoNormal;
    procedure DoSpecial;
    procedure DoDecrypt4;
    procedure SaveNormal;
    procedure SaveSpecial;
    procedure SaveEncrypt4;
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation


{$R *.dfm}


procedure TForm1.SpeedButton1Click(Sender: TObject);
begin
  if ( OpenDialog1.Execute ) then
  begin
    edtIniFileName.Text := OpenDialog1.FileName;
    memContent.Clear;
    Application.ProcessMessages;
    if ( RadioButton3.Checked ) then
      DoSpecial
    else if ( RadioButton4.Checked ) then
      DoDecrypt4    
    else
      DoNormal;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.BitBtn1Click(Sender: TObject);
begin
  Self.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.BitBtn2Click(Sender: TObject);
begin
  if ( RadioButton3.Checked ) then
    SaveSpecial
  else if ( RadioButton4.Checked ) then
    SaveEncrypt4
  else
    SaveNormal;  
  Application.MessageBox( PChar( '檔案已儲存。'), '訊息', MB_OK + MB_ICONINFORMATION );
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.BitBtn3Click(Sender: TObject);
begin
  memContent.Clear;
  Application.ProcessMessages;
  if ( RadioButton3.Checked ) then
    DoSpecial
  else if ( RadioButton4.Checked ) then
    DoDecrypt4
  else
    DoNormal;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.DoNormal;
var
  aKey, aTmp: String;
  aIndex: Integer;
  aList: TStringList;
begin
  aKey := 'CS';
  if RadioButton1.Checked then
    FIntf := CoPassword.Create
  else
    FIntf := CoEncrypt.Create;
  try  
    memContent.Lines.Clear;
    aList := TStringList.Create;
    try
      edtIniFileName.Text := OpenDialog1.FileName;
      aList.LoadFromFile( edtIniFileName.Text );
      for aIndex:=0 to aList.Count-1 do
      begin
        aTmp := aList.Strings[aIndex];
        if (Copy(aTmp,1,2)<>'//') and (aTmp<>'') and (Copy(aTmp,1,1)<>';') and
           (Copy(aTmp,1,8) <> 'COMPCODE') then
        begin
          if RadioButton1.Checked then
            aTmp := FIntf.Decrypt(aTmp,aKey)
          else
            aTmp := FIntf.Decrypt(aTmp );
        end;
        memContent.Lines.Add( aTmp )
      end;
    finally
      aList.Free;
    end;
  finally
    FIntf := Null;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.DoSpecial;
var
  aSections, aValues: TStringList;
  aIndex, aPos: Integer;
  aText, aTmp, aDecrypt: String;
begin
  aSections := TStringList.Create;
  aValues := TStringList.Create;
  try
    aSections.LoadFromFile( edtIniFileName.Text );
    FIntf := CoEncrypt.Create;
    try
      for aIndex := 0 to aSections.Count - 1 do
      begin
        aTmp := aSections[aIndex];
        aPos:= 
          ( AnsiPos( '[', aTmp ) + AnsiPos( ']', aTmp ) + 
            AnsiPos( ';', aTmp ) + AnsiPos( '//', aTmp )+ AnsiPos( 'DefaultDB', aTmp ) );
        if ( aPos = 0 ) and ( aTmp <> EmptyStr ) then
        begin
          aPos := AnsiPos( '=', aTmp );
          aText := Copy( aTmp, aPos + 1, Length( aTmp ) - aPos + 1 );
          aDecrypt := FIntf.Decrypt( aText );
          aTmp := Copy( aTmp, 1, AnsiPos( '=', aTmp ) - 1 ) + '=' + aDecrypt;
        end;        
        memContent.Lines.Append( aTmp );
      end;
    finally
      FIntf := Null;
    end;
  finally
    aValues.Free;
    aSections.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.DoDecrypt4;
var
  aSections, aValues: TStringList;
  aIndex, aPos: Integer;
  aText, aTmp, aDecrypt: String;
begin
  aSections := TStringList.Create;
  aValues := TStringList.Create;
  try
    aSections.LoadFromFile( edtIniFileName.Text );
    FIntf := CoCrypt32.Create;
    try
      for aIndex := 0 to aSections.Count - 1 do
      begin
        aTmp := aSections[aIndex];
        aPos:=
          ( AnsiPos( '[', aTmp ) + AnsiPos( ']', aTmp ) + 
            AnsiPos( ';', aTmp ) + AnsiPos( '//', aTmp )+ AnsiPos( 'DefaultDB', aTmp ) );
        if ( aPos = 0 ) and ( aTmp <> EmptyStr ) then
        begin
          aPos := AnsiPos( '=', aTmp );
          aText := Copy( aTmp, aPos + 1, Length( aTmp ) - aPos + 1 );
          aDecrypt := FIntf.DValue( aText, EmptyParam, 'CsMisk' );
          if ( aPos > 0 ) then
            aTmp := Copy( aTmp, 1, AnsiPos( '=', aTmp ) - 1 ) + '=' + aDecrypt
          else
            aTmp := aDecrypt;  
        end;
        memContent.Lines.Append( aTmp );
      end;
    finally
      FIntf := Null;
    end;
  finally
    aValues.Free;
    aSections.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.SaveNormal;
var
  aList: TStringList;
  aTmp, aKey: String;
  aIndex: Integer;
begin
  aList := TStringList.Create;
  try
    aKey := 'CS';
    if RadioButton1.Checked then
      FIntf := CoPassword.Create
    else
      FIntf := CoEncrypt.Create;
    try
      for aIndex := 0 to memContent.Lines.Count-1 do
      begin
        aTmp := memContent.Lines.Strings[aIndex];
        if ( Copy( aTmp, 1, 2 ) <> '//' ) and ( aTmp <> '' ) and 
           ( Copy( aTmp, 1, 1 ) <> ';' ) and ( Copy( aTmp, 1, 8 ) <> 'COMPCODE' ) then
        begin
          if RadioButton1.Checked  then
            aTmp := FIntf.Encrypt( aTmp,aKey )
          else
            aTmp := FIntf.Encrypt( aTmp );
        end;
        aList.Add( aTmp )
      end;
      aList.SaveToFile( edtIniFileName.Text );
    finally
      FIntf := Null;
    end;
  finally
    aList.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.SaveSpecial;
var
  aList: TStringList;
  aIndex, aPos: Integer;
  aTmp, aText, aEncrypt: String;
begin
  aList := TStringList.Create;
  try
    FIntf := CoEncrypt.Create;
    try
      for aIndex := 0 to memContent.Lines.Count - 1 do
      begin
        aTmp := memContent.Lines.Strings[aIndex];
        aPos:=
          ( AnsiPos( '[', aTmp ) + AnsiPos( ']', aTmp ) +
          AnsiPos( ';', aTmp ) + AnsiPos( '//', aTmp )+ AnsiPos( 'DefaultDB', aTmp ) );
        if ( aPos = 0 ) and ( aTmp <> EmptyStr ) then
        begin
          aPos := AnsiPos( '=', aTmp );
          aText := Copy( aTmp, aPos + 1, Length( aTmp ) - aPos + 1 );
          aEncrypt := FIntf.Encrypt( aText );
          aTmp := Copy( aTmp, 1, AnsiPos( '=', aTmp ) - 1 ) + '=' + aEncrypt;
        end;
        aList.Add( aTmp );
      end;
      aList.SaveToFile( edtIniFileName.Text );
    finally
      FIntf := Null;
    end;
  finally
    aList.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.SaveEncrypt4;
var
  aList: TStringList;
  aIndex: Integer;
  aText, aEncrypt: String;
begin
  aList := TStringList.Create;
  try
    FIntf := CoCrypt32.Create;
    try
      for aIndex := 0 to memContent.Lines.Count - 1 do
      begin
        aText := memContent.Lines.Strings[aIndex];
        aEncrypt := FIntf.BValue( aText, EmptyParam, 'CsMisk' );
        aList.Add( aEncrypt );
      end;
      aList.SaveToFile( edtIniFileName.Text );
    finally
      FIntf := Null;
    end;
  finally
    aList.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
