unit frmMainU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Buttons;     //Encryption_TLB

type
  TForm1 = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    Label1: TLabel;
    edtIniFileName: TEdit;
    SpeedButton1: TSpeedButton;
    memContent: TMemo;
    OpenDialog1: TOpenDialog;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    procedure SpeedButton1Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
  private
    { Private declarations }
    //function decrypt(I_Class: _Password; sI_String:String):String;
  public
    { Public declarations }

  end;

var
  Form1: TForm1;

implementation

uses DevPower_TLB;    //, DevPowerEncrypt_TLB


{$R *.dfm}

procedure TForm1.SpeedButton1Click(Sender: TObject);
var
    L_StrList : TStringList;
    ii : Integer;
    sL_Tmp,sL_Tmp2 : WideString;
    //L_Intf : _Password;
    L_Intf : _Encrypt;
    sL_EncKey : WideString;
begin
    if (OpenDialog1.Execute) then
    begin
      sL_EncKey := 'CS';
      //L_Intf := CoPassword.Create;
      L_Intf := CoEncrypt.Create;

      memContent.Lines.Clear;
      L_StrList := TStringList.Create;
      edtIniFileName.Text := OpenDialog1.FileName;
      L_StrList.LoadFromFile(OpenDialog1.FileName);
      for ii:=0 to L_StrList.Count-1 do
      begin
        sL_Tmp := L_StrList.Strings[ii];
        //if (Copy(sL_Tmp,1,2)<>'//') and (sL_Tmp<>'') then
        //begin
          //sL_Tmp := L_Intf.Decrypt(sL_Tmp,sL_EncKey);
          sL_Tmp2 := L_Intf.Decrypt(sL_Tmp);
          if (Copy(sL_Tmp,1,8)='COMPNAME') or (Copy(sL_Tmp,1,8)='COMPCODE')
          or (Copy(sL_Tmp,1,1)=';') or (sL_Tmp='')
          or (Copy(sL_Tmp,1,8)='DATAPATH') or (Copy(sL_Tmp,1,23)='[CURRENT_LANGUAGE_TYPE]')
          or (Copy(sL_Tmp,1,21)='CURRENT_LANGUAGE_TYPE') or (Copy(sL_Tmp,1,9)='[SYSINFO]')
          or (Copy(sL_Tmp,1,19)='NIGHT_RUN_EXEC_TIME') or (Copy(sL_Tmp,1,14)='CANCEL_CHANNEL')
          or (Copy(sL_Tmp,1,9)='HTTP_PORT') or (Copy(sL_Tmp,1,11)='WINZIP_PATH') then
            memContent.Lines.Add(sL_Tmp)    //¤£¥[±K
          else
            memContent.Lines.Add(sL_Tmp2)   //¥[±K
        //end;

      end;
      L_StrList.Free;
      L_Intf := nil;
    end;
end;

procedure TForm1.BitBtn1Click(Sender: TObject);
begin
    Application.Terminate;
end;
{
function TForm1.decrypt(I_Class: _Password; sI_String: String): String;
var
    sL_EncKey : WideString;
begin
    sL_EncKey := 'CS';
    result := I_Class.Decrypt('abc',sL_EncKey);
//    result := I_Class.Decrypt(sI_String, sL_EncKey);
end;
}
procedure TForm1.BitBtn2Click(Sender: TObject);
var
    ii : Integer;
    sL_Tmp,sL_Tmp2 : WideString;
    L_StrList : TStringList;
    L_Intf : _Encrypt;
    sL_EncKey : WideString;
begin
      L_StrList := TStringList.Create;
      sL_EncKey := 'CS';
      //L_Intf := CoPassword.Create;
      L_Intf := CoEncrypt.Create;

      for ii:=0 to memContent.Lines.Count-1 do
      begin
        sL_Tmp := memContent.Lines.Strings[ii];
//        if (sL_Tmp='') then continue;

          if (Copy(sL_Tmp,1,8)='COMPNAME') or (Copy(sL_Tmp,1,8)='COMPCODE')
              or (Copy(sL_Tmp,1,1)=';') or (sL_Tmp='')
              or (Copy(sL_Tmp,1,8)='DATAPATH') or (Copy(sL_Tmp,1,23)='[CURRENT_LANGUAGE_TYPE]')
              or (Copy(sL_Tmp,1,21)='CURRENT_LANGUAGE_TYPE') or (Copy(sL_Tmp,1,9)='[SYSINFO]')
              or (Copy(sL_Tmp,1,19)='NIGHT_RUN_EXEC_TIME') or (Copy(sL_Tmp,1,14)='CANCEL_CHANNEL')
              or (Copy(sL_Tmp,1,9)='HTTP_PORT') or (Copy(sL_Tmp,1,11)='WINZIP_PATH') then
              L_StrList.Add(sL_Tmp)
          else
          begin
              sL_Tmp2 := L_Intf.Encrypt(sL_Tmp);
              L_StrList.Add(sL_Tmp2);
          end;



      end;
      L_StrList.SaveToFile(edtIniFileName.Text);
      L_StrList.Free;
      MessageDlg('OK!', mtInformation, [mbOK], 0);
end;

end.
