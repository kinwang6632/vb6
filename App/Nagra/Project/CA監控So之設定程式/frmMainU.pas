unit frmMainU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, inifiles, ExtCtrls;

type
  TfrmMain = class(TForm)
    Panel1: TPanel;
    BitBtn1: TBitBtn;
    CheckBox1: TCheckBox;
    CheckBox2: TCheckBox;
    CheckBox3: TCheckBox;
    CheckBox4: TCheckBox;
    CheckBox5: TCheckBox;
    CheckBox6: TCheckBox;
    CheckBox7: TCheckBox;
    CheckBox8: TCheckBox;
    CheckBox9: TCheckBox;
    CheckBox10: TCheckBox;
    CheckBox11: TCheckBox;
    CheckBox12: TCheckBox;
    CheckBox13: TCheckBox;
    CheckBox14: TCheckBox;
    Button1: TButton;
    StaticText1: TStaticText;
    edtIniFileName: TEdit;
    SpeedButton1: TSpeedButton;
    OpenDialog1: TOpenDialog;
    CheckBox15: TCheckBox;
    StaticText2: TStaticText;
    cobCallbackDataAlias: TComboBox;
    StaticText3: TStaticText;
    cobIccDataAlias: TComboBox;
    procedure Button1Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
  private
    { Private declarations }
    function getALias(nI_Tag:Integer):String;
    procedure writeToIni;
    procedure eraseIniSectionContent;
    procedure ReadIni;
    procedure disableAllCheckBox;
    procedure enableCheckBoxStatus(nI_CompCode:Integer);    
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

{$R *.dfm}

procedure TfrmMain.Button1Click(Sender: TObject);
begin
    Application.Terminate;
end;

procedure TfrmMain.ReadIni;
var
    sL_CallbackDataAlias,sL_IccDataAlias,  sL_IniFileName : String;
    L_IniFile : TIniFile;
    nL_DbCount,ii, nL_CompCode, nL_Key : Integer;
begin
    sL_IniFileName := edtIniFileName.Text;

    if FileExists(sL_IniFileName) then
    begin
      L_IniFile := TIniFile.Create(sL_IniFileName);


      nL_DbCount := L_IniFile.ReadInteger('DBINFO','DB_COUNT',0);
      for ii:=0 to nL_DbCount-1 do
      begin
        nL_Key := ii + 1;
        nL_CompCode := L_IniFile.ReadInteger('DBINFO','COMPCODE_' + IntToStr(nL_Key) ,0);
        enableCheckBoxStatus(nL_CompCode);

      end;
      sL_CallbackDataAlias := L_IniFile.ReadString('CALLBACKSVR', 'CALLBACK_DATA_ALIAS', '');
      if UpperCase(sL_CallbackDataAlias)='CATVN' then
        cobCallbackDataAlias.ItemIndex := 0
      else
        cobCallbackDataAlias.ItemIndex := 1;

      sL_IccDataAlias := L_IniFile.ReadString('CALLBACKSVR', 'COM_ALIAS', '');
      if UpperCase(sL_IccDataAlias)='CATVN' then
        cobIccDataAlias.ItemIndex := 0
      else
        cobIccDataAlias.ItemIndex := 1;




//      L_IniFile.DeleteKey('DBINFO', 'USERID_3');
//    L_IniFile.EraseSection('DBINFO');

      L_IniFile.Free;
    end;
end;

procedure TfrmMain.SpeedButton1Click(Sender: TObject);
begin
    if OpenDialog1.Execute then
    begin
      edtIniFileName.Text := OpenDialog1.FileName;
      disableAllCheckBox;
      ReadIni;
    end;
end;

procedure TfrmMain.FormShow(Sender: TObject);
begin
    CheckBox1.Tag := 9 ;//��ܶ����s�����q�O�O 9
    CheckBox2.Tag := 10 ;//��ܷs�x�_�����q�O�O 10
    CheckBox3.Tag := 12 ;//��ܤj�w��s�����q�O�O 12
    CheckBox4.Tag := 11 ;//��ܪ��W�D�����q�O�O 11
    CheckBox5.Tag := 8 ;//��ܥ��p�����q�O�O 8
    CheckBox6.Tag := 14 ;//��ܤj�s�������q�O�O 14
    CheckBox7.Tag := 13 ;//��ܷs�𫰪����q�O�O 13
    CheckBox8.Tag := 7 ;//��ܮ��D�����q�O�O 7
    CheckBox9.Tag := 6 ;//����׷������q�O�O 6
    CheckBox10.Tag := 5 ;//��ܷs�W�D�����q�O�O 5
    CheckBox11.Tag := 3 ;//��ܫn�Ѫ����q�O�O 3
    CheckBox12.Tag := 1 ;//����[�@�����q�O�O 1
    CheckBox13.Tag := 2 ;//��ܫ̫n�����q�O�O 2
    CheckBox14.Tag := 20 ;//��ܭܺު����q�O�O 20
    CheckBox15.Tag := 16 ;//��ܥ_��骺���q�O�O 16

    disableAllCheckBox;
end;

procedure TfrmMain.enableCheckBoxStatus(nI_CompCode: Integer);
var
    ii : Integer;
begin
    for ii:=0 to Panel1.ControlCount -1 do
    begin
      if Panel1.Controls[ii] IS TCheckBox then
      begin
        if (Panel1.Controls[ii] AS TCheckBox).Tag = nI_CompCode then
          (Panel1.Controls[ii] AS TCheckBox).Checked := true;
      end;
    end;
end;

procedure TfrmMain.disableAllCheckBox;
var
    ii : Integer;
begin
    for ii:=0 to Panel1.ControlCount -1 do
    begin
      if Panel1.Controls[ii] IS TCheckBox then
        (Panel1.Controls[ii] AS TCheckBox).Checked := false;

    end;
end;

procedure TfrmMain.BitBtn1Click(Sender: TObject);
begin
    eraseIniSectionContent;
    writeToIni;
end;

procedure TfrmMain.eraseIniSectionContent;
var
    sL_IniFileName : String;
    L_IniFile : TIniFile;
    nL_DbCount,ii, nL_CompCode, nL_Key : Integer;
begin
    sL_IniFileName := edtIniFileName.Text;

    if FileExists(sL_IniFileName) then
    begin
      L_IniFile := TIniFile.Create(sL_IniFileName);
      L_IniFile.EraseSection('DBINFO');
      L_IniFile.EraseSection('CALLBACKSVR');      
      {
      nL_DbCount := L_IniFile.ReadInteger('DBINFO','DB_COUNT',0);
      for ii:=0 to nL_DbCount-1 do
      begin
        nL_Key := ii + 1;
        L_IniFile.DeleteKey('DBINFO', 'ALIAS_' + IntToStr(nL_Key));
        L_IniFile.DeleteKey('DBINFO', 'PASSWORD_' + IntToStr(nL_Key));
        L_IniFile.DeleteKey('DBINFO', 'USERID_' + IntToStr(nL_Key));
        L_IniFile.DeleteKey('DBINFO', 'COMPCODE_' + IntToStr(nL_Key));
      end;
      }
      L_IniFile.Free;
    end;
end;

procedure TfrmMain.writeToIni;
var
    sL_ActivePlatformAlias : String;
    sL_IniFileName, sL_UserID, sL_Passwd, sL_Alias : String;
    ii, nL_Tag, nL_CompCode, nL_Count : Integer;
    L_IniFile : TIniFile;
begin
    sL_IniFileName := edtIniFileName.Text;

    if FileExists(sL_IniFileName) then
    begin
      L_IniFile := TIniFile.Create(sL_IniFileName);

      case cobCallbackDataAlias.ItemIndex of
        0: //�N callback data �g�J�_���x
          sL_ActivePlatformAlias := 'catvn';
        1: //�N callback data �g�J�����x
           sL_ActivePlatformAlias := 'catvc';
      end;
      L_IniFile.WriteString('CALLBACKSVR', 'CALLBACK_DATA_ALIAS', sL_ActivePlatformAlias);
      L_IniFile.WriteString('CALLBACKSVR', 'CALLBACK_DATA_USERID', 'com');
      L_IniFile.WriteString('CALLBACKSVR', 'CALLBACK_DATA_PASSWORD', 'com');


      case cobIccDataAlias.ItemIndex of
        0: //�Ҧ��� ICC ��Ƨ��s��b�_���x
          sL_ActivePlatformAlias := 'catvn';
        1: //�Ҧ��� ICC ��Ƨ��s��b�����x
           sL_ActivePlatformAlias := 'catvc';
      end;
      L_IniFile.WriteString('CALLBACKSVR', 'COM_ALIAS', sL_ActivePlatformAlias);
      L_IniFile.WriteString('CALLBACKSVR', 'COM_USERID', 'com');
      L_IniFile.WriteString('CALLBACKSVR', 'COM_PASSWORD', 'com');


      nL_Count := 0;

      for ii:=0 to Panel1.ControlCount -1 do
      begin
        if Panel1.Controls[ii] IS TCheckBox then
        begin
          if (Panel1.Controls[ii] AS TCheckBox).Checked then
          begin
            nL_Tag := (Panel1.Controls[ii] AS TCheckBox).Tag;
            sL_Alias := getALias(nL_Tag);
            if sL_Alias<>'' then
            begin
              Inc(nL_Count);
              nL_CompCode := nL_Tag;
              sL_UserID := 'com';
              sL_Passwd := 'com';

              L_IniFile.WriteString('DBINFO', 'ALIAS_' + IntToStr(nL_Count),sL_Alias);
              L_IniFile.WriteString('DBINFO', 'USERID_' + IntToStr(nL_Count),sL_UserID);
              L_IniFile.WriteString('DBINFO', 'PASSWORD_' + IntToStr(nL_Count),sL_Passwd);
              L_IniFile.WriteString('DBINFO', 'COMPCODE_' + IntToStr(nL_Count),IntToStr(nL_CompCode));
            end;

          end;
        end;

      end;
      L_IniFile.WriteString('DBINFO', 'DB_COUNT',IntToStr(nL_Count));
      L_IniFile.Free;
      MessageDlg('�]�w����!', mtInformation,[mbOK],0);
    end;
end;

function TfrmMain.getALias(nI_Tag: Integer):String;
var
    sL_Alias : String;
begin
    case nI_Tag of
      1:
        sL_Alias := 'emcks';
      2:
        sL_Alias := 'emcpn';
      3:
        sL_Alias := 'emcnt';
      5:
        sL_Alias := 'emcncc';
      6:
        sL_Alias := 'emcfm';
      7:
        sL_Alias := 'emcct';
      8:
        sL_Alias := 'emcuc';
      9:
        sL_Alias := 'emcyms';
      10:
        sL_Alias := 'emcntp';
      11:
        sL_Alias := 'emckp';
      12:
        sL_Alias := 'emcdw';
      13:
        sL_Alias := 'emctc';
      14:
        sL_Alias := 'emcst';
      20:
        sL_Alias := 'emcyms';
      16:
        sL_Alias := 'emcnty';
      else
        sL_Alias := '';
    end;

    result := sL_Alias;
end;

end.
