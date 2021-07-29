unit frmInvA09U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Mask, ExtCtrls, Buttons, fraYMDU, fraYMU, Grids,
  DBGrids, DB;

type
  TfrmInvA09 = class(TForm)
    Panel1: TPanel;
    Label6: TLabel;
    Panel2: TPanel;
    btnExit: TBitBtn;
    btnQuery: TBitBtn;
    Panel3: TPanel;
    fraInvYM: TfraYM;
    Label2: TLabel;
    Label10: TLabel;
    medPrefix: TMaskEdit;
    chkBusinessID: TCheckBox;
    chkJump: TCheckBox;
    chkNumber: TCheckBox;
    Panel4: TPanel;
    chkAmount: TCheckBox;
    Label1: TLabel;
    fraEInvDate1: TfraYMD;
    fraSInvDate1: TfraYMD;
    chkCoverObsolete: TCheckBox;
    Panel5: TPanel;
    chkMisData: TCheckBox;
    chkMisData1: TCheckBox;
    chkMisData2: TCheckBox;
    chkMisData4: TCheckBox;
    chkMisData5: TCheckBox;
    Label3: TLabel;
    Label4: TLabel;
    fraSInvDate2: TfraYMD;
    fraEInvDate2: TfraYMD;
    chkMisData3: TCheckBox;
    procedure FormShow(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnQueryClick(Sender: TObject);
    procedure medPrefixExit(Sender: TObject);
    procedure chkJumpClick(Sender: TObject);
    procedure fraSInvDate1Exit(Sender: TObject);
    procedure fraSInvDate2Exit(Sender: TObject);
  private
    { Private declarations }
    function isDataOK : Boolean;
  public
    { Public declarations }
  end;

var
  frmInvA09: TfrmInvA09;

implementation

uses frmMainU, dtmMainU, UdateTimeu, Ustru, dtmMainJU, frmInvA09_1U;

{$R *.dfm}

procedure TfrmInvA09.FormShow(Sender: TObject);
begin
    self.Caption := frmMain.GetFormTitleString('A09','�o������ˮ�');
    medPrefix.Enabled := false;
end;

procedure TfrmInvA09.btnExitClick(Sender: TObject);
begin
    Close;
end;

function TfrmInvA09.isDataOK: Boolean;
var
    sL_InvYM,sL_Prefix,sL_SInvDate1,sL_EInvDate1,sL_SInvDate2,sL_EInvDate2 : String;
begin
    Result := true;
    if chkNumber.Checked or chkAmount.Checked or chkMisData.Checked then
    begin
      //�νs�ˮ�
      if chkNumber.Checked then
      begin
        if chkBusinessID.Checked or chkJump.Checked then
        begin
          sL_InvYM := Trim(fraInvYM.getYM);
          sL_Prefix := Trim(medPrefix.Text);

          if sL_InvYM = EMPTY_YM_STR then
          begin
            MessageDlg('�п�J�o���~��!',mtError,[mbOk],0);
            fraInvYM.mseYM.SetFocus;
            Result := false;
            Exit;
          end;

          if chkJump.Checked then
          begin
            if sL_Prefix = '' then
            begin
              MessageDlg('�п�J�r�y!',mtError,[mbOk],0);
              medPrefix.SetFocus;
              Result := false;
              Exit;
            end;

            if Length(sL_Prefix) <> 2 then
            begin
              MessageDlg('�п�J���T���r�y�榡!',mtError,[mbOk],0);
              medPrefix.SetFocus;
              Result := false;
              Exit;
            end;
          end;
        end
        else
        begin
          MessageDlg('���I��n�ˮ֪��νs�ˮֶ���!',mtError,[mbOk],0);
          chkBusinessID.SetFocus;
          Result := false;
          Exit;
        end;
      end;



      //�D���ɪ��B�ˬd
      if chkAmount.Checked then
      begin
        sL_SInvDate1 := Trim(fraSInvDate1.getYMD);
        sL_EInvDate1 := Trim(fraEInvDate1.getYMD);

        if sL_SInvDate1=EMPTY_DATE_STR then
        begin
          MessageDlg('�п�J�o���}�l���',mtError,[mbOk],0);
          fraSInvDate1.mseYMD.SetFocus;
          Result := false;
          Exit;
        end;

        if sL_EInvDate1=EMPTY_DATE_STR then
        begin
          MessageDlg('�п�J�o���������',mtError,[mbOk],0);
          fraEInvDate1.mseYMD.SetFocus;
          Result := false;
          Exit;
        end;

        if sL_SInvDate1 > sL_EInvDate1 then
        begin
          MessageDlg('�o���}�l��������p�󵲧����',mtError,[mbOk],0);
          fraEInvDate1.mseYMD.SetFocus;
          Result := false;
          Exit;
        end;
      end;



      //�o���}�߻P�ȪA��Ƥ��
      if chkMisData.Checked then
      begin
        if chkMisData1.Checked or chkMisData2.Checked or chkMisData3.Checked or
           chkMisData4.Checked or chkMisData5.Checked then
        begin
          sL_SInvDate2 := Trim(fraSInvDate2.getYMD);
          sL_EInvDate2 := Trim(fraEInvDate2.getYMD);

          if sL_SInvDate2=EMPTY_DATE_STR then
          begin
            MessageDlg('�п�J�o���}�l���',mtError,[mbOk],0);
            fraSInvDate2.mseYMD.SetFocus;
            Result := false;
            Exit;
          end;

          if sL_EInvDate2=EMPTY_DATE_STR then
          begin
            MessageDlg('�п�J�o���������',mtError,[mbOk],0);
            fraEInvDate2.mseYMD.SetFocus;
            Result := false;
            Exit;
          end;

          if sL_SInvDate2 > sL_EInvDate2 then
          begin
            MessageDlg('�o���}�l��������p�󵲧����',mtError,[mbOk],0);
            fraEInvDate2.mseYMD.SetFocus;
            Result := false;
            Exit;
          end;
        end
        else
        begin
          MessageDlg('���I����@�o���}�߻P�ȪA��Ƥ�ﶵ��!',mtError,[mbOk],0);
          chkBusinessID.SetFocus;
          Result := false;
          Exit;
        end;
      end;
    end
    else
    begin
      MessageDlg('�Цܤֿ�ܤ@���ˮֶ���!',mtError,[mbOk],0);
      chkNumber.SetFocus;
      Result := false;
      Exit;
    end;
end;

procedure TfrmInvA09.btnQueryClick(Sender: TObject);
var
    sL_CompID,sL_InvYM,sL_Prefix,sL_Year,sL_Month,sL_DaysOfMonth : String;
    sL_CheckNumStrDate,sL_CheckNumEndDate,sL_SInvDate1,sL_EInvDate1 : String;
    sL_SInvDate2,sL_EInvDate2,sL_SelectItem : String;
    nL_ErrorCounts : Integer;
begin
    if isDataOK then
    begin
      //���浥�ݪ��A
      dtmMain.setWaitingCursor;

      nL_ErrorCounts := 0;
      sL_CompID := dtmMain.getCompID;

      if not dtmMainJ.cdsCheckBusinessID.Active then
        dtmMainJ.cdsCheckBusinessID.CreateDataSet;
      dtmMainJ.cdsCheckBusinessID.EmptyDataSet;

      if not dtmMainJ.cdsCheckJumpInvID.Active then
        dtmMainJ.cdsCheckJumpInvID.CreateDataSet;
      dtmMainJ.cdsCheckJumpInvID.EmptyDataSet;

      if not dtmMainJ.cdsCheckAmount.Active then
        dtmMainJ.cdsCheckAmount.CreateDataSet;
      dtmMainJ.cdsCheckAmount.EmptyDataSet;

      if not dtmMainJ.cdsCheckMisData.Active then
        dtmMainJ.cdsCheckMisData.CreateDataSet;
      dtmMainJ.cdsCheckMisData.EmptyDataSet;


      //************************************************************************
      //�νs�ˮ�
      if chkNumber.Checked then
      begin
        sL_InvYM := TUstr.replaceStr(Trim(fraInvYM.getYM),'/','');
        sL_Prefix := Trim(medPrefix.Text);
        //�Τ@�s���ˮ�
        if chkBusinessID.Checked then
        begin
          sL_Year := copy(sL_InvYM,1,4);
          sL_Month := copy(sL_InvYM,5,2);
          sL_DaysOfMonth := IntToStr(TUdateTime.DaysOfMonth(IntToStr(StrToInt(sL_Year)-1911)+sL_Month+'01'));

          sL_CheckNumStrDate := sL_Year+'/'+sL_Month+'/01';
          sL_CheckNumEndDate := sL_Year+'/'+sL_Month+'/'+sL_DaysOfMonth;

          nL_ErrorCounts := nL_ErrorCounts + dtmMainJ.handleCheckBusinessID(sL_CompID,sL_CheckNumStrDate,sL_CheckNumEndDate);
        end;

        //�����ˮ�
        if chkJump.Checked then
        begin
          sL_InvYM := dtmMainJ.changePrefixYM(sL_InvYM);
          nL_ErrorCounts := nL_ErrorCounts + dtmMainJ.handleCheckJumpInvID(sL_CompID,sL_InvYM,sL_Prefix);
        end;
      end;



      //************************************************************************
      //�D���ɪ��B�ˬd
      if chkAmount.Checked then
      begin
        sL_SInvDate1 := Trim(fraSInvDate1.getYMD);
        sL_EInvDate1 := Trim(fraEInvDate1.getYMD);

        nL_ErrorCounts := nL_ErrorCounts + dtmMainJ.handleChkAmountData(StrToDateTime(sL_SInvDate1), StrToDateTime(sL_EInvDate1), chkCoverObsolete.checked,sL_CompID);
      end;



      //************************************************************************
      //�o���}�߻P�ȪA��Ƥ��
      if chkMisData.Checked then
      begin
        sL_SInvDate2 := Trim(fraSInvDate2.getYMD);
        sL_EInvDate2 := Trim(fraEInvDate2.getYMD);

        if chkMisData1.Checked then
        begin
          //�o������H�P�ȪA�D�ɩm�W���P�λP�b������H���P
          sL_SelectItem := '1';
          nL_ErrorCounts := nL_ErrorCounts + dtmMainJ.handleChkMisData(sL_SelectItem,sL_CompID,sL_SInvDate2,sL_EInvDate2);
        end;


        if chkMisData2.Checked then
        begin
          //�����T�p�}���G�p(���b��)�������T�p�}���G�p
          sL_SelectItem := '2';
          nL_ErrorCounts := nL_ErrorCounts + dtmMainJ.handleChkMisData(sL_SelectItem,sL_CompID,sL_SInvDate2,sL_EInvDate2);
        end;


        if chkMisData3.Checked then
        begin
          //�h�ӫȤ�}�ߦb�P�@�i�o���W
          sL_SelectItem := '3';
          nL_ErrorCounts := nL_ErrorCounts + dtmMainJ.handleChkMisData(sL_SelectItem,sL_CompID,sL_SInvDate2,sL_EInvDate2);
        end;


        if chkMisData4.Checked then
        begin
          //�o�����Y�P�ȪA�ҫؤ��o�����Y���P
          sL_SelectItem := '4';
          nL_ErrorCounts := nL_ErrorCounts + dtmMainJ.handleChkMisData(sL_SelectItem,sL_CompID,sL_SInvDate2,sL_EInvDate2);
        end;


        if chkMisData5.Checked then
        begin
          //�o���}�ߤ��νs�P�ȪA�t�Ωҫؤ��νs���P
          sL_SelectItem := '5';
          nL_ErrorCounts := nL_ErrorCounts + dtmMainJ.handleChkMisData(sL_SelectItem,sL_CompID,sL_SInvDate2,sL_EInvDate2);
        end;
      end;
      //************************************************************************

      //�^�_�쪬�A
      dtmMain.setDefaultCursor;
            
      if nL_ErrorCounts <> 0 then
      begin
        frmInvA09_1 := TfrmInvA09_1.Create(Application);
        frmInvA09_1.sG_InvYM := sL_InvYM;
        frmInvA09_1.sG_Prefix := sL_Prefix;
        frmInvA09_1.sG_SInvDate1 := sL_SInvDate1;
        frmInvA09_1.sG_EInvDate1 := sL_EInvDate1;
        frmInvA09_1.sG_SInvDate2 := sL_SInvDate2;
        frmInvA09_1.sG_EInvDate2 := sL_EInvDate2;
        frmInvA09_1.bG_IncludeObsolete := chkCoverObsolete.Checked;
        frmInvA09_1.ShowModal;
        frmInvA09_1.Free;
      end
      else
        MessageDlg('�S�����`���!',mtInformation,[mbOk],0);
    end;
end;

procedure TfrmInvA09.medPrefixExit(Sender: TObject);
begin
    medPrefix.Text := UpperCase(Trim(medPrefix.Text));
end;

procedure TfrmInvA09.chkJumpClick(Sender: TObject);
begin
    if chkJump.Checked then
      medPrefix.Enabled := true
    else
      medPrefix.Enabled := false;
end;

procedure TfrmInvA09.fraSInvDate1Exit(Sender: TObject);
begin
    fraEInvDate1.setYMD(Trim(fraSInvDate1.getYMD));
end;

procedure TfrmInvA09.fraSInvDate2Exit(Sender: TObject);
begin
    fraEInvDate2.setYMD(Trim(fraSInvDate2.getYMD));
end;

end.
