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
    self.Caption := frmMain.GetFormTitleString('A09','發票資料檢核');
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
      //統編檢核
      if chkNumber.Checked then
      begin
        if chkBusinessID.Checked or chkJump.Checked then
        begin
          sL_InvYM := Trim(fraInvYM.getYM);
          sL_Prefix := Trim(medPrefix.Text);

          if sL_InvYM = EMPTY_YM_STR then
          begin
            MessageDlg('請輸入發票年月!',mtError,[mbOk],0);
            fraInvYM.mseYM.SetFocus;
            Result := false;
            Exit;
          end;

          if chkJump.Checked then
          begin
            if sL_Prefix = '' then
            begin
              MessageDlg('請輸入字軌!',mtError,[mbOk],0);
              medPrefix.SetFocus;
              Result := false;
              Exit;
            end;

            if Length(sL_Prefix) <> 2 then
            begin
              MessageDlg('請輸入正確的字軌格式!',mtError,[mbOk],0);
              medPrefix.SetFocus;
              Result := false;
              Exit;
            end;
          end;
        end
        else
        begin
          MessageDlg('請點選要檢核的統編檢核項目!',mtError,[mbOk],0);
          chkBusinessID.SetFocus;
          Result := false;
          Exit;
        end;
      end;



      //主副檔金額檢查
      if chkAmount.Checked then
      begin
        sL_SInvDate1 := Trim(fraSInvDate1.getYMD);
        sL_EInvDate1 := Trim(fraEInvDate1.getYMD);

        if sL_SInvDate1=EMPTY_DATE_STR then
        begin
          MessageDlg('請輸入發票開始日期',mtError,[mbOk],0);
          fraSInvDate1.mseYMD.SetFocus;
          Result := false;
          Exit;
        end;

        if sL_EInvDate1=EMPTY_DATE_STR then
        begin
          MessageDlg('請輸入發票結束日期',mtError,[mbOk],0);
          fraEInvDate1.mseYMD.SetFocus;
          Result := false;
          Exit;
        end;

        if sL_SInvDate1 > sL_EInvDate1 then
        begin
          MessageDlg('發票開始日期必須小於結束日期',mtError,[mbOk],0);
          fraEInvDate1.mseYMD.SetFocus;
          Result := false;
          Exit;
        end;
      end;



      //發票開立與客服資料比對
      if chkMisData.Checked then
      begin
        if chkMisData1.Checked or chkMisData2.Checked or chkMisData3.Checked or
           chkMisData4.Checked or chkMisData5.Checked then
        begin
          sL_SInvDate2 := Trim(fraSInvDate2.getYMD);
          sL_EInvDate2 := Trim(fraEInvDate2.getYMD);

          if sL_SInvDate2=EMPTY_DATE_STR then
          begin
            MessageDlg('請輸入發票開始日期',mtError,[mbOk],0);
            fraSInvDate2.mseYMD.SetFocus;
            Result := false;
            Exit;
          end;

          if sL_EInvDate2=EMPTY_DATE_STR then
          begin
            MessageDlg('請輸入發票結束日期',mtError,[mbOk],0);
            fraEInvDate2.mseYMD.SetFocus;
            Result := false;
            Exit;
          end;

          if sL_SInvDate2 > sL_EInvDate2 then
          begin
            MessageDlg('發票開始日期必須小於結束日期',mtError,[mbOk],0);
            fraEInvDate2.mseYMD.SetFocus;
            Result := false;
            Exit;
          end;
        end
        else
        begin
          MessageDlg('請點選任一發票開立與客服資料比對項目!',mtError,[mbOk],0);
          chkBusinessID.SetFocus;
          Result := false;
          Exit;
        end;
      end;
    end
    else
    begin
      MessageDlg('請至少選擇一項檢核項目!',mtError,[mbOk],0);
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
      //執行等待狀態
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
      //統編檢核
      if chkNumber.Checked then
      begin
        sL_InvYM := TUstr.replaceStr(Trim(fraInvYM.getYM),'/','');
        sL_Prefix := Trim(medPrefix.Text);
        //統一編號檢核
        if chkBusinessID.Checked then
        begin
          sL_Year := copy(sL_InvYM,1,4);
          sL_Month := copy(sL_InvYM,5,2);
          sL_DaysOfMonth := IntToStr(TUdateTime.DaysOfMonth(IntToStr(StrToInt(sL_Year)-1911)+sL_Month+'01'));

          sL_CheckNumStrDate := sL_Year+'/'+sL_Month+'/01';
          sL_CheckNumEndDate := sL_Year+'/'+sL_Month+'/'+sL_DaysOfMonth;

          nL_ErrorCounts := nL_ErrorCounts + dtmMainJ.handleCheckBusinessID(sL_CompID,sL_CheckNumStrDate,sL_CheckNumEndDate);
        end;

        //跳號檢核
        if chkJump.Checked then
        begin
          sL_InvYM := dtmMainJ.changePrefixYM(sL_InvYM);
          nL_ErrorCounts := nL_ErrorCounts + dtmMainJ.handleCheckJumpInvID(sL_CompID,sL_InvYM,sL_Prefix);
        end;
      end;



      //************************************************************************
      //主副檔金額檢查
      if chkAmount.Checked then
      begin
        sL_SInvDate1 := Trim(fraSInvDate1.getYMD);
        sL_EInvDate1 := Trim(fraEInvDate1.getYMD);

        nL_ErrorCounts := nL_ErrorCounts + dtmMainJ.handleChkAmountData(StrToDateTime(sL_SInvDate1), StrToDateTime(sL_EInvDate1), chkCoverObsolete.checked,sL_CompID);
      end;



      //************************************************************************
      //發票開立與客服資料比對
      if chkMisData.Checked then
      begin
        sL_SInvDate2 := Trim(fraSInvDate2.getYMD);
        sL_EInvDate2 := Trim(fraEInvDate2.getYMD);

        if chkMisData1.Checked then
        begin
          //發票收件人與客服主檔姓名不同或與帳號收件人不同
          sL_SelectItem := '1';
          nL_ErrorCounts := nL_ErrorCounts + dtmMainJ.handleChkMisData(sL_SelectItem,sL_CompID,sL_SInvDate2,sL_EInvDate2);
        end;


        if chkMisData2.Checked then
        begin
          //應為三聯開成二聯(有帳號)或應為三聯開成二聯
          sL_SelectItem := '2';
          nL_ErrorCounts := nL_ErrorCounts + dtmMainJ.handleChkMisData(sL_SelectItem,sL_CompID,sL_SInvDate2,sL_EInvDate2);
        end;


        if chkMisData3.Checked then
        begin
          //多個客戶開立在同一張發票上
          sL_SelectItem := '3';
          nL_ErrorCounts := nL_ErrorCounts + dtmMainJ.handleChkMisData(sL_SelectItem,sL_CompID,sL_SInvDate2,sL_EInvDate2);
        end;


        if chkMisData4.Checked then
        begin
          //發票抬頭與客服所建之發票抬頭不同
          sL_SelectItem := '4';
          nL_ErrorCounts := nL_ErrorCounts + dtmMainJ.handleChkMisData(sL_SelectItem,sL_CompID,sL_SInvDate2,sL_EInvDate2);
        end;


        if chkMisData5.Checked then
        begin
          //發票開立之統編與客服系統所建之統編不同
          sL_SelectItem := '5';
          nL_ErrorCounts := nL_ErrorCounts + dtmMainJ.handleChkMisData(sL_SelectItem,sL_CompID,sL_SInvDate2,sL_EInvDate2);
        end;
      end;
      //************************************************************************

      //回復原狀態
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
        MessageDlg('沒有異常資料!',mtInformation,[mbOk],0);
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
