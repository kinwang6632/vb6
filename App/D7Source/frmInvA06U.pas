unit frmInvA06U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ExtCtrls, fraYMDU;

type
  TfrmInvA06 = class(TForm)
    Label2: TLabel;
    fraStartChargeDate: TfraYMD;
    fraEndChargeDate: TfraYMD;
    btnExit: TBitBtn;
    btnQuery: TBitBtn;
    Panel1: TPanel;
    Label5: TLabel;
    Label1: TLabel;
    MsgLable: TLabel;
    Bevel1: TBevel;
    procedure FormShow(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure fraStartChargeDateExit(Sender: TObject);
    procedure btnQueryClick(Sender: TObject);
    procedure fraStartChargeDatemseYMDChange(Sender: TObject);
  private
    { Private declarations }
    sG_StartChargeDate,sG_EndChargeDate : String;
    function isDataOK : Boolean;
  public
    { Public declarations }
    sG_UserID,sG_DataArea,sG_CompID : String;

  end;

var
  frmInvA06: TfrmInvA06;

implementation

uses dtmMainU, frmInvA06_1U, frmMainU, dtmMainJU;

{$R *.dfm}

procedure TfrmInvA06.FormShow(Sender: TObject);
begin
   Self.Caption := frmMain.GetFormTitleString( 'A06', '未開發票資料查詢' );
   MsgLable.Caption := '';
end;

procedure TfrmInvA06.btnExitClick(Sender: TObject);
begin
    Close;
end;

function TfrmInvA06.isDataOK: Boolean;
var
    nL_Len : Integer;
    dL_StartDate,dL_EndDate : TDate;
begin

    sG_StartChargeDate := Trim(fraStartChargeDate.getYMD);
    sG_EndChargeDate := Trim(fraEndChargeDate.getYMD);

    nL_Len := Length(sG_StartChargeDate);
    if (sG_StartChargeDate = '') OR (nL_Len <> 10) then
    begin
      MessageDlg('請填入開始收費日期',mtError,[mbOk],0);
      fraStartChargeDate.mseYMD.SetFocus;
      Result := false;
      Exit;
    end;

    nL_Len := Length(sG_EndChargeDate);
    if (sG_EndChargeDate = '') OR (nL_Len <> 10) then
    begin
      MessageDlg('請填入截止收費日期',mtError,[mbOk],0);
      fraEndChargeDate.mseYMD.SetFocus;
      Result := false;
      Exit;
    end;

    dL_StartDate := StrToDate(sG_StartChargeDate);
    dL_EndDate := StrToDate(sG_EndChargeDate);

    if dL_EndDate < dL_StartDate then
    begin
      MessageDlg('截止日期必須大於等於開始日期',mtError,[mbOk],0);
      fraEndChargeDate.mseYMD.SetFocus;
      Result := false;
      Exit;
    end;

  Result := True;

end;

procedure TfrmInvA06.fraStartChargeDateExit(
  Sender: TObject);
begin
    fraEndChargeDate.setYMD(fraStartChargeDate.getYMD);
end;

procedure TfrmInvA06.btnQueryClick(Sender: TObject);
var
    fL_SaleAmount,fL_TaxAmount,fL_InvAmount : Double;
    nL_TotalCounts : Integer;
begin
   MsgLable.Caption := '';
   
   if not isDataOK then Exit;

   sG_CompID := dtmMain.getCompID;

   nL_TotalCounts := dtmMainJ.getInv016Data( QUERY_TYPE_ALL, '','', 
     sG_CompID, sG_StartChargeDate, sG_EndChargeDate, '', '',  fL_SaleAmount,
     fL_TaxAmount, fL_InvAmount );

   if nL_TotalCounts <= 0 then
   begin
     MsgLable.Caption := '查無資料，請重新輸入收費日期。';
     Exit;
   end;

   frmInvA06_1 := TfrmInvA06_1.Create( Application );
   try
     frmInvA06_1.sG_UserID := sG_UserID;
     frmInvA06_1.sG_DataArea := sG_DataArea;
     frmInvA06_1.sG_StartChargeDate := sG_StartChargeDate;
     frmInvA06_1.sG_EndChargeDate := sG_EndChargeDate;
     frmInvA06_1.sG_CompID := sG_CompID;
     frmInvA06_1.fG_SaleAmount := fL_SaleAmount;
     frmInvA06_1.fG_TaxAmount := fL_TaxAmount;
     frmInvA06_1.fG_InvAmount := fL_InvAmount;
     frmInvA06_1.fG_Counts := nL_TotalCounts;
     frmInvA06_1.sGMinSeq := EmptyStr;
     frmInvA06_1.sGMaxSeq := EmptyStr;
     frmInvA06_1.ShowModal;
   finally
     frmInvA06_1.Free;
   end;
end;

procedure TfrmInvA06.fraStartChargeDatemseYMDChange(Sender: TObject);
begin
  MsgLable.Caption := '';
end;

end.
