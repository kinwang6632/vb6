unit frmMainU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, ComCtrls, StdCtrls, ExtCtrls, DB, ADODB;

type
  TfrmMain = class(TForm)
    pnlProgStatus: TPanel;
    chbActivate: TCheckBox;
    MainMenu1: TMainMenu;
    N11: TMenuItem;
    A1: TMenuItem;
    B1: TMenuItem;
    N1: TMenuItem;
    X1: TMenuItem;
    timPeriodCount: TTimer;
    ADOQuery1: TADOQuery;
    memErrorLog: TMemo;
    StatusBar1: TStatusBar;
    pgcLV: TPageControl;
    TabSheet1: TTabSheet;
    livStatus006: TListView;
    TabSheet2: TTabSheet;
    livStatus007: TListView;
    TabSheet3: TTabSheet;
    livStatusSO151: TListView;
    TabSheet4: TTabSheet;
    livStatusSO153: TListView;
    timProgStatus: TTimer;
    procedure FormCreate(Sender: TObject);
    procedure B1Click(Sender: TObject);
    procedure X1Click(Sender: TObject);
    procedure A1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure chbActivateClick(Sender: TObject);
    procedure timPeriodCountTimer(Sender: TObject);
    procedure timProgStatusTimer(Sender: TObject);
  private
    { Private declarations }
    sG_UserID, sG_UserName : String;
    sG_SysName : String;
    nG_ProcessPeriod : Integer;
    bG_LogData, bG_ShowUI, bG_AcceptNullData : boolean;
    nG_006LivItemPos,nG_007LivItemPos,nG_SO151LivItemPos,nG_SO153LivItemPos,  nG_006UiSeqNo, nG_007UiSeqNo, nG_So151UiSeqNo, nG_SO153UiSeqNo : Integer;
    bG_LoginOk : boolean;
    procedure processEmcEbt006;
    procedure processEmcEbt007;    
    procedure prepare006ListItem(var I_ListItem:TListItem);
    procedure prepare007ListItem(var I_ListItem:TListItem);

    procedure prepareSo153ListItem(var I_ListItem:TListItem);

    procedure createLivItem;
    procedure getSysParam;
    function startProcess: Boolean;
    procedure clearStatusItem(I_Lv:TListView; nI_ItemIndex:Integer);
    procedure setEMC_EBT006Ui(I_DataSet:TDataSet; var sI_UiSeqNo:String);
    procedure setEMC_EBT007Ui(I_DataSet:TDataSet; var sI_UiSeqNo:String);

    function is006DataOk(I_DataSet:TDataSet; var sI_ErrorCode, sI_ErrorMsg:String):boolean;
    function is007DataOk(I_DataSet:TDataSet; var sI_ErrorCode, sI_ErrorMsg:String):boolean;
    procedure showErrorOn006Ui(I_ListItem:TListItem; sI_ErrorCode, sI_ErrorMsg, sI_ProcessFlag:String);
    procedure showErrorOn007Ui(I_ListItem:TListItem; sI_ErrorCode, sI_ErrorMsg, sI_ProcessFlag:String);
    procedure prepareSo151ListItem(var I_ListItem: TListItem);
    procedure setProgStatus;
  public
    { Public declarations }
    procedure write006Error(I_DataSet:TDataSet;sI_ErrorCode, sI_ErrorMsg, sI_UiSeqNo:String);
    procedure write007Error(I_DataSet:TDataSet;sI_ErrorCode, sI_ErrorMsg, sI_UiSeqNo:String);
    procedure showErrorOnMemo(sI_ErrorMsg : String);
    procedure setSo151Ui(I_DataSet:TDataSet);
    procedure setSo153Ui(I_DataSet:TDataSet);
    function getFrmCaption(sI_FrmDesc:String):String;
    function getIfAcceptNullData:boolean;
  end;

var
  frmMain: TfrmMain;

implementation

uses frmLoginU, frmUserU, frmParamU, dtmMainU, ConstObjU;

{$R *.dfm}

procedure TfrmMain.FormCreate(Sender: TObject);
var

    L_frmLogin: TfrmLogin;
begin
    L_frmLogin := TfrmLogin.Create(Application);
    with L_frmLogin do
    begin

      if ShowModal=mrCancel then
      begin
        bG_LoginOk := false;
      end
      else
      begin
        bG_LoginOk := true;
        self.sG_UserID := L_frmLogin.sG_UserID;
        self.sG_UserName := L_frmLogin.sG_UserName;
        Close;
        free;
      end;
    end;



    if bG_LoginOk then
    begin

    end
    else
    begin
      L_frmLogin.Free;    
      Application.Terminate;
    end;

end;

procedure TfrmMain.B1Click(Sender: TObject);
var
    L_frmUser : TfrmUser;
begin
    L_frmUser := TfrmUser.Create(Application);
    L_frmUser.ShowModal;
    L_frmUser.Free;
end;

procedure TfrmMain.X1Click(Sender: TObject);
begin
    if MessageDlg('是否確認要結束程式?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      dtmMain.CloseDbConn;
      Application.Terminate;
    end;
end;

procedure TfrmMain.A1Click(Sender: TObject);
var
    L_frmParam : TfrmParam;
begin
    L_frmParam := TfrmParam.Create(Application);
    L_frmParam.ShowModal;
    L_frmParam.Free;
end;

function TfrmMain.getFrmCaption(sI_FrmDesc: String): String;
var
    sL_Result : String;
begin

    if sG_SysName<>'' then
      sL_Result := sG_SysName + ' - ' + sI_FrmDesc + ' [' + sG_UserName + ']'
    else
      sL_Result := sI_FrmDesc + ' [' + sG_UserName + ']';

    result :=  sL_Result;
end;

procedure TfrmMain.FormShow(Sender: TObject);
var
    sL_Result : String;
begin
    if bG_LoginOk then
    begin
      dtmMain.GetCompanyCodeToStringList;
      dtmMain.getCompMappingData;
      getSysParam;
      Caption := getFrmCaption('資料交換');
      nG_006UiSeqNo := 1;
      nG_007UiSeqNo := 1;
      nG_SO151UiSeqNo := 1;
      nG_SO153UiSeqNo := 1;            
      createLivItem;
    end;
end;

procedure TfrmMain.getSysParam;
begin
    sG_SysName := dtmMain.getSysName;
    nG_ProcessPeriod := dtmMain.getProcessPeriod;
    bG_ShowUI := dtmMain.getIfShowUI;
    bG_LogData := dtmMain.getIfLogData;
    bG_AcceptNullData := dtmMain.getIfAcceptNullData;

end;

procedure TfrmMain.chbActivateClick(Sender: TObject);
begin
    if chbActivate.Checked then
    begin
      if nG_ProcessPeriod<=0 then
      begin
        chbActivate.Checked := false;
        Application.MessageBox(
          '系統參數中資料處理週期設定錯誤!無法啟動資料處理機制。', '警告',
           MB_OK + MB_ICONWARNING );
      end
      else
      begin
        pnlProgStatus.Caption := ACTIVATE_CAPTION;        
        timProgStatus.Enabled := true;
        StartProcess;      
      end;
    end
    else
    begin
      pnlProgStatus.Caption := DEACTIVATE_CAPTION;    
      timPeriodCount.Enabled := false;
      timProgStatus.Enabled := false;
    end;
end;

function TfrmMain.startProcess: Boolean;
var
    L_DataSet : TDataSet;
    sL_ErrorCode, sL_ErrorMsg:String;
    sL_UiSeqNo : String;
    sL_Result: string;
begin
    Result := True;
    { 先將觸發的 startProcess 的 Timer 關掉, 全部處理完後再打開 }
    timPeriodCount.Enabled := False;
    try
       timPeriodCount.Interval := nG_ProcessPeriod*60*1000;
       pnlProgStatus.Caption := DISCONNECT_FROM_DB;
       pnlProgStatus.Update;
       dtmMain.CloseDbConn;
       Sleep( 2000 );
       pnlProgStatus.Caption := DISCONNECT_FROM_DB_OK;
       pnlProgStatus.Update;
       Sleep( 2000 );
       pnlProgStatus.Caption := CONNECT_TO_DB;
       pnlProgStatus.Update;
       sL_Result := dtmMain.connectToDB;
       if sL_Result <> '' then
       begin
         showErrorOnMemo( sL_Result );
         Result := False;
         Exit;
       end;
       pnlProgStatus.Caption := ACTIVATE_CAPTION;
       pnlProgStatus.Update;
      //down, 開始處理各項資料... EBT-->EMC
      try
        processEmcEbt006;
      except
        on E: Exception do
          showErrorOnMemo('處理EMC_EBT006時發生錯誤! ' + 'HelpContext:' + IntToStr(E.HelpContext)+ ' Error Message:' + E.Message );
      end;
      //EBT-->EMC
      try
        processEmcEbt007;
      except
        on E: Exception do
          showErrorOnMemo('處理EMC_EBT007時發生錯誤! ' + 'HelpContext:' + IntToStr(E.HelpContext)+ ' Error Message:' + E.Message );
      end;
     
      //其它系統台資料 EMC --> EBT
      dtmMain.processSoData;
    finally
      timPeriodCount.Enabled := True;
    end;
end;

procedure TfrmMain.timPeriodCountTimer(Sender: TObject);
begin
    if chbActivate.Checked then
      startProcess;
end;

procedure TfrmMain.setEMC_EBT006Ui(I_DataSet: TDataSet;var sI_UiSeqNo:String);
var
    nL_NextItemIndex : Integer;
    L_ListItem : TListItem;
    sL_EmcCustID,sL_CatvValid : String;
    sL_ErrorCode,sL_ErrorMsg, sL_ProcessFlag : String;
begin
    if bG_ShowUI then
    begin
      prepare006ListItem(L_ListItem);

      sI_UiSeqNo := IntToStr(nG_006UiSeqNo);
      L_ListItem.Caption := sI_UiSeqNo;



      sL_EmcCustID := I_DataSet.FieldByName('EmcCustID').AsString;
      sL_CatvValid := UpperCase(I_DataSet.FieldByName('CATVVALID').AsString);

      if sL_CatvValid='N' then
        L_ListItem.SubItems[0] := '待賦予客編'
      else
        L_ListItem.SubItems[0] := '完工資料';

      L_ListItem.SubItems[1] := I_DataSet.FieldByName('EmcCustID').AsString;
      L_ListItem.SubItems[2] := I_DataSet.FieldByName('EbtCatvID').AsString;
      L_ListItem.SubItems[3] := I_DataSet.FieldByName('EbtContractNo').AsString;
      L_ListItem.SubItems[4] := I_DataSet.FieldByName('EbtCustID').AsString;
      L_ListItem.SubItems[5] := I_DataSet.FieldByName('EbtContractBDate').AsString;
      L_ListItem.SubItems[6] := I_DataSet.FieldByName('EbtContractEDate').AsString;
      L_ListItem.SubItems[7] := I_DataSet.FieldByName('EbtCustCName').AsString;
      L_ListItem.SubItems[8] := I_DataSet.FieldByName('EbtCustContactPhone').AsString;
      L_ListItem.SubItems[9] := I_DataSet.FieldByName('EbtCustContactMobile').AsString;
      L_ListItem.SubItems[10] := I_DataSet.FieldByName('EbtCompOwnerName').AsString;

      L_ListItem.SubItems[11] := I_DataSet.FieldByName('EbtContactPhone').AsString;
      L_ListItem.SubItems[12] := I_DataSet.FieldByName('EbtItContactName').AsString;
      L_ListItem.SubItems[13] := I_DataSet.FieldByName('EbtItContactPhone').AsString;

      L_ListItem.SubItems[14] := I_DataSet.FieldByName('EbtItContactMobile').AsString;
      L_ListItem.SubItems[15] := I_DataSet.FieldByName('EbtItEMail').AsString;
      L_ListItem.SubItems[16] := I_DataSet.FieldByName('EbtInstAddr').AsString;
      L_ListItem.SubItems[17] := I_DataSet.FieldByName('EbtCustAddr').AsString;
      L_ListItem.SubItems[18] := I_DataSet.FieldByName('EbtBillAddr').AsString;


      L_ListItem.SubItems[19] := I_DataSet.FieldByName('EbtContractStatusDesc').AsString;
      L_ListItem.SubItems[20] := I_DataSet.FieldByName('EbtFeePeriodDesc').AsString;
      L_ListItem.SubItems[21] := I_DataSet.FieldByName('EbtServiceType').AsString;
      L_ListItem.SubItems[22] := I_DataSet.FieldByName('EbtAgentName').AsString;
      L_ListItem.SubItems[23] := I_DataSet.FieldByName('EbtAgentPhone').AsString;

      L_ListItem.SubItems[24] := I_DataSet.FieldByName('EbtAgentID').AsString;
      L_ListItem.SubItems[25] := I_DataSet.FieldByName('EbtAgentAddress').AsString;

      L_ListItem.SubItems[26] := I_DataSet.FieldByName('EbtIdCardId').AsString;
      L_ListItem.SubItems[27] := I_DataSet.FieldByName('EbtCompanyOwnerId').AsString;
      L_ListItem.SubItems[28] := I_DataSet.FieldByName('EbtNonProfitCompanyId').AsString;
      L_ListItem.SubItems[29] := I_DataSet.FieldByName('EbtCompanyId').AsString;
      sL_ProcessFlag := I_DataSet.FieldByName('ProcessFlag').AsString;

      sL_ErrorCode := I_DataSet.FieldByName('ErrorCode').AsString;
      sL_ErrorMsg := I_DataSet.FieldByName('ErrorMsg').AsString;
      showErrorOn006Ui(L_ListItem,sL_ErrorCode,sL_ErrorMsg, sL_ProcessFlag);


      L_ListItem.SubItems[33] := I_DataSet.FieldByName('EbtNotes').AsString;
      L_ListItem.SubItems[34] := I_DataSet.FieldByName('EmcNotes').AsString;
      L_ListItem.SubItems[35] := I_DataSet.FieldByName('CATVVALID').AsString;
      L_ListItem.SubItems[36] := I_DataSet.FieldByName('UpdTime').AsString;
      L_ListItem.SubItems[37] := I_DataSet.FieldByName('UpdEn').AsString;
      L_ListItem.SubItems[38] := I_DataSet.FieldByName('TAG').AsString;

      
      Inc(nG_006UiSeqNo);

    end;




end;

procedure TfrmMain.createLivItem;
var
    ii, jj : Integer;
    L_ListItem : TListItem;

begin
    if bG_ShowUI then
    begin
      nG_006LivItemPos := 0;
      nG_007LivItemPos := 0;
      nG_SO151LivItemPos := 0;
      nG_SO153LivItemPos := 0;

      pgcLV.Visible := true;
      pgcLV.Align := alClient;

      frmMain.livStatus006.Items.Clear;
      for ii:=0 to CA_UI_MAX_COUNT -1 do
      begin
        L_ListItem := frmMain.livStatus006.Items.Add;

        for jj:=0 to LISTVIEW006_SUB_ITEM_COUNT-1 do
         L_ListItem.SubItems.Add('');

      end;

      frmMain.livStatus007.Items.Clear;
      for ii:=0 to CA_UI_MAX_COUNT -1 do
      begin
        L_ListItem := frmMain.livStatus007.Items.Add;

        for jj:=0 to LISTVIEW007_SUB_ITEM_COUNT-1 do
         L_ListItem.SubItems.Add('');

      end;

      frmMain.livStatusSO151.Items.Clear;
      for ii:=0 to CA_UI_MAX_COUNT -1 do
      begin
        L_ListItem := frmMain.livStatusSO151.Items.Add;

        for jj:=0 to LISTVIEWSO151_SUB_ITEM_COUNT-1 do
         L_ListItem.SubItems.Add('');

      end;


      frmMain.livStatusSO153.Items.Clear;
      for ii:=0 to CA_UI_MAX_COUNT -1 do
      begin
        L_ListItem := frmMain.livStatusSO153.Items.Add;

        for jj:=0 to LISTVIEWSO153_SUB_ITEM_COUNT-1 do
         L_ListItem.SubItems.Add('');

      end;
    end
    else
    begin
      pgcLV.Visible := false;
      memErrorLog.Align := alClient;
    end;
end;

procedure TfrmMain.clearStatusItem(I_Lv:TListView; nI_ItemIndex: Integer);
var
    jj : Integer;
begin
      I_Lv.Items[nI_ItemIndex].Caption := '';
      for jj:=0 to I_Lv.Items[nI_ItemIndex].SubItems.Count -1 do
        I_Lv.Items[nI_ItemIndex].SubItems.Strings[jj] := '';

end;

procedure TfrmMain.prepare006ListItem(var I_ListItem: TListItem);
var
    nL_NextItemIndex : Integer;
begin

      if frmMain.nG_006LivItemPos=CA_UI_MAX_COUNT then
      begin
//        clearStatusItem; //呼叫此 function...似乎會讓 performance 下降...
        frmMain.nG_006LivItemPos := 0;
      end;
      I_ListItem := frmMain.livStatus006.Items[frmMain.nG_006LivItemPos];
      Inc(frmMain.nG_006LivItemPos);

      //down, 將下一個 ListItem 清除成空白
      if frmMain.nG_006LivItemPos=CA_UI_MAX_COUNT then
        nL_NextItemIndex := 0
      else
        nL_NextItemIndex := frmMain.nG_006LivItemPos;

      clearStatusItem(livStatus006, nL_NextItemIndex);


end;

function TfrmMain.is006DataOk(I_DataSet: TDataSet; var sI_ErrorCode, sI_ErrorMsg:String): boolean;
var
    bL_Result : boolean;
    sL_EbtCustCName,  sL_EbtCustContactPhone, sL_EbtCustContactMobile : String;
    sL_EmcCustID, sL_EbtInstAddr,   sL_EbtServiceType, sL_CatvValid : String;

    sL_EbtCatvID,sL_EbtContractNo,sL_EbtCustID : String;
begin
    bL_Result := true;


    sL_EbtContractNo := I_DataSet.FieldByName('EbtContractNo').AsString;
    sL_EbtCatvID := I_DataSet.FieldByName('EbtCatvID').AsString;
    sL_EbtCustCName := I_DataSet.FieldByName('EbtCustCName').AsString;
    sL_EbtCustContactPhone := I_DataSet.FieldByName('EbtCustContactPhone').AsString;
    sL_EbtCustContactMobile := I_DataSet.FieldByName('EbtCustContactMobile').AsString;
    sL_EbtInstAddr := I_DataSet.FieldByName('EbtInstAddr').AsString;
    sL_EbtServiceType := I_DataSet.FieldByName('EbtServiceType').AsString;
    sL_CatvValid := I_DataSet.FieldByName('CatvValid').AsString;
    sL_EmcCustID := I_DataSet.FieldByName('EmcCustID').AsString;
    sL_EbtCustID := I_DataSet.FieldByName('EbtCustID').AsString;

    sI_ErrorCode := '';
    sI_ErrorMsg := '';
    if sL_EbtCustCName='' then
    begin
      sI_ErrorCode := IntToStr(ERR_CODE_01);
      sI_ErrorMsg := ERR_MSG_01;
    end
    else if (sL_EbtCustContactPhone='') and (sL_EbtCustContactMobile='') then
    begin
      sI_ErrorCode := IntToStr(ERR_CODE_02);
      sI_ErrorMsg := ERR_MSG_02;
    end
    else if sL_EbtInstAddr='' then
    begin
      sI_ErrorCode := IntToStr(ERR_CODE_03);
      sI_ErrorMsg := ERR_MSG_03;
    end
    else if sL_EbtCatvID='' then
    begin
      sI_ErrorCode := IntToStr(ERR_CODE_08);
      sI_ErrorMsg := ERR_MSG_08;
    end
    else if sL_EbtContractNo='' then
    begin
      sI_ErrorCode := IntToStr(ERR_CODE_09);
      sI_ErrorMsg := ERR_MSG_09;
    end
    else if sL_EbtCustID='' then
    begin
      sI_ErrorCode := IntToStr(ERR_CODE_11);
      sI_ErrorMsg := ERR_MSG_11;
    end

    else if sL_EbtServiceType='' then
    begin
      sI_ErrorCode := IntToStr(ERR_CODE_04);
      sI_ErrorMsg := ERR_MSG_04;
    end
    else if sL_CatvValid='' then
    begin
      sI_ErrorCode := IntToStr(ERR_CODE_05);
      sI_ErrorMsg := ERR_MSG_05;
    end
    else if (sL_CatvValid='Y') and (sL_EmcCustID='') then
    begin
      sI_ErrorCode := IntToStr(ERR_CODE_06);
      sI_ErrorMsg := ERR_MSG_06;
    end;
    if sI_ErrorCode<>'' then
      bL_Result := false;
    result := bL_Result;
end;

procedure TfrmMain.write006Error(I_DataSet: TDataSet; sI_ErrorCode,
  sI_ErrorMsg,sI_UiSeqNo: String);
var
    sL_EbtContractNo, sL_ProcessFlag : String;
    L_ListItem : TListItem;
begin
    L_ListItem := livStatus006.FindCaption(0,sI_UiSeqNo,false,true,false);
    if L_ListItem<>nil then
    begin
      showErrorOn006Ui(L_ListItem,sI_ErrorCode,sI_ErrorMsg, sL_ProcessFlag);
    end;

    sL_EbtContractNo := I_DataSet.FieldByName('EbtContractNo').AsString ;
    dtmMain.updateEmcEbt006(sL_EbtContractNo,  sI_ErrorCode, sI_ErrorMsg, PROCESS_FLAG_4);
end;

procedure TfrmMain.showErrorOn006Ui(I_ListItem: TListItem; sI_ErrorCode,
  sI_ErrorMsg,sI_ProcessFlag: String);
begin
    if (I_ListItem<>nil) then
    begin
      if sI_ProcessFlag<>'' then
        I_ListItem.SubItems[30] := sI_ProcessFlag;
      I_ListItem.SubItems[31] := sI_ErrorCode;
      I_ListItem.SubItems[32] := sI_ErrorMsg;
    end;
end;

procedure TfrmMain.processEmcEbt006;
var
    L_DataSet : TDataSet;
    sL_UiSeqNo, sL_ErrorCode, sL_ErrorMsg : String;
begin
    L_DataSet := dtmMain.getEmcEbt006Data;

    with L_DataSet do
    begin
      while not Eof do
      begin
        setEMC_EBT006Ui(L_DataSet, sL_UiSeqNo);
        Application.ProcessMessages;
        if is006DataOk(L_DataSet,sL_ErrorCode, sL_ErrorMsg) then //若此筆data沒有缺任何欄位...
        begin
          dtmMain.dispatch006DataToSo(L_DataSet);
        end
        else
        begin
          write006Error(L_DataSet,sL_ErrorCode, sL_ErrorMsg, sL_UiSeqNo);
        end;
        Next;
        Application.ProcessMessages;
      end;
      Close;
    end;

end;

procedure TfrmMain.processEmcEbt007;
var
    L_DataSet : TDataSet;
    sL_UiSeqNo, sL_ErrorCode, sL_ErrorMsg : String;
begin
    L_DataSet := dtmMain.getEmcEbt007Data;

    with L_DataSet do
    begin
      while not Eof do
      begin
        setEMC_EBT007Ui(L_DataSet, sL_UiSeqNo);
        Application.ProcessMessages;
        if is007DataOk(L_DataSet,sL_ErrorCode, sL_ErrorMsg) then //若此筆data沒有缺任何欄位...
        begin
          dtmMain.dispatch007DataToSo(L_DataSet);
        end
        else
        begin
          write007Error(L_DataSet,sL_ErrorCode, sL_ErrorMsg, sL_UiSeqNo);
        end;
        Next;
        Application.ProcessMessages;
      end;
      Close;
    end;

end;

procedure TfrmMain.setEMC_EBT007Ui(I_DataSet: TDataSet;
  var sI_UiSeqNo: String);
var
    nL_NextItemIndex : Integer;
    L_ListItem : TListItem;
    sL_EmcCustID,sL_CatvValid : String;
    sL_ErrorCode,sL_ErrorMsg, sL_ProcessFlag : String;

    sL_IfMoveToOtherSo : String;
begin
    if bG_ShowUI then
    begin
      prepare007ListItem(L_ListItem);

      sI_UiSeqNo := IntToStr(nG_007UiSeqNo);
      L_ListItem.Caption := sI_UiSeqNo;

      dtmMain.EBT007_UIDataSet.Append;
      dtmMain.EBT007_UIDataSet.FieldByName( 'ROWID' ).AsString :=
        I_DataSet.FieldByName( 'ROWIDS' ).AsString; 
      dtmMain.EBT007_UIDataSet.FieldByName( 'SEQ' ).AsString :=
        IntToStr( nG_007UiSeqNo );
      dtmMain.EBT007_UIDataSet.Post;


      sL_EmcCustID := I_DataSet.FieldByName('EmcCustID').AsString;
      sL_IfMoveToOtherSo := UpperCase(I_DataSet.FieldByName('IfMoveToOtherSo').AsString);




      if sL_IfMoveToOtherSo='Y' then
        L_ListItem.SubItems[0] := '跨系統台移機'
      else if sL_IfMoveToOtherSo='N' then
        L_ListItem.SubItems[0] := '移機'
      else
        L_ListItem.SubItems[0] := '資料異動';

      L_ListItem.SubItems[1] := I_DataSet.FieldByName('EmcCustID').AsString;
      L_ListItem.SubItems[2] := I_DataSet.FieldByName('EbtCatvID').AsString;
      L_ListItem.SubItems[3] := I_DataSet.FieldByName('EbtContractNo').AsString;
      L_ListItem.SubItems[4] := I_DataSet.FieldByName('EbtModifySerialNo').AsString;
      L_ListItem.SubItems[5] := I_DataSet.FieldByName('EbtCustID').AsString;
      L_ListItem.SubItems[6] := I_DataSet.FieldByName('EbtContractBDate').AsString;
      L_ListItem.SubItems[7] := I_DataSet.FieldByName('EbtContractEDate').AsString;
      L_ListItem.SubItems[8] := I_DataSet.FieldByName('EbtCustCName').AsString;
      L_ListItem.SubItems[9] := I_DataSet.FieldByName('EbtCustContactPhone').AsString;
      L_ListItem.SubItems[10] := I_DataSet.FieldByName('EbtCustContactMobile').AsString;
      L_ListItem.SubItems[11] := I_DataSet.FieldByName('EbtCompOwnerName').AsString;

      L_ListItem.SubItems[12] := I_DataSet.FieldByName('EbtContactPhone').AsString;
      L_ListItem.SubItems[13] := I_DataSet.FieldByName('EbtItContactName').AsString;
      L_ListItem.SubItems[14] := I_DataSet.FieldByName('EbtItContactPhone').AsString;

      L_ListItem.SubItems[15] := I_DataSet.FieldByName('EbtItContactMobile').AsString;
      L_ListItem.SubItems[16] := I_DataSet.FieldByName('EbtItEMail').AsString;
      L_ListItem.SubItems[17] := I_DataSet.FieldByName('EbtInstAddr').AsString;
      L_ListItem.SubItems[18] := I_DataSet.FieldByName('EbtCustAddr').AsString;
      L_ListItem.SubItems[19] := I_DataSet.FieldByName('EbtBillAddr').AsString;


      L_ListItem.SubItems[20] := I_DataSet.FieldByName('EbtContractStatusDesc').AsString;
      L_ListItem.SubItems[21] := I_DataSet.FieldByName('EbtFeePeriodDesc').AsString;
      L_ListItem.SubItems[22] := I_DataSet.FieldByName('EbtServiceType').AsString;
      L_ListItem.SubItems[23] := I_DataSet.FieldByName('EbtAgentName').AsString;
      L_ListItem.SubItems[24] := I_DataSet.FieldByName('EbtAgentPhone').AsString;

      L_ListItem.SubItems[25] := I_DataSet.FieldByName('EbtAgentID').AsString;
      L_ListItem.SubItems[26] := I_DataSet.FieldByName('EbtAgentAddress').AsString;

      L_ListItem.SubItems[27] := I_DataSet.FieldByName('EbtIdCardId').AsString;
      L_ListItem.SubItems[28] := I_DataSet.FieldByName('EbtCompanyOwnerId').AsString;
      L_ListItem.SubItems[29] := I_DataSet.FieldByName('EbtNonProfitCompanyId').AsString;
      L_ListItem.SubItems[30] := I_DataSet.FieldByName('EbtCompanyId').AsString;

      L_ListItem.SubItems[31] := I_DataSet.FieldByName('EbtReqType').AsString;
      L_ListItem.SubItems[32] := I_DataSet.FieldByName('OldEmcCustID').AsString;
      L_ListItem.SubItems[33] := I_DataSet.FieldByName('IfMoveToOtherSo').AsString;

      sL_ProcessFlag := I_DataSet.FieldByName('ProcessFlag').AsString;

      sL_ErrorCode := I_DataSet.FieldByName('ErrorCode').AsString;
      sL_ErrorMsg := I_DataSet.FieldByName('ErrorMsg').AsString;
      showErrorOn007Ui(L_ListItem,sL_ErrorCode,sL_ErrorMsg, sL_ProcessFlag);


      L_ListItem.SubItems[37] := I_DataSet.FieldByName('EbtNotes').AsString;
      L_ListItem.SubItems[38] := I_DataSet.FieldByName('CATVVALID').AsString;




      L_ListItem.SubItems[39] := I_DataSet.FieldByName('UpdTime').AsString;
      L_ListItem.SubItems[40] := I_DataSet.FieldByName('UpdEn').AsString;
      L_ListItem.SubItems[41] := I_DataSet.FieldByName('TAG').AsString;


      Inc(nG_007UiSeqNo);

    end;

end;

procedure TfrmMain.prepare007ListItem(var I_ListItem: TListItem);
var
    nL_NextItemIndex : Integer;
begin

      if frmMain.nG_007LivItemPos=CA_UI_MAX_COUNT then
      begin
//        clearStatusItem; //呼叫此 function...似乎會讓 performance 下降...
        frmMain.nG_007LivItemPos := 0;
      end;
      I_ListItem := frmMain.livStatus007.Items[frmMain.nG_007LivItemPos];
      Inc(frmMain.nG_007LivItemPos);

      //down, 將下一個 ListItem 清除成空白
      if frmMain.nG_007LivItemPos=CA_UI_MAX_COUNT then
        nL_NextItemIndex := 0
      else
        nL_NextItemIndex := frmMain.nG_007LivItemPos;

      clearStatusItem(livStatus007, nL_NextItemIndex);


end;

procedure TfrmMain.showErrorOn007Ui(I_ListItem: TListItem; sI_ErrorCode,
  sI_ErrorMsg, sI_ProcessFlag: String);
begin
    if (I_ListItem<>nil) then
    begin
      if sI_ProcessFlag<>'' then
        I_ListItem.SubItems[34] := sI_ProcessFlag;
      I_ListItem.SubItems[35] := sI_ErrorCode;
      I_ListItem.SubItems[36] := sI_ErrorMsg;
    end;
end;

function TfrmMain.is007DataOk(I_DataSet: TDataSet; var sI_ErrorCode,
  sI_ErrorMsg: String): boolean;
var
    bL_Result : boolean;
    sL_EbtCustCName,  sL_EbtCustContactPhone, sL_EbtCustContactMobile : String;
    sL_EmcCustID, sL_EbtInstAddr,   sL_EbtServiceType, sL_CatvValid : String;

    sL_EbtCustID,sL_EbtReqId,sL_EbtReqType, sL_OldEmcCustID : String;
    sL_IfMoveToOtherSo,sL_EbtCatvID,sL_EbtContractNo,sL_EbtModifySerialNo : String;
begin
    bL_Result := true;

    sL_EbtCustID := I_DataSet.FieldByName('EbtCustID').AsString;
    sL_EbtCustCName := I_DataSet.FieldByName('EbtCustCName').AsString;
    sL_EbtCustContactPhone := I_DataSet.FieldByName('EbtCustContactPhone').AsString;
    sL_EbtCustContactMobile := I_DataSet.FieldByName('EbtCustContactMobile').AsString;
    sL_EbtInstAddr := I_DataSet.FieldByName('EbtInstAddr').AsString;
    sL_EbtCatvID := I_DataSet.FieldByName('EbtServiceType').AsString;
    sL_EbtServiceType := I_DataSet.FieldByName('EbtCatvID').AsString;
    sL_CatvValid := I_DataSet.FieldByName('CatvValid').AsString;
    sL_EmcCustID := I_DataSet.FieldByName('EmcCustID').AsString;
    sL_EbtContractNo := I_DataSet.FieldByName('EbtContractNo').AsString;
    sL_EbtModifySerialNo := I_DataSet.FieldByName('EbtModifySerialNo').AsString;

    sL_EbtReqId := I_DataSet.FieldByName('EbtReqId').AsString;
    sL_EbtReqType := I_DataSet.FieldByName('EbtReqType').AsString;


    sL_IfMoveToOtherSo := UpperCase(I_DataSet.FieldByName('IfMoveToOtherSo').AsString);
    sL_OldEmcCustID := UpperCase(I_DataSet.FieldByName('OldEmcCustID').AsString);


    sI_ErrorCode := '';
    sI_ErrorMsg := '';
    if (sL_IfMoveToOtherSo='') then
    begin
      sI_ErrorCode := IntToStr(ERR_CODE_07);
      sI_ErrorMsg := ERR_MSG_07;
    end
    else if (sL_IfMoveToOtherSo<>'Y') and (sL_EmcCustID='')then
    begin
      sI_ErrorCode := IntToStr(ERR_CODE_13);
      sI_ErrorMsg := ERR_MSG_13;
    end
    else if (sL_IfMoveToOtherSo='Y') and (sL_OldEmcCustID='')then
    begin
      sI_ErrorCode := IntToStr(ERR_CODE_14);
      sI_ErrorMsg := ERR_MSG_14;
    end
    else if (sL_EbtCatvID='') then
    begin
      sI_ErrorCode := IntToStr(ERR_CODE_08);
      sI_ErrorMsg := ERR_MSG_08;
    end
    else if (sL_EbtContractNo='') then
    begin
      sI_ErrorCode := IntToStr(ERR_CODE_09);
      sI_ErrorMsg := ERR_MSG_09;
    end
    else if (sL_EbtModifySerialNo='') then
    begin
      sI_ErrorCode := IntToStr(ERR_CODE_10);
      sI_ErrorMsg := ERR_MSG_10;
    end
    else if (sL_EbtCustID='') then
    begin
      sI_ErrorCode := IntToStr(ERR_CODE_11);
      sI_ErrorMsg := ERR_MSG_11;
    end
    else if (sL_EbtReqId='') or (sL_EbtReqType='') then
    begin
      sI_ErrorCode := IntToStr(ERR_CODE_12);
      sI_ErrorMsg := ERR_MSG_12;
    end
    else if (sL_EbtCustCName='') then
    begin
      sI_ErrorCode := IntToStr(ERR_CODE_01);
      sI_ErrorMsg := ERR_MSG_01;
    end
    else if (sL_IfMoveToOtherSo='Y')and(sL_EbtCustContactPhone='') and (sL_EbtCustContactMobile='') then
    begin
      sI_ErrorCode := IntToStr(ERR_CODE_02);
      sI_ErrorMsg := ERR_MSG_02;
    end
    else if (sL_EbtInstAddr='') then
    begin
      sI_ErrorCode := IntToStr(ERR_CODE_03);
      sI_ErrorMsg := ERR_MSG_03;
    end
    else if sL_EbtServiceType='' then
    begin
      sI_ErrorCode := IntToStr(ERR_CODE_04);
      sI_ErrorMsg := ERR_MSG_04;
    end
    else if (sL_IfMoveToOtherSo='Y') and (sL_CatvValid='') then
    begin
      sI_ErrorCode := IntToStr(ERR_CODE_05);
      sI_ErrorMsg := ERR_MSG_05;
    end
    else if (sL_CatvValid='Y') and (sL_EmcCustID='') then
    begin
      sI_ErrorCode := IntToStr(ERR_CODE_06);
      sI_ErrorMsg := ERR_MSG_06;
    end;
    if sI_ErrorCode<>'' then
      bL_Result := false;
    result := bL_Result;
end;

procedure TfrmMain.write007Error(I_DataSet: TDataSet; sI_ErrorCode,
  sI_ErrorMsg, sI_UiSeqNo: String);
var
    sL_ProcessFlag : String;
    L_ListItem : TListItem;
begin
    L_ListItem := livStatus007.FindCaption(0,sI_UiSeqNo,false,true,false);
    if L_ListItem<>nil then
    begin
      showErrorOn007Ui(L_ListItem,sI_ErrorCode,sI_ErrorMsg, sL_ProcessFlag);
    end;

    dtmMain.updateEmcEbt007( I_DataSet.FieldByName('EbtContractNo').AsString,
      I_DataSet.FieldByName('EbtModifySerialNo').AsString,
      sI_ErrorCode, sI_ErrorMsg, PROCESS_FLAG_4 );
end;

procedure TfrmMain.prepareSo151ListItem(var I_ListItem: TListItem);
var
    nL_NextItemIndex : Integer;
begin

      if frmMain.nG_SO151LivItemPos=CA_UI_MAX_COUNT then
      begin
//        clearStatusItem; //呼叫此 function...似乎會讓 performance 下降...
        frmMain.nG_SO151LivItemPos := 0;
      end;
      I_ListItem := frmMain.livStatusSO151.Items[frmMain.nG_SO151LivItemPos];
      Inc(frmMain.nG_SO151LivItemPos);

      //down, 將下一個 ListItem 清除成空白
      if frmMain.nG_SO151LivItemPos=CA_UI_MAX_COUNT then
        nL_NextItemIndex := 0
      else
        nL_NextItemIndex := frmMain.nG_SO151LivItemPos;

      clearStatusItem(livStatusSO151, nL_NextItemIndex);


end;

procedure TfrmMain.setSo151Ui(I_DataSet: TDataSet);
var
    nL_NextItemIndex : Integer;
    L_ListItem : TListItem;
    sL_EmcCustID,sL_CatvValid : String;
    sL_ErrorCode,sL_ErrorMsg, sL_ProcessFlag : String;

    sL_IfMoveToOtherSo : String;
begin
    if bG_ShowUI then
    begin
      prepareSo151ListItem(L_ListItem);

      L_ListItem.Caption := IntToStr(nG_So151UiSeqNo);



      L_ListItem.SubItems[0] := I_DataSet.FieldByName('SeqNo').AsString;

      L_ListItem.SubItems[1] := I_DataSet.FieldByName('EmcCompCode').AsString;
      L_ListItem.SubItems[2] := I_DataSet.FieldByName('EmcCustID').AsString;
      L_ListItem.SubItems[3] := I_DataSet.FieldByName('EmcNewCustStatusDesc').AsString;
      L_ListItem.SubItems[4] := I_DataSet.FieldByName('EmcOldCustStatusDesc').AsString;
      L_ListItem.SubItems[5] := I_DataSet.FieldByName('EbtCustID').AsString;
      L_ListItem.SubItems[6] := I_DataSet.FieldByName('EbtContractNo').AsString;
      L_ListItem.SubItems[7] := I_DataSet.FieldByName('EbtServiceType').AsString;
      L_ListItem.SubItems[8] := I_DataSet.FieldByName('ProcessFlag').AsString;
      L_ListItem.SubItems[9] := I_DataSet.FieldByName('UpdTime').AsString;

      L_ListItem.SubItems[10] := I_DataSet.FieldByName('UpdEn').AsString;

      Inc(nG_So151UiSeqNo);

    end;

end;

procedure TfrmMain.setSo153Ui(I_DataSet: TDataSet);
var
    nL_NextItemIndex : Integer;
    L_ListItem : TListItem;
    sL_EmcCustID : String;
    sL_ErrorCode,sL_ErrorMsg, sL_ProcessFlag : String;
begin
    if bG_ShowUI then
    begin
      prepareSo153ListItem(L_ListItem);

      L_ListItem.Caption := IntToStr(nG_So153UiSeqNo);



      sL_EmcCustID := I_DataSet.FieldByName('EmcCustID').AsString;


      L_ListItem.SubItems[0] := '賦予客編完成';

      L_ListItem.SubItems[1] := I_DataSet.FieldByName('EmcCustID').AsString;
      L_ListItem.SubItems[2] := I_DataSet.FieldByName('EbtCatvID').AsString;
      L_ListItem.SubItems[3] := I_DataSet.FieldByName('EbtContractNo').AsString;
      L_ListItem.SubItems[4] := I_DataSet.FieldByName('EbtCustID').AsString;
      L_ListItem.SubItems[5] := I_DataSet.FieldByName('EbtContractBDate').AsString;
      L_ListItem.SubItems[6] := I_DataSet.FieldByName('EbtContractEDate').AsString;
      L_ListItem.SubItems[7] := I_DataSet.FieldByName('EbtCustCName').AsString;
      L_ListItem.SubItems[8] := I_DataSet.FieldByName('EbtCustContactPhone').AsString;
      L_ListItem.SubItems[9] := I_DataSet.FieldByName('EbtCustContactMobile').AsString;
      L_ListItem.SubItems[10] := I_DataSet.FieldByName('EbtCompOwnerName').AsString;

      L_ListItem.SubItems[11] := I_DataSet.FieldByName('EbtContactPhone').AsString;
      L_ListItem.SubItems[12] := I_DataSet.FieldByName('EbtItContactName').AsString;
      L_ListItem.SubItems[13] := I_DataSet.FieldByName('EbtItContactPhone').AsString;

      L_ListItem.SubItems[14] := I_DataSet.FieldByName('EbtItContactMobile').AsString;
      L_ListItem.SubItems[15] := I_DataSet.FieldByName('EbtItEMail').AsString;
      L_ListItem.SubItems[16] := I_DataSet.FieldByName('EbtInstAddr').AsString;
      L_ListItem.SubItems[17] := I_DataSet.FieldByName('EbtCustAddr').AsString;
      L_ListItem.SubItems[18] := I_DataSet.FieldByName('EbtBillAddr').AsString;


      L_ListItem.SubItems[19] := I_DataSet.FieldByName('EbtContractStatusDesc').AsString;
      L_ListItem.SubItems[20] := I_DataSet.FieldByName('EbtFeePeriodDesc').AsString;
      L_ListItem.SubItems[21] := I_DataSet.FieldByName('EbtServiceType').AsString;
      L_ListItem.SubItems[22] := I_DataSet.FieldByName('EbtAgentName').AsString;
      L_ListItem.SubItems[23] := I_DataSet.FieldByName('EbtAgentPhone').AsString;

      L_ListItem.SubItems[24] := I_DataSet.FieldByName('EbtAgentID').AsString;
      L_ListItem.SubItems[25] := I_DataSet.FieldByName('EbtAgentAddress').AsString;

      L_ListItem.SubItems[26] := I_DataSet.FieldByName('EbtIdCardId').AsString;
      L_ListItem.SubItems[27] := I_DataSet.FieldByName('EbtCompanyOwnerId').AsString;
      L_ListItem.SubItems[28] := I_DataSet.FieldByName('EbtNonProfitCompanyId').AsString;
      L_ListItem.SubItems[29] := I_DataSet.FieldByName('EbtCompanyId').AsString;
      sL_ProcessFlag := I_DataSet.FieldByName('ProcessFlag').AsString;

      sL_ErrorCode := I_DataSet.FieldByName('ErrorCode').AsString;
      sL_ErrorMsg := I_DataSet.FieldByName('ErrorMsg').AsString;

      L_ListItem.SubItems[30] := sL_ProcessFlag;
      L_ListItem.SubItems[31] := sL_ErrorCode;
      L_ListItem.SubItems[32] := sL_ErrorMsg;


      L_ListItem.SubItems[33] := I_DataSet.FieldByName('EbtNotes').AsString;
      L_ListItem.SubItems[34] := I_DataSet.FieldByName('EmcNotes').AsString;





      L_ListItem.SubItems[35] := I_DataSet.FieldByName('UpdTime').AsString;
      L_ListItem.SubItems[36] := I_DataSet.FieldByName('UpdEn').AsString;


      Inc(nG_So153UiSeqNo);

    end;




end;

procedure TfrmMain.prepareSo153ListItem(var I_ListItem: TListItem);
var
    nL_NextItemIndex : Integer;
begin

      if frmMain.nG_SO153LivItemPos=CA_UI_MAX_COUNT then
      begin
//        clearStatusItem; //呼叫此 function...似乎會讓 performance 下降...
        frmMain.nG_SO153LivItemPos := 0;
      end;
      I_ListItem := frmMain.livStatusSO153.Items[frmMain.nG_SO153LivItemPos];
      Inc(frmMain.nG_SO153LivItemPos);

      //down, 將下一個 ListItem 清除成空白
      if frmMain.nG_SO153LivItemPos=CA_UI_MAX_COUNT then
        nL_NextItemIndex := 0
      else
        nL_NextItemIndex := frmMain.nG_SO153LivItemPos;

      clearStatusItem(livStatusSO153, nL_NextItemIndex);


end;

function TfrmMain.getIfAcceptNullData: boolean;
begin
    result := bG_AcceptNullData;
end;

procedure TfrmMain.setProgStatus;
begin
    if chbActivate.Checked then
      pnlProgStatus.Caption := ACTIVATE_CAPTION
    else
      pnlProgStatus.Caption := DEACTIVATE_CAPTION;

    if pnlProgStatus.Font.Color = clRed then
      pnlProgStatus.Font.Color := clYellow
    else
      pnlProgStatus.Font.Color := clRed;
end;

procedure TfrmMain.timProgStatusTimer(Sender: TObject);
begin
    setProgStatus;
end;

procedure TfrmMain.showErrorOnMemo(sI_ErrorMsg: String);
begin
    memErrorLog.Lines.Insert(0, DateTimeToStr(now) + '=>' +  sI_ErrorMsg);
end;

end.
