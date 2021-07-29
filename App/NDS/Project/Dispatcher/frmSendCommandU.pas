unit frmSendCommandU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Sockets, DB, ExtCtrls, StdCtrls, Buttons, Grids, DBGrids ,ADODB ,TListernerU,
  ComCtrls, ImgList;

type
  TfrmSendCommand = class(TForm)
    Panel2: TPanel;
    Timer1: TTimer;
    DataSource1: TDataSource;
    TcpClient1: TTcpClient;
    TcpServer1: TTcpServer;
    chbAction: TCheckBox;
    imlStatus: TImageList;
    Panel1: TPanel;
    memErrorMsg: TMemo;
    Splitter1: TSplitter;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    livStatus: TListView;
    TabSheet2: TTabSheet;
    Memo1: TMemo;
    btnData: TBitBtn;
    BitBtn1: TBitBtn;
    btnExit: TBitBtn;

    procedure Timer1Timer(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure chbActionClick(Sender: TObject);
    procedure btnDataClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
  private
    { Private declarations }
    bG_ShowUI : Boolean;
    nG_LivItemPos : Integer;
    procedure connectToProcessor(sI_ProcessorIP:String);
    function getCmdSequence(sI_CompCode, sI_SeqNo : String):String;
    procedure createLivItem;
    Procedure StartProcess;
    Procedure getParam;
    Procedure sendDataToProcessor;
    procedure clearStatusItem(nI_ItemIndex:Integer); overload;
  public
    { Public declarations }
    G_ListenerThread : TListerner;
    G_ListenerThread_ExitCode : Cardinal;
    nG_CmdRefreshRate,nG_MaxCommandCount,nG_ReSentTimes : Integer;
    bG_SaveCommandLog,bG_SaveResponseLog : Boolean;
    nG_TestingCmdSeqNo,nG_MaxTestingCmdSeqNo : Integer;
    Procedure sendNdsString(I_AdoQrySendNagraData : TADOQuery;sI_CompName,sI_CompCode,sI_ProcessorIP : String);
    function sendDataToProcessor1(sI_CompCode,sI_CompName,sI_SendString,sI_ProcessorIP : String) : String;
    function sendDataToProcessor2 : String;
    function sendDataToProcessor3 : String;
    Procedure createThread;
    procedure updateReturnDataUI(sI_CompCode, sI_SeqNo, sI_ErrorCode, sI_ErrorMsg: String);
  end;

var
  frmSendCommand: TfrmSendCommand;

implementation

uses dtmMainU, ConstU, Ustru, frmTestDataU, frmMainU;

{$R *.dfm}

procedure TfrmSendCommand.getParam;
begin

    //多少秒讀資料庫一次
    nG_CmdRefreshRate := dtmMain.getCmdRefreshRate;

    //可同時處理幾個指令
    nG_MaxCommandCount := dtmMain.cdsParam.FieldByName('nMaxCommandCount').AsInteger;

    //指令重傳次數
    nG_ReSentTimes := dtmMain.cdsParam.FieldByName('nCmdReSentTimes').AsInteger;

    //指令傳送是否要存Log
    bG_SaveCommandLog := dtmMain.cdsParam.FieldByName('bCommandLog').AsBoolean;

    //指令回應是否要存Log
    bG_SaveResponseLog := dtmMain.cdsParam.FieldByName('bResponseLog').AsBoolean;

    //是否顯示傳送UI [ 會使得指令傳送速度變慢 ]
    bG_ShowUI := dtmMain.cdsParam.FieldByName('bShowUI').Asboolean;

end;

procedure TfrmSendCommand.sendDataToProcessor;
var
    ii,nL_Counts : Integer;
    sL_CompCode,sL_ErrMsg,sL_CompName : String;
    L_ADOQuery : TADOQuery;
    sL_ProcessorIP : String;
begin
    nL_Counts := 0;

    //依各公司將 SEND_NDS 資料取出,傳至各平台
    for ii:=0 to dtmMain.G_CompInfoStrList.Count-1 do
    begin
      sL_CompCode := dtmMain.G_CompInfoStrList.Strings[ii];
      sL_CompName := dtmMain.getCompName(sL_CompCode);

      sL_ProcessorIP := dtmMain.getProcessorIP(sL_CompCode);
      if dtmMain.isDisconnect(sL_ProcessorIP) then
      begin
        memErrorMsg.Lines.Add('負責處理 ' + sL_CompName + ' 的 CA Processor 已經關閉了.所以,無法處理資料!' + ' IP:' + sL_ProcessorIP + ' -- ' + DateTimeToStr(now));
        connectToProcessor(sL_ProcessorIP);
        if not TcpClient1.Connected then
          memErrorMsg.Lines.Add('嚐試重新連接到處理' + sL_CompName + ' 的 CA Processor.但仍無法順利連接!' + ' IP:' + sL_ProcessorIP + ' -- '  + DateTimeToStr(now))
        else
        begin
          dtmMain.hasConnected(sL_ProcessorIP);
          memErrorMsg.Lines.Add('已經順利連接上處理' + sL_CompName + ' 的 CA Processor.' + ' IP:' + sL_ProcessorIP + ' -- ' + DateTimeToStr(now));
        end;
        continue;
      end;

      L_ADOQuery := dtmMain.getSendNdsData(sL_CompCode,sL_ErrMsg);

      //複製至實體AdoQuery
      dtmMain.adoQrySendNdsData.Clone(L_ADOQuery,ltUnspecified);

      if sL_ErrMsg = '' then
      begin
        //SHOWMESSAGE(sL_CompName + '   ' + IntToStr(dtmMain.adoQrySendNagraData.RecordCount));

{//暫時先不用JACKAL920428
        //若第一家公司沒資料,傳假資料至 Nds
        if (dtmMain.adoQrySendNdsData.RecordCount = 0) and (ii=0) then
        begin
          nG_TestingCmdSeqNo := nG_TestingCmdSeqNo + 1;

          dtmMain.adoQrySendNdsData.Append;
          dtmMain.adoQrySendNdsData.FieldByName('SEQNO').AsInteger := nG_TestingCmdSeqNo;
          dtmMain.adoQrySendNdsData.FieldByName('CMDSTATUS').AsString := 'W';
          dtmMain.adoQrySendNdsData.FieldByName('COMMANDID').AsString := CA_TESTING_CMD_ID;
          dtmMain.adoQrySendNdsData.FieldByName('COMPCODE').AsString := TESTING_CMD_COMP_CODE;
          dtmMain.adoQrySendNdsData.Post;

          if nG_TestingCmdSeqNo = nG_MaxTestingCmdSeqNo then
            nG_TestingCmdSeqNo := 0;

        end;
}

        dtmMain.adoQrySendNdsData.First;
        while not dtmMain.adoQrySendNdsData.Eof do
        begin

          //串出要送Nagra的

          sendNdsString(dtmMain.adoQrySendNdsData,sL_CompName,sL_CompCode,sL_ProcessorIP);

          dtmMain.adoQrySendNdsData.Next;

          //每各公司只處理固定筆數
          nL_Counts := nL_Counts + 1;



          if nL_Counts > nG_MaxCommandCount then
          begin
            nL_Counts := 0;
            dtmMain.adoQrySendNdsData.Close;
          end;
        end;
        dtmMain.adoQrySendNdsData.Close;
      end
      else
      begin
        MessageDlg(sL_ErrMsg,mtError, [mbOK],0);
        exit;
      end;

    end;


end;

procedure TfrmSendCommand.StartProcess;
var
    nL_CmdRefreshRate : Integer;
begin

    if chbAction.Checked then
    begin
      Timer1.Enabled := false;

      nL_CmdRefreshRate := nG_CmdRefreshRate;

      Timer1.Interval := nL_CmdRefreshRate * 1000;

      //將資料傳至 Processor
      sendDataToProcessor;

      Timer1.Enabled := true;
    end
    else
    begin
      nL_CmdRefreshRate := 0;

      Timer1.Interval := nL_CmdRefreshRate * 1000;
    end;

end;

procedure TfrmSendCommand.Timer1Timer(Sender: TObject);
begin
    StartProcess;
end;

procedure TfrmSendCommand.FormCreate(Sender: TObject);
var
    sL_Msg : String;
begin
    dtmMain.getDbConnInfo;
    sL_Msg := dtmMain.connectToDB;
    if sL_Msg = '' then
    begin
      // 取得相關參數
      getParam;

      //開啟接收 Processor Tcp Server
      //openTcpServer;
      createThread;

      //傳送各ProcessorIP的CompCode字串給所屬Processor
      sendDataToProcessor2;

      //若無資料傳假資料 , SeqNo 最大不行超過此nG_MaxTestingCmdSeqNo
      nG_MaxTestingCmdSeqNo := StrToIntDef(TUstr.AddString('','9',true,SEQ_NUM_LENGTH),999999);
    end
    else
    begin
      showmessage(sL_Msg);
      Close;
    end;

end;

procedure TfrmSendCommand.sendNdsString(I_AdoQrySendNagraData: TADOQuery;
  sI_CompName, sI_CompCode, sI_ProcessorIP: String);
var
    sL_SeqNo,sL_CompCode,sL_CompName,sL_CommandID,sL_SubscriberID : String;
    sL_CardID,sL_SubBeginDate,sL_SubEndDate,sL_PinCode,sL_PopulationID : String;
    sL_RegionKey,sL_ReportbackAvailability,sL_ReportBackDate : String;
    sL_Operator,sL_Notes,sL_FullNotes,sL_PinControl : String;

    sL_CmdSequence : String;
    L_NewListITem : TListItem;
    
    sL_Msg,sL_TotalString : String;
    nL_StrLength : Integer;
    fL_FunReturn,fL_RetCode : Double;

begin
{
    '00000002  7    CSV3.0 12312345678    ICCNO0012003042820031231 78966  Region188    JACKAL   9Notes:NDSJ';
}
    with I_AdoQrySendNagraData do
    begin
      sL_SeqNo := dtmMain.fillStringLength(FieldByName('SEQNO').AsString,ZERO_STR,SEQ_NO_LENGTH);

      sL_CompCode := dtmMain.fillStringLength(FieldByName('COMPCODE').AsString,SPACE_STR,COMP_CODE_LENGTH);

      sL_CompName := dtmMain.fillStringLength(FieldByName('COMPNAME').AsString,SPACE_STR,COMP_NAME_LENGTH);

      sL_CommandID := dtmMain.fillStringLength(FieldByName('COMMANDID').AsString,SPACE_STR,HIGH_LEVEL_CMD_ID_LENGTH);

      sL_SubscriberID := dtmMain.fillStringLength(FieldByName('SUBSCRIBERID').AsString,SPACE_STR,SUBSCRIBER_ID_LENGTH);
      
      sL_CardID := dtmMain.fillStringLength(FieldByName('ICCNO').AsString,SPACE_STR,ICC_NO_LENGTH);

      sL_SubBeginDate := Copy(TUstr.replaceStr(FieldByName('SUBBEGINDATE').AsString,'/',''),1,8);
      sL_SubBeginDate := dtmMain.fillStringLength(sL_SubBeginDate,SPACE_STR,SUB_BEGIN_DATE_LENGTH);

      sL_SubEndDate := Copy(TUstr.replaceStr(FieldByName('SUBENDDATE').AsString,'/',''),1,8);
      sL_SubEndDate := dtmMain.fillStringLength(sL_SubEndDate,SPACE_STR,SUB_END_DATE_LENGTH);

      sL_PinCode := dtmMain.fillStringLength(FieldByName('PINCODE').AsString,SPACE_STR,PIN_CODE_LENGTH);

      sL_PopulationID := dtmMain.fillStringLength(FieldByName('POPULATIONID').AsString,SPACE_STR,INT_POPULATION_ID_LENGTH);

      sL_RegionKey := dtmMain.fillStringLength(FieldByName('REGIONKEY').AsString,SPACE_STR,REGION_KEY_LENGTH);

      sL_ReportbackAvailability := dtmMain.fillStringLength(FieldByName('REPORTBACKAVAILABILITY').AsString,SPACE_STR,REPORTBACK_AVAILABILITY_LENGTH);

      sL_ReportBackDate := FieldByName('REPORTBACKDATE').AsString;
      if sL_ReportBackDate = '' then
        sL_ReportBackDate := '0';
      sL_ReportBackDate := dtmMain.fillStringLength(sL_ReportBackDate,ZERO_STR,REPORTBACK_DATE_LENGTH);

      sL_Operator := dtmMain.fillStringLength(FieldByName('OPERATOR').AsString,SPACE_STR,OPERATOR_LENGTH);

      sL_Notes := FieldByName('NOTES').AsString;
      nL_StrLength := Length(sL_Notes);
      sL_FullNotes := dtmMain.fillStringLength(IntToStr(nL_StrLength),SPACE_STR,NOTES_LENGTH)
                      + sL_Notes;


      sL_PinControl := dtmMain.fillStringLength(FieldByName('PINCONTROL').AsString,SPACE_STR,PIN_CONTROL_LENGTH);

      sL_TotalString := IntToStr(SEND_TO_DISPATCHER_INFO_TYPE_1) + '~' + sL_SeqNo + sL_CompCode + sL_CompName +
                        sL_CommandID + sL_SubscriberID + sL_CardID +
                        sL_SubBeginDate + sL_SubEndDate + sL_PinCode +
                        sL_PopulationID + sL_RegionKey + sL_ReportbackAvailability +
                        sL_ReportBackDate + sL_Operator + sL_FullNotes +
                        sL_PinControl;

//sL_Data := '00123456*  9*    陽明山*  A0*    2DC3*  1234567890*        *20030430*8899*23*98760054*E*28*Howard    *  25*4~20040531~N;5~20041231~A* ';
//sL_Data := '00000002*  7*    CSV3.0* 123*12345678*    ICCNO001*20030428*20031231* 789*66*  Region*1*88*    JACKAL*  25*4~20040531~N;5~20041231~A*J';
//           '1~00000001  7     A公司  A0    2DC3  1234567890200304292003042988992398760054E28         A  254~20040531~N;5~20041231~AL'

      fL_FunReturn := 10.5;

      //呼叫  將CMD_STATUS  W 改為 P
      dtmMain.runSF1(sI_CompCode,'P','','',Trim(sL_SEQNO),fL_FunReturn,fL_RetCode,sL_Msg);

      fL_FunReturn := fL_FunReturn * 2;

      //down, 處理 UI 之顯示...
      if bG_ShowUI then
      begin
        sL_CmdSequence := getCmdSequence(sL_CompCode, sL_SeqNo);

        L_NewListItem := livStatus.Items[nG_LivItemPos];
        L_NewListItem.Caption := sL_CmdSequence;
        L_NewListItem.SubItems[0] := sL_CompName;
        L_NewListItem.SubItems[1] := sL_CommandID;
        L_NewListItem.SubItems[6] := frmMain.sG_OperatorName;
        L_NewListItem.SubItems[7] := DateTimeToStr(now);
        L_NewListItem.SubItemImages[2] := 1;
        Inc(nG_LivItemPos);
        if nG_LivItemPos=CA_UI_MAX_COUNT then
          nG_LivItemPos := 0;
        //clearStatusItem(nG_LivItemPos);
      end;
      //up, 處理 UI 之顯示...

      //將資料傳至Processor
      sendDataToProcessor1(sI_CompCode,sL_CompName,sL_TotalString,sI_ProcessorIP);

      Application.ProcessMessages;

{
      dtmMain.cdsTemp.Append;
      dtmMain.cdsTemp.FieldByName('COL1').AsString := sL_SEQNO;
      dtmMain.cdsTemp.FieldByName('COL4').AsString := sI_ProcessorIP;
      dtmMain.cdsTemp.FieldByName('COL3').AsString := sI_CompName;

      if sL_Msg = '' then
        sL_Msg := 'ReturnNO:  ' + FloatToStr(fL_FunReturn) + ' sL_SEQNO:  ' + sL_SEQNO
      else
        sL_Msg := 'ReturnNO:  ' + FloatToStr(fL_FunReturn) + '  sL_Msg: ' + sL_Msg;

      dtmMain.cdsTemp.FieldByName('COL2').AsString := sL_Msg;
      dtmMain.cdsTemp.Post;
}

    end;

end;

function TfrmSendCommand.sendDataToProcessor1(sI_CompCode,sI_CompName, sI_SendString,
  sI_ProcessorIP: String): String;
begin

    try
      try
        connectToProcessor(sI_ProcessorIP);
        if TcpClient1.Connected then
          sleep(500);
          TcpClient1.Sendln(sI_SendString);
      except
        memErrorMsg.Lines.Add('[1]負責處理 ' + sI_CompName + ' 的 CA Processor 已經關閉了.所以無法傳送資料!' + ' IP:' + sI_ProcessorIP + ' -- [' + DateTimeToStr(now) + ']');
      end;
    finally
      if not TcpClient1.Connected then
        memErrorMsg.Lines.Add('[1]負責處理 ' + sI_CompName + ' 的 CA Processor 已經關閉了.所以無法傳送資料!' + ' IP:' + sI_ProcessorIP + ' -- [' + DateTimeToStr(now) + ']')
      else
        TcpClient1.Disconnect;
    end;


end;

procedure TfrmSendCommand.BitBtn1Click(Sender: TObject);
var
    sL_Temp : String;
    nL_Len : Integer;
begin
    //sL_Temp := '1C01903123000000009000001';
    //sL_Temp := '20190002A1006howard00900000301000302574428920030408N2003040720030409U1160194446005211911825900000';
    //sL_Temp := '30190090000041000000008972050257000344289200303081000009000004000000000000000000000009';
    //dtmMain.separateCaReturnString(sL_Temp);
    sL_Temp := '00123456 9 ??? A0 2DC3 1234567890 2003043088992398760054E28Howard 254~20040531~N;5~20041231~A ';
    nL_Len := Length(sL_Temp);
    showmessage(IntToStr(nL_Len));
    sL_Temp := '00000002  7    CSV3.0 12312345678    ICCNO0012003042820031231 78966  Region188    JACKAL   9Notes:NDSJ';
    nL_Len := Length(sL_Temp);
    showmessage(IntToStr(nL_Len));    

end;

procedure TfrmSendCommand.createThread;
begin
    G_ListenerThread := TListerner.Create(dtmMain.sG_DispatcherIP, dtmMain.sG_DispatcherListenPort);
    if GetExitCodeThread(G_ListenerThread.ThreadID,G_ListenerThread_ExitCode) then
    begin
      MessageDlg('執行緒 1 產生失敗!請重新執行程式' , mtError, [mbOK],0);

      Application.Terminate;
      Exit;
    end;
end;

function TfrmSendCommand.sendDataToProcessor2: String;
var
    ii : Integer;
    sL_CompCodeString,sL_ProcessorIP : String;
begin
    //串起所有的公司別字串,傳至各平台
    with dtmMain.G_ProcessorIPStrList do
    begin
      for ii:=0 to Count-1 do
      begin
        sL_ProcessorIP := (Objects[ii] as TProcessorIPData).sProcessorIp;
        sL_CompCodeString := (Objects[ii] as TProcessorIPData).sCompCodeString;

        TcpClient1.RemoteHost := sL_ProcessorIP;
        TcpClient1.RemotePort := dtmMain.sG_ProcessorListenPort;

        try
          try
            if TcpClient1.Connect then
              TcpClient1.Sendln(IntToStr(SEND_TO_DISPATCHER_INFO_TYPE_2) + '-' + sL_CompCodeString);
          except
            memErrorMsg.Lines.Add('[2]負責處理IP: ' + sL_ProcessorIP + ' 的 CA Processor 已經關閉了.所以無法傳送資料!' + ' -- [' + DateTimeToStr(now) + ']');
          end;
        finally
          if not TcpClient1.Connected then
            memErrorMsg.Lines.Add('[2]負責處理IP: ' + sL_ProcessorIP + ' 的 CA Processor 已經關閉了.所以無法傳送資料!' + ' -- [' + DateTimeToStr(now) + ']')
          else
            TcpClient1.Disconnect;
        end;
      end;
    end;

end;

procedure TfrmSendCommand.chbActionClick(Sender: TObject);
begin
    if chbAction.Checked then
      Timer1.Enabled := true
    else
      Timer1.Enabled := false;
end;

procedure TfrmSendCommand.btnDataClick(Sender: TObject);
begin
    frmTestData := TfrmTestData.Create(Application);
    frmTestData.ShowModal;
    frmTestData.Free;
end;

procedure TfrmSendCommand.createLivItem;
var
    ii, jj : Integer;
    L_ListItem : TListItem;

begin
    if bG_ShowUI then
    begin
      nG_LivItemPos := 0;
      memErrorMsg.Align := alBottom;
      memErrorMsg.Height := 120;


      Splitter1.Visible := true;
      Splitter1.Parent := Panel1;
      Splitter1.Height := 4;
      Splitter1.Align := alBottom;

      PageControl1.Visible := true;
      PageControl1.Align := alClient;

      TabSheet1.Caption := '指令處理狀態';


      frmSendCommand.livStatus.Items.Clear;
      for ii:=0 to CA_UI_MAX_COUNT -1 do
      begin
        L_ListItem := frmSendCommand.livStatus.Items.Add;

        for jj:=0 to LISTVIEW_SUB_ITEM_COUNT-1 do
         L_ListItem.SubItems.Add('');

      end;
    end
    else
    begin
      PageControl1.Visible := false;
      TabSheet1.Caption := '處理狀態';
      Splitter1.Visible := false;
      memErrorMsg.Align := alClient;
    end;

end;

procedure TfrmSendCommand.FormShow(Sender: TObject);
begin
    createLivItem;
    frmSendCommand.Caption := 'NDS GateWay --[分派指令]';

    TabSheet2.TabVisible := false;

    if Uppercase(ParamStr(1)) = 'TESTING' then
      TabSheet2.TabVisible := true;

//    dtmMain.cdsTemp.EmptyDataSet;
end;

procedure TfrmSendCommand.clearStatusItem(nI_ItemIndex: Integer);
var
    ii,jj : Integer;
begin
      livStatus.Items[nI_ItemIndex].Caption := '';
      livStatus.Items[nI_ItemIndex].SubItemImages[2] := -1;
      livStatus.Items[nI_ItemIndex].SubItemImages[3] := -1;
      livStatus.Items[nI_ItemIndex].SubItemImages[4] := -1;
      for jj:=0 to livStatus.Items[nI_ItemIndex].SubItems.Count -1 do
        livStatus.Items[nI_ItemIndex].SubItems.Strings[jj] := '';
{
    for ii:=0 to livStatus.Items.Count -1 do
    begin

      livStatus.Items[ii].Caption := '';
      livStatus.Items[ii].SubItemImages[2] := -1;
      livStatus.Items[ii].SubItemImages[3] := -1;
      livStatus.Items[ii].SubItemImages[4] := -1;
      for jj:=0 to livStatus.Items[ii].SubItems.Count -1 do
        livStatus.Items[ii].SubItems.Strings[jj] := '';

    end;
}
end;

procedure TfrmSendCommand.updateReturnDataUI(sI_CompCode, sI_SeqNo,
  sI_ErrorCode, sI_ErrorMsg: String);
var
    L_ListItem : TListItem;
    sL_CmdSequence : String;
begin
    sL_CmdSequence := getCmdSequence(sI_CompCode, sI_SeqNo);
    L_ListItem := livStatus.FindCaption(0,sL_CmdSequence, true, true, false );
    if L_ListItem<>nil then
    begin
      L_ListItem.SubItems[4] := sI_ErrorCode;
      L_ListItem.SubItems[5] := sI_ErrorMsg;
      L_ListItem.SubItemImages[3] := 1;
    end;
end;

function TfrmSendCommand.getCmdSequence(sI_CompCode,
  sI_SeqNo: String): String;
begin
    result := Format('%.2d',[StrToInt(sI_CompCode)]) + Format('%.8d',[StrToInt(sI_SeqNo)]);
end;

procedure TfrmSendCommand.connectToProcessor(sI_ProcessorIP: String);
begin
    try
      TcpClient1.Disconnect;
      TcpClient1.RemoteHost := sI_ProcessorIP;
      TcpClient1.RemotePort := dtmMain.sG_ProcessorListenPort;
      TcpClient1.Connect;
    except

    end;  
end;

function TfrmSendCommand.sendDataToProcessor3: String;
var
    ii : Integer;
    sL_CompCodeString,sL_ProcessorIP : String;
begin
    //告訴 Processor ..此 Dispatcher 要斷線了..
    with dtmMain.G_ProcessorIPStrList do
    begin
      for ii:=0 to Count-1 do
      begin
        sL_ProcessorIP := (Objects[ii] as TProcessorIPData).sProcessorIp;


        TcpClient1.RemoteHost := sL_ProcessorIP;
        TcpClient1.RemotePort := dtmMain.sG_ProcessorListenPort;

        try
          try
            if TcpClient1.Connect then
              TcpClient1.Sendln(IntToStr(SEND_TO_DISPATCHER_INFO_TYPE_3));
          except
            memErrorMsg.Lines.Add('[3]負責處理IP: ' + sL_ProcessorIP + ' 的 CA Processor 已經關閉了.所以無法傳送資料!' + ' -- [' + DateTimeToStr(now) + ']')
          end;
        finally
          if not TcpClient1.Connected then
            memErrorMsg.Lines.Add('[3]負責處理IP: ' + sL_ProcessorIP + ' 的 CA Processor 已經關閉了.所以無法傳送資料!' + ' -- [' + DateTimeToStr(now) + ']')
          else
            TcpClient1.Disconnect;
        end;
      end;
    end;

end;

procedure TfrmSendCommand.btnCloseClick(Sender: TObject);
begin
    if MessageDlg('是否確定要結束指令分派?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      sendDataToProcessor3;
      Close;
    end;
end;

procedure TfrmSendCommand.btnExitClick(Sender: TObject);
begin
    if MessageDlg('是否確定要結束指令分派?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      sendDataToProcessor3;

      Sleep(1000);
      Close;
    end;
end;

end.
