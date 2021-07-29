//Syn OK
unit HandleUIU;

interface

uses
  Classes, ComCtrls, SysUtils;

type
  THandleUI = class(TThread)
  private
    { Private declarations }


  protected
    procedure Execute; override;
  public
    G_ListItem: TListItem;
    sG_ErrorCode, sG_ErrorMsg : String;
    procedure clearStatusItem; overload;
    procedure clearStatusItem(nI_ItemIndex:Integer); overload;    
    constructor Create;overload;
    procedure createListItem(sI_TransNo, sI_CommandID, sI_Operator, sI_Now:String);
    procedure updateListItemStatus(I_ListItem:TListItem; nI_Status:Integer; sI_TransNum, sI_ErrorCode, sI_ErrorMsg:String);

    procedure setListItemStatus1;
    procedure setListItemStatus2;
    procedure setListItemStatus3(sI_ErrorCode, sI_ErrorMsg:String);
    procedure setListItemStatus4;
  end;

implementation

uses frmMainU, ConstObjU;

{ Important: Methods and properties of objects in VCL or CLX can only be used
  in a method called using Synchronize, for example,

      Synchronize(UpdateCaption);

  and UpdateCaption could look like,

    procedure THandleUI.UpdateCaption;
    begin
      Form1.Caption := 'Updated in a thread';
    end; }

{ THandleUI }

procedure THandleUI.clearStatusItem;
var
    ii,jj : Integer;
begin
    for ii:=0 to frmMain.livStatus.Items.Count -1 do
    begin

      frmMain.livStatus.Items[ii].Caption := '';
      frmMain.livStatus.Items[ii].SubItemImages[2] := -1;
      frmMain.livStatus.Items[ii].SubItemImages[3] := -1;
      frmMain.livStatus.Items[ii].SubItemImages[4] := -1;
      for jj:=0 to frmMain.livStatus.Items[ii].SubItems.Count -1 do
        frmMain.livStatus.Items[ii].SubItems.Strings[jj] := '';

    end;
end;

procedure THandleUI.clearStatusItem(nI_ItemIndex: Integer);
var
    jj : Integer;
begin
      frmMain.livStatus.Items[nI_ItemIndex].Caption := '';
      frmMain.livStatus.Items[nI_ItemIndex].SubItemImages[2] := -1;
      frmMain.livStatus.Items[nI_ItemIndex].SubItemImages[3] := -1;
      frmMain.livStatus.Items[nI_ItemIndex].SubItemImages[4] := -1;
      for jj:=0 to frmMain.livStatus.Items[nI_ItemIndex].SubItems.Count -1 do
        frmMain.livStatus.Items[nI_ItemIndex].SubItems.Strings[jj] := '';

end;

constructor THandleUI.Create;
begin
    inherited Create(false);
end;

procedure THandleUI.createListItem(sI_TransNo, sI_CommandID,
  sI_Operator, sI_Now: String);
var
    jj, nL_NextItemIndex : Integer;
    L_ListItem : TListItem;
begin
    try
      if StrToIntDef(Copy(sI_TransNo,1,3),0)=(StrToInt(TESTING_CMD_COMP_CODE)) then Exit;
      if sI_Operator = BATCH_OPERATOR then Exit; // 批次作業 所產生的指令不顯示 UI
//      if StrToIntDef(Copy(sI_TransNo,4,6),0)=0 then Exit;
      if frmMain.nG_LivItemPos=CA_UI_MAX_COUNT then
      begin
//        clearStatusItem; //呼叫此 function...似乎會讓 performance 下降...
        frmMain.nG_LivItemPos := 0;
      end;
      L_ListItem := frmMain.livStatus.Items[frmMain.nG_LivItemPos];


{
// 因為 frmMain 中已經 create 好所有的 listview item 了, 所以此處不需要再 create...
      for jj:=0 to LISTVIEW_SUB_ITEM_COUNT-1 do
       L_ListItem.SubItems.Add('');
}
      Inc(frmMain.nG_LivItemPos);

      //down, 將下一個 ListItem 清除成空白
      if frmMain.nG_LivItemPos=CA_UI_MAX_COUNT then
        nL_NextItemIndex := 0
      else
        nL_NextItemIndex := frmMain.nG_LivItemPos;
      clearStatusItem(nL_NextItemIndex);
      //up, 將下一個 ListItem 清除成空白

//      L_ListItem.StateIndex := 0;
      L_ListItem.Caption := sI_TransNo;
      L_ListItem.SubItems[0] := sI_TransNo;
      L_ListItem.SubItems[1] := sI_CommandID;
      L_ListItem.SubItemImages[2] := 2;//打勾之icon
      L_ListItem.SubItemImages[3] := -1; //清除 icon
      L_ListItem.SubItemImages[4] := -1; //清除 icon
      L_ListItem.SubItems[5] := '30%';
      L_ListItem.SubItems[6] := sI_Operator;
      L_ListItem.SubItems[9] := sI_Now;




    except

    end;
end;

{
function THandleUI.createListItem(sI_TransNo, sI_CommandID,
  sI_Operator, sI_Now: String):TListItem;
var
    jj : Integer;
    L_ListItem : TListItem;
begin
    try
      if (frmMain.livStatus.Items<>nil) and (frmMain.livStatus.Items.Count>CA_UI_MAX_COUNT) then
        Synchronize(clearStatusItem);


      L_ListItem := frmMain.livStatus.Items.Insert(0);

      for jj:=1 to 10 do
        L_ListItem.SubItems.Add('');
      L_ListItem.StateIndex := 0;
      L_ListItem.Caption := sI_TransNo;
      L_ListItem.SubItems[0] := sI_TransNo;
      L_ListItem.SubItems[1] := sI_CommandID;
      L_ListItem.SubItemImages[2] := 2;//打勾之icon
      L_ListItem.SubItems[5] := '30%';
      L_ListItem.SubItems[6] := sI_Operator;
      L_ListItem.SubItems[9] := sI_Now;



      result := L_ListItem;
    except
      result := nil;
    end;
end;

}

procedure THandleUI.Execute;
begin
  { Place thread code here }
    self.FreeOnTerminate := true;
    self.OnTerminate := frmMain.HandleUIThreadDone;
end;

procedure THandleUI.setListItemStatus1;
begin
    if (Assigned(G_ListItem)) and (Assigned(G_ListItem.SubItems))then
    begin
      if G_ListItem.SubItems.Count>=4 then
        G_ListItem.SubItemImages[3] := 2;//打勾之icon

      if G_ListItem.SubItems.Count>=6 then
        G_ListItem.SubItems[5] := '60%';
    end;
end;

procedure THandleUI.setListItemStatus2;
begin
    if (Assigned(G_ListItem)) and (Assigned(G_ListItem.SubItems))then
    begin
      if G_ListItem.SubItems.Count>=5 then
        G_ListItem.SubItemImages[4] := 2;//打勾之icon
      if G_ListItem.SubItems.Count>=6 then
        G_ListItem.SubItems[5] := '100%';
    end;
end;

procedure THandleUI.setListItemStatus3(sI_ErrorCode, sI_ErrorMsg:String);
begin
    if (Assigned(G_ListItem)) and (Assigned(G_ListItem.SubItems))then
    begin
      if G_ListItem.SubItems.Count>=8 then
        G_ListItem.SubItems[7] := sI_ErrorCode;
      if G_ListItem.SubItems.Count>=9 then
        G_ListItem.SubItems[8] := sI_ErrorMsg;
      if G_ListItem.SubItems.Count>=5 then
        G_ListItem.SubItemImages[4] := 2;//打勾之icon
    end;
end;

procedure THandleUI.setListItemStatus4;
begin
    if (Assigned(G_ListItem)) and (Assigned(G_ListItem.SubItems))then
    begin
      if G_ListItem.SubItems.Count>=8 then
        G_ListItem.SubItems[7] := sG_ErrorCode;
      if G_ListItem.SubItems.Count>=9 then
        G_ListItem.SubItems[8] := sG_ErrorMsg;
      if G_ListItem.SubItems.Count>=4 then
        G_ListItem.SubItemImages[3] := 3; //錯誤之 icon
    end;  
end;

procedure THandleUI.updateListItemStatus(I_ListItem: TListItem;
  nI_Status: Integer; sI_TransNum, sI_ErrorCode, sI_ErrorMsg:String);
var
    nL_Pos,ii : Integer;
begin
    try
      if StrToIntDef(Copy(sI_TransNum,1,3),0)=(StrToInt(TESTING_CMD_COMP_CODE)) then Exit;

      case nI_Status of
        1://傳送指令ok
         begin
           if StrToIntDef(Copy(sI_TransNum,1,3),0)<>(StrToInt(TESTING_CMD_COMP_CODE)) then
           begin
             nL_Pos := 0;
             G_ListItem := frmMain.livStatus.FindCaption(nL_Pos,sI_TransNum,true,true,false);
             if G_ListItem<>nil then
             {
             if I_ListItem=nil then Exit;
             G_ListItem := I_ListItem;
             }

             if frmMain.G_HandleUIThread<>nil then
//               Synchronize(setListItemStatus1);
               setListItemStatus1;
           end;
         end;
        2://接收回應ok
         begin
           if StrToIntDef(Copy(sI_TransNum,1,3),0)<>(StrToInt(TESTING_CMD_COMP_CODE)) then
           begin
             nL_Pos := 0;
//             for ii:=0 to frmMain.nG_CmdResentTimes do
//             begin
               G_ListItem := frmMain.livStatus.FindCaption(nL_Pos,sI_TransNum,true,true,false);
//               nL_Pos := L_ListItem.Index;
               if G_ListItem<>nil then
               begin

                 if frmMain.G_HandleUIThread<>nil then
//                   Synchronize(setListItemStatus2);
                   setListItemStatus2;
               end;
//             end;
           end;
         end;
        3://回應1001
         begin
           if StrToIntDef(Copy(sI_TransNum,1,3),0)<>(StrToInt(TESTING_CMD_COMP_CODE)) then
//           if StrToIntDef(Copy(sI_TransNum,4,6),0)<>0 then
           begin
             nL_Pos := 0;
//             for ii:=0 to frmMain.nG_CmdResentTimes do
//             begin

               G_ListItem := frmMain.livStatus.FindCaption(nL_Pos,sI_TransNum,true,true,false);
//               nL_Pos := L_ListItem.Index;
               if G_ListItem<>nil then
               begin
                 {
                 frmMain.G_HandleUIThread.sG_ErrorCode := sI_ErrorCode;
                 frmMain.G_HandleUIThread.sG_ErrorMsg := sI_ErrorMsg;
                 }
                 if frmMain.G_HandleUIThread<>nil then
                 begin
{
                   sG_ErrorCode := sI_ErrorCode;
                   sG_ErrorMsg := sI_ErrorMsg;
}                   
//                   Synchronize(setListItemStatus3);
                   setListItemStatus3(sI_ErrorCode, sI_ErrorMsg);
                 end;
               end;
//             end;
           end;
         end;
        4://指令傳送錯誤
         begin
           if StrToIntDef(Copy(sI_TransNum,1,3),0)<>(StrToInt(TESTING_CMD_COMP_CODE)) then
//           if StrToIntDef(Copy(sI_TransNum,4,6),0)<>0 then
           begin
             nL_Pos := 0;
//             for ii:=0 to frmMain.nG_CmdResentTimes do
//             begin

               G_ListItem := frmMain.livStatus.FindCaption(nL_Pos,sI_TransNum,true,true,false);
//               nL_Pos := L_ListItem.Index;
               if G_ListItem<>nil then
               begin
                 {
                 frmMain.G_HandleUIThread.sG_ErrorCode := sI_ErrorCode;
                 frmMain.G_HandleUIThread.sG_ErrorMsg := sI_ErrorMsg;
                 }
                 if frmMain.G_HandleUIThread<>nil then
                 begin
                   sG_ErrorCode := sI_ErrorCode;
                   sG_ErrorMsg := sI_ErrorMsg;
//                   Synchronize(setListItemStatus4);
                   setListItemStatus4;
                 end;
               end;
//             end;
           end;
         end;
      end;
    except

    end;
end;

end.
