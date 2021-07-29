unit HandleUIU;

interface

uses
  Classes, ComCtrls, SysUtils;

  procedure ClearStatusItem; overload;
  procedure ClearStatusItem(nI_ItemIndex:Integer); overload;
  procedure CreateListItem(sI_TransNo, sI_CommandID, sI_Operator, sI_Now:String);
  procedure UpdateListItemStatus(I_ListItem:TListItem; nI_Status:Integer; sI_TransNum, sI_ErrorCode, sI_ErrorMsg:String);
  procedure SetListItemStatus1;
  procedure SetListItemStatus2;
  procedure SetListItemStatus3(sI_ErrorCode, sI_ErrorMsg:String);
  procedure SetListItemStatus4;

implementation

uses frmMainU, ConstObjU;


{ THandleUI }

var
  G_ListItem: TListItem;
  sG_ErrorCode, sG_ErrorMsg : String;


procedure ClearStatusItem;
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

procedure ClearStatusItem(nI_ItemIndex: Integer);
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

{ ---------------------------------------------------------------------------- }

procedure CreateListItem(sI_TransNo, sI_CommandID,
  sI_Operator, sI_Now: String);
var
    jj, nL_NextItemIndex : Integer;
    L_ListItem : TListItem;
begin
    try
      if StrToIntDef(Copy(sI_TransNo,1,3),0)=(StrToInt(TESTING_CMD_COMP_CODE)) then Exit;
      // �妸�@�~ �Ҳ��ͪ����O����� UI
      if sI_Operator = BATCH_OPERATOR then Exit;
      if frmMain.nG_LivItemPos=CA_UI_MAX_COUNT then
      begin
        //clearStatusItem; //�I�s�� function...���G�|�� performance �U��...
        frmMain.nG_LivItemPos := 0;
      end;
      L_ListItem := frmMain.livStatus.Items[frmMain.nG_LivItemPos];


      Inc( frmMain.nG_LivItemPos );

      //down, �N�U�@�� ListItem �M�����ť�
      if frmMain.nG_LivItemPos=CA_UI_MAX_COUNT then
        nL_NextItemIndex := 0
      else
        nL_NextItemIndex := frmMain.nG_LivItemPos;
      clearStatusItem(nL_NextItemIndex);
      //up, �N�U�@�� ListItem �M�����ť�


      L_ListItem.Caption := sI_TransNo;
      L_ListItem.SubItems[0] := sI_TransNo;
      L_ListItem.SubItems[1] := sI_CommandID;
      L_ListItem.SubItemImages[2] := 2;//���Ĥ�icon
      L_ListItem.SubItemImages[3] := -1; //�M�� icon
      L_ListItem.SubItemImages[4] := -1; //�M�� icon
      L_ListItem.SubItems[5] := '30%';
      L_ListItem.SubItems[6] := sI_Operator;
      L_ListItem.SubItems[9] := sI_Now;
    except

    end;
end;

{ ---------------------------------------------------------------------------- }

procedure SetListItemStatus1;
begin
  if (Assigned(G_ListItem)) and (Assigned(G_ListItem.SubItems))then
  begin
    if G_ListItem.SubItems.Count>=4 then
      G_ListItem.SubItemImages[3] := 2;//���Ĥ�icon

    if G_ListItem.SubItems.Count>=6 then
      G_ListItem.SubItems[5] := '60%';
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure SetListItemStatus2;
begin
    if (Assigned(G_ListItem)) and (Assigned(G_ListItem.SubItems))then
    begin
      if G_ListItem.SubItems.Count>=5 then
        G_ListItem.SubItemImages[4] := 2;  //���Ĥ�icon
      if G_ListItem.SubItems.Count>=6 then
        G_ListItem.SubItems[5] := '100%';
    end;
end;

{ ---------------------------------------------------------------------------- }

procedure SetListItemStatus3(sI_ErrorCode, sI_ErrorMsg:String);
begin
  if (Assigned(G_ListItem)) and (Assigned(G_ListItem.SubItems))then
  begin
    if G_ListItem.SubItems.Count>=8 then
      G_ListItem.SubItems[7] := sI_ErrorCode;
    if G_ListItem.SubItems.Count>=9 then
      G_ListItem.SubItems[8] := sI_ErrorMsg;
    if G_ListItem.SubItems.Count>=5 then
      G_ListItem.SubItemImages[4] := 2;//���Ĥ�icon
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure SetListItemStatus4;
begin
  if (Assigned(G_ListItem)) and (Assigned(G_ListItem.SubItems))then
  begin
    if G_ListItem.SubItems.Count>=8 then
      G_ListItem.SubItems[7] := sG_ErrorCode;
    if G_ListItem.SubItems.Count>=9 then
      G_ListItem.SubItems[8] := sG_ErrorMsg;
    if G_ListItem.SubItems.Count>=4 then
      G_ListItem.SubItemImages[3] := 3; //���~�� icon
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure UpdateListItemStatus(I_ListItem: TListItem;
  nI_Status: Integer; sI_TransNum, sI_ErrorCode, sI_ErrorMsg:String);
var
    nL_Pos,ii : Integer;
begin
G_ListItem := nil;
    try
      if StrToIntDef(Copy(sI_TransNum,1,3),0)=(StrToInt(TESTING_CMD_COMP_CODE)) then Exit;

      case nI_Status of
        1://�ǰe���Ook
         begin
           if StrToIntDef(Copy(sI_TransNum,1,3),0)<>(StrToInt(TESTING_CMD_COMP_CODE)) then
           begin
             G_ListItem := frmMain.livStatus.FindCaption(0,sI_TransNum,true,true,false);
             if G_ListItem <> nil then
               HandleUIU.setListItemStatus1;
           end;
         end;
        2://�����^��ok
         begin
           if StrToIntDef(Copy(sI_TransNum,1,3),0)<>(StrToInt(TESTING_CMD_COMP_CODE)) then
           begin
             G_ListItem := frmMain.livStatus.FindCaption(0,sI_TransNum,true,true,false);
             if G_ListItem<>nil then
             begin
               HandleUIU.SetListItemStatus2;
             end;
           end;
         end;
        3://�^��1001
         begin
           if StrToIntDef(Copy(sI_TransNum,1,3),0)<>(StrToInt(TESTING_CMD_COMP_CODE)) then
           begin
             G_ListItem := frmMain.livStatus.FindCaption( 0,sI_TransNum,true,true,false);
             if G_ListItem <> nil then
               HandleUIU.SetListItemStatus3(sI_ErrorCode, sI_ErrorMsg);
           end;
         end;
        4://���O�ǰe���~
         begin
           if StrToIntDef(Copy(sI_TransNum,1,3),0)<>(StrToInt(TESTING_CMD_COMP_CODE)) then
           begin
             G_ListItem := frmMain.livStatus.FindCaption( 0,sI_TransNum,true,true,false);
             if G_ListItem<>nil then
             begin
                sG_ErrorCode := sI_ErrorCode;
                sG_ErrorMsg := sI_ErrorMsg;
                HandleUIU.setListItemStatus4;
             end;
           end;
         end;
       end;
    except
      { ... }
    end;
    if Assigned( G_ListItem ) then G_ListItem.MakeVisible( True );
end;

end.
