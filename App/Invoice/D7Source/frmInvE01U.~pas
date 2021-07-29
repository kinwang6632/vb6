unit frmInvE01U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, ExtCtrls, StdCtrls, Buttons, ImgList,
  xmldom, XMLIntf, msxmldom, XMLDoc, DB, ADODB, cxButtons, Menus,
  cxLookAndFeelPainters, cxGraphics, cxControls, cxContainer, cxEdit,
  cxTextEdit, cxMaskEdit, cxDropDownEdit, dxSkinsCore,
  dxSkinsDefaultPainters;

type
  TfrmInvE01 = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    StaticText1: TStaticText;
    chbGroupInfo: TcxComboBox;
    BitBtn1: TcxButton;
    BitBtn2: TcxButton;
    livNonCompetenceItem: TListView;
    ImageList1: TImageList;
    livHasCompetenceItem: TListView;
    StaticText2: TStaticText;
    StaticText3: TStaticText;
    btnExit: TcxButton;
    BitBtn4: TcxButton;
    BitBtn3: TcxButton;
    XMLDocument: TXMLDocument;
    Timer1: TTimer;
    procedure FormShow(Sender: TObject);
    procedure livHasCompetenceItemDragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure livHasCompetenceItemDragOver(Sender, Source: TObject; X, Y: Integer;
      State: TDragState; var Accept: Boolean);
    procedure BitBtn1Click(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure livNonCompetenceItemDragOver(Sender, Source: TObject; X,
      Y: Integer; State: TDragState; var Accept: Boolean);
    procedure livNonCompetenceItemDragDrop(Sender, Source: TObject; X,
      Y: Integer);
    procedure BitBtn4Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
  private
    { Private declarations }
    sG_GroupID : STring;
    procedure sortListView;
  public
    { Public declarations }
    procedure addListviewItem(sI_CompetenceID, sI_CompetenceDesc:String; I_TargetListView : TListView; nI_ImgIndex:Integer);
    procedure removeListviewItem(sI_CompetenceID:String; I_TargetListView : TListView);
  end;

var
  frmInvE01: TfrmInvE01;

implementation

uses dtmMainU, xmlU, Ustru, cbUtilis,
     frmMainU, dtmMainHU;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TfrmInvE01.FormCreate(Sender: TObject);
begin
  Timer1.Enabled := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvE01.FormShow(Sender: TObject);
begin
  Caption := frmMain.GetFormTitleString( 'E01', '權限設定' );
  Timer1.Enabled := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvE01.Timer1Timer(Sender: TObject);
var
  aParam: TComboBoxCreateParam;
begin
  Timer1.Enabled := False;
  Screen.Cursor := crSQLWait;
  try
    aParam.Sql :=
      ' SELECT GROUPID AS CODENO, DESCRIPTION FROM INV012 ORDER BY GROUPID ';
    aParam.KeyField := 'CODENO';
    aParam.DescField := 'DESCRIPTION';
    aParam.AddAllText := False;
    CreateCxComboBoxItem( chbGroupInfo, aParam );
    if ( chbGroupInfo.Properties.Items.Count > 0 ) then
    begin
      chbGroupInfo.ItemIndex := 0;
      BitBtn1.Click;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvE01.livHasCompetenceItemDragDrop(Sender, Source: TObject; X,
  Y: Integer);
var
  sL_Id, sL_Desc, sL_Caption:String;
begin
  if (Sender is TListView) and (Source is TListView) then
  begin
    sL_Caption := (Source as TListView).Selected.Caption;
    sL_Id := Copy(sL_Caption,1,3);
    sL_Desc := Copy(sL_Caption,4,length(sL_Caption)-3);
    addListviewItem(sL_Id,sL_Desc, livHasCompetenceItem,0);
    removeListviewItem(sL_Id, livNonCompetenceItem);
    sortListView;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvE01.livHasCompetenceItemDragOver(Sender, Source: TObject; X,
  Y: Integer; State: TDragState; var Accept: Boolean);
begin
  Accept := Source is TListView;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvE01.BitBtn1Click(Sender: TObject);
var
  aParam: TComboBoxItemParam;
begin
  if chbGroupInfo.ItemIndex < 0 then
  begin
    WarningMsg( '請先選擇群組!' );
    Exit;
  end;
  livNonCompetenceItem.Clear;
  livHasCompetenceItem.Clear;
  ZeroMemory( @aParam, SizeOf( aParam ) );
  //sG_GroupID := TUCommonFun.getComboCode(chbGroupInfo);
  GetCxComboBoxItemValue( chbGroupInfo, aParam );
  sG_GroupID := aParam.KeyValue;
  //dtmMainH.initCompetenceItem(sG_GroupID, self);
  dtmMainH.initCompetenceItem( aParam.KeyValue, Self );
  sortListView;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvE01.addListviewItem(sI_CompetenceID,
  sI_CompetenceDesc: String; I_TargetListView: TListView; nI_ImgIndex:Integer);
var
    L_ListItem : TListItem;
begin
    L_ListItem := I_TargetListView.Items.Add;
    L_ListItem.Caption := sI_CompetenceID + sI_CompetenceDesc;
    L_ListItem.ImageIndex := nI_ImgIndex;
end;

procedure TfrmInvE01.btnExitClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmInvE01.removeListviewItem(sI_CompetenceID: String;
  I_TargetListView: TListView);
var
    L_ListItem : TListItem;
begin
    L_ListItem := I_TargetListView.FindCaption(0,sI_CompetenceID,true,true,false);
    if L_ListItem<> nil then
      I_TargetListView.Items.Delete(L_ListItem.Index);

end;

procedure TfrmInvE01.livNonCompetenceItemDragOver(Sender, Source: TObject;
  X, Y: Integer; State: TDragState; var Accept: Boolean);
begin
    Accept := Source is TListView;
end;

procedure TfrmInvE01.livNonCompetenceItemDragDrop(Sender, Source: TObject;
  X, Y: Integer);
var
    sL_Id, sL_Desc, sL_Caption:String;
begin
  if (Sender is TListView) and (Source is TListView) then
  begin
    sL_Caption := (Source as TListView).Selected.Caption;
    sL_Id := Copy(sL_Caption,1,3);
    sL_Desc := Copy(sL_Caption,4,length(sL_Caption)-3);
    addListviewItem(sL_Id,sL_Desc, livNonCompetenceItem,1);
    removeListviewItem(sL_Id, livHasCompetenceItem);
    sortListView;
  end;
end;

procedure TfrmInvE01.sortListView;
begin
      livNonCompetenceItem.AlphaSort;
      livHasCompetenceItem.AlphaSort;
end;

procedure TfrmInvE01.BitBtn4Click(Sender: TObject);
var
    ii : Integer;
    sL_Caption,sL_Id, sL_Desc: String;
begin
   for ii:=0 to livNonCompetenceItem.Items.Count -1 do
   begin
     sL_Caption := livNonCompetenceItem.Items[ii].Caption;
     sL_Id := Copy(sL_Caption,1,3);
     sL_Desc := Copy(sL_Caption,4,length(sL_Caption)-3);
     addListviewItem(sL_Id,sL_Desc, livHasCompetenceItem,0);
   end;
   livNonCompetenceItem.Clear;
   sortListView;

end;

procedure TfrmInvE01.BitBtn3Click(Sender: TObject);
var
    ii : Integer;
    sL_Caption,sL_Id, sL_Desc: String;
begin
  for ii:=0 to livHasCompetenceItem.Items.Count -1 do
  begin
    sL_Caption := livHasCompetenceItem.Items[ii].Caption;
    sL_Id := Copy(sL_Caption,1,3);
    sL_Desc := Copy(sL_Caption,4,length(sL_Caption)-3);
    addListviewItem(sL_Id,sL_Desc, livNonCompetenceItem,1);
  end;
  livHasCompetenceItem.Clear;
  sortListView;
end;

procedure TfrmInvE01.BitBtn2Click(Sender: TObject);
var
  aIndex: Integer;
  aResult, aId : String;
  aRootNode, aChildNode : IXmlNode;
  aSQL: String;
begin
   if sG_GroupID='' then
   begin
     WarningMsg( '請先查詢出群組之權限!' );
     Exit;
   end;
   
   XMLDocument.Active := False;
   XMLDocument.Active := true;
   XMLDocument.ChildNodes.Clear;
   XMLDocument.Encoding := 'Big5';

   aRootNode := TMyXML.createRootNode( XMLDocument, 'GROUP', '', aResult );
   
   TMyXML.setAttribute( aRootNode, 'ID', sG_GroupID );

   for aIndex := 0 to livHasCompetenceItem.Items.Count -1 do
   begin
     aId := Copy( livHasCompetenceItem.Items[aIndex].Caption, 1, 3 );
     aChildNode := TMyXML.addChild( aRootNode, 'FUN', '' );
     TMyXML.setAttribute( aChildNode, 'ID', aId );
     TMyXML.setAttribute( aChildNode, 'Q', 'Y' );
     TMyXML.setAttribute( aChildNode, 'I', 'Y' );
     TMyXML.setAttribute( aChildNode, 'D', 'Y' );
     TMyXML.setAttribute( aChildNode, 'U', 'Y' );
     TMyXML.setAttribute( aChildNode, 'E', 'Y' );
   end;

   aSQL := Format( ' DELETE INV013 WHERE GROUPID = ''%s'' ', [sG_GroupID] );
   dtmMainH.runSQL( IUD_MODE, aSQL, dtmMainH.adoComm );

   aSQL :=
     ' INSERT INTO INV013                               ' +
     '           ( IDENTIFYID1, IDENTIFYID2, GROUPID,   ' +
     '             COMPETENCE, UPTTIME, UPTEN )         ' +
     '    VALUES ( ''%s'', ''%s'', ''%s'',              ' +
     '             ''%s'', SYSDATE, ''%s''    )         ';

    aSQL := Format( aSQL, [IDENTIFYID1, IDENTIFYID2, sG_GroupID,
       XMLDocument.XML.Text, dtmMain.getLoginUser] );

   dtmMainH.runSQL( IUD_MODE, aSQL, dtmMainH.adoComm );

   XMLDocument.Active := False;

end;

end.
