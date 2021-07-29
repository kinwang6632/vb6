unit frmBrowInvU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, cxStyles, dxSkinsCore, dxSkinBlack, dxSkinBlue, dxSkinCaramel,
  dxSkinCoffee, dxSkinsDefaultPainters, dxSkinscxPCPainter, cxCustomData,
  cxGraphics, cxFilter, cxData, cxDataStorage, cxEdit, DB, cxDBData,
  cxGridLevel, cxClasses, cxControls, cxGridCustomView,
  cxGridCustomTableView, cxGridTableView, cxGridDBTableView, cxGrid,
  ExtCtrls, DBClient, cxLookAndFeels, cxTextEdit, Menus,
  cxLookAndFeelPainters, StdCtrls, cxButtons;

type
  TfrmBrowINV = class(TForm)
    dsINV008: TDataSource;
    Panel1: TPanel;
    cdsBrowInv: TClientDataSet;
    cxStyleRepository1: TcxStyleRepository;
    cxStyle1: TcxStyle;
    Panel2: TPanel;
    cxGrid: TcxGrid;
    cxGridBrow: TcxGridDBTableView;
    cxColName: TcxGridDBColumn;
    cxColDesc: TcxGridDBColumn;
    cxGridLevel: TcxGridLevel;
    btnLog: TcxButton;
    procedure FormCreate(Sender: TObject);
    procedure cxGridBrowCustomDrawCell(Sender: TcxCustomGridTableView;
      ACanvas: TcxCanvas; AViewInfo: TcxGridTableDataCellViewInfo;
      var ADone: Boolean);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure btnLogClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
    procedure GetData();
  public
    { Public declarations }
  end;

var
  frmBrowINV: TfrmBrowINV;

implementation
  uses CommU,dtmCommU,frmBrowLogU;
{$R *.dfm}

procedure TfrmBrowINV.FormCreate(Sender: TObject);
 var aIndex : Integer;
   aFieldName,aFieldDesc : String;
begin
   // SetWindowPos(Handle,HWND_TOPMOST,Left,Top,Width,Height,SWP_NOACTIVATE or SWP_NOMOVE or SWP_NOSIZE);
    GetData;
    if (cdsBrowInv.FieldDefs.Count <= 0 )then
    begin
      cdsBrowInv.FieldDefs.Add( 'FldName', ftString, 30 );
      cdsBrowInv.FieldDefs.Add( 'FldDesc', ftString, 70);
      cdsBrowInv.CreateDataSet;
    end;
    cdsBrowInv.EmptyDataSet;
    {}
    if dtmComm.adoINV008.IsEmpty then
    begin
      MessageDlg('�䤣�����o����ơI',mtWarning,[mbOK],0);
      Application.Terminate;
      Exit;
    end;
    {}
    try
      with dtmComm do
      begin
        adoINV008.First;
        cdsBrowInv.ReadOnly := False;
        for aIndex := 0 to adoINV008.FieldCount -1 do
        begin
          aFieldName := EmptyStr;
          aFieldDesc := EmptyStr;
          aFieldName := GetIniField(FInvDefFile, adoINV008.Recordset.Fields.Item[aIndex].Name ,
                        0,EmptyStr);
          if aFieldName <> EmptyStr then
          begin
            cdsBrowInv.Append;
            cdsBrowInv.FieldByName( 'FldName' ).AsString := aFieldName;
            aFieldDesc := GetIniField( FInvDefFile,adoINV008.Recordset.Fields.Item[aIndex].Name,
                    1,VarToStr( adoINV008.Fields[aIndex].Value ));
            if aFieldDesc = EmptyStr then
            begin
              aFieldDesc := VarToStr( adoINV008.Fields[aIndex].Value );
            end;

            cdsBrowInv.FieldByName( 'FldDesc' ).AsString := aFieldDesc;


            {
            cdsBrowInv.FieldByName( 'FldName' ).AsString :=
                adoINV008.Recordset.Fields.Item[aIndex].Name;
            if adoINV008.Recordset.Fields.Item[aIndex].Name = '�o���榡' then
            begin
              case adoINV008.Fields[ aIndex ].AsInteger  of
                1:cdsBrowInv.FieldByName( 'FldDesc').AsString :='�q�l' ;
                2:cdsBrowInv.FieldByName( 'FldDesc').AsString :='��G' ;
                3:cdsBrowInv.FieldByName( 'FldDesc').AsString :='��T';
              else
                cdsBrowInv.FieldByName( 'FldDesc').AsString :='';
              end;
            end
            else if adoINV008.Recordset.Fields.Item[aIndex].Name = '�|�O' then
            begin
              case adoINV008.Fields[ aIndex ].AsInteger  of
                1:cdsBrowInv.FieldByName( 'FldDesc').AsString :='���|' ;
                2:cdsBrowInv.FieldByName( 'FldDesc').AsString :='�s�|�v' ;
                3:cdsBrowInv.FieldByName( 'FldDesc').AsString :='�K�|';
              else
                cdsBrowInv.FieldByName( 'FldDesc').AsString :='';
              end;
            end
            else if adoINV008.Recordset.Fields.Item[aIndex].Name ='�o���}�ߺ���' then
            begin
              case adoINV008.Fields[ aIndex ].AsInteger of
                1:cdsBrowInv.FieldByName( 'FldDesc').AsString :='�q�l�o��';
              else
                cdsBrowInv.FieldByName( 'FldDesc').AsString :='�q�l�p����o��';
              end;

            end else
            begin
              cdsBrowInv.FieldByName( 'FldDesc' ).AsString := VarToStr( adoINV008.Fields[aIndex].Value );
            end;
            }
            cdsBrowInv.Post;
          end;
        end;
        cdsBrowInv.First;
      end;
    except
      MessageDlg('��J�o����ƿ��~�I',mtError,[mbOK],0);
    end;
end;
{----------------------------------------------------------------------------------------------------------------------}
procedure TfrmBrowINV.GetData;
  var aSQL,aOwner,aSid,aFields : String;
begin
  try
    aFields := dtmComm.GetIniField(dtmComm.FInvDefFile ,EmptyStr,2,EmptyStr);
    aOwner := dtmComm.GetOwner;
    aSid := dtmComm.GetSID;

    aSQL := Format( 'SELECT INV007.INVID �o�����X ,INV007.INVDATE �o�����,              ' +
                    'INV007.INVFORMAT �o���榡,                                          ' +
                    'INV007.INVUSEDESC �o���γ~,INV007.CUSTID �Ƚs,                      ' +
                    'INV007.CUSTSNAME �Ȥ�m�W,INV007.INVTITLE �o�����Y,                 ' +
                    'INV007.BUSINESSID �Τ@�s��,INV007.INVADDR �o���a�},                 ' +
                    'INV007.MAILADDR �l�H�a�},INV007.SALEAMOUNT �P����B,                ' +
                    'INV007.TAXAMOUNT �|�B,INV007.TAXTYPE �|�O,                          ' +
                    'INV007.INVAMOUNT �o�����B,                                          ' +
                    'INV007.PRINTTIME �C�L�ɶ�,                                          ' +
                    'INV007.INVOICEKIND �o���}�ߺ���                                     ' +
                    'FROM INV007                                                         ' +
                    'Where  INV007.IdentifyId1 = ''%s''                                  ' +
                    'AND INV007.IdentifyId2 = %s AND INV007.INVID = ''%s''               ' ,
                    [ '1','0',ParamStr(3) ] );
    aSQL := Format( 'SELECT %s                                                           ' +
                    'FROM INV007                                                         ' +
                    'Where  INV007.IdentifyId1 = ''%s''                                  ' +
                    'AND INV007.IdentifyId2 = %s AND INV007.INVID = ''%s''               ' ,
                    [ aFields, '1','0',ParamStr(3) ] );
    dtmComm.adoINV008.Active := False;
    dtmComm.adoINV008.Close;
    dtmComm.adoINV008.SQL.Clear;
    dtmComm.adoINV008.SQL.Text := aSQL;
    dtmComm.adoINV008.Open;
  except
    MessageDlg('�}�ҵo����ƥ��ѡI',mtWarning,[mbOK],0);
    Application.Terminate;
  end;
end;
{----------------------------------------------------------------------------------------------------------------------}
procedure TfrmBrowINV.cxGridBrowCustomDrawCell(
  Sender: TcxCustomGridTableView; ACanvas: TcxCanvas;
  AViewInfo: TcxGridTableDataCellViewInfo; var ADone: Boolean);
  var
   aRect: TRect;
begin

  if AViewInfo.Item = cxColName then
  begin
    aRect := AViewInfo.ContentBounds;
    ACanvas.FillRect( aRect,clBtnFace );
  end;

end;
{----------------------------------------------------------------------------------------------------------------------}
procedure TfrmBrowINV.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = 27 then
    Application.Terminate;
end;

procedure TfrmBrowINV.btnLogClick(Sender: TObject);
var
  frm : TfrmLog;
begin
  try
    if dtmComm.GetLogData(ParamStr(3),dtmComm.GetCompID) then
    begin
      if dtmComm.adoLog.RecordCount = 0 then
      begin
        MessageDlg( '�L����ǰe�O���I',mtInformation,[mbOk],0);
        Exit;
      end;
      try
        frm := TfrmLog.Create( Application );
        frm.ShowModal;
      finally
        frm.Free;
      end;
    end;
  finally

  end;

end;

procedure TfrmBrowINV.FormShow(Sender: TObject);
begin
  if not dtmComm.CanUseSend then
    btnLog.Enabled := False;
end;

end.
