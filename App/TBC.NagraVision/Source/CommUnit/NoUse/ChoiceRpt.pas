{ ---------------------------------------------------------------------------- }
{                                                                              }
{ PC HOME ONLINE Copyright (c) 2002-2003                                       }
{                                                                              }
{ Project: PC home online 網路家庭( EC2 ) ERP Program                          }
{ Unit: Report Chioce Form                                                     }
{ Author: Bug                                                                  }
{ Date: 2003/08/20                                                             }
{                                                                              }
{ ---------------------------------------------------------------------------- }

unit ChoiceRpt;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls,
  { EC2 ERP Project Units }
  ApClass, ApConst,
  { Raize CodeSite, http://www.raize.com }
  {$IFDEF DEBUG} CsIntf, {$ENDIF}
  { Developer Express http://www.devexpress.com }
  { Common }
  cxClasses, cxStyles, cxCustomData, cxGraphics, cxFilter, cxData, cxDBData,
  cxControls, cxContainer, cxLookAndFeels, cxDataStorage, cxFormats,
  { Express Editor }
  cxEdit, cxTextEdit, cxButtonEdit, cxCheckBox, cxMaskEdit, cxDropDownEdit,
  cxSpinEdit, cxCurrencyEdit, cxHyperLinkEdit, cxImageComboBox, cxMemo,
  cxRadioGroup, cxLookAndFeelPainters, cxButtons;

type
  TfrmChoiceRpt = class(TForm)
    Panel2: TPanel;
    btnOK: TcxButton;
    btnCancel: TcxButton;
    Panel1: TPanel;
    ClientBox: TScrollBox;
    Image1: TImage;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
    FList: TList;
  public
    { Public declarations }
    procedure AddReportName(const AReportName: String;
      const AEnabled: Boolean = True);
  end;

var
  frmChoiceRpt: TfrmChoiceRpt;

implementation

uses Main, ApUtilis;

{$R *.dfm}

{ TfrmChoiceRpt }

{ ---------------------------------------------------------------------------- }

procedure TfrmChoiceRpt.AddReportName(const AReportName: String;
  const AEnabled: Boolean = True);
var
  ARadioButton: TcxRadioButton;
  ATop: Integer;
begin
  if AReportName = '' then
    raise Exception.Create( '指定報表名稱不可空白!' );
  if FList.Count = 0 then
    ATop := 50
  else
    ATop := TcxRadioButton( FList.Items[FList.Count - 1] ).Top + 28;
  if ATop >= Panel1.Height then Self.Height := Self.Height + 40;
  ARadioButton := TcxRadioButton.Create( Self );
  ARadioButton.Caption := AReportName;
  ARadioButton.LookAndFeel.Kind := lfFlat;
  ARadioButton.Enabled := AEnabled;
  ARadioButton.Parent := Panel1;
  ARadioButton.Top := ATop;
  ARadioButton.Left := 95;
  ARadioButton.Width := ( ClientBox.Width - ARadioButton.Left - 5 );
  ARadioButton.Checked := ( FList.Add( ARadioButton ) = 0 );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmChoiceRpt.FormCreate(Sender: TObject);
begin
  FList := TList.Create;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmChoiceRpt.FormClose(Sender: TObject; var Action: TCloseAction);
var
  AIndex: Integer;
begin
  for AIndex := 0 to FList.Count - 1 do
    TcxRadioButton( FList[AIndex] ).Free;
  Flist.Free;
end;

{ ---------------------------------------------------------------------------- }

end.
