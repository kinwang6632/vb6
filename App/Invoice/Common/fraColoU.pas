unit fraColoU;

interface

uses 
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls, StdCtrls, Mask, DBCtrls, EnumDBRadioGroup, Db, Buttons,
  DBTables;

type
  TfraColo = class(TFrame)
    lblPercent: TLabel;
    lblAmount: TLabel;
    lblUnitPrice: TLabel;
    lblUnit: TLabel;
    lblCoinID: TLabel;
    lblQuantity: TLabel;
    dblFromTo: TDBLookupComboBox;
    dbcCoload: TDBCheckBox;
    edbrInOut: TEnumDBRadioGroup;
    dbcIsSplit: TDBCheckBox;
    dbePercent: TDBEdit;
    dbeAmount: TDBEdit;
    dbeUnitPrice: TDBEdit;
    dblFTCoinName: TDBLookupComboBox;
    dbeQuantity: TDBEdit;
    dsrComp: TDataSource;
    dsrCoin: TDataSource;
    BitBtn1: TBitBtn;
    cobFTUnit: TComboBox;
    procedure dbeAmountEnter(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

implementation

{$R *.DFM}

procedure TfraColo.dbeAmountEnter(Sender: TObject);
begin
    if dbeAmount.DataSource.State <> dsEdit then
      dbeAmount.DataSource.DataSet.Edit;
    dbeAmount.DataSource.DataSet.FieldByName('fFTAmount').AsFloat :=
      dbeAmount.DataSource.DataSet.FieldByName('fFTQuantity').AsFloat *
      dbeAmount.DataSource.DataSet.FieldByName('fFTUPrice').AsFloat ; 
end;

end.
