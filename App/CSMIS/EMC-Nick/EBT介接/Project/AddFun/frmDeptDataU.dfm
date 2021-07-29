inherited frmDeptData: TfrmDeptData
  Caption = 'frmDeptData'
  OldCreateOrder = True
  PixelsPerInch = 96
  TextHeight = 20
  inherited Panel1: TPanel
    inherited Panel2: TPanel
      inherited btnPrint: TBitBtn
        Visible = False
      end
    end
  end
  inherited Panel3: TPanel
    inherited pnlSingleData: TPanel
      object StaticText1: TStaticText
        Left = 24
        Top = 16
        Width = 36
        Height = 24
        Caption = #20195#30908
        TabOrder = 0
      end
      object StaticText2: TStaticText
        Left = 216
        Top = 16
        Width = 36
        Height = 24
        Caption = #21517#31281
        TabOrder = 1
      end
      object DBRadioGroup1: TDBRadioGroup
        Left = 392
        Top = 8
        Width = 185
        Height = 57
        Caption = ' '#26159#21542#20572#29992' '
        Columns = 2
        DataField = 'STOPFLAG'
        DataSource = dsrSTabMaintain
        Items.Strings = (
          #26159
          #21542)
        TabOrder = 2
        Values.Strings = (
          '1'
          '0')
      end
      object DBEdit1: TDBEdit
        Left = 64
        Top = 16
        Width = 121
        Height = 28
        DataField = 'CODENO'
        DataSource = dsrSTabMaintain
        TabOrder = 3
      end
      object DBEdit2: TDBEdit
        Left = 256
        Top = 16
        Width = 121
        Height = 28
        DataField = 'DESCRIPTION'
        DataSource = dsrSTabMaintain
        TabOrder = 4
      end
    end
    inherited pnlMultiData: TPanel
      inherited DBGrid1: TDBGrid
        DataSource = dsrSTabMaintain
        Columns = <
          item
            Expanded = False
            FieldName = 'CODENO'
            Title.Caption = #20195#30908
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'DESCRIPTION'
            Title.Caption = #21517#31281
            Width = 64
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'STOPFLAGDESC'
            Title.Caption = #26159#21542#20572#29992
            Width = 64
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'UPDTIME'
            Title.Caption = #30064#21205#26178#38291
            Width = 64
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'UPDEN'
            Title.Caption = #30064#21205#20154#21729
            Width = 64
            Visible = True
          end>
      end
    end
  end
  inherited dsrSTabMaintain: TDataSource
    DataSet = dtmDeptData.ClientDataSet1
  end
end
