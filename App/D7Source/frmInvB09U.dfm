object frmInvB09: TfrmInvB09
  Left = 443
  Top = 286
  ActiveControl = txtInvDate1
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsDialog
  Caption = '[B09] '#19978#20659#37325#26032#25480#27402
  ClientHeight = 188
  ClientWidth = 327
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  PixelsPerInch = 96
  TextHeight = 13
  object pnlMain: TPanel
    Left = 0
    Top = 0
    Width = 327
    Height = 188
    Align = alClient
    TabOrder = 0
    object mskInvId1: TcxMaskEdit
      Left = 87
      Top = 40
      ParentFont = False
      Properties.CharCase = ecUpperCase
      Properties.MaskKind = emkRegExprEx
      Properties.EditMask = '[a-zA-Z][a-zA-Z]\d\d\d\d\d\d\d\d'
      Properties.MaxLength = 0
      Properties.OnValidate = mskInvId1PropertiesValidate
      Style.Font.Charset = ANSI_CHARSET
      Style.Font.Color = clWindowText
      Style.Font.Height = -12
      Style.Font.Name = 'Tahoma'
      Style.Font.Style = []
      Style.IsFontAssigned = True
      TabOrder = 0
      Width = 101
    end
    object mskInvId2: TcxMaskEdit
      Left = 215
      Top = 40
      ParentFont = False
      Properties.CharCase = ecUpperCase
      Properties.MaskKind = emkRegExprEx
      Properties.EditMask = '[a-zA-Z][a-zA-Z]\d\d\d\d\d\d\d\d'
      Properties.MaxLength = 0
      Properties.OnValidate = mskInvId1PropertiesValidate
      Style.Font.Charset = ANSI_CHARSET
      Style.Font.Color = clWindowText
      Style.Font.Height = -12
      Style.Font.Name = 'Tahoma'
      Style.Font.Style = []
      Style.IsFontAssigned = True
      TabOrder = 1
      Width = 101
    end
    object cxLabel1: TcxLabel
      Left = 6
      Top = 16
      Caption = #30332#31080#26085#26399#36215#36804
    end
    object cxLabel2: TcxLabel
      Left = 196
      Top = 16
      Caption = '~'
    end
    object cxLabel3: TcxLabel
      Left = 6
      Top = 41
      Caption = #30332#31080#34399#30908#36215#36804
    end
    object cxLabel4: TcxLabel
      Left = 197
      Top = 43
      Caption = '~'
    end
    object btnQuery: TcxButton
      Left = 12
      Top = 150
      Width = 71
      Height = 27
      Caption = #30906#23450
      TabOrder = 6
      OnClick = btnQueryClick
      Glyph.Data = {
        36040000424D3604000000000000360000002800000010000000100000000100
        2000000000000004000000000000000000000000000000000000000000000000
        000000000000000000000000000A000000250000003300000033000000330000
        0033000000250000000A00000000000000000000000000000000000000000000
        00000000000000000022001D105C006738C9008C4BFF008B4AFF008B4AFF008C
        4BFF006738C9001D105C0000001E000000000000000000000000000000000000
        00000000001E005E33BB009050FF01A169FF01AB76FF01AC79FF01AC79FF01AB
        76FF01A169FF009050FF00532DAA0000001E0000000000000000000000000000
        000A00532DAA009152FF02AC77FF00C38CFF00D79BFF00DA9CFF00DA9CFF00D7
        9CFF01C38CFF01AB76FF009253FF00532DAA0000000A0000000000000000001C
        1051009051FF0FB483FF00D298FF00D598FF00D192FF00CF90FF00D091FF00D3
        96FF00D69BFF00D198FF01AB76FF009050FF001D105100000000000000000067
        36C916AB78FF10C996FF00D397FF00CD8CFFFFFFFFFFFFFFFFFFFFFFFFFF00CC
        8CFF00D195FF00D59BFF01C18CFF01A169FF006838C90000000000000000008A
        48FF39C49DFF00D198FF00CB8CFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
        FFFF00CA8CFF00CF96FF00D29BFF01AB76FF008C4BFF00000000000000000089
        46FF52D2B0FF00CC92FFFFFFFFFFFFFFFFFFFFFFFFFF00C484FFFFFFFFFFFFFF
        FFFFFFFFFFFF00C88DFF00D09AFF00AD79FF008B4AFF00000000000000000088
        45FF68DDBEFF00C991FFFFFFFFFFFFFFFFFF00C68CFF00C891FF00C58BFFFFFF
        FFFFFFFFFFFFFFFFFFFF00CC96FF00AD78FF008B4AFF00000000000000000088
        46FF76E0C6FF00CB98FF00C590FF00C691FF00C895FF00C997FF00C894FF00C3
        8CFFFFFFFFFFFFFFFFFF00C792FF00AB75FF008C4BFF00000000000000000065
        34BE59C9A4FF49DEBCFF00C794FF00C897FF00C998FF00C999FF00C998FF00C7
        94FF00C38EFFFFFFFFFF00BD8AFF00A067FF006838BF0000000000000000001C
        0F330A9458FFADF8E9FF18D0A7FF00C495FF00C697FF00C698FF00C798FF00C7
        98FF00C697FF00C596FF12B585FF008F50FF001C0F3300000000000000000000
        0000005C30AA199C63FFBCFFF7FF5EE4C9FF00C59AFF00C396FF00C497FF00C5
        9AFF22CAA2FF2FC196FF029355FF00522C950000000000000000000000000000
        00000000000000512A950E9659FF74D5B6FFA0F4E1FF94EFDCFF7CE6CCFF5ED6
        B5FF2EB587FF039152FF005D33AA000000000000000000000000000000000000
        00000000000000000000001C0F33006433BB008744FF008743FF008744FF0089
        46FF006636BB001C0F3300000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000}
    end
    object btnExit: TcxButton
      Left = 246
      Top = 150
      Width = 71
      Height = 27
      Caption = #38626#38283
      TabOrder = 7
      OnClick = btnExitClick
      Glyph.Data = {
        36040000424D3604000000000000360000002800000010000000100000000100
        2000000000000004000000000000000000000000000000000000000000000000
        00000000000000000000000000060000133C0000338C00004DB200004CB20000
        328F000013460000000C00000000000000000000000000000000000000000000
        0000000000000000194900007AE20000BFFF0000E1FF0000DCFF0000D6FF0000
        CFFF0000AAFF000070E500001956000000030000000000000000000000000000
        00000000266C0000C2FF0000F6FF0000F4FF0000F0FF0000EAFF0000E4FF0000
        DCFF0000D4FF0000CAFF0000A0FF000026750000000300000000000000000000
        204F1D1DC5FF0808FFFF0707FFFF0404FEFF0101FCFF0000F8FF0000F1FF0000
        E9FF0000E0FF0000D6FF0000CBFF0000A0FF0000206200000000000000003434
        92DF2727FFFF1919FFFF1212E2FF7272C5FF1717CFFF0707FEFF0101FCFF2F2F
        DCFF6F6FB8FF0000C8FF0000D5FF0000CAFF000076E50000000C0000265F4F4F
        E2FF2828FFFF2929FFFF9191C5FFFFFFFFFFCFCFE2FF1B1BCFFF3737E2FFEFEF
        F5FFFFFFFFFF7F7FC2FF0000DFFF0000D2FF0000B5FF0000266C1313529F5555
        FFFF3838FFFF3939FFFF4646BCFFEFEFF5FFFFFFFFFFCFCFE2FFEFEFF5FFFFFF
        FFFFEFEFF5FF2F2FC0FF0000E7FF0000DAFF0000CDFF00004BA8282873BF5C5C
        FFFF4949FFFF4949FFFF4747FFFF4A4ABCFFEFEFF5FFFFFFFFFFFFFFFFFFEFEF
        F5FF3333BCFF0000FBFF0000EFFF0000E2FF0000D4FF000066CC292973BF6969
        FFFF5959FFFF5959FFFF5757FFFF6161E2FFEFEFF5FFFFFFFFFFFFFFFFFFCFCF
        E2FF1A1ACFFF0505FEFF0000F6FF0000E8FF0000D9FF000066CC202062AF8080
        FFFF6969FFFF6A6AFFFF6F6FE2FFEFEFF5FFFFFFFFFFEFEFF5FFEFEFF5FFFFFF
        FFFFCFCFE2FF1515CFFF0000FAFF0000ECFF0000DEFF00005AB200002C6F9696
        F5FF7878FFFF7E7EFFFFB9B9D8FFFFFFFFFFEFEFF5FF5959BCFF5656BCFFEFEF
        F5FFFFFFFFFF9090C5FF0202FDFF0000F0FF0000D2FF00002C6F0000060F5F5F
        ACEFB2B2FFFFB4B4FFFF7F7FCFFFAFAFCFFF6B6BBCFF9191FFFF8787FFFF6262
        BCFFB7B7D8FF6666EBFF5555FEFF2C2CF4FF00009FEF00000613000000000C0C
        3C7FAAAAE2FFCFCFFFFFC7C7FFFFC0C0FFFFB8B8FFFFB1B1FFFFA9A9FFFFA2A2
        FFFF9A9AFFFF9393FFFF8C8CFEFF6565DDFF00003B7F00000000000000000000
        00001818529FB3B3E2FFDEDEFFFFD7D7FFFFD3D3FFFFCFCFFFFFCACAFFFFC6C6
        FFFFC2C2FFFFBFBFFFFF8B8BE2FF0707499F0000000000000000000000000000
        0000000000000B0B3C7F7777BCFFD6D6F5FFF0F0FFFFEEEEFFFFECECFFFFEBEB
        FFFFC9C9F5FF6464BCFF08083C7F000000000000000000000000000000000000
        0000000000000000000000000C1F0B0B438F303073BF5E5EA2EF6060ACEF2B2B
        73BF0909438F00000C1F00000000000000000000000000000000}
    end
    object txtInvDate1: TcxDateEdit
      Tag = 3
      Left = 87
      Top = 13
      ParentFont = False
      Properties.DateButtons = [btnToday]
      Style.Font.Charset = DEFAULT_CHARSET
      Style.Font.Color = clWindowText
      Style.Font.Height = -12
      Style.Font.Name = 'MS Sans Serif'
      Style.Font.Style = []
      Style.StyleController = dtmMain.cxEditStyle
      Style.IsFontAssigned = True
      TabOrder = 9
      Width = 102
    end
    object txtInvDate2: TcxDateEdit
      Tag = 3
      Left = 215
      Top = 13
      ParentFont = False
      Properties.DateButtons = [btnToday]
      Style.Font.Charset = DEFAULT_CHARSET
      Style.Font.Color = clWindowText
      Style.Font.Height = -12
      Style.Font.Name = 'MS Sans Serif'
      Style.Font.Style = []
      Style.StyleController = dtmMain.cxEditStyle
      Style.IsFontAssigned = True
      TabOrder = 10
      Width = 102
    end
    object grpDiscount: TGroupBox
      Left = 8
      Top = 70
      Width = 311
      Height = 75
      Enabled = False
      TabOrder = 11
      object txtDiscDate1: TcxDateEdit
        Tag = 3
        Left = 96
        Top = 18
        ParentFont = False
        Properties.DateButtons = [btnToday]
        Style.Font.Charset = DEFAULT_CHARSET
        Style.Font.Color = clWindowText
        Style.Font.Height = -12
        Style.Font.Name = 'MS Sans Serif'
        Style.Font.Style = []
        Style.StyleController = dtmMain.cxEditStyle
        Style.IsFontAssigned = True
        TabOrder = 0
        Width = 93
      end
      object cxLabel5: TcxLabel
        Left = 193
        Top = 24
        Caption = '~'
      end
      object txtDiscDate2: TcxDateEdit
        Tag = 3
        Left = 210
        Top = 18
        ParentFont = False
        Properties.DateButtons = [btnToday]
        Style.Font.Charset = DEFAULT_CHARSET
        Style.Font.Color = clWindowText
        Style.Font.Height = -12
        Style.Font.Name = 'MS Sans Serif'
        Style.Font.Style = []
        Style.StyleController = dtmMain.cxEditStyle
        Style.IsFontAssigned = True
        TabOrder = 2
        Width = 93
      end
      object cxLabel6: TcxLabel
        Left = 6
        Top = 20
        Caption = #25240#35731#21934#26085#26399#36215#36804
      end
      object cxLabel7: TcxLabel
        Left = 6
        Top = 45
        Caption = #25240#35731#21934#34399#30908#36215#36804
      end
      object txtDiscNoSt: TcxMaskEdit
        Left = 96
        Top = 44
        ParentFont = False
        Properties.CharCase = ecUpperCase
        Properties.MaxLength = 0
        Style.Font.Charset = ANSI_CHARSET
        Style.Font.Color = clWindowText
        Style.Font.Height = -12
        Style.Font.Name = 'Tahoma'
        Style.Font.Style = []
        Style.IsFontAssigned = True
        TabOrder = 5
        Width = 94
      end
      object cxLabel8: TcxLabel
        Left = 193
        Top = 48
        Caption = '~'
      end
      object txtDiscNoEd: TcxMaskEdit
        Left = 210
        Top = 44
        ParentFont = False
        Properties.CharCase = ecUpperCase
        Properties.MaxLength = 0
        Style.Font.Charset = ANSI_CHARSET
        Style.Font.Color = clWindowText
        Style.Font.Height = -12
        Style.Font.Name = 'Tahoma'
        Style.Font.Style = []
        Style.IsFontAssigned = True
        TabOrder = 6
        Width = 94
      end
    end
    object chkDiscount: TcxCheckBox
      Left = 16
      Top = 65
      Caption = #25240#35731#35387#35352
      Properties.OnEditValueChanged = chkDiscountPropertiesEditValueChanged
      TabOrder = 8
      Width = 71
    end
  end
end
