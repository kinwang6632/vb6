object frmInvA07_1: TfrmInvA07_1
  Left = 136
  Top = 141
  Width = 809
  Height = 482
  Caption = 'frmInvA07_1'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  ShowHint = True
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 14
  object Bevel1: TBevel
    Left = 0
    Top = 39
    Width = 801
    Height = 3
    Align = alTop
    Shape = bsSpacer
  end
  object ScrollBox1: TScrollBox
    Left = 0
    Top = 0
    Width = 801
    Height = 39
    Align = alTop
    BevelInner = bvNone
    BorderStyle = bsNone
    TabOrder = 0
    object btnInsert: TcxButton
      Left = 15
      Top = 5
      Width = 85
      Height = 27
      Caption = #26032#22686#25240#35731#21934
      TabOrder = 0
      OnClick = btnInsertClick
      Glyph.Data = {
        36060000424D3606000000000000360000002800000020000000100000000100
        1800000000000006000000000000000000000000000000000000FF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        9B7C6B9D7E6D9C7E6D9C7E6D9C7E6D9C7D6D9C7D6CBFABA1007500007000006D
        00BFABA1FF00FFFF00FFFF00FFFF00FFA47878A47878A47878A47878A47878A4
        7878A47878BEA3A3000000000000000000BEA3A3FF00FFFF00FFFF00FFFF00FF
        9B7766FFFFFFFAF4E9FAF4E9FAF4E9FAF3E8FAF3E7FCF8F1007D0044DD770072
        00BDA99DFF00FFFF00FFFF00FFFF00FF857769FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FF000000FF00FF000000BEA3A3FF00FFFF00FFFF00FFFF00FF
        A27F6FFFFFFFDDC2B5DDC2B5DCC2B5E9D6CD00870000850000810048E17B007A
        00007500007000FF00FFFF00FFFF00FFA47878FF00FFDCC1C1DCC1C1DCC1C1E9
        D4D4000000000000000000FF00FF000000000000000000FF00FFFF00FFFF00FF
        A38070FFFFFFDBC3BBDBC3BADBC2B8EBDCD5008D005EF7915AF38D53EC8648E1
        7B45DE78007800FF00FFFF00FFFF00FFA47878FF00FFDCC1C1DCC1C1DCC1C1EB
        D7D7000000FF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFF00FFFF00FF
        A98778FFFFFFDBC7C2DBC6C1DBC4BCEBDDD7009100008D00008B005AF38D0083
        00008100007D00FF00FFFF00FFFF00FFA98888FF00FFDCC1C1DCC1C1DCC1C1EB
        D7D7000000000000000000FF00FF000000000000000000FF00FFFF00FFFF00FF
        AB897AFFFFFFDBC7C3DBC7C2DBC5BEEBDED9EBDCD6EBDBD3008F005EF7910089
        00C8B7AEFF00FFFF00FFFF00FFFF00FFA98888FF00FFDCC1C1DCC1C1DCC1C1EC
        D9D9EBD7D7EBD7D7000000FF00FF000000C2AEAEFF00FFFF00FFFF00FFFF00FF
        AF8E7FFFFFFFDCC5C0DCC5BFDBC4BDDBC2B9DBBFB4EBDBD3009300009100008D
        00BFABA1FF00FFFF00FFFF00FFFF00FFB18B8BFF00FFDCC1C1DCC1C1DCC1C1DC
        C1C1DABBBBEBD7D7000000000000000000BEA3A3FF00FFFF00FFFF00FFFF00FF
        AF8F80FFFFFFFEFEFEFEFEFEFEFDFDFEFCFBFDFBF8FEFCFAFDFCF8FDFBF7FCF6
        ED9B7C6CFF00FFFF00FFFF00FFFF00FFB18B8BFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFA47878FF00FFFF00FFFF00FFFF00FF
        AF8F80FFFFFFDFCECCDFCDCBDECAC6DEC6C0DEC4BADEC1B4DEBEADDEBEABFCF6
        EE9C7C6DFF00FFFF00FFFF00FFFF00FFB18B8BFF00FFDDCBCBDDCBCBE0C7C7DC
        C1C1DEBDBDDEBDBDDEBDBDDEBDBDFF00FFA47878FF00FFFF00FFFF00FFFF00FF
        B19080FFFFFFFFFFFFFFFFFFFFFFFFFEFEFEFEFDFBFDFBF8FCF9F4F9F4EEF0E8
        E09D7F6FFF00FFFF00FFFF00FFFF00FFB18B8BFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFA47878FF00FFFF00FFFF00FFFF00FF
        B79787FFFFFFFFFFFFFFFFFFFFFFFFFEFEFEFEFDFBFDFBF9A78270A78270A782
        70A78270FF00FFFF00FFFF00FFFF00FFB39292FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFA47878A47878A47878A47878FF00FFFF00FFFF00FFFF00FF
        B89888FFFFFFFFFFFFFFFFFFFEFEFEFEFEFDFEFCFAFDFBF9A78270F5E2D9B18E
        7EAB9E98FF00FFFF00FFFF00FFFF00FFBD9494FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFA47878FF00FFB18B8BAA9F9EFF00FFFF00FFFF00FFFF00FF
        B89888FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFA78270B18E7EAB9E
        98FF00FFFF00FFFF00FFFF00FFFF00FFBD9494FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFA47878B18B8BAA9F9EFF00FFFF00FFFF00FFFF00FFFF00FF
        B89888B89888B49383B49383B08E7DB08E7DAC8877AC8877A78270AB9E98FF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFBD9494BD9494B39292B39292B18B8BB1
        8B8BA98888A98888A47878AA9F9EFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF}
      NumGlyphs = 2
    end
    object btnEdit: TcxButton
      Left = 711
      Top = 25
      Width = 85
      Height = 27
      Caption = #20462#25913
      TabOrder = 1
      Visible = False
      OnClick = btnEditClick
      Glyph.Data = {
        36060000424D3606000000000000360000002800000020000000100000000100
        1800000000000006000000000000000000000000000000000000FF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        9B7C6B9D7E6D9C7E6D9C7E6D9C7E6D9C7D6D9C7D6C9B7C6B9B7C6B9A7C6A997B
        689B7C6BFF00FFFF00FFFF00FFFF00FFA47878A47878A47878A47878A47878A4
        7878A47878A47878A47878857769857769A47878FF00FFFF00FFFF00FFFF00FF
        9B7766FFFFFFFAF4E9FAF4E9EEE9DEE8E2D8F7F0E4FAF2E6FAF1E4F9EFE0F8ED
        DA977967FF00FFFF00FFFF00FFFF00FF857769FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FF857769FF00FFFF00FFFF00FFFF00FF
        A27F6FFFFFFFDDC2B5DDC2B5B5A9A4B1A19ADBC2B4DCBBA9DCBAA5DCBAA3FAF1
        E2987968FF00FFFF00FFFF00FFFF00FFA47878FF00FFDCC1C1DCC1C1B1A6A6B3
        9F9FDCC1C1DABBBBDABBBBDABBBBFF00FF857769FF00FFFF00FFFF00FFFF00FF
        A38070FFFFFFDBC3BBEADDD76F6D71928B96C3CDB9E8D6CCDAB8A5DCB9A5FAF3
        E6997A6AFF00FFFF00FFFF00FFFF00FFA47878FF00FFDCC1C1EBD7D76D6D6C94
        9292C5C2C3E8D0D0DABBBBDABBBBFF00FF857769FF00FFFF00FFFF00FFFF00FF
        A98778FFFFFFDBC7C2E9DDDAA8BBCE34B3562CB44471B46FECDBD2DCBBA7FBF4
        E8997B6BFF00FFFF00FFFF00FFFF00FFA98888FF00FFDCC1C1ECD9D9B6B6B659
        585C424141737373ECD9D9DABBBBFF00FF857769FF00FFFF00FFFF00FFFF00FF
        AB897AFFFFFFDBC7C3E6D7D499D0A766FF985AEC862EAD4687BE81EAD8CCFBF5
        EA9A7C6BFF00FFFF00FFFF00FFFF00FFA98888FF00FFDCC1C1E9D4D4A9A9A999
        99998786874A4B4B878687EAD6D6FF00FFA47878FF00FFFF00FFFF00FFFF00FF
        AF8E7FFFFFFFDCC5C0DEC9C3DBE4D657E27F6AFF9D55E17C2AA43C9CC494FCF8
        F29B7C6BFF00FFFF00FFFF00FFFF00FFB18B8BFF00FFDCC1C1E0C7C7DCD7D781
        81819999998181813A3A3A9D9592FF00FFA47878FF00FFFF00FFFF00FFFF00FF
        AF8F80FFFFFFFEFEFEFEFEFEFEFDFDD1F8DC54EE8368FF9B50DC77249E38B7DC
        B6B69F94FF00FFFF00FFFF00FFFF00FFB18B8BFF00FFFF00FFFF00FFFF00FFD7
        D8D88282839999997777773A3A3AB6B6B6B79F9FFF00FFFF00FFFF00FFFF00FF
        AF8F80FFFFFFDFCECCDFCDCBDECAC6E9D9D5ADEAB95AF68864FF9742DA693487
        3EC6BDB6FF00FFFF00FFFF00FFFF00FFB18B8BFF00FFDDCBCBDDCBCBE0C7C7EA
        D6D6B6B6B6878687979797686868424141C6BFC0FF00FFFF00FFFF00FFFF00FF
        B19080FFFFFFFFFFFFFFFFFFFFFFFFFEFEFEFEFEFD91EFAC55FC889AC1A4CDBB
        B66E6D8CFF00FFFF00FFFF00FFFF00FFB18B8BFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFABACAC878687A9A9A9D2BCBC707274FF00FFFF00FFFF00FFFF00FF
        B79787FFFFFFFFFFFFFFFFFFFFFFFFFEFEFEFEFDFBFBFCFAA5C5A6FFFFFF7892
        F5203DDC292AA1FF00FFFF00FFFF00FFB39292FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFA9A9A9FF00FF9492923A3A3A3A3A3AFF00FFFF00FFFF00FF
        B89888FFFFFFFFFFFFFFFFFFFEFEFEFEFEFDFEFCFAFDFBF9DED0C98C99DE4277
        FF2D4AD81F209FFF00FFFF00FFFF00FFBD9494FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFDDCBCB9999997777774A4B4B3A3A3AFF00FFFF00FFFF00FF
        B89888FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFB39384C1B9D03545
        C52C30A7FF00FFFF00FFFF00FFFF00FFBD9494FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFB39292C5C2C34241413A3A3AFF00FFFF00FFFF00FFFF00FF
        B89888B89888B49383B49383B08E7DB08E7DAC8877AC8877A78270B3A8A2FF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFBD9494BD9494B39292B39292B18B8BB1
        8B8BA98888A98888A47878B1A6A6FF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF}
      NumGlyphs = 2
    end
    object btnCancel: TcxButton
      Left = 549
      Top = 29
      Width = 85
      Height = 27
      Cancel = True
      Caption = #21462#28040
      TabOrder = 2
      Visible = False
      OnClick = btnCancelClick
      Glyph.Data = {
        36060000424D3606000000000000360000002800000020000000100000000100
        1800000000000006000000000000000000000000000000000000FF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFC58634D48A
        2CC99350FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFB18B8B808080808080FF00FFFF00FFFF00FFFF00FF
        C09363FF00FFFF00FFFF00FFFF00FFFF00FFC28A4DD48729DF9239FFD0A0FFDB
        B5D68C2CFF00FFFF00FFFF00FFFF00FFBD9494FF00FFFF00FFFF00FFFF00FFFF
        00FFB18B8B808080808080FF00FFFF00FF808080FF00FFFF00FFFF00FFB78A5C
        BE854FFF00FFFF00FFFF00FFC78C4FD6842BEAA04CFFC98EFFC88EFFC78FFFCF
        A0DE8E2AFF00FFFF00FFFF00FFB18B8BB18B8BFF00FFFF00FFFF00FF80808080
        8080808080FF00FFFF00FFFF00FFFF00FF808080FF00FFFF00FFFF00FFA96D44
        BD8450FF00FFFF00FFFF00FFE2933EFFE2B6FFD19DFFCD99FFCC9AFFC48BFFCB
        95E79E4FBC8F5AFF00FFFF00FFA47878A47878FF00FFFF00FFFF00FF808080FF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FF808080B18B8BFF00FFFF00FFB96727
        B58360FF00FFFF00FFFF00FFC29162C28241B6753AD78628FFC587FFC286FFC7
        8EF7BA79BF833BFF00FFFF00FF8C6060A47878FF00FFFF00FFFF00FF808080A4
        7878A47878C7A386FF00FFFF00FFFF00FF808080A47878FF00FFFF00FFCC6A18
        B67A4EFF00FFFF00FFFF00FFFF00FFFF00FFC0813DFFBB77FFC282FEBE7DFFCE
        9BFFCF98C27C2AFF00FFFF00FF8C6060A47878FF00FFFF00FFFF00FFFF00FFFF
        00FFA47878FF00FFFF00FFFF00FFFF00FFFF00FFA47878FF00FFFF00FFCC6813
        CA732BFF00FFFF00FFFF00FFFF00FFFF00FFD9872DFFBF79FFC483D58930D890
        3EFFD8AAC9832FFF00FFFF00FF8C6060A47878FF00FFFF00FFFF00FFFF00FFFF
        00FF808080FF00FFFF00FF808080808080FF00FF808080FF00FFFF00FFBF651D
        EB7F16FF00FFFF00FFFF00FFFF00FFC67621FFBB6EFFBB71FCB86FC08340B587
        5BFFD7AAD48730FF00FFFF00FF8C6060A47878FF00FFFF00FFFF00FFFF00FFA4
        7878FF00FFFF00FFFF00FFA47878B18B8BFF00FF808080FF00FFFF00FFA8673D
        FFA223DB7314BC7231BC7539CD7720FFB35CFFB262FFBD73C67C2EFF00FFFF00
        FFD08635C48D4FFF00FFFF00FF808080D9AF8BA47878A47878A47878A4787880
        8080FF00FFFF00FFA47878FF00FFFF00FF808080808080FF00FFFF00FFFF00FF
        E27A1BFF9F2AFF9F2DFFA740FFAA4BFFAB50FFC986D27D2CB89979FF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFA47878FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFA47878BD9494FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        B6744DEC8C2FFFC679FFC37CFFCB8BFFC886CA7326B99471FF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFA47878FF00FFFF00FFFF00FFFF00FFFF
        00FFA47878BD9494FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFB28361B05F24C66618BF671DB07240FF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFA478788C60608C60608C6060A4
        7878FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF}
      NumGlyphs = 2
    end
    object btnSave: TcxButton
      Left = 469
      Top = 21
      Width = 85
      Height = 27
      Caption = #23384#27284
      TabOrder = 3
      Visible = False
      OnClick = btnSaveClick
      Glyph.Data = {
        36060000424D3606000000000000360000002800000020000000100000000100
        1800000000000006000000000000000000000000000000000000FF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        9B7C6B9D7E6D9C7E6D9C7E6D9F8271C7BAB0709F642B811E1E710831781D6C96
        61C6BBAFFF00FFFF00FFFF00FFFF00FFA47878A47878A47878A47878A47878A4
        78787072743A3A3A3A3A3A3A3A3A6D6D6C857769FF00FFFF00FFFF00FFFF00FF
        9B7766FFFFFFFAF4E9FAF4E9F6F5EE41A43E009A0000A2070097000E89002E71
        003C7C2DFF00FFFF00FFFF00FFFF00FF857769FF00FFFF00FFFF00FFFF00FF42
        41410000000000000000000000003A3A3A3A3A3AFF00FFFF00FFFF00FFFF00FF
        A27F6FFFFFFFDDC2B5E6D2C882B7790DA71A2CB53FFDFBFF6FD4880098000095
        003070005A8A5AFF00FFFF00FFFF00FFA47878FF00FFDCC1C1E6CECE8181813A
        3A3A424141FF00FF8786870000000000003A3A3A59585CFF00FFFF00FFFF00FF
        A38070FFFFFFDBC3BBECDFDA43AA472EB33CF5ECF4FFF6FFFFFFFF6FD588009A
        001288002A751AFF00FFFF00FFFF00FFA47878FF00FFDCC1C1ECD9D94A4B4B3A
        3A3AFF00FFFF00FFFF00FF8786870000000000003A3A3AFF00FFFF00FFFF00FF
        A98778FFFFFFDBC7C2EEE4E23AAE445DB964C0D6BC03AD219DDBA7FFFFFF6DD4
        870095001B7106FF00FFFF00FFFF00FFA98888FF00FFDCC1C1EEE3E34241415E
        6062C5C2C33A3A3AA9A9A9FF00FF8786870000003A3A3AFF00FFFF00FFFF00FF
        AB897AFFFFFFDBC7C3ECE1DE4EB55743CD6D19BB4225C04F0BB22E9DD9A5FFFF
        FF67CF7C237A1BFF00FFFF00FFFF00FFA98888FF00FFDCC1C1EFDEDE59585C6D
        6D6C4241414A4B4B3A3A3AA9A9A9FF00FF8181813A3A3AFF00FFFF00FFFF00FF
        AF8E7FFFFFFFDCC5C0E5D4CF85C07E72DD9635CE6A2EC75E25BE4B09AF278BD0
        944EBB595C925CFF00FFFF00FFFF00FFB18B8BFF00FFDCC1C1E7D2D287868797
        97976A6A6A5C5F624A4B4B3A3A3A97979759585C5C5F62FF00FFFF00FFFF00FF
        AF8F80FFFFFFFEFEFEFEFEFEF9FBF95AC06071DD964AD37A29C2521CB73B0DAA
        22369530FF00FFFF00FFFF00FFFF00FFB18B8BFF00FFFF00FFFF00FFFF00FF5C
        5F629797977777774A4B4B3A3A3A3A3A3A3A3A3AFF00FFFF00FFFF00FFFF00FF
        AF8F80FFFFFFDFCECCDFCDCBDFCBC7E9DFDA86C07E52B85B3CB14844AC4889C4
        86C7BAAFFF00FFFF00FFFF00FFFF00FFB18B8BFF00FFDDCBCBDDCBCBE0C7C7EC
        D9D987868759585C4A4B4B4A4B4B878687A47878FF00FFFF00FFFF00FFFF00FF
        B19080FFFFFFFFFFFFFFFFFFFFFFFFFEFEFEFEFEFCFEFDFBFEFCFAFCF9F6F4EE
        E89E7F71FF00FFFF00FFFF00FFFF00FFB18B8BFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFA47878FF00FFFF00FFFF00FFFF00FF
        B79787FFFFFFFFFFFFFFFFFFFFFFFFFEFEFEFEFDFBFDFBF9A78270A78270A782
        70A78270FF00FFFF00FFFF00FFFF00FFB39292FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFA47878A47878A47878A47878FF00FFFF00FFFF00FFFF00FF
        B89888FFFFFFFFFFFFFFFFFFFEFEFEFEFEFDFEFCFAFDFBF9A78270F5E2D9B18E
        7EAB9E98FF00FFFF00FFFF00FFFF00FFBD9494FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFA47878FF00FFB18B8BA47878FF00FFFF00FFFF00FFFF00FF
        B89888FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFA78270B18E7EAB9E
        98FF00FFFF00FFFF00FFFF00FFFF00FFBD9494FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFA47878B18B8BA47878FF00FFFF00FFFF00FFFF00FFFF00FF
        B89888B89888B49383B49383B08E7DB08E7DAC8877AC8877A78270AB9E98FF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFBD9494BD9494B39292B39292B18B8BB1
        8B8BA98888A98888A47878A47878FF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF}
      NumGlyphs = 2
    end
    object btnQuit: TcxButton
      Left = 287
      Top = 5
      Width = 85
      Height = 27
      Cancel = True
      Caption = #32080#26463
      TabOrder = 4
      OnClick = btnQuitClick
      Glyph.Data = {
        36060000424D3606000000000000360000002800000020000000100000000100
        1800000000000006000000000000000000000000000000000000FF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFCB6611FF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FF8C6060FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFCB6611CB6611FF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FF8C60608C6060FF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFCB6611FFE0B9CB66
        11FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FF8C6060FF00FF8C6060FF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFCB6611CB6611CB6611CB6611CB6611CB6611CB6611FFE0B9FFE0
        B9CB6611FF00FFFF00FFFF00FFFF00FFFF00FFFF00FF8C60608C60608C60608C
        60608C60608C60608C6060FF00FFFF00FF8C6060FF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFCB6611FFE0B9FFE0B9FFE0B9FFE0B9FFE0B9FFE0B9FFE0B9FFE0
        B9FFE0B9CB6611FF00FFFF00FFFF00FFFF00FFFF00FF8C6060FF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF8C6060FF00FFAC5D27AC5D27
        AC5D27AC5D27CB6611FFE0B9FFE0B9FFE0B9FFE0B9FFE0B9FFE0B9FFE0B9FFE0
        B9FFE0B9FFE0B9CB66118857578857578857578857578C6060FF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF8C6060AC5D27FFFFFF
        FFFFFFFFFFFFAC5D27FFE0B9FFE0B9FFE0B9FFE0B9FFE0B9FFE0B9FFE0B9FFE0
        B9FFE0B9CB6611FF00FF885757FF00FFFF00FFFF00FF885757FF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF8C6060FF00FFAC5D27FFFFFF
        B58868B58868CB6611CB6611CB6611CB6611CB6611CB6611CB6611FFE0B9FFE0
        B9CB6611FF00FFFF00FF885757FF00FFB18B8BB18B8B8C60608C60608C60608C
        60608C60608C60608C6060FF00FFFF00FF8C6060FF00FFFF00FFAC5D27FFFFFF
        FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFAC5D27FF00FFCB6611FFE0B9CB66
        11FF00FFFF00FFFF00FF885757FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FF885757FF00FF8C6060FF00FF8C6060FF00FFFF00FFFF00FFAC5D27FFFFFF
        B58868B58868B58868B58868B58868FFFFFFAC5D27FF00FFCB6611CB6611FF00
        FFFF00FFFF00FFFF00FF885757FF00FFB18B8BB18B8BB18B8BB18B8BB18B8BFF
        00FF885757FF00FF8C60608C6060FF00FFFF00FFFF00FFFF00FFAC5D27FFFFFF
        FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFAC5D27FF00FFCB6611FF00FFFF00
        FFFF00FFFF00FFFF00FF885757FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FF885757FF00FF8C6060FF00FFFF00FFFF00FFFF00FFFF00FFAC5D27AC5D27
        AC5D27AC5D27AC5D27AC5D27AC5D27AC5D27AC5D27FF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FF88575788575788575788575788575788575788575788
        5757885757FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFAC5D27AC5D27
        AC5D27AC5D27AC5D27AC5D27AC5D27AC5D27AC5D27FF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FF88575788575788575788575788575788575788575788
        5757885757FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF}
      NumGlyphs = 2
    end
    object btnQuery: TcxButton
      Left = 103
      Top = 5
      Width = 85
      Height = 27
      Caption = #26597#35426
      TabOrder = 5
      OnClick = btnQueryClick
      Glyph.Data = {
        36040000424D3604000000000000360000002800000010000000100000000100
        2000000000000004000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000FF018B8A7AC385827DA26B6B6B3F00000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        00000000FF01094FFF974392F6FFEEE9DFFF86827FA700000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        FF020D51FF9B439AFFFF6ADAFFFF5DADF5FF998F7ED900000000000000000000
        00000000000000000000000000000000000000000000000000006B7E690F0041
        FC9F469FFFFF6FDAFFFF50ACFFFF1357FFB9002AFF0A00000000000000000000
        000000000000000000006164682E515459524C4E523F1617180698928B9E7091
        B2FF61D3FFFF4EAAFFFF1657FFB4001FFF070000000000000000000000000000
        000076777A257B7E80C8B3A081FED2B588FDC3AA83FD83817AE566686CEAFFF7
        F0FF6B93BDFF084AFEAF0028FF06000000000000000000000000000000007B7D
        7F188D8A82E6F5CB84FEF5CB84FFF1C885FFF5CE8EFFFCD08CFFAE9E85FE7071
        75F3A9A193BB3956B20A0000000000000000000000000000000000000000767A
        7F95ECCB8EFFF3D192FFEECE92FFEDCC8EFFECC784FFEDC687FFFDD28FFF8783
        7CDB000000010000000000000000000000000000000000000000000000009894
        8BE1FEDFA1FFF2DAA5FFF2DBA7FFF1D79FFFEFD097FFECC886FFF6D298FFC1A9
        87FF5254592A000000000000000000000000000000000000000000000000A7A1
        91F4FFEBB9FFF8ECC6FFF7EBC0FFF5E3B2FFF2DAA3FFEECF94FFF4D093FFCFB6
        8DFF5153583D0000000000000000000000000000000000000000000000009996
        92D0FFF7CBFFFDFAE6FFFDF9E7FFF8EDC5FFF4E0B0FFF0D49BFFFAD594FFAF9F
        85FE6264671A000000000000000000000000000000000000000000000000A0A1
        A56FD5CCB1FFFFFFF2FFFFFFF2FFFBF3D2FFF6E2B4FFF7D99EFFF6D393FF7D7D
        81B800000000000000000000000000000000000000000000000000000000D9D9
        DA04A1A1A2B8D6CFB6FFFFFFDCFFFFF6CAFFFFEBB1FFEDD49CFE8C8983E47677
        7A20000000000000000000000000000000000000000000000000000000000000
        0000EBECED059C9CA07B979893E0A8A397FE97948CF07679809E7C7D7F1A0000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000}
    end
    object btnReport: TcxButton
      Left = 193
      Top = 5
      Width = 85
      Height = 27
      Caption = #22577#34920
      TabOrder = 6
      OnClick = btnReportClick
      Glyph.Data = {
        36060000424D3606000000000000360000002800000020000000100000000100
        1800000000000006000000000000000000000000000000000000FF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        C9C6C28068508060508060507060507058407058407050407050406048306048
        30604830A29A92FF00FFFF00FFFF00FFC9C6C280685080605080605070605070
        5840705840705040705040604830604830604830A29A92FF00FFFF00FFC1C4C3
        A38D85E0D0C0B0A090B0A090B0A090B0A090B0A090B0A090B0A090B0A090B0A0
        90B0A090604830FF00FFFF00FFC1C4C3A38D85FF00FFB0A090B0A090B0A090B0
        A090B0A090B0A090B0A090B0A090B0A090B0A090604830FF00FFFF00FFBCB7B0
        B29C94FFE8E0FFF8F0FFF0E0FFE8E0F0D8D0F0D0B0F0C0A000A00000A00000A0
        00705840604830857767FF00FFBCB7B0B29C94FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FF705840FF00FF857767FF00FFB29485
        E0D8D0FFFFFFFFFFFFFFFFFFFFFFFFFFF0F0F0E0E0F0D8C000FF1000FFB000A0
        00806850705040604830FF00FFB29485E0D8D0FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FF806850FF00FF604830FF00FFB29485
        F0E8E0FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF8F0F0E8E000FF1000FF1000A0
        00907060706050604830FF00FFB29485F0E8E0FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FF907060FF00FF604830FF00FFB09880
        D0C0B0D0C0B0C0B0A0B29C94B09880A088809080709070608068608060507058
        50908070806860705840FF00FFB09880D0C0B0D0C0B0C0B0A0B29C94B09880A0
        8880908070907060806860806050705850908070FF00FF705840FF00FFC0A090
        FFF8F0FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF8F0FFF0F0F0F0F0F0E8
        E0A38D85907860806050FF00FFC0A090FF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFA38D85FF00FF806050FF00FFCEC9C3
        B6A18CD0B0A0C0A8A0D0B0A0C0A090B29485A080709070608060507060508070
        60B0A090A08870806050FF00FFCEC9C3B6A18CD0B0A0C0A8A0D0B0A0C0A090B2
        9485A08070907060806050706050807060B0A090A08870806050FF00FFFF00FF
        C9C6C2C0B0A0E0C8C0FFFFFFFFF8FFFFF8FFFFF0F0F0F0E0F0E0E0C0B0A08060
        50A08880C0B0A0806050FF00FFFF00FFC9C6C2C0B0A0E0C8C0FF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFC0B0A0806050A08880C0B0A0806050FF00FFFF00FF
        FF00FFDEDFDDC0B0A0FFFFFFF0E8E0D0C8C0D0C8C0D0B8B0D0B8B0E0D0D08068
        60806050B29C94B0A090FF00FFFF00FFFF00FFDEDFDDC0B0A0FF00FFF0E8E0D0
        C8C0D0C8C0D0B8B0D0B8B0E0D0D0806860806050B29C94B0A090FF00FFFF00FF
        FF00FFFF00FFD8CBBCF0E8E0FFFFFFFFFFFFFFFFFFFFFFFFFFF0F0F0E0D0D0B8
        B0806050BAADA8FF00FFFF00FFFF00FFFF00FFFF00FFD8CBBCFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFD0B8B0FF00FFBAADA8FF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFD0C0B0FFFFFFF0F0F0D0C8C0D0C8C0D0B8B0C0B0B0E0D8
        D0857767806050FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFD0C0B0FF00FFFF
        00FFD0C8C0D0C8C0D0B8B0C0B0B0E0D8D0857767806050FF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFE0D0D0E0D0D0FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF0E8
        E0D0B8B0806050FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFE0D0D0E0D0D0FF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFD0B8B0806050FF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFDED5D3D0C0B0D0B8B0D0C0B0D0C0B0D0B8B0D0C0
        B0D0B8B0D0C0B0FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFDED5D3D0
        C0B0D0B8B0D0C0B0D0C0B0D0B8B0D0C0B0D0B8B0D0C0B0FF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF}
      NumGlyphs = 2
    end
    object btnDelete: TcxButton
      Left = 639
      Top = 27
      Width = 85
      Height = 27
      Cancel = True
      Caption = #21034#38500
      TabOrder = 7
      Visible = False
      OnClick = btnDeleteClick
      Glyph.Data = {
        36060000424D3606000000000000360000002800000020000000100000000100
        18000000000000060000F00A0000F00A00000000000000000000FF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF0205C3
        0A14AF867EB8D4C8C32E2DB32526B8B7ACC3A588799B7C6B9B7C6B9A7C6A997B
        689B7C6BFF00FFFF00FFFF00FFB79787B79787B79787D4C8C3B79787B79787B7
        9787A588799B7C6B9B7C6B9A7C6A997B689B7C6BFF00FFFF00FFFF00FF0203AD
        2C72FF1534D42135C4174EFF155CFF363AC5FDF9F3FAF2E6FAF1E4F9EFE0F8ED
        DA977967FF00FFFF00FFFF00FFB79787FF00FFB79787B79787FF00FFFF00FFB7
        9787FF00FFFF00FFFF00FFFF00FFFF00FF977967FF00FFFF00FFFF00FF0000C4
        1225C52C67FF255DFF205BFF141FBAD3C9D8DFC3B5DCBBA9DCBAA5DCBAA3FAF1
        E2987968FF00FFFF00FFFF00FFB79787B79787FF00FFFF00FFFF00FFB79787D3
        C9D8DFC3B5DCBBA9DCBAA5DCBAA3FF00FF987968FF00FFFF00FFFF00FFFF00FF
        A79EC11022BF2D66FF1C49F47471C4ECDDD6DBBDAFDABBAADAB8A5DCB9A5FAF3
        E6997A6AFF00FFFF00FFFF00FFFF00FFA79EC1B79787FF00FFB79787B79787FF
        00FFDBBDAFDABBAADAB8A5DCB9A5FF00FF997A6AFF00FFFF00FFFF00FFFF00FF
        2725A63F7DFF1C3FDD2961FF1223C4D2CADADEC4B8DBBCADDABAA8DCBBA7FBF4
        E8997B6BFF00FFFF00FFFF00FFFF00FFB79787FF00FFFF00FFFF00FFB79787FF
        00FFDEC4B8DBBCADDABAA8DCBBA7FF00FF997B6BFF00FFFF00FFFF00FFFF00FF
        151CC0396DFF8683CC3738B42B6DFF292ABDEEE2DCDBBDAFDBBAA9DCBCA9FBF5
        EA9A7C6BFF00FFFF00FFFF00FFFF00FFB79787FF00FFB79787B79787FF00FFB7
        9787FF00FFDBBDAFDBBAA9DCBCA9FF00FF9A7C6BFF00FFFF00FFFF00FFFF00FF
        A59BCA4242C4F0E6E4CAC3DA3231BBBDB8E5E7D4CCDBBDAFDBBAAADDBBA9FBF5
        EC9B7C6BFF00FFFF00FFFF00FFFF00FFA59BCAB79787FF00FFFF00FFB79787FF
        00FFFF00FFDBBDAFDBBAAADDBBA9FF00FF9B7C6BFF00FFFF00FFFF00FFFF00FF
        BCA194FFFFFFFEFEFEFEFEFEFFFEFEFEFDFCFDFBF8FDFAF6FCF9F3FCF7F0FCF6
        ED9B7C6CFF00FFFF00FFFF00FFFF00FFBCA194FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FF9B7C6CFF00FFFF00FFFF00FFFF00FF
        AF8F80FFFFFFDFCECCDFCDCBDFCBC7DEC6C0DEC4BADEC1B4DEBEADDEBEABFCF6
        EE9C7C6DFF00FFFF00FFFF00FFFF00FFAF8F80FF00FFDFCECCDFCDCBDFCBC7DE
        C6C0DEC4BADEC1B4DEBEADDEBEABFF00FF9C7C6DFF00FFFF00FFFF00FFFF00FF
        B19080FFFFFFFFFFFFFFFFFFFFFFFFFEFEFEFEFDFBFDFBF8FCF9F4F9F4EEF0E8
        E09F8070FF00FFFF00FFFF00FFFF00FFB19080FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FF9F8070FF00FFFF00FFFF00FFFF00FF
        B79787FFFFFFFFFFFFFFFFFFFFFFFFFEFEFEFEFDFBFDFBF9A78270A78270A782
        70A78270FF00FFFF00FFFF00FFFF00FFB79787FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFA78270A78270A78270A78270FF00FFFF00FFFF00FFFF00FF
        B89888FFFFFFFFFFFFFFFFFFFEFEFEFEFEFDFEFCFAFDFBF9A78270F5E2D9B18E
        7EA78270FF00FFFF00FFFF00FFFF00FFB89888FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFA78270FF00FFB18E7EA78270FF00FFFF00FFFF00FFFF00FF
        B89888FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFA78270B18E7EA782
        70FF00FFFF00FFFF00FFFF00FFFF00FFB89888FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFA78270B18E7EA78270FF00FFFF00FFFF00FFFF00FFFF00FF
        B89888B89888B49383B49383B08E7DB08E7DAC8877AC8877A78270A78270FF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFB89888B89888B49383B49383B08E7DB0
        8E7DAC8877AC8877A78270A78270FF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF}
      NumGlyphs = 2
    end
  end
  object Panel1: TPanel
    Left = 0
    Top = 42
    Width = 801
    Height = 407
    Align = alClient
    BevelOuter = bvNone
    BorderWidth = 2
    TabOrder = 1
    object gvGrid: TcxGrid
      Left = 2
      Top = 2
      Width = 797
      Height = 403
      Hint = #28369#40736#24038#37749#40670'2'#19979', '#36914#20837' "'#25240#35731#36039#26009#26126#32048'"'#30059#38754
      Align = alClient
      TabOrder = 0
      object gvView: TcxGridDBBandedTableView
        OnDblClick = gvViewDblClick
        NavigatorButtons.ConfirmDelete = False
        DataController.DataSource = dsMaster
        DataController.KeyFieldNames = 'IDENTIFYID1;IDENTIFYID2;COMPID;ALLOWANCENO'
        DataController.Summary.DefaultGroupSummaryItems = <>
        DataController.Summary.FooterSummaryItems = <>
        DataController.Summary.SummaryGroups = <>
        OptionsCustomize.ColumnFiltering = False
        OptionsCustomize.ColumnHiding = True
        OptionsCustomize.ColumnSorting = False
        OptionsCustomize.BandHiding = True
        OptionsData.Deleting = False
        OptionsData.DeletingConfirmation = False
        OptionsData.Editing = False
        OptionsData.Inserting = False
        OptionsSelection.CellSelect = False
        OptionsView.CellEndEllipsis = True
        OptionsView.GroupByBox = False
        OptionsView.HeaderEndEllipsis = True
        OptionsView.Indicator = True
        OptionsView.RowSeparatorWidth = 5
        Styles.Background = dtmMain.cxGridBackGroundStyle
        Styles.Content = dtmMain.cxGridBackGroundStyle
        Styles.Inactive = dtmMain.cxGridInActiveStyle
        Styles.Selection = dtmMain.cxGridActiveStyle
        Bands = <
          item
            Width = 535
          end
          item
            Width = 398
          end
          item
            Width = 77
          end>
        object gvViewCOMPSNAME: TcxGridDBBandedColumn
          Caption = #20844#21496#21029
          DataBinding.FieldName = 'COMPSNAME'
          Width = 85
          Position.BandIndex = 0
          Position.ColIndex = 0
          Position.RowIndex = 0
        end
        object gvViewCUSTID: TcxGridDBBandedColumn
          Caption = #23458#32232
          DataBinding.FieldName = 'CUSTID'
          Width = 78
          Position.BandIndex = 0
          Position.ColIndex = 1
          Position.RowIndex = 1
        end
        object gvViewCUSTSNAME: TcxGridDBBandedColumn
          Caption = #31777#31281
          DataBinding.FieldName = 'CUSTNAME'
          Width = 257
          Position.BandIndex = 0
          Position.ColIndex = 2
          Position.RowIndex = 1
        end
        object gvViewPAPERNO: TcxGridDBBandedColumn
          Caption = #21934#25818#32232#34399
          DataBinding.FieldName = 'PAPERNO'
          Width = 112
          Position.BandIndex = 0
          Position.ColIndex = 2
          Position.RowIndex = 0
        end
        object gvViewPAPERDATE: TcxGridDBBandedColumn
          Caption = #21934#25818#26085#26399
          DataBinding.FieldName = 'PAPERDATE'
          Width = 82
          Position.BandIndex = 0
          Position.ColIndex = 3
          Position.RowIndex = 0
        end
        object gvViewBUSINESSID: TcxGridDBBandedColumn
          Caption = #32113#32232
          DataBinding.FieldName = 'BUSINESSID'
          Width = 121
          Position.BandIndex = 0
          Position.ColIndex = 3
          Position.RowIndex = 1
        end
        object gvViewYEARMONTH2: TcxGridDBBandedColumn
          Caption = #30003#22577#24180#26376
          DataBinding.FieldName = 'YEARMONTH2'
          Width = 88
          Position.BandIndex = 0
          Position.ColIndex = 4
          Position.RowIndex = 0
        end
        object gvViewINVID: TcxGridDBBandedColumn
          Caption = #30332#31080#34399#30908
          DataBinding.FieldName = 'INVID'
          Width = 90
          Position.BandIndex = 1
          Position.ColIndex = 0
          Position.RowIndex = 0
        end
        object gvViewINVFORMAT: TcxGridDBBandedColumn
          Caption = #26684#24335
          DataBinding.FieldName = 'INVFORMATDESC'
          Width = 59
          Position.BandIndex = 1
          Position.ColIndex = 2
          Position.RowIndex = 0
        end
        object gvViewINVDATE: TcxGridDBBandedColumn
          Caption = #30332#31080#26085#26399
          DataBinding.FieldName = 'INVDATE'
          Width = 81
          Position.BandIndex = 1
          Position.ColIndex = 1
          Position.RowIndex = 0
        end
        object gvViewTAXTYPE: TcxGridDBBandedColumn
          Caption = #31237#21029
          DataBinding.FieldName = 'TAXTYPEDESC'
          Width = 71
          Position.BandIndex = 1
          Position.ColIndex = 4
          Position.RowIndex = 0
        end
        object gvViewSALEAMOUNT: TcxGridDBBandedColumn
          Caption = #25240#35731#37559#21806#38989
          DataBinding.FieldName = 'SALEAMOUNT'
          PropertiesClassName = 'TcxCurrencyEditProperties'
          Properties.DecimalPlaces = 0
          Properties.DisplayFormat = ',0;-$,0'
          Properties.NullString = '0'
          HeaderAlignmentHorz = taRightJustify
          Width = 103
          Position.BandIndex = 1
          Position.ColIndex = 0
          Position.RowIndex = 1
        end
        object gvViewTAXAMOUNT: TcxGridDBBandedColumn
          Caption = #25240#35731#31237#38989
          DataBinding.FieldName = 'TAXAMOUNT'
          PropertiesClassName = 'TcxCurrencyEditProperties'
          Properties.DecimalPlaces = 0
          Properties.DisplayFormat = ',0;-$,0'
          Properties.NullString = '0'
          HeaderAlignmentHorz = taRightJustify
          Width = 112
          Position.BandIndex = 1
          Position.ColIndex = 1
          Position.RowIndex = 1
        end
        object gvViewINVAMOUNT: TcxGridDBBandedColumn
          Caption = #25240#35731#30332#31080#38989
          DataBinding.FieldName = 'INVAMOUNT'
          PropertiesClassName = 'TcxCurrencyEditProperties'
          Properties.DecimalPlaces = 0
          Properties.DisplayFormat = ',0;-$,0'
          Properties.NullString = '0'
          HeaderAlignmentHorz = taRightJustify
          Width = 110
          Position.BandIndex = 1
          Position.ColIndex = 2
          Position.RowIndex = 1
        end
        object gvViewUPTEN: TcxGridDBBandedColumn
          Caption = #26356#26032#20154#21729
          DataBinding.FieldName = 'UPTEN'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Vert = taVCenter
          HeaderAlignmentHorz = taCenter
          HeaderAlignmentVert = vaCenter
          Position.BandIndex = 2
          Position.ColIndex = 0
          Position.LineCount = 2
          Position.RowIndex = 0
        end
        object gvViewALLOWANCENO: TcxGridDBBandedColumn
          Caption = #25240#35731#21934#34399
          DataBinding.FieldName = 'ALLOWANCENO'
          Width = 110
          Position.BandIndex = 0
          Position.ColIndex = 1
          Position.RowIndex = 0
        end
        object gvViewSOURCE: TcxGridDBBandedColumn
          Caption = #21934#25818#20358#28304
          DataBinding.FieldName = 'SOURCE2'
          Width = 99
          Position.BandIndex = 0
          Position.ColIndex = 0
          Position.RowIndex = 1
        end
        object gvViewKind: TcxGridDBBandedColumn
          Caption = #31278#39006
          DataBinding.FieldName = 'INVOICEKIND'
          Width = 97
          Position.BandIndex = 1
          Position.ColIndex = 3
          Position.RowIndex = 0
        end
        object gvViewUploadFlag: TcxGridDBBandedColumn
          Caption = #19978#20659#35387#35352
          DataBinding.FieldName = 'UPLOADFLAG'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taCenter
          Properties.Alignment.Vert = taVCenter
          Width = 58
          Position.BandIndex = 0
          Position.ColIndex = 5
          Position.RowIndex = 0
        end
      end
      object gvLevel1: TcxGridLevel
        GridView = gvView
      end
    end
  end
  object dsMaster: TDataSource
    DataSet = cdsMaster
    Left = 228
    Top = 352
  end
  object cdsMaster: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'DataSetProvider1'
    AfterScroll = cdsMasterAfterScroll
    OnNewRecord = cdsMasterNewRecord
    Left = 268
    Top = 352
  end
end
