object frmPrincipal: TfrmPrincipal
  Left = 558
  Top = 189
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'Relat'#243'rios Personalizados (v4.27.7) - '
  ClientHeight = 535
  ClientWidth = 479
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label3: TLabel
    Left = 8
    Top = 102
    Width = 116
    Height = 13
    Caption = 'Relat'#243'rio Personalizados'
  end
  object lbl_Usuario: TLabel
    Left = 160
    Top = 104
    Width = 52
    Height = 13
    Caption = 'lbl_Usuario'
    Visible = False
  end
  object Label2: TLabel
    Left = 8
    Top = 366
    Width = 108
    Height = 13
    Caption = 'Descri'#231#227'o do Relat'#243'rio'
  end
  object lblExporta: TLabel
    Left = 440
    Top = 104
    Width = 6
    Height = 13
    Caption = 'F'
    Visible = False
  end
  object lblImprime: TLabel
    Left = 400
    Top = 104
    Width = 6
    Height = 13
    Caption = 'F'
    Visible = False
  end
  object lblTitulo: TLabel
    Left = 136
    Top = 104
    Width = 7
    Height = 13
    Caption = 'A'
    Visible = False
  end
  object lblCodRelatorio: TLabel
    Left = 336
    Top = 104
    Width = 6
    Height = 13
    Caption = '0'
    Visible = False
  end
  object dbgRelatorios: TDBGrid
    Left = 8
    Top = 120
    Width = 465
    Height = 241
    DataSource = dsRelatorios
    ReadOnly = True
    TabOrder = 0
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    OnCellClick = dbgRelatoriosCellClick
    OnDrawColumnCell = dbgRelatoriosDrawColumnCell
    OnDblClick = dbgRelatoriosDblClick
    OnKeyDown = dbgRelatoriosKeyDown
    OnKeyUp = dbgRelatoriosKeyUp
    Columns = <
      item
        Expanded = False
        FieldName = 'ID'
        Title.Caption = 'C'#243'digo'
        Width = 53
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'TITULO'
        Title.Caption = 'T'#237'tulo'
        Width = 374
        Visible = True
      end>
  end
  object GroupBox1: TGroupBox
    Left = 8
    Top = 8
    Width = 465
    Height = 89
    Caption = ' Filtrar Relat'#243'rio '
    TabOrder = 1
    object Label1: TLabel
      Left = 104
      Top = 27
      Width = 94
      Height = 13
      Caption = 'Par'#226'metro para filtro'
    end
    object SpeedButton1: TSpeedButton
      Left = 376
      Top = 24
      Width = 73
      Height = 49
      Caption = '&Filtrar'
      Glyph.Data = {
        42040000424D4204000000000000420000002800000010000000100000000100
        200003000000000400006F0000006F00000000000000000000000000FF0000FF
        0000FF000000FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF000000
        001C000000A7FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF000000001D0000
        00B90000001CFFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF000000001D000000BA0000
        001DFFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF000000001E000000BC0000001DFFFF
        FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF0000000018000000480000
        002A00000002FFFFFF00FFFFFF000000001F000000BD0000001EFFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF000000003B000000AF00000091000000650000
        0081000000AC0000007400000021000000BE0000001FFFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00000000690000008D00000008FFFFFF00FFFFFF00FFFF
        FF00FFFFFF0000000054000000DA00000022FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF000000002A000000A0FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF000000005400000075FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF000000009F00000014FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF00000000AC00000002FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00000000A9FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF00000000810000002AFFFFFF00FFFFFF00FFFF
        FF00FFFFFF00000000A8FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF000000006500000048FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00000000AAFFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF000000009100000018FFFFFF00FFFFFF00FFFF
        FF00FFFFFF000000008400000039FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF0000000008000000B0FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF000000000B000000B500000009FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF000000008E0000003FFFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF0000000032000000B500000039FFFFFF00FFFFFF00FFFF
        FF0000000014000000A10000006FFFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00FFFFFF00FFFFFF000000000B00000084000000AA000000A80000
        00A90000009F0000002BFFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
        FF00FFFFFF00}
      OnClick = SpeedButton1Click
    end
    object edtFiltroPesquisa: TEdit
      Left = 104
      Top = 43
      Width = 257
      Height = 21
      TabOrder = 0
      OnKeyPress = edtFiltroPesquisaKeyPress
    end
    object rbtFiltroTitulo: TRadioButton
      Left = 7
      Top = 44
      Width = 65
      Height = 17
      Caption = 'T'#237'tulo'
      TabOrder = 1
      OnClick = rbtFiltroTituloClick
    end
    object rbtFiltroChave: TRadioButton
      Left = 7
      Top = 68
      Width = 97
      Height = 17
      Caption = 'Palavra Chave'
      TabOrder = 2
      OnClick = rbtFiltroChaveClick
    end
    object rbtFiltroCodigo: TRadioButton
      Left = 7
      Top = 20
      Width = 65
      Height = 17
      Caption = 'C'#243'digo'
      TabOrder = 3
      OnClick = rbtFiltroCodigoClick
    end
  end
  object Button1: TButton
    Left = 8
    Top = 504
    Width = 113
    Height = 25
    Caption = '&Abrir'
    TabOrder = 2
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 360
    Top = 504
    Width = 113
    Height = 25
    Caption = 'F&echar'
    TabOrder = 3
    OnClick = Button2Click
  end
  object memDescricao: TMemo
    Left = 8
    Top = 384
    Width = 465
    Height = 113
    Color = clScrollBar
    Lines.Strings = (
      '')
    ReadOnly = True
    ScrollBars = ssVertical
    TabOrder = 4
  end
  object memSQL: TMemo
    Left = 264
    Top = 499
    Width = 26
    Height = 33
    Lines.Strings = (
      'me'
      'mS'
      'QL')
    TabOrder = 5
    Visible = False
  end
  object memSQL2: TMemo
    Left = 292
    Top = 499
    Width = 26
    Height = 33
    Lines.Strings = (
      'me'
      'mS'
      'QL')
    TabOrder = 6
    Visible = False
  end
  object qryRelatorios: TADOQuery
    Connection = DM.conTI
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      ''
      'SELECT r.*, p.usuario'
      'FROM ti.relatorios r, ti.relatorios_permissao p'
      'WHERE r.id = p.id_relatorio')
    Left = 128
    Top = 504
    object qryRelatoriosID: TIntegerField
      FieldName = 'ID'
    end
    object qryRelatoriosTITULO: TWideStringField
      FieldName = 'TITULO'
      Size = 80
    end
    object qryRelatoriosCHAVE: TWideStringField
      FieldName = 'CHAVE'
      Size = 150
    end
    object qryRelatoriosFORMULARIO: TWideStringField
      FieldName = 'FORMULARIO'
      Size = 40
    end
    object qryRelatoriosUSUARIO: TWideStringField
      FieldName = 'USUARIO'
      Size = 40
    end
    object qryRelatoriosDESCRICAO: TWideStringField
      FieldName = 'DESCRICAO'
      Size = 800
    end
  end
  object dsRelatorios: TDataSource
    DataSet = qryRelatorios
    Left = 160
    Top = 504
  end
  object qryExec: TADOQuery
    Connection = DM.conTI
    Parameters = <>
    Left = 320
    Top = 504
  end
  object XPManifest1: TXPManifest
    Left = 240
    Top = 96
  end
  object ApplicationEvents1: TApplicationEvents
    OnMessage = ApplicationEvents1Message
    Left = 272
    Top = 96
  end
end
