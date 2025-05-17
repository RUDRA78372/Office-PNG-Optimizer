object Form1: TForm1
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'Office PNG Optimizer'
  ClientHeight = 180
  ClientWidth = 624
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  OnCreate = FormCreate
  TextHeight = 15
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 624
    Height = 180
    Align = alClient
    DragMode = dmAutomatic
    TabOrder = 0
    object Label1: TLabel
      Left = 8
      Top = 32
      Width = 60
      Height = 15
      Caption = 'Source File:'
    end
    object Label2: TLabel
      Left = 8
      Top = 77
      Width = 84
      Height = 15
      Caption = 'Destination File:'
    end
    object Label3: TLabel
      Left = 8
      Top = 125
      Width = 85
      Height = 15
      Caption = 'Command Line:'
    end
    object Label4: TLabel
      Left = 1
      Top = 1
      Width = 622
      Height = 29
      Align = alTop
      Alignment = taCenter
      Caption = 'Office PNG Optimizer'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -24
      Font.Name = 'Calibri'
      Font.Style = []
      ParentFont = False
      ExplicitWidth = 205
    end
    object Label5: TLabel
      Left = 8
      Top = 11
      Width = 73
      Height = 15
      Caption = 'Author: 78372'
    end
    object Edit1: TEdit
      Left = 8
      Top = 48
      Width = 521
      Height = 23
      TabOrder = 0
    end
    object Button1: TButton
      Left = 542
      Top = 47
      Width = 75
      Height = 25
      Caption = 'Browse'
      TabOrder = 1
      OnClick = Button1Click
    end
    object Button2: TButton
      Left = 542
      Top = 95
      Width = 75
      Height = 25
      Caption = 'Browse'
      TabOrder = 2
      OnClick = Button2Click
    end
    object Edit2: TEdit
      Left = 8
      Top = 96
      Width = 521
      Height = 23
      TabOrder = 3
    end
    object Edit3: TEdit
      Left = 8
      Top = 146
      Width = 521
      Height = 23
      TabOrder = 4
      Text = '--quality=45-85 '
    end
    object Button3: TButton
      Left = 542
      Top = 144
      Width = 75
      Height = 25
      Caption = 'Optimize'
      TabOrder = 5
      OnClick = Button3Click
    end
  end
  object OpenDialog1: TOpenDialog
    Left = 592
    Top = 8
  end
  object SaveDialog1: TSaveDialog
    Left = 552
    Top = 8
  end
  object DropFileTarget1: TDropFileTarget
    DragTypes = [dtCopy, dtLink]
    OnDrop = DropFileTarget1Drop
    Target = Panel1
    OptimizedMove = True
    Left = 504
    Top = 8
  end
end
