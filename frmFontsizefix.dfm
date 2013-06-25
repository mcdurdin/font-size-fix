object Form1: TForm1
  Left = 192
  Top = 114
  Caption = 'Font Size Fix'
  ClientHeight = 505
  ClientWidth = 897
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Button1: TButton
    Left = 12
    Top = 12
    Width = 93
    Height = 25
    Caption = 'Switch to 100%'
    TabOrder = 0
    OnClick = Button1Click
  end
  object Memo1: TMemo
    Left = 12
    Top = 44
    Width = 881
    Height = 453
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Lucida Console'
    Font.Style = []
    Lines.Strings = (
      
        'This program attempts to make the font scaling information on th' +
        'e '
      
        'system consistent in the event that it has gone out of sync.  It' +
        ' will'
      
        'correct current user and local machine registry settings, and in' +
        'stall '
      'the correct Windows bitmap fonts for the selected font scale.'
      ''
      
        'http://marc.durdin.net/2011/07/fixing-windows-font-scaling-witho' +
        'ut.html'
      ''
      '*** No support is available, use at your own risk! ***'
      '')
    ParentFont = False
    ReadOnly = True
    ScrollBars = ssBoth
    TabOrder = 1
    WordWrap = False
  end
  object Button2: TButton
    Left = 112
    Top = 12
    Width = 93
    Height = 25
    Caption = 'Switch to 125%'
    TabOrder = 2
    OnClick = Button2Click
  end
end
