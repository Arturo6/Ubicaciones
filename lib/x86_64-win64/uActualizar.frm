object fActualizar: TfActualizar
  Left = 344
  Height = 124
  Top = 256
  Width = 324
  BorderIcons = []
  BorderStyle = bsSingle
  Caption = 'Actualizando'
  ClientHeight = 124
  ClientWidth = 324
  OnActivate = FormActivate
  Position = poScreenCenter
  LCLVersion = '7.5'
  object lblActualizar: TLabel
    Left = 23
    Height = 25
    Top = 17
    Width = 278
    Caption = 'Actualizando las bases de datos'
    Font.Color = clBlue
    Font.Height = -19
    Font.Style = [fsBold]
    ParentFont = False
  end
  object lblEsperar: TLabel
    Left = 27
    Height = 25
    Top = 54
    Width = 270
    Caption = 'Por favor, espera un momento'
    Font.Color = clBlue
    Font.Height = -19
    Font.Style = [fsBold]
    ParentFont = False
  end
  object pbActualizar: TProgressBar
    Left = 23
    Height = 13
    Top = 96
    Width = 278
    Smooth = True
    TabOrder = 0
    BarShowText = True
  end
end
