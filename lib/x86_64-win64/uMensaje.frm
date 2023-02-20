object fMensaje: TfMensaje
  Left = 344
  Height = 50
  Top = 256
  Width = 393
  BorderIcons = []
  BorderStyle = bsNone
  ClientHeight = 50
  ClientWidth = 393
  FormStyle = fsStayOnTop
  Position = poScreenCenter
  LCLVersion = '7.5'
  object Panel: TPanel
    Left = 0
    Height = 50
    Top = 0
    Width = 393
    Align = alClient
    AutoSize = True
    BevelInner = bvLowered
    BorderStyle = bsSingle
    Font.Color = clRed
    Font.Height = -19
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 0
  end
  object Reloj: TTimer
    Enabled = False
    OnTimer = RelojTimer
    Left = 229
    Top = 16
  end
end
