unit uLicencia;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, Buttons, ExtCtrls;

type
{ TfLicencia }
  TfLicencia = class ( TForm )
    btnCerrar: TBitBtn;
    lblTextoGPL: TLabel;
    Memo: TMemo;
    pnlScrool: TPanel;
    Reloj: TTimer;
    procedure btnCerrarClick(Sender: TObject);
    procedure RelojTimer(Sender: TObject);
  private

  public

  end;

var
   fLicencia: TfLicencia;

implementation

{$R *.frm}

{ TfLicencia }

procedure TfLicencia.RelojTimer(Sender: TObject);
begin
  if lblTextoGPL.BoundsRect.Bottom <= 0 then begin //reset position
    lblTextoGPL.Top := pnlScrool.Height;
  end else begin // scroooolll!
    lblTextoGPL.Top := lblTextoGPL.Top - 1;
  end;
end;

procedure TfLicencia.btnCerrarClick(Sender: TObject);
begin
  Close;
end;

end.