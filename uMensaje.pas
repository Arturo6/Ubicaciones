{
  uMensaje: unidad que muestra un mensaje en pantalla durante un tiempo especificado.

  Versión: 0.1 17/05/2022
  Copyright (C) 2022 <Arturo Molina> <amolinaj@gmx.es>

  Parámetros:
    Mensaje = Mensaje a mostrar en la ventana
	Tiempo  = Tiempo en segundo que permanece en pantalla
	
  Uso:
    - Agregar la unidad a la cláusula uses de proyecto
	- Agregar una variable del tipo ventana:TfMensaje en la sección var del formulario principal
	- Crear una variable general del tipo TfMensaje en el evento Create de la aplicación
	  ventana := TfMensaje.Create( Application );

  Ejemplo:
    ventana.Mensaje( 'Saludos', 5 );  
	
  Este código es software libre; puede redistribuirlo y/o modificarlo bajo los términos de la
  licencia publicada por la Free Software Foundation; ya sea la versión 3 o cualquier versión posterior.

  Este código se distribuye con la esperanza de que sea útil, pero SIN NINGUNA GARANTÍA; sin siquiera la
  garantía de COMERCIABILIDAD o IDONEIDAD PARA UN PROPÓSITO PARTICULAR.
  Consulte la Licencia pública general de GNU para obtener más información y/o detalles.

  Una copia de la Licencia Pública General GNU está disponible en la página web
  <http://www.gnu.org/copyleft/gpl.html>. También puede obtenerlo escribiendo a la Free Software Foundation, Inc., 51
  Franklin Street - Fifth Floor, Boston, MA 02110-1335, USA.
}

unit uMensaje;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, Forms, ExtCtrls;

type

  { TfMensaje }

  TfMensaje = class(TForm)
    Panel: TPanel;
    Reloj: TTimer;
    procedure RelojTimer( Sender: TObject );
  public
    procedure Mensaje( sMensaje: string; Tiempo: integer );
  end;

var
  fMensaje: TfMensaje;

implementation

{$R *.frm}

{ TfMensaje }

procedure TfMensaje.Mensaje( sMensaje: string; Tiempo: integer );
begin
  fMensaje.Width := Length( sMensaje ) * 10;
  Panel.Caption  := sMensaje;
  Reloj.Interval := Tiempo * 1000;

  Reloj.Enabled := true;
end;

procedure TfMensaje.RelojTimer(Sender: TObject);
begin
  Reloj.Enabled:=false;
  fMensaje.Close;
end;

end.

