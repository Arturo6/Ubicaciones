{
  Copyright (C) 2020 Arturo Molina amolinaj@gmail.com

  UBICACIONES es un programa desarrollado para acceder al catálogo de libros
  editados por el servicio de publicaciones de la Universidad de Alicante y
  facilitar el trabajo de localización de títulos, ubicaciones, portadas y cantidades
  de los ejemplares existentes en el sistema de almacenamiendo de dicho servicio.

  Este programa es Software Libre; Usted puede redistribuirlo y/o modificarlo
  bajo los términos de la "GNU General Public License (GPL)" tal y como ha sido
  públicada por la Free Software Foundation; o bien la versión 3 de la Licencia,
  o (a su opción) cualquier versión posterior.

  Este programa se distribuye con la esperanza de que sea útil, pero SIN NINGUNA
  GARANTÍA; tampoco las implícitas garantías de MERCANTILIDAD o ADECUACIÓN A UN
  PROPÓSITO PARTICULAR. Consulte la "GNU General Public License (GPL)" para más
  detalles. Usted debe recibir una copia de la GNU General Public License (GPL)
  junto con este programa; si no, escriba a la Free Software Foundation Inc.
  51 Franklin Street, 5º Piso, Boston, MA 02110-1301, USA.
}

unit uPrincipal;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, DB, Forms, Controls, Graphics, Dialogs, ExtCtrls, ComCtrls,
  windows, Buttons, StdCtrls, RichMemo, FZCommon, FZBase, ZDataset, ZConnection,
  Grids, DBGrids, LCLType, LCLIntf, DateUtils, DefaultTranslator, UniqueInstance,
  uLicencia, uActualizar, IniFiles;

type

  { TfUbicaciones }

  TfUbicaciones = class(TForm)
    btnAyuda: TSpeedButton;
    btnBuscar: TSpeedButton;
    btnPrincipal: TSpeedButton;
    lblPrecio: TLabel;
    lblPrecioDB: TLabel;
    lblLicenciaGNU: TLabel;
    lblAutor: TLabel;
    lblAutorDB: TLabel;
    lblSubitulo: TLabel;
    lblSubtituloDB: TLabel;
    LogoPublicaciones: TImage;
    Abrir: TOpenDialog;
    LogoPublicaciones1: TImage;
    PanelLogo: TPanel;
    Rejilla: TDBGrid;
    DS: TDataSource;
    edtTitulo: TEdit;
    lblTituloBuscar: TLabel;
    Paginas: TFZPageControl;
    tsPrincipal: TFZVirtualPage;
    tsBuscar: TFZVirtualPage;
    tsAyuda: TFZVirtualPage;
    edtISBN: TEdit;
    lblCantidad: TLabel;
    lblCantidadDB: TLabel;
    lblISBN: TLabel;
    lblTitulo: TLabel;
    lblTituloDB: TLabel;
    lblUbicacion: TLabel;
    lblUbicacionDB: TLabel;
    Portada: TImage;
    pnlBotones: TPanel;
    btnCerrar: TSpeedButton;
    txtAyuda: TRichMemo;
    Query: TZQuery;
    Conexion: TZConnection;
    UniqueInstance: TUniqueInstance;
    procedure btnAyudaClick(Sender: TObject);
    procedure btnBuscarClick(Sender: TObject);
    procedure btnCerrarClick(Sender: TObject);
    procedure btnPrincipalClick(Sender: TObject);
    procedure edtISBNKeyPress ( Sender: TObject; var Key: char ) ;
    procedure edtTituloKeyPress ( Sender: TObject; var Key: char ) ;
    procedure FormActivate(Sender: TObject);
    procedure FormCreate ( Sender: TObject ) ;
    procedure FormKeyDown ( Sender: TObject; var Key: Word; Shift: TShiftState ) ;
    procedure lblAutorDBClick ( Sender: TObject ) ;
    procedure lblLicenciaGNUClick ( Sender: TObject ) ;
    procedure lblTituloDBClick ( Sender: TObject ) ;
    procedure LogoPublicaciones1DblClick(Sender: TObject);
    procedure LogoPublicacionesDblClick(Sender: TObject);
    procedure RejillaCellClick ( Column: TColumn ) ;
    procedure RejillaPrepareCanvas ( sender: TObject; DataCol: Integer; Column: TColumn; AState: TGridDrawState ) ;
    function BuscarArchivo(): String;
  private
    procedure BuscaPortada(Archivo: string);
    procedure DBGridOnGetText(Sender: TField; var aText: string; DisplayText: Boolean);
    function  GetFileTime(const AFileName: String; var FileTime: TDateTime): Boolean;
    function CrearBaseDatos( const Fichero: String ): Boolean;
  public

  end;

var
  fUbicaciones: TfUbicaciones;
  Lugar: String;

resourcestring
  ErrorAlConectar = 'Error al conectar la base de datos';
  ActualizarBase = 'La base de datos podría no estar actualizada';
  PSActualizarBase = 'Por favor, actualiza la base de datos de libros';
  Procesados = ' registros procesados';
  Modificados = ' registros modificados';

implementation

{$R *.frm}

{ TfUbicaciones }

procedure TfUbicaciones.FormCreate ( Sender: TObject ) ;
begin
  if not FileExists( ExtractFilePath( Application.ExeName ) + 'lista.db' ) then
    if not CrearBaseDatos( ExtractFilePath( Application.ExeName ) + 'lista.db' ) then begin
      MessageDlg('Hubo un error al crear la base de datos'+LineEnding+'', mtError, [],0 );
      Close;
    end else begin
      // Muestra la pantalla de actualizar
      with TfActualizar.Create(nil) do
        try
          ShowModal
        finally
          Free
        end;

      MessageDlg('Base de datos actualizada'+LineEnding+'', mtInformation, [],0 );
    end;

  // Conecta con la base de datos
  Conexion.Disconnect;
  Conexion.Database := ExtractFilePath( Application.ExeName ) + 'lista.db';
  Try
    Conexion.Connected := true;
  except
    MessageDlg( ErrorAlConectar+LineEnding+'', mtError, [],0 );
    Application.Terminate;
  end;
end;

procedure TfUbicaciones.FormActivate(Sender: TObject);
var
  Mydir: String;
  FileTime1, FileTime2: TDateTime;
  FicheroINI: TIniFile;
begin
  FicheroINI := TiniFile.Create( ChangeFileExt( Application.ExeName, '.ini' ) );
  // Muestra la página principal
  Paginas.ActivePage := tsPrincipal;
  // Y entrega el foco al ISBN
  edtISBN.SetFocus;

  // Busca la ubicacion del archivo de excel con la lista de libros
  if FicheroINI.ReadString( 'Archivos', 'Lista', '' ) = '' then begin
    fUbicaciones.Abrir.Title := 'Abrir una hoja de datos con LIBROS';
    MyDir := BuscarArchivo();
    if MyDir <> '' then begin
      FicheroINI.WriteString( 'Archivos', 'Lista', MyDir );
      FicheroINI.Free;
    end else begin
      FicheroINI.Free;
      Exit;
    end;
  end else
      MyDir := FicheroINI.ReadString( 'Archivos', 'Lista', '' );
//  MyDir := '\\172.16.204.63\Archivo\Lista.xls';
  if GetFileTime( MyDir, FileTime1 ) then begin
    GetFileTime( ExtractFilePath( Application.ExeName ) + 'lista.db', FileTime2 );
    if  (FileTime1 ) > ( FileTime2 ) then
       if MessageDlg('El archivo de datos está desactualizado'+#13+'¿deseas actualizarlo?', mtInformation, mbYesNo,0) = mrYes then begin
         // Muestra la pantalla de actualizar
         with TfActualizar.Create(nil) do
           try
             ShowModal;
             MessageDlg('Base de datos actualizada'+LineEnding+'', mtInformation, [],0 );
           finally
             Free
           end
       end;
  end;
end;


procedure TfUbicaciones.btnPrincipalClick(Sender: TObject);
begin
  // Muestra la página principal
  Paginas.ActivePage := tsPrincipal;
  // Y entrega el foco al ISBN
  edtISBN.SetFocus;
end;

procedure TfUbicaciones.btnBuscarClick(Sender: TObject);
begin
  // Muestra la página de búsquedas
  Paginas.ActivePage := tsBuscar;
  // Cierra la consulta
  Query.Close;
  // Limpia el campo de título y le da el foco
  edtTitulo.Clear;
  edtTitulo.SetFocus;
end;

procedure TfUbicaciones.btnAyudaClick(Sender: TObject);
begin
  // Muestra la página de ayuda
  Paginas.ActivePage := tsAyuda;
end;

procedure TfUbicaciones.btnCerrarClick(Sender: TObject);
begin
  // Cierra la conexion y el programa
  Conexion.Disconnect;
  Close;
end;

procedure TfUbicaciones.edtISBNKeyPress ( Sender: TObject; var Key: char ) ;
begin
  // Se ha pulsado RETURN en el ISBN
  if Key = #13 then begin
    if edtISBN.Text = '' then begin
      edtISBN.Text := '';
      lblAutorDB.Caption := '';
      lblTituloDB.Caption := '';
      lblUbicacionDB.Caption := '';
      lblCantidadDB.Caption := '';
      Portada.Picture.LoadFromFile( 'sin imagen.jpg' );
      exit;
	end;
	// Prepara la búsqueda del isbn
  Query.Active := False;
  Query.SQL.Text := 'SELECT * FROM libros WHERE isbn = :isbn1;';
  Query.ParamByName('isbn1').AsString := edtISBN.Text;
  Query.Open;

  // Si hay algún resultado...
  if Query.RecordCount > 0 then begin
    lblAutorDB.Caption := Query.FieldByName('autor').AsString;
    lblAutorDB.Hint := lblAutorDB.Caption;
    lblTituloDB.Caption := Query.FieldByName('titulo').AsString;
    lblTituloDB.Hint := lblTituloDB.Caption;
    lblSubtituloDB.Caption := Query.FieldByName('subtitulo').AsString;
    lblSubtituloDB.Hint := lblSubtituloDB.Caption;
    lblUbicacionDB.Caption := Query.FieldByName('ubicacion').AsString;
    if Query.FieldByName('cantidad').AsInteger = 1 then
      lblCantidadDB.Caption := IntToStr( Query.FieldByName('cantidad').AsInteger ) + ' unidad'
    else
      lblCantidadDB.Caption := IntToStr( Query.FieldByName('cantidad').AsInteger ) + ' unidades';
//    lblPrecioDB.Caption := FormatFloat ('###,##0.00 €;-###,##0.00 €;0', StrToFloat( Query.FieldByName('precio').AsString ) / 10000 );
    lblPrecioDB.Caption := FormatFloat ('###,##0.00 €;-###,##0.00 €;0', StrToFloat( Query.FieldByName('precio').AsString ) );

    // Muestra la portada del libro
    BuscaPortada( edtISBN.Text );
  end else
    MessageDlg('Ese ISBN no está en la base de datos'+LineEnding+'', mtInformation, [],0 );
end;

  // Toma el foco de nuevo
  edtISBN.SetFocus;
end;

procedure TfUbicaciones.edtTituloKeyPress ( Sender: TObject; var Key: char ) ;
begin
  // Se ha pulsado RETURN en el título
  if Key = #13 then
    if Copy( edtTitulo.Text, 1, 1 ) = '-' then begin
      edtTitulo.Text := Copy( edtTitulo.Text, 2, Length( edtTitulo.Text ) );
      Query.Active := False;
      Query.SQL.Text := 'SELECT * FROM libros WHERE autor LIKE ' + quotedstr( '%' + edtTitulo.Text + '%' ) + ';';
      Query.Open;
    end else begin
      // Prepara la búsqueda del isbn
      Query.Active := False;
      Query.SQL.Text := 'SELECT * FROM libros WHERE titulo LIKE ' + quotedstr( '%' + edtTitulo.Text + '%' ) + ';';
      Query.Open;
    end;
end;

procedure TfUbicaciones.FormKeyDown ( Sender: TObject; var Key: Word; Shift: TShiftState ) ;
begin
  // Se ha pulsado F5, mostramos la pantalla buscar titulo
  if Key = VK_F5 then begin
    Paginas.ActivePage := tsBuscar;
    edtTitulo.SetFocus;
  end;

  // Se ha pulsado ESCAPE, salimos del programa
  if Key = VK_ESCAPE then begin
    // Cierra la conexion
    Conexion.Disconnect;
    Close;
  end;

  // Se ha pulsado F1, se muestra la pantalla de ayuda
  if Key = VK_F1 then
    Paginas.ActivePage := tsAyuda;

  // Se ha pulsado F2, se actualizan las bases de datos
  if Key = VK_F2 then begin
    // Muestra la pantalla de actualizar
    with TfActualizar.Create(nil) do
      try
        ShowModal
      finally
        Free
      end;

    MessageDlg('Base de datos actualizada'+LineEnding+'', mtInformation, [],0 );
  end;
end;

procedure TfUbicaciones.RejillaPrepareCanvas ( sender: TObject; DataCol: Integer; Column: TColumn; AState: TGridDrawState ) ;
var
   MyTextStyle: TTextStyle;
begin
  // Elegimos la columna del DBGrid que sera afectada
  if (Datacol = 0) or (DataCol = 1) or (Datacol = 2) then begin
    // Lo siguiente no es necesario, pero puedes usarlo para ajustar la apariencia de tu texto.
    // puedes cambiar los colores, la fuente, el tamaño y otros parámetros.
    MyTextStyle := TDBGrid(Sender).Canvas.TextStyle;
    MyTextStyle.SingleLine := False;
    MyTextStyle.Wordbreak := False;
    TDBGrid(Sender).Canvas.TextStyle := MyTextStyle;

    // Aquí cómo mostrar cualquier texto:
    // simplemente asigne un procedimiento de evento a OnGetText del campo.
    Column.Field.OnGetText := @fUbicaciones.DBGridOnGetText;
  end;
end;

procedure TfUbicaciones.RejillaCellClick ( Column: TColumn ) ;
var
  Tecla: Char;
begin
  Tecla:=#13; // Caracter que representa el RETURN

  // Muestra la página principal
  Paginas.ActivePage := tsPrincipal;
  // Rescata el isbn del registro seleccionado
  edtISBN.Text := Query.FieldByName('isbn').AsString;

  // Busca el ISBN
  edtISBNKeyPress( self, Tecla );
end;

procedure TfUbicaciones.lblLicenciaGNUClick ( Sender: TObject ) ;
begin
  // Muestra la licencia GNU 2.0
  with TfLicencia.Create(nil) do
    try
      ShowModal
    finally
      Free
    end
end;

procedure TfUbicaciones.lblTituloDBClick ( Sender: TObject ) ;
begin
  // Muestra el título completo
  MessageDlg(lblTituloDB.Hint, mtInformation, [mbOK],0);
end;

procedure TfUbicaciones.lblAutorDBClick ( Sender: TObject ) ;
begin
  // Muestra el autor completo
  MessageDlg(lblAutorDB.Hint, mtInformation, [mbOK],0);
end;

procedure TfUbicaciones.LogoPublicacionesDblClick(Sender: TObject);
begin
  OpenURL('https://www.ua.es/');
end;

procedure TfUbicaciones.LogoPublicaciones1DblClick(Sender: TObject);
begin
  OpenURL('https://publicaciones.ua.es/');
end;

//***************************************************************//
//***************************************************************//
//                                                               //
//                                                               //
// Funciones auxiliares no vinculadas con el flujo del programa  //
//                                                               //
//                                                               //
//***************************************************************//
//***************************************************************//

function TfUbicaciones.BuscarArchivo(): String;
begin
  if Abrir.Execute then
   Result := Abrir.FileName
  else
    Result := '';
end;

procedure TfUbicaciones.BuscaPortada(Archivo: string);
begin
  // Intenta cargar el archivo grande de la portada
  try
     Portada.Picture.LoadFromFile( '\\172.16.204.63\Archivo\PORTADAS\' + Archivo + '_L38_04_g.jpg' );
  except
    // No ha podido cargar el archivo grande, así que intenta cargar el archivo mediano de la portada
    try
      Portada.Picture.LoadFromFile( '\\172.16.204.63\Archivo\PORTADAS\' + Archivo + '_L38_04_m.jpg' );
    except
      // No se ha podido localizar la portada, se usa la de sin imagen
      Portada.Picture.LoadFromFile( 'sin imagen.jpg' );
    end;
  end;
end;

procedure TfUbicaciones.DBGridOnGetText(Sender: TField; var aText: string; DisplayText: Boolean);
begin
  if (DisplayText) then
		aText := Sender.AsString;
end;

function TfUbicaciones.GetFileTime(const AFileName: String; var FileTime: TDateTime): Boolean;
var
  SR: TSearchRec;
  LocalFileTime: TFileTime;
  SystemTime: TSystemTime;
begin
  // Devuelve la fecha del archivo pasado como parámetro
  Result := False;
  if FindFirst(AFileName, faAnyFile, SR) = 0 then
  begin
    if FileTimeToLocalFileTime(SR.FindData.ftLastWriteTime, LocalFileTime) then
    begin
      if FileTimeToSystemTime(LocalFileTime, SystemTime) then
       begin
         FileTime := SystemTimeToDateTime(SystemTime);
         Result := True;
       end;
    end;
    SysUtils.FindClose(SR);
  end;
end;

function TfUbicaciones.CrearBaseDatos( const Fichero: String ): Boolean;
begin
  // Crea la base de datos
  Conexion.Disconnect;
  Conexion.Database := Fichero;
  Try
    Conexion.Connected := true;

    // Crea la tabla de libros
    if fUbicaciones.Query.Active then
      fUbicaciones.Query.Close;
    fUbicaciones.Query.SQL.Clear;
    fUbicaciones.Query.SQL.Text := 'CREATE TABLE libros (isbn	TEXT(13) NOT NULL, titulo	TEXT(2048) NOT NULL, subtitulo	TEXT(2048) NOT NULL, autor TEXT(1024), ubicacion	TEXT(3) NOT NULL, cantidad	INTEGER NOT NULL, precio VARCHAR NOT NULL )';
    fUbicaciones.Query.ExecSQL;
    Result := true;
  except
    MessageDlg( ErrorAlConectar + LineEnding+'', mtError, [],0 );
    Application.Terminate;
  end;

end;

end.
