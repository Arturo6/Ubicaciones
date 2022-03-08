{
  Copyright (C) 2020 Arturo Molina amolinaj@gmail.com

  UBICACIONES es un programa desarrollado para acceder al catálogo de libros
  editados por el servicio de publicaciones de la Universidad de Alicante y
  facilitar el trabajo de localización de títulos, ubicaciones, portadas y cantidades
  de los ejemplares existentes en el sistema de almacenamiendo de dicho servicio.

  Este programa es Software Libre; Usted puede redistribuirlo y/o modificarlo
  bajo los términos de la "GNU General Public License (GPL)" tal y como ha sido
  públicada por la Free Software Foundation; o bien la versión 2 de la Licencia,
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
  Grids, DBGrids, fpspreadsheetctrls, fpstypes, fpspreadsheet, fpsallformats,
  LCLType, IniFiles, DateUtils, DefaultTranslator, uLicencia, UniqueInstance ;

type

  { TfUbicaciones }

  TfUbicaciones = class(TForm)
    btnActualizaArchivo: TBitBtn;
    btnActualizar: TSpeedButton;
    btnAyuda: TSpeedButton;
    btnBuscar: TSpeedButton;
    btnPrincipal: TSpeedButton;
    FechaArchivo: TLabel;
    lblInfo: TLabel;
    lblLicenciaGNU: TLabel;
    lblProcesados: TLabel;
    lblCambios: TLabel;
    lblFechaArchivo: TLabel;
    lblProgreso: TLabel;
    lblAutor: TLabel;
    lblAutorDB: TLabel;
    LogoPublicaciones: TImage;
    Abrir: TOpenDialog;
    PanelLogo: TPanel;
    Progreso: TProgressBar;
    Rejilla: TDBGrid;
    DS: TDataSource;
    edtTitulo: TEdit;
    lblTituloBuscar: TLabel;
    Paginas: TFZPageControl;
    tsPrincipal: TFZVirtualPage;
    tsBuscar: TFZVirtualPage;
    tsActualizar: TFZVirtualPage;
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
    procedure btnActualizaArchivoClick ( Sender: TObject ) ;
    procedure btnActualizarClick(Sender: TObject);
    procedure btnAyudaClick(Sender: TObject);
    procedure btnBuscarArchivoClick ( Sender: TObject ) ;
    procedure btnBuscarAutoresClick ( Sender: TObject ) ;
    procedure btnBuscarClick(Sender: TObject);
    procedure btnCerrarClick(Sender: TObject);
    procedure btnPrincipalClick(Sender: TObject);
    procedure edtISBNKeyPress ( Sender: TObject; var Key: char ) ;
    procedure edtTituloKeyPress ( Sender: TObject; var Key: char ) ;
    procedure FormCloseQuery ( Sender: TObject; var CanClose: boolean ) ;
    procedure FormCreate ( Sender: TObject ) ;
    procedure FormKeyDown ( Sender: TObject; var Key: Word; Shift: TShiftState ) ;
    procedure FormShow(Sender: TObject);
    procedure lblAutorDBClick ( Sender: TObject ) ;
    procedure lblLicenciaGNUClick ( Sender: TObject ) ;
    procedure lblTituloDBClick ( Sender: TObject ) ;
    procedure RejillaCellClick ( Column: TColumn ) ;
    procedure RejillaPrepareCanvas ( sender: TObject; DataCol: Integer; Column: TColumn; AState: TGridDrawState ) ;
  private
    procedure BuscaPortada(Archivo: string);
    procedure DBGridOnGetText(Sender: TField; var aText: string; DisplayText: Boolean);
    procedure Actualizar;
    function  QuitaGuion( texto: string ): String;
    function  GetFileTime(const AFileName: String; var FileTime: TDateTime): Boolean;
    procedure Autores();
  public

  end;

var
  fUbicaciones: TfUbicaciones;
  Lugar: String;

const
  OUTPUT_FORMAT = sfExcel8;

resourcestring
  HelloMessage = 'Hello World!';
  CloseMessage = 'Closing your app... bye bye!';
  GreetMessage = 'Hey %0:s from %1:s country!';
  ErrorAlConectar = 'Error al conectar';
  ActualizarBase = 'La base de datos podría no estar actualizada';
  PSActualizarBase = 'Por favor, actualiza la base de datos de libros';
  Procesados = ' registros procesados';
  Modificados = ' registros modificados';

implementation

{$R *.frm}

{ TfUbicaciones }

procedure TfUbicaciones.FormCreate ( Sender: TObject ) ;
var
  FicheroIni: TIniFile;
begin
  // Conecta con la base de datos
  Conexion.Disconnect;
  Conexion.Database := ExtractFilePath( Application.ExeName ) + 'lista.db';
  Try
    Conexion.Connected := true;
  except
    MessageDlg( ErrorAlConectar, mtError, [mbOK],0);
    Application.Terminate;
  end;

  // Parámetros de la etiqueta de progreso
  lblProgreso.Parent := Progreso;
  lblProgreso.AutoSize := False;
  lblProgreso.Transparent := True;
  lblProgreso.Top :=  0;
  lblProgreso.Left :=  0;
  lblProgreso.Width := Progreso.ClientWidth;
  lblProgreso.Height := Progreso.ClientHeight;
  lblProgreso.Alignment := taCenter;
  lblProgreso.Layout := tlCenter;

  // Comprobamos la fecha de la base de datos
  FicheroIni := TiniFile.Create( ChangeFileExt( Application.ExeName, '.ini' ) );
  if FicheroIni.ReadString( 'Archivo', 'Fecha', '' ) <> '' then begin
    if StrToDate( FicheroIni.ReadString( 'Archivo', 'Fecha', '' ) ) < Date then
      MessageDlg( ActualizarBase, mtWarning, [mbOK],0);
  end else
    MessageDlg( PSActualizarBase, mtWarning, [mbOK],0);
  FicheroIni.Free;
end;

procedure TfUbicaciones.FormShow(Sender: TObject);
begin
  // Muestra la página principal
  Paginas.ActivePage := tsPrincipal;
  // Y entrega el foco al ISBN
  edtISBN.SetFocus;
end;

procedure TfUbicaciones.FormCloseQuery ( Sender: TObject; var CanClose: boolean	) ;
begin
  // Confirma que quiere salir
  //if Application.MessageBox ( '¿Cerrar el programa?', 'Cerrar', MB_YESNO ) = idNo then
  //   CanClose := false;
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

procedure TfUbicaciones.btnActualizarClick(Sender: TObject);
begin
  // Muestra la página de actualización
  Paginas.ActivePage := tsActualizar;
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
    lblUbicacionDB.Caption := Query.FieldByName('ubicacion').AsString;
    lblCantidadDB.Caption := IntToStr( Query.FieldByName('cantidad').AsInteger );

    // Muestra la portada del libro
    BuscaPortada( edtISBN.Text );
  end else
    MessageDlg('Ese ISBN no está en la base de datos', mtInformation, [mbOK],0);
end;

  // Toma el foco de nuevo
  edtISBN.SetFocus;
end;

procedure TfUbicaciones.btnBuscarArchivoClick ( Sender: TObject ) ;
begin
  //// Se ha pulsado en el botón de buscar archivo excel de libros
  //Abrir.Title := 'Abrir el archivo de LIBROS';
  //if Abrir.Execute then
  //  edtArchivo.Text := Abrir.FileName;
end;

procedure TfUbicaciones.btnBuscarAutoresClick ( Sender: TObject ) ;
begin
  //// Se ha pulsado en el botón de buscar archivo excel de autores
  //Abrir.Title := 'Abrir el archivo de AUTORES';
  //if Abrir.Execute then
  //  edtAutores.Text := Abrir.FileName;
end;

procedure TfUbicaciones.edtTituloKeyPress ( Sender: TObject; var Key: char ) ;
begin
  // Se ha pulsado RETURN en el título
  if Key = #13 then begin
    // Prepara la búsqueda del isbn
    Query.Active := False;
    Query.SQL.Text := 'SELECT * FROM libros WHERE titulo LIKE ' + quotedstr( '%' + edtTitulo.Text + '%' ) + ';';
    Query.Open;

    // No encuentra el título, así que lo intenta con el autor
    if ( Query.RecordCount < 1 ) then begin
      Query.Active := False;
      Query.SQL.Text := 'SELECT * FROM libros WHERE autor LIKE ' + quotedstr( '%' + edtTitulo.Text + '%' ) + ';';
      Query.Open;
    end;
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

  // Se ha pulsado F2, se muestra la pantalla de actualización
  if Key = VK_F2 then begin
    Paginas.ActivePage := tsActualizar;
    btnActualizaArchivo.SetFocus;
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

procedure TfUbicaciones.btnActualizaArchivoClick ( Sender: TObject ) ;
begin
  // Muestra la barra de progreso
  Progreso.Show;
  // Si no se muestra el path del archivo excel, se pulsa el botón de buscarlo
  //if edtArchivo.Text = '' then
  //  btnBuscarArchivoClick( self );
  //
  Actualizar;
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

//***************************************************************//
//***************************************************************//
//                                                               //
//                                                               //
// Funciones auxiliares no vinculadas con el flujo del programa  //
//                                                               //
//                                                               //
//***************************************************************//
//***************************************************************//

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

procedure TfUbicaciones.Actualizar;
var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  Cell: PCell;
  MyDir: String;
  FileTime: TDateTime;
  FicheroIni: TIniFile;
  row: Cardinal;
  ContM: integer;
begin
  // Muestra el archivo que está procesando
  lblInfo.Caption := 'Procesando las ubicaciones...';

  // Prepara el contador de cambios
  contM := 0;

  // Limpia la base de datos
  if Query.Active then
    Query.Close;
  Query.SQL.Clear;
  Query.SQL.Text := 'DELETE FROM libros';
  Query.ExecSQL;

  // Busca el archivo de los libros
  MyDir := '\\172.16.204.63\Archivo\Lista.xls';

  // Si no hay archivo de libros cancela la opreación
  if Mydir = '' then
     Exit;

  try
    // Prepara el archivo de excel
    MyWorkbook := TsWorkbook.Create;

    // Cambia el cursor al reloj de arena
    fUbicaciones.Cursor := crHourGlass;

    // Abre el archivo excel
    MyWorkbook.ReadFromFile( MyDir );
    MyWorksheet := MyWorkbook.ActiveWorksheet;

    // Prepara el Query para la inserción de libros
    if Query.Active then
       Query.Close;
    Query.SQL.Text := 'INSERT INTO libros VALUES ( :isbn1, :autor1, :titulo1, :ubicacion1, :cantidad1 );';

    // Prepara la barra de progreso
    Progreso.Max := MyWorksheet.GetLastRowIndex;

    // Prepara la inserción del registro

    // Empieza en la línea 1 para saltar el encabezado de las columnas
    for row := 1 to MyWorksheet.GetLastRowIndex do begin
      // Prepara los campos del registro
      // La primera columna es la 0
      cell := MyWorksheet.FindCell( row, 1 ); // ISBN
      Query.Params.ParamByName('isbn1').AsString := QuitaGuion( MyWorksheet.ReadAsText( cell ) );

      cell := MyWorksheet.FindCell( row, 2 ); // TÍTULO
      Query.Params.ParamByName('titulo1').AsString := MyWorksheet.ReadAsText( cell );

      cell := MyWorksheet.FindCell( row, 5 ); // UBICACIÓN
      Query.Params.ParamByName('ubicacion1').AsString := MyWorksheet.ReadAsText( cell );

      cell := MyWorksheet.FindCell( row, 6 ); // STOCK
      Query.Params.ParamByName('cantidad1').AsInteger := Round( MyWorkSheet.ReadAsNumber( cell ) );

      // Prepara el campo de autor
      Query.Params.ParamByName('autor1').AsString := '';

      // Actualiza la barra de progreso y los registros procesados
      Application.ProcessMessages;
      Progreso.Position := row;
      lblProgreso.Caption := IntToStr( Progreso.Position * 100 div Progreso.Max) + '%';
      lblProcesados.Caption := IntToStr( row ) + Procesados;

      // Guarda el registro si es diferente de "NO_UBI_1, NO_UBI_3 y cajas de archivo numeradas
      if Length( Query.Params.ParamByName('ubicacion1').AsString ) <> 3 then
         Continue;
      Query.ExecSQL;

      // Actualiza los registros modificados
      Inc( ContM );
      lblCambios.Caption := IntToStr( ContM ) + Modificados;
    end;
  finally
    MyWorkbook.Free;
    fUbicaciones.Cursor := crDefault;
  end;

  lblInfo.Caption := '';
  // Carga los autores
  Autores();

  // Cierra la consulta a la base de datos
  Query.Close;

  // Obtiene la fecha del archivo excel de libros y la muestra
  if GetFileTime( MyDir, FileTime ) then begin
    lblFechaArchivo.Visible:=true;
    if DayOfTheMonth( FileTime ) = DayOfTheMonth( Now() ) then
       FechaArchivo.Font.Color := clBlue
    else
      FechaArchivo.Font.Color := clRed;

    FechaArchivo.Caption := FormatDateTime( 'dd "de" mmmm "de" yyyy', FileTime )
  end;

  // Guarda en el archivo INI la fecha y ubicación del excel con los libros
  FicheroIni := TiniFile.Create(  ChangeFileExt( Application.ExeName, '.ini' ) );
  FicheroIni.WriteString( 'Archivo', 'Fecha', DateToStr( FileTime ) );
  FicheroIni.Free;

  // Avisa de que ha finalizado la actualización
  MessageDlg('La actualización se ha realizado con éxito', mtInformation, [mbOK],0);

  // Vuelve a la página principal
  btnPrincipalClick( Self );
end;

function TfUbicaciones.QuitaGuion( texto: string ): String;
var
  i: Integer;
begin
  // Elimina los guiones del ISBN
  Result := '';
  for i := 1 to Length( Texto ) do
    if Texto[ i ] in ['0'..'9'] then
      Result := Result + Texto[ i ];
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

procedure TfUbicaciones.Autores();
var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  Cell: PCell;
  row: Cardinal;
  sAutores: String;
begin
  sAutores := '\\172.16.204.63\Archivo\Autores.xls';

  // Si no hay archivo de autores cancela la opreación
  if sAutores = '' then
     Exit;

  lblCambios.Visible := false;
    // Muestra el archivo que está procesando
  lblInfo.Caption := 'Procesando los autores...';

  if FileExists( sAutores ) then begin
    try
      // Prepara el archivo de excel
      MyWorkbook := TsWorkbook.Create;

      // Cambia el cursor al reloj de arena
      fUbicaciones.Cursor := crHourGlass;

      // Abre el archivo excel
      MyWorkbook.ReadFromFile( sAutores );
      MyWorksheet := MyWorkbook.ActiveWorksheet;

      // Prepara el Query para la inserción de libros
      if Query.Active then
         Query.Close;
      Query.SQL.Text := 'UPDATE libros SET autor = :autor1 WHERE isbn = :isbn1;';

      // Prepara la barra de progreso
      Progreso.Max := MyWorksheet.GetLastRowIndex;

      // Prepara la inserción del registro

      // Empieza en la línea 1 para saltar el encabezado de las columnas
      for row := 1 to MyWorksheet.GetLastRowIndex do begin
        // Prepara los campos del registro
        // La primera columna es la 0
        cell := MyWorksheet.FindCell( row, 3 ); // ISBN
        Query.Params.ParamByName('isbn1').AsString := MyWorksheet.ReadAsText( cell );

        cell := MyWorksheet.FindCell( row, 6 ); // AUTOR
        Query.Params.ParamByName('autor1').AsString := MyWorksheet.ReadAsText( cell );

        // Actualiza la barra de progreso y los registros procesados
        Application.ProcessMessages;
        Progreso.Position := row;
        lblProgreso.Caption := IntToStr( Progreso.Position * 100 div Progreso.Max) + '%';
        lblProcesados.Caption := IntToStr( row ) + Procesados;

        Query.ExecSQL;
      end;
    finally
      MyWorkbook.Free;
      fUbicaciones.Cursor := crDefault;
    end;
  end;

  lblInfo.Caption := '';
end;

end.
