unit uActualizar;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, ComCtrls,
  fpspreadsheetctrls, fpstypes, fpspreadsheet, fpsallformats, IniFiles;

type

  { TfActualizar }

  TfActualizar = class(TForm)
    lblActualizar: TLabel;
    lblEsperar: TLabel;
    pbActualizar: TProgressBar;
    procedure FormActivate(Sender: TObject);
  private
    procedure Autores();
  public

  end;

var
  fActualizar: TfActualizar;

const
  OUTPUT_FORMAT = sfExcel8;

implementation
uses uPrincipal;

{$R *.frm}

{ TfActualizar }

procedure TfActualizar.FormActivate(Sender: TObject);
var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  Cell: PCell;
  MyDir: String;
  row: Cardinal;
  ContM: integer;
  FicheroINI: TIniFile;
begin
  FicheroINI := TiniFile.Create( ChangeFileExt( Application.ExeName, '.ini' ) );
  // Prepara el contador de cambios
  contM := 0;

  // Limpia la base de datos
  if fUbicaciones.Query.Active then
    fUbicaciones.Query.Close;
  fUbicaciones.Query.SQL.Clear;
  fUbicaciones.Query.SQL.Text := 'DELETE FROM libros';
  fUbicaciones.Query.ExecSQL;

  // Busca el archivo de los libros
  // Busca la ubicacion del archivo de excel con la lista de libros
  if FicheroINI.ReadString( 'Archivos', 'Lista', '' ) = '' then begin
    fUbicaciones.Abrir.Title := 'Abrir una hoja de datos con LIBROS';
    MyDir := fUbicaciones.BuscarArchivo();
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
  try
    // Prepara el archivo de excel
    MyWorkbook := TsWorkbook.Create;

    // Cambia el cursor al reloj de arena
    Cursor := crHourGlass;

    // Abre el archivo excel
    MyWorkbook.ReadFromFile( MyDir );
    MyWorksheet := MyWorkbook.ActiveWorksheet;

    // Prepara el Query para la inserción de libros
    if fUbicaciones.Query.Active then
       fUbicaciones.Query.Close;
    fUbicaciones.Query.SQL.Text := 'INSERT INTO libros VALUES ( :isbn1, :autor1, :titulo1, :ubicacion1, :cantidad1 );';

    // Prepara la barra de progreso
    pbActualizar.Max := MyWorksheet.GetLastRowIndex;

    // Prepara la inserción del registro

    // Empieza en la línea 1 para saltar el encabezado de las columnas
    for row := 1 to MyWorksheet.GetLastRowIndex do begin
      // Prepara los campos del registro
      // La primera columna es la 0
      cell := MyWorksheet.FindCell( row, 1 ); // ISBN
      // Elimina los guiones del ISBN
      fUbicaciones.Query.Params.ParamByName('isbn1').AsString := StringReplace( MyWorksheet.ReadAsText( cell ), '-', '', [rfReplaceAll] );

      cell := MyWorksheet.FindCell( row, 2 ); // TÍTULO
      fUbicaciones.Query.Params.ParamByName('titulo1').AsString := MyWorksheet.ReadAsText( cell );

      cell := MyWorksheet.FindCell( row, 5 ); // UBICACIÓN
      fUbicaciones.Query.Params.ParamByName('ubicacion1').AsString := MyWorksheet.ReadAsText( cell );

      cell := MyWorksheet.FindCell( row, 7 ); // STOCK
      fUbicaciones.Query.Params.ParamByName('cantidad1').AsInteger := Round( MyWorkSheet.ReadAsNumber( cell ) );

      // Prepara el campo de autor
      fUbicaciones.Query.Params.ParamByName('autor1').AsString := '';

      // Actualiza la barra de progreso y los registros procesados
      Application.ProcessMessages;
      pbActualizar.Position := row;
      pbActualizar.Caption := IntToStr( pbActualizar.Position * 100 div pbActualizar.Max) + '%';

      // Guarda el registro si es diferente de "NO_UBI_1, NO_UBI_3 y cajas de archivo numeradas
      if Length( fUbicaciones.Query.Params.ParamByName('ubicacion1').AsString ) <> 3 then
         Continue;
      fUbicaciones.Query.ExecSQL;

      // Actualiza los registros modificados
      Inc( ContM );
    end;

    // Actualiza la fecha del archivo de excel
    FileSetDate( MyDir, DateTimeToFileDate( Now ) );

  finally
    MyWorkbook.Free;
    Cursor := crDefault;
  end;

  // Carga los autores
  Autores();

  // Cierra la consulta a la base de datos
  fUbicaciones.Query.Close;

  // Avisa de que ha finalizado la actualización
  MessageDlg( 'La actualización se ha realizado con éxito', mtInformation, [mbOK],0);

  Close;
end;

procedure TfActualizar.Autores();
var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  Cell: PCell;
  row: Cardinal;
  sAutores: String;
  FicheroINI: TIniFile;
begin
  FicheroINI := TiniFile.Create( ChangeFileExt( Application.ExeName, '.ini' ) );
  // Busca la ubicacion del archivo de excel con la lista de autores
  if FicheroINI.ReadString( 'Archivos', 'Autores', '' ) = '' then begin
    fUbicaciones.Abrir.Title := 'Abrir una hoja de datos con AUTORES';
    sAutores := fUbicaciones.BuscarArchivo();
    if sAutores <> '' then begin
      FicheroINI.WriteString( 'Archivos', 'Autores', sAutores );
      FicheroINI.Free;
    end else begin
      FicheroINI.Free;
      Exit;
    end;
  end else
      sAutores := FicheroINI.ReadString( 'Archivos', 'Autores', '' );
//  sAutores := '\\172.16.204.63\Archivo\Autores.xls';

  try
    // Prepara el archivo de excel
    MyWorkbook := TsWorkbook.Create;

    // Cambia el cursor al reloj de arena
    Cursor := crHourGlass;

    // Abre el archivo excel
    MyWorkbook.ReadFromFile( sAutores );
    MyWorksheet := MyWorkbook.ActiveWorksheet;

    // Prepara el Query para la inserción de libros
    if fUbicaciones.Query.Active then
       fUbicaciones.Query.Close;
    fUbicaciones.Query.SQL.Text := 'UPDATE libros SET autor = :autor1 WHERE isbn = :isbn1;';

    // Prepara la barra de progreso
    pbActualizar.Max := MyWorksheet.GetLastRowIndex;

    // Prepara la inserción del registro

    // Empieza en la línea 1 para saltar el encabezado de las columnas
    for row := 1 to MyWorksheet.GetLastRowIndex do begin
      // Prepara los campos del registro
      // La primera columna es la 0
      cell := MyWorksheet.FindCell( row, 3 ); // ISBN
      fUbicaciones.Query.Params.ParamByName('isbn1').AsString := MyWorksheet.ReadAsText( cell );

      cell := MyWorksheet.FindCell( row, 6 ); // AUTOR
      fUbicaciones.Query.Params.ParamByName('autor1').AsString := MyWorksheet.ReadAsText( cell );

      // Actualiza la barra de progreso y los registros procesados
      Application.ProcessMessages;
      pbActualizar.Position := row;
      pbActualizar.Caption := IntToStr( pbActualizar.Position * 100 div pbActualizar.Max) + '%';

      fUbicaciones.Query.ExecSQL;
    end;

    // Actualiza la fecha del archivo de excel
    FileSetDate( sAutores, DateTimeToFileDate( Now ) );
  finally
    MyWorkbook.Free;
    Cursor := crDefault;
  end;
end;

end.

