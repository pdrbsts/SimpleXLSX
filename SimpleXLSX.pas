unit SimpleXLSX;

interface

uses
  System.SysUtils,
  System.Classes,
  System.Zip,
  System.Generics.Collections;

type
  TRowData = class
    RowIndex: Integer;
    Cells: TStringList;
    constructor Create(ARowIndex: Integer);
    destructor Destroy; override;
  end;

  TSimpleExcel = class
  private
    FWorkbookXml: TStringStream;
    FSheetXml: TStringList;
    FCoreXml: TStringStream;
    FRelsXml: TStringStream;
    FContentTypesXml: TStringStream;
    FWorkbookRelsXml: TStringStream;
    FMaxRow, FMaxCol: Integer;
    FRows: TObjectList<TRowData>;
    function ColumnNumberToLetter(Col: Integer): string;
	function EscapeXML(const S: string): string;
  public
    constructor Create;
    destructor Destroy; override;
    procedure AddData(Row, Col: Integer; const Value: string);
    procedure SaveToFile(const FileName: string);
  end;

implementation

constructor TRowData.Create(ARowIndex: Integer);
begin
  RowIndex := ARowIndex;
  Cells := TStringList.Create;
end;

destructor TRowData.Destroy;
begin
  Cells.Free;
  inherited;
end;

constructor TSimpleExcel.Create;
begin
  FMaxRow := 1;
  FMaxCol := 1;
  FRows := TObjectList<TRowData>.Create(True);
  
  // ContentTypes XML
  FContentTypesXml := TStringStream.Create(
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
  '</Types>', TEncoding.UTF8);

  // Workbook XML
  FWorkbookXml := TStringStream.Create(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets>' +
        '<sheet name="Sheet1" sheetId="1" r:id="rId1"/>' +
      '</sheets>' +
    '</workbook>', TEncoding.UTF8);

  // Core Properties XML
  FCoreXml := TStringStream.Create(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
    'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
    'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      '<dc:title>Generated Document</dc:title>' +
      '<dc:creator>SimpleXLSX</dc:creator>' +
      '<dcterms:created xsi:type="dcterms:W3CDTF">' + FormatDateTime('yyyy-mm-dd"T"hh:nn:ss"Z"', Now) + '</dcterms:created>' +
    '</cp:coreProperties>', TEncoding.UTF8);
	
  // Workbook Rels XML
  FWorkbookRelsXml := TStringStream.Create(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
    '</Relationships>', TEncoding.UTF8);
	
  // Rels XML
  FRelsXml := TStringStream.Create(
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
  '</Relationships>', TEncoding.UTF8);
  
  FRows.Clear;
  FMaxRow := 1;
  FMaxCol := 1;
end;

destructor TSimpleExcel.Destroy;
begin
  FWorkbookXml.Free;
  FCoreXml.Free;
  FRelsXml.Free;
  FContentTypesXml.Free;
  FWorkbookRelsXml.Free;
  FRows.Free;
  inherited;
end;

function TSimpleExcel.EscapeXML(const S: string): string;
begin
  Result := S.Replace('&', '&amp;', [rfReplaceAll])
             .Replace('<', '&lt;', [rfReplaceAll])
             .Replace('>', '&gt;', [rfReplaceAll])
             .Replace('"', '&quot;', [rfReplaceAll])
             .Replace('''', '&apos;', [rfReplaceAll]);
end;

function TSimpleExcel.ColumnNumberToLetter(Col: Integer): string;
begin
  Result := '';
  while Col > 0 do
  begin
    Col := Col - 1;
    Result := Chr(65 + Col mod 26) + Result;
    Col := Col div 26;
  end;
end;

procedure TSimpleExcel.AddData(Row, Col: Integer; const Value: string);
var
  ColumnLetter, CellData: string;
  RowData: TRowData;
  i: Integer;
begin
  ColumnLetter := ColumnNumberToLetter(Col);
  CellData := Format('<c r="%s%d" t="inlineStr"><is><t>%s</t></is></c>', [ColumnLetter, Row, EscapeXml(Value)]);

  RowData := nil;
  for i := 0 to FRows.Count - 1 do
  begin
    if FRows[i].RowIndex = Row then
    begin
      RowData := FRows[i];
      Break;
    end;
  end;

  if RowData = nil then
  begin
    RowData := TRowData.Create(Row);
    FRows.Add(RowData);
  end;

  RowData.Cells.Add(CellData);

  if Row > FMaxRow then
    FMaxRow := Row;
  if Col > FMaxCol then
    FMaxCol := Col;
end;

procedure TSimpleExcel.SaveToFile(const FileName: string);
var
  Zip: TZipFile;
  MaxCellRef, CellsText: string;
  SheetXml: TStringList;
  SheetStream: TStringStream;
  i: Integer;
begin
  MaxCellRef := ColumnNumberToLetter(FMaxCol) + IntToStr(FMaxRow);
  SheetXml := TStringList.Create;
  SheetStream := TStringStream.Create('', TEncoding.UTF8);
  try
    SheetXml.Add('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    SheetXml.Add('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">');
    SheetXml.Add('  <sheetPr/>');
    SheetXml.Add(Format('  <dimension ref="A1:%s"/>', [MaxCellRef]));
    SheetXml.Add('  <sheetData>');

    for i := 0 to FRows.Count - 1 do
    begin
      CellsText := String.Join('', FRows[i].Cells.ToStringArray);
      SheetXml.Add(Format('    <row r="%d">%s</row>', [FRows[i].RowIndex, CellsText]));
    end;

    SheetXml.Add('  </sheetData>');
    SheetXml.Add('</worksheet>');
    SheetStream.WriteString(SheetXml.Text);

    Zip := TZipFile.Create;
    try
      Zip.Open(FileName + '.xlsx', zmWrite);
      try
        FContentTypesXml.Position := 0;
        FWorkbookXml.Position := 0;
        FCoreXml.Position := 0;
        FRelsXml.Position := 0;
        FWorkbookRelsXml.Position := 0;
        SheetStream.Position := 0;

        Zip.Add(FContentTypesXml, '[Content_Types].xml');
        Zip.Add(FRelsXml, '_rels/.rels');
        Zip.Add(FCoreXml, 'docProps/core.xml');
        Zip.Add(FWorkbookRelsXml, 'xl/_rels/workbook.xml.rels');
        Zip.Add(FWorkbookXml, 'xl/workbook.xml');
        Zip.Add(SheetStream, 'xl/worksheets/sheet1.xml');
      finally
        Zip.Close;
      end;
    finally
      Zip.Free;
    end;
  finally
    SheetXml.Free;
    SheetStream.Free;
  end;
end;

end.
