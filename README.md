# SimpleXLSX
Basic library to create simple Microsoft Excel XLSX files

 **Usage**
procedure TForm1.Button1Click(Sender: TObject);
var
  XLS: TSimpleExcel;
begin
  XLS := TSimpleExcel.Create;
  try
    XLS.AddData(1, 1, 'Test1');
    XLS.AddData(1, 2, 'Test2');
    XLS.AddData(2, 1, 'Value1');
    XLS.AddData(2, 2, 'Value2');
    XLS.SaveToFile('teste');
  finally
    XLS.Free;
end;
end;	
