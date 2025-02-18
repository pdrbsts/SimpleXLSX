## SimpleXLSX
Basic library to create simple Microsoft Excel XLSX files

#Sample
procedure TForm1.Button1Click(Sender: TObject);
var
  XLS: TSimpleExcel;
begin
  XLS := TSimpleExcel.Create; // Correct way to create the object
  try
    XLS.AddData(1, 1, 'Teste1');
    XLS.AddData(1, 2, 'Teste2');
    XLS.AddData(2, 1, 'Valor1');
    XLS.AddData(2, 2, 'Valor2');
    XLS.SaveToFile('teste');
  finally
    XLS.Free; // Ensure the object is freed to avoid memory leaks
  end;
end;
