unit ExportFunctionsU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ActnList, StdCtrls, TeEngine, Series, ExtCtrls, TeeProcs, Chart,
  Excel97, Excel2000, OLECtrls, OLEServer, Spin, ActiveX, ComCtrls, Math,
  Menus, ExperementDataU, ShellApi;

procedure ExportOneToExcel(ExperementData: TList; CurrentExperementDataIndex: Integer; speSampleNoValue: Integer);
procedure ExportAllToExcel(ExperementData: TList; CurrentExperementDataIndex: Integer);
procedure ExportAllToTXT(ExperementData: TList);
procedure ExportAllExperementsToExcel(ExperementData: TList);
function generateTempFileName(): string;

implementation

const
  kt = 1.96;

// перевірка чи даний OLE-об’єкт зареєстрований в системі

function IsOLEObjectInstalled(Name: string): Boolean;
var
  ClassID: TCLSID;
  Res: HRESULT;
begin
  // шукаємо CLSID OLE-об’єкта
  Res := CLSIDFromProgID(PWideChar(WideString(Name)), ClassID);

  if (Res = S_OK)
    then Result := true
  else Result := false;
end;


procedure ExportAllToExcel(ExperementData: TList; CurrentExperementDataIndex: Integer);
var
        // сервер Excel
  ExcelApp: TExcelApplication;
  // робоча книга
  ExcelBook: ExcelWorkBook;
  // робочий листок
  ExcelSheet: ExcelWorkSheet;
  // діапазон комірок
  Range: OLEVariant;
  // допоміжний масив
  V: OLEVariant;
  // допоміжні змінні
  i: Integer;

  function IntToColumnName(const Int: Integer): string;
  var
    n: Integer;

    function IntToLetter(const Int: Integer): string;
    begin
      if (Int <> 0)
        then Result := Chr(Int + 64)
      else Result := '';
    end;

  begin
    n := Int mod 26;
    if (n = 0)
      then Result := IntToLetter(Int div 26 - 1) + IntToLetter(26)
    else Result := IntToLetter(Int div 26) + IntToLetter(n);
  end;

begin
  if not IsOLEObjectInstalled('Excel.Application') then begin
    Application.MessageBox('Не вдається знайти та запустити програму Excel.'#13 +
      'Можливо вона не встановлена.', 'Помилка', mb_OK);
    Exit;
  end;

  // створення сервера
  ExcelApp := TExcelApplication.Create(ExcelApp);
  // і задання початкових налаштувань
  ExcelApp.ConnectKind := ckRunningOrNew;

  // створення нової книги
  ExcelBook := ExcelApp.Workbooks.Add(EmptyParam, LOCALE_USER_DEFAULT);

  ExcelSheet := ExcelApp.ActiveWorkbook.Worksheets[1] as ExcelWorkSheet;
  ExcelSheet.Name := 'MAE statistics';

  // налаштування параметрів сторінки
  try
    with ExcelSheet.PageSetup do begin
      BottomMargin := 28;
      FooterMargin := 20;
      HeaderMargin := 20;
      LeftMargin := 28;
      RightMargin := 28;
      TopMargin := 28;
      Orientation := xlLandscape;
    end;
  except
  end;

  // створення таблиці
  V := VarArrayCreate([1, Max(TExperementData(ExperementData[CurrentExperementDataIndex]).SamplesCount + 1, 15), 1, 9], varVariant);

  // наповнення таблиці даними
  V[1, 1] := 'Data #';
  V[1, 2] := 'A Max';
  V[1, 3] := 'A Sum';
  V[1, 4] := 'N Sum';
  V[1, 5] := 'A Sqr Sum';

  for i := 0 to TExperementData(ExperementData[CurrentExperementDataIndex]).SamplesCount - 1 do begin
    V[i + 2, 1] := i + 1;
    V[i + 2, 2] := TExperementData(ExperementData[CurrentExperementDataIndex]).AMax[i];
    V[i + 2, 3] := FloatToStr(TExperementData(ExperementData[CurrentExperementDataIndex]).ASum[i]);
    V[i + 2, 4] := FloatToStr(TExperementData(ExperementData[CurrentExperementDataIndex]).NSum[i]);
    V[i + 2, 5] := FloatToStr(TExperementData(ExperementData[CurrentExperementDataIndex]).ASumDivNSum[i]);
  end;

  V[1, 7] := 'Середні значення';
  V[2, 7] := 'Амплітуда =';
  V[3, 7] := 'Сума амплітуд =';
  V[4, 7] := 'Підсум рахунок =';
  V[5, 7] := 'Сума квадр. амплітуд =';
  V[2, 9] := TExperementData(ExperementData[CurrentExperementDataIndex]).AMaxMeanValue;
  V[3, 9] := TExperementData(ExperementData[CurrentExperementDataIndex]).ASumMeanValue;
  V[4, 9] := TExperementData(ExperementData[CurrentExperementDataIndex]).NSumMeanValue;
  V[5, 9] := TExperementData(ExperementData[CurrentExperementDataIndex]).ASumDivNSumMeanValue;
  V[6, 7] := 'Довірчі інтервали';
  V[7, 7] := 'Амплітуда =';
  V[8, 7] := 'Сума амплітуд =';
  V[9, 7] := 'Підсум рахунок =';
  V[10, 7] := 'Сума квадр. амплітуд =';
  V[7, 8] := StrToFloat(FloatToStrF(TExperementData(ExperementData[CurrentExperementDataIndex]).AMaxMeanValue - kt * TExperementData(ExperementData[CurrentExperementDataIndex]).AMaxRMS / Sqrt(TExperementData(ExperementData[CurrentExperementDataIndex]).k), ffFixed, 6, 1));
  V[7, 9] := StrToFloat(FloatToStrF(TExperementData(ExperementData[CurrentExperementDataIndex]).AMaxMeanValue + kt * TExperementData(ExperementData[CurrentExperementDataIndex]).AMaxRMS / Sqrt(TExperementData(ExperementData[CurrentExperementDataIndex]).k), ffFixed, 6, 1));
  V[8, 8] := StrToFloat(FloatToStrF(TExperementData(ExperementData[CurrentExperementDataIndex]).ASumMeanValue - kt * TExperementData(ExperementData[CurrentExperementDataIndex]).ASumRMS / Sqrt(TExperementData(ExperementData[CurrentExperementDataIndex]).k), ffFixed, 6, 1));
  V[8, 9] := StrToFloat(FloatToStrF(TExperementData(ExperementData[CurrentExperementDataIndex]).ASumMeanValue + kt * TExperementData(ExperementData[CurrentExperementDataIndex]).ASumRMS / Sqrt(TExperementData(ExperementData[CurrentExperementDataIndex]).k), ffFixed, 6, 1));
  V[9, 8] := StrToFloat(FloatToStrF(TExperementData(ExperementData[CurrentExperementDataIndex]).NSumMeanValue - kt * TExperementData(ExperementData[CurrentExperementDataIndex]).NSumRMS / Sqrt(TExperementData(ExperementData[CurrentExperementDataIndex]).k), ffFixed, 6, 1));
  V[9, 9] := StrToFloat(FloatToStrF(TExperementData(ExperementData[CurrentExperementDataIndex]).NSumMeanValue + kt * TExperementData(ExperementData[CurrentExperementDataIndex]).NSumRMS / Sqrt(TExperementData(ExperementData[CurrentExperementDataIndex]).k), ffFixed, 6, 1));
  V[10, 8] := StrToFloat(FloatToStrF(TExperementData(ExperementData[CurrentExperementDataIndex]).ASumDivNSumMeanValue - kt * TExperementData(ExperementData[CurrentExperementDataIndex]).ASumDivNSumRMS / Sqrt(TExperementData(ExperementData[CurrentExperementDataIndex]).k), ffFixed, 6, 1));
  V[10, 9] := StrToFloat(FloatToStrF(TExperementData(ExperementData[CurrentExperementDataIndex]).ASumDivNSumMeanValue + kt * TExperementData(ExperementData[CurrentExperementDataIndex]).ASumDivNSumRMS / Sqrt(TExperementData(ExperementData[CurrentExperementDataIndex]).k), ffFixed, 6, 1));
  V[11, 7] := 'Середні квадратичні відхилення';
  V[12, 7] := 'Амплітуда =';
  V[13, 7] := 'Сума амплітуд =';
  V[14, 7] := 'Підсум рахунок =';
  V[15, 7] := 'Сума квадр. амплітуд =';
  V[12, 9] := TExperementData(ExperementData[CurrentExperementDataIndex]).AMaxRMS;
  V[13, 9] := TExperementData(ExperementData[CurrentExperementDataIndex]).ASumRMS;
  V[14, 9] := TExperementData(ExperementData[CurrentExperementDataIndex]).NSumRMS;
  V[15, 9] := TExperementData(ExperementData[CurrentExperementDataIndex]).ASumDivNSumRMS;

  // налаштування колонок
  ExcelSheet.Range['A1', 'E1'].ColumnWidth := 11;
  ExcelSheet.Range['G1', 'G1'].ColumnWidth := 23;
  ExcelSheet.Range['H1', 'I1'].ColumnWidth := 8;
  ExcelSheet.Range['A1', 'E1'].Font.Bold := true;
  ExcelSheet.Range['G1', 'G15'].Font.Bold := true;
  ExcelSheet.Range['G1', 'G1'].Font.Underline := true;
  ExcelSheet.Range['G6', 'G6'].Font.Underline := true;
  ExcelSheet.Range['G11', 'G11'].Font.Underline := true;

  // визначаємо діапазон комірок, в які передавати масив V
  Range := ExcelSheet.Range['A1', 'I' + IntToStr(TExperementData(ExperementData[CurrentExperementDataIndex]).SamplesCount + 1)];
  // вивід даних в Excel
  Range.Value := V;

  // робимо активним перший аркуш і видимим Excel
  ExcelSheet := ExcelApp.ActiveWorkbook.Worksheets[1] as ExcelWorkSheet;
  ExcelSheet.Activate(LOCALE_USER_DEFAULT);
  ExcelApp.Visible[LOCALE_USER_DEFAULT] := true;
end;


// експорт в Excel

procedure ExportOneToExcel(ExperementData: TList; CurrentExperementDataIndex: Integer;
  speSampleNoValue: Integer);
var
  // сервер Excel
  ExcelApp: TExcelApplication;
  // робоча книга
  ExcelBook: ExcelWorkBook;
  // робочий листок
  ExcelSheet: ExcelWorkSheet;
  // діапазон комірок
  Range: OLEVariant;
  // допоміжний масив
  V: OLEVariant;
  // допоміжні змінні
  i: Integer;
  TmpStr: string;

  function IntToColumnName(const Int: Integer): string;
  var
    n: Integer;

    function IntToLetter(const Int: Integer): string;
    begin
      if (Int <> 0)
        then Result := Chr(Int + 64)
      else Result := '';
    end;

  begin
    n := Int mod 26;
    if (n = 0)
      then Result := IntToLetter(Int div 26 - 1) + IntToLetter(26)
    else Result := IntToLetter(Int div 26) + IntToLetter(n);
  end;

begin
  if not IsOLEObjectInstalled('Excel.Application') then begin
    Application.MessageBox('Не вдається знайти та запустити програму Excel.'#13 +
      'Можливо вона не встановлена.', 'Помилка', mb_OK);
    Exit;
  end;

  // створення сервера
  ExcelApp := TExcelApplication.Create(ExcelApp);
  // і задання початкових налаштувань
  ExcelApp.ConnectKind := ckRunningOrNew;

  // створення нової книги
  ExcelBook := ExcelApp.Workbooks.Add(EmptyParam, LOCALE_USER_DEFAULT);

  ExcelSheet := ExcelApp.ActiveWorkbook.Worksheets[1] as ExcelWorkSheet;
  ExcelSheet.Name := 'Sample #' + IntToStr(speSampleNoValue);

  // налаштування параметрів сторінки
  try
    with ExcelSheet.PageSetup do begin
      BottomMargin := 28;
      FooterMargin := 20;
      HeaderMargin := 20;
      LeftMargin := 28;
      RightMargin := 28;
      TopMargin := 28;
      Orientation := xlPortrait;
    end;
  except
  end;

  // створення таблиці
  V := VarArrayCreate([1, High(TExperementData(ExperementData[CurrentExperementDataIndex]).Samples[speSampleNoValue - 1]) + 1, 1, 4], varVariant);

  // наповнення таблиці даними

  for i := 0 to High(TExperementData(ExperementData[CurrentExperementDataIndex]).Samples[speSampleNoValue - 1]) do begin
    V[i + 1, 1] := TExperementData(ExperementData[CurrentExperementDataIndex]).Samples[speSampleNoValue - 1, i];
  end;

  V[1, 3] := 'Назва файлу:';
  V[2, 3] := 'Номер виборки:';

  // "витягуємо" ім'я файлу
  TmpStr := TExperementData(ExperementData[CurrentExperementDataIndex]).FileName;
  SetLength(TmpStr, Length(TmpStr) - 4);

  V[1, 4] := TmpStr;
  V[2, 4] := speSampleNoValue;

  // налаштування колонок
  ExcelSheet.Range['A1', 'A1'].ColumnWidth := 11;
  ExcelSheet.Range['C1', 'C1'].ColumnWidth := 14;
  ExcelSheet.Range['D1', 'D1'].ColumnWidth := 11;

  // визначаємо діапазон комірок, в які передавати масив V
  Range := ExcelSheet.Range['A1', 'D' + IntToStr(High(TExperementData(ExperementData[CurrentExperementDataIndex]).Samples[speSampleNoValue - 1]) + 1)];
  // вивід даних в Excel
  Range.Value := V;

  // робимо активним перший аркуш і видимим Excel
  ExcelSheet := ExcelApp.ActiveWorkbook.Worksheets[1] as ExcelWorkSheet;
  ExcelSheet.Activate(LOCALE_USER_DEFAULT);
  ExcelApp.Visible[LOCALE_USER_DEFAULT] := true;
end;

procedure ExportAllExperementsToExcel(ExperementData: TList);
var
          // сервер Excel
  ExcelApp: TExcelApplication;
  // робоча книга
  ExcelBook: ExcelWorkBook;
  // робочий листок
  ExcelSheet: ExcelWorkSheet;
  // діапазон комірок
  Range: OLEVariant;
  // допоміжний масив
  V: OLEVariant;
  // допоміжні змінні
  i: Integer;
  TmpStr: string;
begin
  if not IsOLEObjectInstalled('Excel.Application') then begin
    Application.MessageBox('Не вдається знайти та запустити програму Excel.'#13 +
      'Можливо вона не встановлена.', 'Помилка', mb_OK);
    Exit;
  end;

  // створення сервера
  ExcelApp := TExcelApplication.Create(ExcelApp);
  // і задання початкових налаштувань
  ExcelApp.ConnectKind := ckRunningOrNew;

  // створення нової книги
  ExcelBook := ExcelApp.Workbooks.Add(EmptyParam, LOCALE_USER_DEFAULT);

  ExcelSheet := ExcelApp.ActiveWorkbook.Worksheets[1] as ExcelWorkSheet;
  //ExcelSheet.Name := 'Sample #' + IntToStr(speSampleNoValue);

  // налаштування параметрів сторінки
  try
    with ExcelSheet.PageSetup do begin
      BottomMargin := 28;
      FooterMargin := 20;
      HeaderMargin := 20;
      LeftMargin := 28;
      RightMargin := 28;
      TopMargin := 28;
      Orientation := xlPortrait;
    end;
  except
  end;

  /////////////////////////////////////////////////////////////
  // створення таблиці
  V := VarArrayCreate([1, 6, 1, ExperementData.Count + 1], varVariant);

  V[1, 1] := 'Експеременти:';
  V[2, 1] := 'Напруженість:';
  V[3, 1] := 'Середні значення';
  V[4, 1] := 'Амплітуда:';
  V[5, 1] := 'Сума амплітуд:';
  V[6, 1] := 'Підсум рахунок:';

  // наповнення таблиці даними
  for i := 0 to ExperementData.Count - 1 do begin
        V[1, 2 + i] := TExperementData(ExperementData[i]).FileName;
        V[2, 2 + i] := TExperementData(ExperementData[i]).magneticFieldStrength;
        //V[3, 2 + i] := 'Середні значення';
        V[4, 2 + i] := TExperementData(ExperementData[i]).AMaxMeanValue;
        V[5, 2 + i] := TExperementData(ExperementData[i]).ASumMeanValue;
        V[6, 2 + i] := TExperementData(ExperementData[i]).NSumMeanValue;
  end;

  /////////////////////////////////////////////////////////////

  // налаштування колонок
  //ExcelSheet.Range['A1', 'A1'].ColumnWidth := 11;
  //ExcelSheet.Range['C1', 'C1'].ColumnWidth := 14;
  //ExcelSheet.Range['D1', 'D1'].ColumnWidth := 11;

  // визначаємо діапазон комірок, в які передавати масив V
  Range := ExcelSheet.Range[ExcelSheet.Cells.Item[1,1], ExcelSheet.Cells.Item[6,ExperementData.Count + 1]];
  // вивід даних в Excel
  Range.Value := V;

  // робимо активним перший аркуш і видимим Excel
  ExcelSheet := ExcelApp.ActiveWorkbook.Worksheets[1] as ExcelWorkSheet;
  ExcelSheet.Activate(LOCALE_USER_DEFAULT);
  ExcelApp.Visible[LOCALE_USER_DEFAULT] := true;
end;

procedure ExportAllToTXT(ExperementData: TList);
var
  sl: TStringList; tmpFile: string; tmpSrt: string; i: Integer;
  minV: Double; maxV: Double;
begin
  tmpFile := generateTempFileName();
  sl := TStringList.Create;
  try // add the text in here

    tmpSrt := 'Експеременти';
    for i := 0 to ExperementData.Count - 1 do begin
      tmpSrt := tmpSrt + #9 + TExperementData(ExperementData[i]).FileName;
    end;
    sl.Add(tmpSrt);

    tmpSrt := 'Напруженість';
    for i := 0 to ExperementData.Count - 1 do begin
      tmpSrt := tmpSrt + #9 + FloatToStr(TExperementData(ExperementData[i]).magneticFieldStrength);
    end;
    sl.Add(tmpSrt);

    sl.Add('Середні значення');

    tmpSrt := 'Амплітуда';
    for i := 0 to ExperementData.Count - 1 do begin
      tmpSrt := tmpSrt + #9 + FloatToStr(TExperementData(ExperementData[i]).AMaxMeanValue);
    end;
    sl.Add(tmpSrt);

    tmpSrt := 'Сума амплітуд';
    for i := 0 to ExperementData.Count - 1 do begin
      tmpSrt := tmpSrt + #9 + FloatToStr(TExperementData(ExperementData[i]).ASumMeanValue);
    end;
    sl.Add(tmpSrt);

    tmpSrt := 'Підсум рахунок';
    for i := 0 to ExperementData.Count - 1 do begin
      tmpSrt := tmpSrt + #9 + FloatToStr(TExperementData(ExperementData[i]).NSumMeanValue);
    end;
    sl.Add(tmpSrt);


    {
    tmpSrt := 'Сума квадр. амплітуд';
    for i := 0 to ExperementData.Count - 1 do begin
        tmpSrt := tmpSrt + #9 + #9 + FloatToStr(TExperementData(ExperementData[i]).ASumDivNSumMeanValue);
    end;
    sl.Add(tmpSrt);


    sl.Add('Середні квадратичні відхилення');

    tmpSrt := 'Амплітуда';
    for i := 0 to ExperementData.Count - 1 do begin
        tmpSrt := tmpSrt + #9 + #9 + FloatToStr(TExperementData(ExperementData[i]).AMaxRMS);
    end;
    sl.Add(tmpSrt);

    tmpSrt := 'Сума амплітуд';
    for i := 0 to ExperementData.Count - 1 do begin
        tmpSrt := tmpSrt + #9 + #9 + FloatToStr(TExperementData(ExperementData[i]).ASumRMS);
    end;
    sl.Add(tmpSrt);

    tmpSrt := 'Підсум рахунок';
    for i := 0 to ExperementData.Count - 1 do begin
        tmpSrt := tmpSrt + #9 + #9 + FloatToStr(TExperementData(ExperementData[i]).NSumRMS);
    end;
    sl.Add(tmpSrt);

    tmpSrt := 'Сума квадр. амплітуд';
    for i := 0 to ExperementData.Count - 1 do begin
        tmpSrt := tmpSrt + #9 + #9 + FloatToStr(TExperementData(ExperementData[i]).ASumDivNSumRMS);
    end;
    sl.Add(tmpSrt);


    sl.Add('Довірчі інтервали');

    tmpSrt := 'Амплітуда';
    for i := 0 to ExperementData.Count - 1 do begin
        minV := TExperementData(ExperementData[i]).AMaxMeanValue - kt*TExperementData(ExperementData[i]).AMaxRMS/Sqrt(TExperementData(ExperementData[i]).k);
        maxV := TExperementData(ExperementData[i]).AMaxMeanValue + kt*TExperementData(ExperementData[i]).AMaxRMS/Sqrt(TExperementData(ExperementData[i]).k);
        tmpSrt := tmpSrt + #9 + FloatToStr(minV) + #9 + FloatToStr(maxV);
    end;
    sl.Add(tmpSrt);

    tmpSrt := 'Сума амплітуд';
    for i := 0 to ExperementData.Count - 1 do begin
        minV := TExperementData(ExperementData[i]).ASumMeanValue - kt*TExperementData(ExperementData[i]).ASumRMS/Sqrt(TExperementData(ExperementData[i]).k);
        maxV := TExperementData(ExperementData[i]).ASumMeanValue + kt*TExperementData(ExperementData[i]).ASumRMS/Sqrt(TExperementData(ExperementData[i]).k);
        tmpSrt := tmpSrt + #9 + FloatToStr(minV) + #9 + FloatToStr(maxV);
    end;
    sl.Add(tmpSrt);

    tmpSrt := 'Підсум рахунок';
    for i := 0 to ExperementData.Count - 1 do begin
        minV := TExperementData(ExperementData[i]).NSumMeanValue - kt*TExperementData(ExperementData[i]).NSumRMS/Sqrt(TExperementData(ExperementData[i]).k);
        maxV := TExperementData(ExperementData[i]).NSumMeanValue + kt*TExperementData(ExperementData[i]).NSumRMS/Sqrt(TExperementData(ExperementData[i]).k);
        tmpSrt := tmpSrt + #9 + FloatToStr(minV) + #9 + FloatToStr(maxV);
    end;
    sl.Add(tmpSrt);

    tmpSrt := 'Сума квадр. амплітуд';
    for i := 0 to ExperementData.Count - 1 do begin
        minV := TExperementData(ExperementData[i]).ASumDivNSumMeanValue - kt*TExperementData(ExperementData[i]).ASumDivNSumRMS/Sqrt(TExperementData(ExperementData[i]).k);
        maxV := TExperementData(ExperementData[i]).ASumDivNSumMeanValue + kt*TExperementData(ExperementData[i]).ASumDivNSumRMS/Sqrt(TExperementData(ExperementData[i]).k);
        tmpSrt := tmpSrt + #9 + FloatToStr(minV) + #9 + FloatToStr(maxV);
    end;
    sl.Add(tmpSrt);
    }

    sl.SaveToFile(tmpFile); //use the 'Savetofile' function of the stringlist
  finally
    sl.Free //Always free a stringlist when you are done using it
  end;

  ShellExecute(0, 'open', 'c:\windows\notepad.exe', PChar(tmpFile), nil, SW_SHOWNORMAL);

end;

function generateTempFileName(): string;
var
  TempFile: array[0..MAX_PATH - 1] of Char;
  TempPath: array[0..MAX_PATH - 1] of Char;
begin
  GetTempPath(MAX_PATH, TempPath);
  if GetTempFileName(TempPath, PChar('abc'), 0, TempFile) = 0 then
    raise Exception.Create(
      'GetTempFileName API failed. ' + SysErrorMessage(GetLastError)
      );
  Result := TempFile;
end;

end.
