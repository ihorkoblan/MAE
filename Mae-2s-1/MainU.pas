unit MainU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ActnList, StdCtrls, TeEngine, Series, ExtCtrls, TeeProcs, Chart,
  Excel97, Excel2000, OLECtrls, OLEServer, Spin, ActiveX, ComCtrls, Math,
  Menus, ExperementDataU, ExportFunctionsU, CheckLst;

type
  TfmMain = class(TForm)
    OpnDlg: TOpenDialog;
    actlMain: TActionList;

    actLoadFromFile: TAction;
    actExportAll: TAction;
    actCalc: TAction;
    StatusBar1: TStatusBar;
    Chart1: TChart;
    pnlMark: TPanel;
    Series1: TLineSeries;
    Series2: TPointSeries;
    Series3: TLineSeries;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    gbxSettings: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label7: TLabel;
    speSampleNo: TSpinEdit;
    edEndPoint: TEdit;
    edStartPoint: TEdit;
    gbxCurrentDataSample: TGroupBox;
    lb11: TLabel;
    lb12: TLabel;
    lb13: TLabel;
    lb14: TLabel;
    gbxMeanValues: TGroupBox;
    lb21: TLabel;
    lb22: TLabel;
    lb23: TLabel;
    lb24: TLabel;
    GroupBox1: TGroupBox;
    lb31: TLabel;
    lb32: TLabel;
    lb33: TLabel;
    lb34: TLabel;
    ppmExport: TPopupMenu;
    pmiExportAll: TMenuItem;
    pmiExportOne: TMenuItem;
    actExportOne: TAction;
    edThreshold: TSpinEdit;
    gbxExperements: TGroupBox;
    ListBoxExperements: TListBox;
    actExportToTXT: TAction;
    ButtonShow: TButton;
    ButtonCalcAll: TButton;
    CheckBoxReadNewFormat: TCheckBox;
    Series4: TLineSeries;
    TabSheet3: TTabSheet;
    btn1: TButton;
    btn2: TButton;
    mm1: TMainMenu;
    N1: TMenuItem;
    Edit1: TMenuItem;
    Open1: TMenuItem;
    Export1: TMenuItem;
    Exit1: TMenuItem;
    ExportsampleExcell1: TMenuItem;
    All1: TMenuItem;
    ExportallexperimentstoTXT1: TMenuItem;
    ExportallexperimentstoExscel1: TMenuItem;
    Delete1: TMenuItem;
    Deleteall1: TMenuItem;
    grp1: TGroupBox;
    chklst1: TCheckListBox;
    movingAvgWidthSpinEdit: TSpinEdit;
    lbl1: TLabel;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    //************************************************************
    procedure speSampleNoChange(Sender: TObject);
    //************************************************************
    //************************************************************
    procedure edThresholdChange(Sender: TObject);
    procedure ListBoxExperementsClick(Sender: TObject);
    procedure ButtonShowClick(Sender: TObject);
    procedure ButtonCalcAllClick(Sender: TObject);
    procedure edStartPointExit(Sender: TObject);
    procedure ButtonDeleteAllClick(Sender: TObject);
    procedure edEndPointExit(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure btn1Click(Sender: TObject);
    procedure btn2Click(Sender: TObject);
    procedure ExportsampleExcell1Click(Sender: TObject);
    procedure All1Click(Sender: TObject);
    procedure ExportallexperimentstoTXT1Click(Sender: TObject);
    procedure ExportallexperimentstoExscel1Click(Sender: TObject);
    procedure Deleteall1Click(Sender: TObject);
    procedure Open1Click(Sender: TObject);
    procedure movingAvgWidthSpinEditChange(Sender: TObject);
 
    //************************************************************

  public

    ExperementData: TList;
    CountExperementData: Integer;
    CurrentExperementData: TExperementData;

    //************************************************************

    // відображення порогу
    procedure DrawThreshold(const Index: Integer);

    // відображення хвильового відображення
    procedure ShowSample(const Index: Integer);
    procedure ShowSquareSample(const Index: Integer);
    procedure MovingAverage(const Index: Integer);
//    function CalcSqAverage(const pIndex: Integer) : RealDoubleArray;
    //************************************************************


  end;

var
  fmMain: TfmMain;
type RealArray = array of Real;
type RealDoubleArray = array of RealArray;

implementation

{$R *.dfm}

const
  kt = 1.96; // коефіціент розподілу Стьюдента

// відображення порогу

procedure TfmMain.DrawThreshold(const Index: Integer);
begin
  Series3.Clear;
  Series3.AddXY(0, CurrentExperementData.Threshold);
  Series3.AddXY(High(CurrentExperementData.Samples[Index]), CurrentExperementData.Threshold);
end;

// відображення хвильового відображення

procedure TfmMain.ShowSample(const Index: Integer);
var
  i,j,offset: Integer;
  avg,sum: Real;
  flag: Boolean;
begin
  // попереднє очищення серій
  Series1.Clear;
  Series2.Clear;
  Series4.Clear;
  flag := false;

  for i := 0 to High(CurrentExperementData.Samples[Index]) do begin
    Series1.AddXY(i + 1, CurrentExperementData.Samples[Index, i]);

    offset:= 10;
    sum:=0;
    avg:= 0;
      for j:=i-offset to i do begin
        if j<0 then continue;
        sum := sum+ CurrentExperementData.Samples[Index, j];
      end;
      avg:=sum/offset;
      Series4.AddXY(i + 1, avg);

    if (CurrentExperementData.Samples[Index, i] = CurrentExperementData.AMax[Index]) and (not flag) and
      (i + 1 >= CurrentExperementData.StartPoint) and (i + 1 <= CurrentExperementData.EndPoint) then begin
      Series2.AddXY(i + 1, CurrentExperementData.Samples[Index, i], 'x=' + IntToStr(i) + ' y=' + FloatToStr(CurrentExperementData.Samples[Index, i]));
      flag := true;
    end;
  end;

  if (CurrentExperementData.TrueValues[Index] = 0)
    then pnlMark.Color := clGreen
  else pnlMark.Color := clRed;
end;

function CalcSqAverage(data: TExperementData; const pIndex: Integer; const pMovingAverageWidth: Integer) : RealDoubleArray;
var avgArray: RealDoubleArray;
  i,j,start,k,NN: Integer;
  summ: Double;
  avg,val: Double;
  picksY : RealArray;
  picksX : RealArray;
  sqY: RealArray;
  aa: RealDoubleArray;
BEGIN
      k:=0;
      NN := pMovingAverageWidth;
      for i:=0 to High(data.Samples[pIndex]) do
      begin
           start:= i * NN;
           for j:= start to start + NN - 1 do
           begin
               if j>High(data.Samples[pIndex])-1 then break;
                    k:=k+1;
                    SetLength(picksX,k+1);
                    SetLength(picksY,k+1);
                    picksY[k]:= data.Samples[pIndex, j];
                    picksX[k]:=j;
           end;
      end;

      for i:=0 to High(picksY)-1 do
      begin
           summ:=0;
           avg:=0;
           for j:= i to i+NN do
           begin
             if j>High(picksX) then break;
             if j>High(picksY) then break;

             summ := summ + Power(picksY[j],2);
           end;
           summ := summ /NN;
           val := sqrt(summ);

           SetLength(sqY, i+1);
           sqY[i]:= val * Sqrt(2);
      end;

      SetLength(Result, 2);
      Result[0]:= picksX;
      Result[1]:= sqY
END;

procedure TfmMain.ShowSquareSample(const Index: Integer);
var i,j,start,k,NN: Integer;
x, avgSQY, avgY, y1, y2, y3, y4, y5, y6, y7, y0: RealArray;
y: RealDoubleArray;
S: Real;
begin
  // попереднє очищення серій
  Series1.Clear;
  Series2.Clear;
  Series3.Clear;
  Series4.Clear;

  x:= CalcSqAverage(CurrentExperementData, 0, CurrentExperementData.MovingAverageFrqmeWidth)[0];
  SetLength(y, CurrentExperementData.SamplesCount);

  for i:=0 to CurrentExperementData.SamplesCount-1 do
  begin
    y[i]:= CalcSqAverage(CurrentExperementData, i, CurrentExperementData.MovingAverageFrqmeWidth)[1];
  end;

  SetLength(avgY,High(CurrentExperementData.Samples[0]));
  S:=0;
  for i:=0 to High(CurrentExperementData.Samples[0]) - 5 do
  begin
    S:=0;
    for j:=0 to CurrentExperementData.SamplesCount - 1 do
    begin
     S:= S + CurrentExperementData.Samples[j,i];
    end;
    avgY[i]:= S / CurrentExperementData.SamplesCount;
    Series1.AddXY(i, avgY[i]);
  end;

  SetLength(avgSQY,High(y[0]));
  S:=0;
  for i:=0 to High(y[0])-5 do
  begin
    S:=0;
    for j:=0 to CurrentExperementData.SamplesCount-1 do
    begin
     S:= S + y[j,i];
    end;
    avgSQY[i]:=S / CurrentExperementData.SamplesCount;
    Series4.AddXY(x[i], avgSQY[i]);
  end;

end;

procedure TfmMain.MovingAverage(const Index: Integer);
var
  i,j,NN: Integer;
  flag: Boolean;
  summ: Double;
  avg: Double;
begin
  // попереднє очищення серій
  Series1.Clear;
  Series2.Clear;
  Series3.Clear;
  Series4.Clear;
  flag := false;

  for i:=0 to High(CurrentExperementData.Samples[Index]) do
  begin
       NN:=5;
       summ:=0;
       avg:=0;
       for j:= i to i+NN do
       begin
          summ := summ + CurrentExperementData.Samples[Index, j];
       end;
       avg:= summ / NN;
       if (avg>2000) then
       begin
        avg :=2000;
       end;
       if (avg<-2000) then avg:=-2000;
       Series4.AddXY(i + 1, avg);
       Series1.AddXY(i + 1, CurrentExperementData.Samples[Index, i]);
  end
end;


procedure TfmMain.FormClose(Sender: TObject; var Action: TCloseAction);
var i: integer;
begin
  for i := 0 to CountExperementData - 1 do begin
    TExperementData(ExperementData[i]).Destroy;
  end;
  Action := caFree;
end;


// експорт в Excel

procedure TfmMain.speSampleNoChange(Sender: TObject);
begin
  exit;
  ShowSample(speSampleNo.Value - 1);
  // відображення даних поточної виборки
  lb11.Caption := 'Амплітуда = ' + FloatToStr(CurrentExperementData.AMax[speSampleNo.Value - 1]);
  lb12.Caption := 'Сума амплітуд = ' + FloatToStr(CurrentExperementData.ASum[speSampleNo.Value - 1]);
  lb13.Caption := 'Підсумковий рахунок = ' + FloatToStr(CurrentExperementData.NSum[speSampleNo.Value - 1]);
  lb14.Caption := 'Сума квадратів амплітуд = ' + FloatToStr(CurrentExperementData.ASumDivNSum[speSampleNo.Value - 1]);

  ButtonCalcAll.Font.Style := [fsBold];
end;

procedure TfmMain.edThresholdChange(Sender: TObject);
begin
  try
    CurrentExperementData.Threshold := edThreshold.Value;
  except
    Exit;
  end;
  ButtonCalcAll.Font.Style := [fsBold];
  //DrawThreshold(speSampleNo.Value-1);
end;

procedure TfmMain.ListBoxExperementsClick(Sender: TObject);
begin
  CurrentExperementData := TExperementData(ExperementData[ListBoxExperements.ItemIndex]);
  speSampleNo.Value := 1;
  speSampleNo.MaxValue := CurrentExperementData.SamplesCount;
end;

procedure TfmMain.ButtonShowClick(Sender: TObject);
begin
  // відображення даних поточної виборки
  lb11.Caption := 'Амплітуда = ' + FloatToStr(CurrentExperementData.AMax[speSampleNo.Value - 1]);
  lb12.Caption := 'Сума амплітуд = ' + FloatToStr(CurrentExperementData.ASum[speSampleNo.Value - 1]);
  lb13.Caption := 'Підсумковий рахунок = ' + FloatToStr(CurrentExperementData.NSum[speSampleNo.Value - 1]);
  lb14.Caption := 'Сума квадратів амплітуд = ' + FloatToStr(CurrentExperementData.ASumDivNSum[speSampleNo.Value - 1]);

  // відображення усереднених значень
  lb21.Caption := 'Амплітуда = ' + FloatToStr(CurrentExperementData.AMaxMeanValue);
  lb22.Caption := 'Сума амплітуд = ' + FloatToStr(CurrentExperementData.ASumMeanValue);
  lb23.Caption := 'Підсумковий рахунок = ' + FloatToStr(CurrentExperementData.NSumMeanValue);
  lb24.Caption := 'Сума квадратів амплітуд = ' + FloatToStr(CurrentExperementData.ASumDivNSumMeanValue);

  // відображення довірчих інтервалів
  lb31.Caption := 'Амплітуда = [' +
    FloatToStrF(CurrentExperementData.AMaxMeanValue - kt * CurrentExperementData.AMaxRMS / Sqrt(CurrentExperementData.k), ffFixed, 6, 1) + '; ' +
    FloatToStrF(CurrentExperementData.AMaxMeanValue + kt * CurrentExperementData.AMaxRMS / Sqrt(CurrentExperementData.k), ffFixed, 6, 1) + ']';
  lb32.Caption := 'Сума амплітуд = [' +
    FloatToStrF(CurrentExperementData.ASumMeanValue - kt * CurrentExperementData.ASumRMS / Sqrt(CurrentExperementData.k), ffFixed, 6, 1) + '; ' +
    FloatToStrF(CurrentExperementData.ASumMeanValue + kt * CurrentExperementData.ASumRMS / Sqrt(CurrentExperementData.k), ffFixed, 6, 1) + ']';
  lb33.Caption := 'Підсумковий рахунок = [' +
    FloatToStrF(CurrentExperementData.NSumMeanValue - kt * CurrentExperementData.NSumRMS / Sqrt(CurrentExperementData.k), ffFixed, 6, 1) + '; ' +
    FloatToStrF(CurrentExperementData.NSumMeanValue + kt * CurrentExperementData.NSumRMS / Sqrt(CurrentExperementData.k), ffFixed, 6, 1) + ']';
  lb34.Caption := 'Сума квадратів амплітуд = [' +
    FloatToStrF(CurrentExperementData.ASumDivNSumMeanValue - kt * CurrentExperementData.ASumDivNSumRMS / Sqrt(CurrentExperementData.k), ffFixed, 6, 1) + '; ' +
    FloatToStrF(CurrentExperementData.ASumDivNSumMeanValue + kt * CurrentExperementData.ASumDivNSumRMS / Sqrt(CurrentExperementData.k), ffFixed, 6, 1) + ']';

  // відображення порогу і хвильового відображення
  ShowSample(speSampleNo.Value - 1);
  DrawThreshold(speSampleNo.Value - 1);
end;

procedure TfmMain.ButtonCalcAllClick(Sender: TObject);
var
  i: Integer;
begin

  for i := 0 to ExperementData.Count - 1 do
  begin

                // calculate zero line for each sample and update data to nulling average
    TExperementData(ExperementData[i]).updateToZeroLine();
                // пошук макс амплітуд, підсумкових рахунків, суми амплітуд тощо
                //TExperementData(ExperementData[i]).PrepareData(speSampleNo.Value, StrToInt(edStartPoint.Text), StrToInt(edEndPoint.Text), StrToInt(edThreshold.Text));
    TExperementData(ExperementData[i]).PrepareData(StrToInt(edStartPoint.Text), StrToInt(edEndPoint.Text), StrToInt(edThreshold.Text));

                // обчислення усереднених значень
    TExperementData(ExperementData[i]).CalcAverage;

                // обчислення середнього квадратичного відхилення
    TExperementData(ExperementData[i]).CalcRMSDeviation;

    TExperementData(ExperementData[i]).k := TExperementData(ExperementData[i]).SamplesCount;
  end;





  ButtonShow.Enabled := True;

  ButtonCalcAll.Font.Style := [];

end;

procedure TfmMain.edStartPointExit(Sender: TObject);
var
  a: integer;
begin
  try
    a := StrToInt(edStartPoint.Text);
  except
    a := 1;
  end;

  if ((a < 1)) then begin
    a := 1;
  end;

  edStartPoint.Text := IntToStr(a);
  ButtonCalcAll.Font.Style := [fsBold];
end;

procedure TfmMain.ButtonDeleteAllClick(Sender: TObject);
var
  i: integer;
begin
  ListBoxExperements.ItemIndex := -1;
  for i := ListBoxExperements.Count - 1 downto 0 do begin
    TExperementData(ExperementData[i]).Destroy;
    ExperementData.Delete(i);
    ListBoxExperements.Items.Delete(i);
  end;

  ButtonCalcAll.Enabled := False;
  ButtonShow.Enabled := False;

  speSampleNo.Enabled := False;
  edThreshold.Enabled := False;
  edStartPoint.Enabled := False;
  edEndPoint.Enabled := False;
end;

procedure TfmMain.edEndPointExit(Sender: TObject);
begin
  ButtonCalcAll.Font.Style := [fsBold];
end;

procedure TfmMain.Button1Click(Sender: TObject);
begin
  // відображення порогу і хвильового відображення

  //DrawThreshold(speSampleNo.Value - 1);
end;

procedure TfmMain.btn1Click(Sender: TObject);
begin
  MovingAverage(speSampleNo.Value - 1);
end;

procedure TfmMain.btn2Click(Sender: TObject);
begin
  ShowSquareSample(speSampleNo.Value - 1);
end;

procedure TfmMain.ExportsampleExcell1Click(Sender: TObject);
begin
  ExportOneToExcel(ExperementData, ListBoxExperements.ItemIndex, speSampleNo.Value);
end;

procedure TfmMain.All1Click(Sender: TObject);
begin
   ExportAllToExcel(ExperementData, ListBoxExperements.ItemIndex);
end;

procedure TfmMain.ExportallexperimentstoTXT1Click(Sender: TObject);
begin
   ExportAllToTXT(ExperementData);
end;

procedure TfmMain.ExportallexperimentstoExscel1Click(Sender: TObject);
begin
   ExportAllExperementsToExcel(ExperementData);
end;

procedure TfmMain.Deleteall1Click(Sender: TObject);
begin
  if not (ListBoxExperements.Count = 0) then begin
    CurrentExperementData.Destroy;
    ExperementData.Delete(ListBoxExperements.ItemIndex);
    ListBoxExperements.Items.Delete(ListBoxExperements.ItemIndex);

    ListBoxExperements.ItemIndex := ListBoxExperements.Count - 1;
  end;
  if (ListBoxExperements.Count = 0) then begin

    ButtonCalcAll.Enabled := False;
    ButtonShow.Enabled := False;

    speSampleNo.Enabled := False;
    edThreshold.Enabled := False;
    edStartPoint.Enabled := False;
    edEndPoint.Enabled := False;
  end;
end;

procedure TfmMain.Open1Click(Sender: TObject);
var TmpStr: string; i: Integer;
begin
  OpnDlg.InitialDir := ExtractFilePath(Application.ExeName);

  //====================================
  //====== ЧИТАННЯ ДАНИХ З ФАЙЛУ =======
  //====================================

  if (OpnDlg.Execute) then begin
    try
      if (ExperementData = nil) then
        ExperementData := TList.Create;

        //TStringList(OpnDlg.Files).Sort;

      for i := 0 to OpnDlg.Files.Count - 1 do begin
        ExperementData.Add(TExperementData.Create);
        CurrentExperementData := TExperementData(ExperementData.Last);
        if (CheckBoxReadNewFormat.Checked = false) then
                CurrentExperementData.LoadFromNewFile(OpnDlg.Files[i])
        else
                CurrentExperementData.LoadFromFile(OpnDlg.Files[i]);
        ListBoxExperements.Items.Add(CurrentExperementData.FileName);
      end;


    except
      on E: Exception do
                //Application.MessageBox(E.Message,'Помилка', mb_OK);
        Exit;
    end;

    ListBoxExperements.ItemIndex := ExperementData.Count - 1;

    speSampleNo.MinValue := 1;
    speSampleNo.MaxValue := CurrentExperementData.SamplesCount;
    speSampleNo.Value := 1;

    // "витягуємо" ім'я файлу
    TmpStr := ExtractFileName(OpnDlg.FileName);
    SetLength(TmpStr, Length(TmpStr) - 4);

    fmMain.Caption := 'МАЕ калібрування | файл: ' + TmpStr + ' | к-сть виборок: ' + InttoStr(CurrentExperementData.SamplesCount);


    ButtonShow.Enabled := False;

    // unblock calculation functions
    ButtonCalcAll.Enabled := true;
    ButtonCalcAll.Font.Style := [fsBold];

    speSampleNo.Enabled := true;
    edThreshold.Enabled := true;
    edStartPoint.Enabled := true;
    edEndPoint.Enabled := true;

  end;
end;

procedure TfmMain.movingAvgWidthSpinEditChange(Sender: TObject);
begin
  CurrentExperementData.MovingAverageFrqmeWidth := movingAvgWidthSpinEdit.Value
end;



end.
