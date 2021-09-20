unit WordReport;

{=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-\
|  WordReport - Компонент для автоматизации создания отчетов в MS Word.        |
|                                                                              |
|  Модуль: WordReport                                                          |
|  Автор: Власов Александр Сергеевич (velesov7493@yandex.ru)                   |
|  Copyright: 2012 Власов Александр Сергеевич                                  |
|  Дата: 10.11.2012                                                            |
|                                                                              |
|  Описание:                                                                   |
|                                                                              |
|  При запуске метода Build экземпляр класса TWordReport достает все имена     |
|  переменных вне секций, имена секций, имена всех переменных в каждой секции  |
|  из шаблона, указанного в TemplateDocFileName. Далее вызывается обработчик   |
|  OnReadMaket для привязки данных к этим именам. Затем данные отправляются    |
|  в готовый отчет, который сохраняется в документе с именем ResultDocFileName |
|                                                                              |
\=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-}

interface

uses
  DB, ComObj, Classes;

const
  MaxFieldsOnBand = 16;
  MaxBandsOnReport = 8;

  bvVariable = 0;
  bvField = 1;
  bvFormat = 2;

type

  PDataSet = ^TDataSet;
  TBandVariables = array [0..2,1..MaxFieldsOnBand] of string;

  PDataBand = ^TDataBand;
  TDataBand = class
  private
    fName: string;
    fRowBookmark: string;                                       // Имя закладки для секции документа
    fVariableNames: TBandVariables;                             // Соответствие между именами полей Dataset`а и именами переменных
    fVariableCount: integer;                                    // Количество переменных в повторяемой секции
    fCounterName: string;                                       // Имя счетчика в секции документа
    fCounterVal: integer;                                       // Значение счетчика секции
    fDataSet: PDataSet;                                         // Указатель на DataSet
    fTableNumber: integer;                                      // Номер таблицы документа, которой принадлежит секция
    fKeyName: string;                                           // Имя сигнального поля НД для группы секций
    fKeyValue: integer;                                         // Значение сигнального поля НД для этой секции. Также используется как уровень вложенности
    fNextBand: PDataBand;                                       // Указатель на следующую секцию в группе
    function FindVariable(VarName:string): integer;
    function VarToPattern(VarIndex:integer):string;
    function getTableNumberByBookmark: integer;
    procedure Prepare;
    procedure PokeData;
  protected
    function GetDimensionByVarName(VarName: string; Index: integer): string;
    procedure SetDimensionByVarName(VarName: string; Index: integer; Value: string);
    property TableNumber: integer read fTableNumber write fTableNumber;
    property KeyName: string read fKeyName write fKeyName;
    property KeyValue: integer read fKeyValue write fKeyValue;
    property NextBand: PDataBand read fNextBand write fNextBand;
  public
    constructor Create(aBookmark,aName,aCounterName:string);
    procedure AssignDataSet(aDataSet: PDataSet);
    procedure SetField(VariableName,FieldName,Format:string);

    property Name: string read fName write fName;
    property Field[VarName:string]: string index 1 read GetDimensionByVarName write SetDimensionByVarName;
    property Format[VarName:string]: string index 2 read GetDimensionByVarName write SetDimensionByVarName;
  end;

  TWordReport = class(TComponent)
  private
    fDataBands: array [1..MaxBandsOnReport] of TDataBand;
    fBandCount: integer;
    fTemplateDocFileName: string;
    fResultDocFileName: string;
    fShowResult: boolean;
    fOnReadMaket: TNotifyEvent;
    procedure ReadVarNames;
    procedure ReadBandNames(Pattern:string; var Band,Counter,Variable: string);
    procedure NewBandVar(aBookmark,aBandName,aCounterName,aVariableName: string);
    procedure CreateMultiBandMaket(Band1,Band2:TDataBand);
    procedure StartProgress;
  protected
    fReportVars: TParams;
    function GetReportDifficulty: integer;
    function GetBandByName(BandName:string):TDataBand;
    function GetBandNumber(BandName:string):integer;
    procedure SetTemplateFileName(Value: string);
    procedure BindVariables;
    procedure ReadMaket;
    procedure SaveToFile(FileName: string);
    procedure CloseDocument;
  public
    procedure JoinBands(BandKeyField:string; BandName1:string; KeyValue1:integer; BandName2:string; KeyValue2:integer);
    procedure SetValue(VariableName:string; Value:Variant);
    procedure Build;
    procedure Quit;
    function BandExists(BandName: string): boolean;
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; reintroduce; overload;
    property Bands[Name:string]: TDataBand read GetBandByName;
    property BandCount: integer read fBandCount;
  published
    property TemplateDocFileName: string read fTemplateDocFileName write SetTemplateFileName;
    property ResultDocFileName: string read fResultDocFileName write fResultDocFileName;
    property ShowResult: boolean read fShowResult write fShowResult default true;
    property OnReadMaket: TNotifyEvent read fOnReadMaket write fOnReadMaket;
  end;

procedure Register;

implementation

uses
  SysUtils, StrUtils, Windows, Messages, Variants, Graphics,
  Controls, Forms, Dialogs, ComCtrls;

type
  TProgressForm = class(TForm)
    ProgressBar1: TProgressBar;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
    procedure DoProgress(Step: integer);
  public
    { Public declarations }
  end;

const
  wdCollapseEnd = 0;
  wdCollapseStart = 1;
  wdFindContinue = 1;

var
  wProgress: TProgressForm;     // переменная формы прогресса
  curWordApp: Variant;          // OLE-объект Word
  curDocument: Variant;         // OLE-объект документа

// Извлечь имя свободной переменной из его шаблона

function PatternToName(ParamPattern:string):string;
var
  i1,i2: integer;
begin
  i1:=Pos('#(',ParamPattern)+2;
  i2:=Length(ParamPattern);
  Result:=MidStr(ParamPattern,i1,i2-i1);
end;

// Получить шаблон из имени

function NameToPattern(ParamName:string):string;
begin
  Result:='#('+ParamName+')';
end;

// Подготовить OLE-объект поиска Word
// Вход:
//   FindText - что искать
//   RegExp - Считать FindText регулярным выражением Word
// Выход:
//   FindObj - подготовленный объект поиска

procedure PrepareFindObject(FindText:string; RegExp:boolean; var FindObj:Variant);
begin
  FindObj:=curWordApp.Selection.Find;
  FindObj.ClearFormatting;
  FindObj.Text:= FindText;
  FindObj.Replacement.Text:='';
  FindObj.Forward:=true;
  FindObj.Wrap:=wdFindContinue;
  FindObj.Format:=false;
  FindObj.MatchCase:=false;
  FindObj.MatchWholeWord:=false;
  FindObj.MatchAllWordForms:=false;
  FindObj.MatchSoundsLike:=false;
  FindObj.MatchWildcards:=RegExp;
end;

//==============================================================================
// TProgress Form
//==============================================================================

{$R wrtprogress.dfm}

procedure TProgressForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action:=caFree;
end;

procedure TProgressForm.DoProgress(Step: integer);
begin
  with wProgress.ProgressBar1 do Position:=Position+Step;
end;

//==============================================================================
// TDataBand
//==============================================================================

constructor TDataBand.Create(aBookmark:string; aName: string; aCounterName: string);
begin
  Name:=aName;
  fCounterName:=aCounterName;
  fRowBookmark:=aBookmark;
  fVariableCount:=0;
  fCounterVal:=1;
  fDataSet:=nil;
end;

// Найти переменную секции, т.е. получить ее индекс
// Возвращает -1, если такой переменной нет

function TDataBand.FindVariable(VarName: string):integer;
var
  i,r: integer;
begin
  r:=-1;
  for i:=1 to fVariableCount do
    if fVariableNames[bvVariable,i]=VarName then begin
      r:=i;
      Break;
    end;
  Result:=r;
end;

// Получить номер таблицы документа, в которой находится
// диапазон закладки секции

function TDataBand.getTableNumberByBookmark: integer;
var
  i,tableStart,tableNumber: integer;
  r,table: Variant;
begin
  tableNumber:=0;
  if not curDocument.Bookmarks.Exists(fRowBookmark) then begin
    Result:=tableNumber;
    Exit;
  end;
  r:=curDocument.Bookmarks.Item(fRowBookmark).Range;
  if r.Tables.Count>0 then begin
    table:=r.Tables.Item(1);
    if table.NestingLevel<>1 then begin
      Result:=tableNumber;
      Exit;
    end;
    tableStart:=table.Range.Start;
    for i:=1 to curDocument.Tables.Count do
      if curDocument.Tables.Item(i).Range.Start=tableStart then begin
        tableNumber:=i;
        Break;
      end;
  end else raise Exception.Create('Закладка не является частью таблицы');
  Result:=tableNumber;
end;

// Получить параметр переменной по имени переменной и индексу этого параметра

function TDataBand.GetDimensionByVarName(VarName: string; Index: integer): string;
var
  VarIndex: integer;
begin
  VarIndex:=FindVariable(VarName);
  if index>0 then Result:=fVariableNames[Index,VarIndex] else Result:='';
end;

// Установить параметр переменной по имени переменной и индексу этого параметра

procedure TDataBand.SetDimensionByVarName(VarName: string; Index: integer; Value: string);
var
  VarIndex: integer;
begin
  VarIndex:=FindVariable(VarName);
  if index<0 then Exit;
  fVariableNames[Index,VarIndex]:=Value;
end;

// Получить шаблон переменной секции по индексу
// индекс=0 у счетчика записей секции

function TDataBand.VarToPattern(VarIndex:integer):string;
var r: string;
begin
  if VarIndex=0 then r:='#('+fCounterName+')' else
    r:='#('+Name+'('+fCounterName+').'+fVariableNames[bvVariable,VarIndex]+')';
  Result:=r;
end;

// Привязать поле набора данных и его формат
// к переменной VariableName

procedure TDataBand.SetField(VariableName,FieldName,Format:string);
var
  VarIndex: integer;
begin
  VarIndex:=FindVariable(VariableName);
  if VarIndex<0 then begin
    raise Exception.Create('Переменная "'+VariableName+'" отсутствует в секции "'+fName+'" документа');
    Exit;
  end;
  fVariableNames[bvField,VarIndex]:=FieldName;
  fVariableNames[bvFormat,VarIndex]:=Format;
end;

// "Втолкнуть" данные
// Отправляет все данные секции в документ

var
  // Прежнее значение сигнального поля
  OldKeyValue: integer;

procedure TDataBand.PokeData;
label
  ValueDefinedS,ValueDefinedM;
var
  Pattern,Value: string;
  cField: TField;
  FindObj: Variant;
  i,j,c: integer;
begin
  if fDataSet=nil then Exit;
  if Length(fKeyName)=0 then begin
    // Если секция не состоит в группе, то обрабатываем ее обыкновенно
    fDataSet.First;
    for i:=1 to fDataset.RecordCount do begin
      Pattern:=VarToPattern(0);
      PrepareFindObject(Pattern,false,FindObj);
      if FindObj.Execute then begin
        Value:=IntToStr(fDataset.RecNo)+'.';
        curWordApp.Selection.Text:=Value;
      end;
      for j:=1 to fVariableCount do begin
        Pattern:=VarToPattern(j);
        PrepareFindObject(Pattern,false,FindObj);
        if FindObj.Execute then begin
          cField:=fDataset.FieldByName(fVariableNames[bvField,j]);
          if cField.IsNull then begin
            Value:='';
            goto ValueDefinedS;
          end;
          if Length(fVariableNames[bvFormat,j])<>0 then begin
            Value:=Trim(SysUtils.Format(fVariableNames[bvFormat,j],[cField.AsFloat]));
            goto ValueDefinedS;
          end;
          Value:=cField.AsString;
ValueDefinedS:
          curWordApp.Selection.Text:=Value;
        end;
      end;
      fDataSet.Next;
      wProgress.DoProgress(1);
    end;
  end else begin
    // Иначе обрабатываем только записи с нашим значением сигнального поля
    c:=fDataset.FieldByName(fKeyName).AsInteger;
    // Если уровень вложенности больше прежнего, то сбросить счетчик
    if fKeyValue>OldKeyValue then fCounterVal:=1;
    while (c=fKeyValue) and (not fDataset.Eof) do begin
      Pattern:=VarToPattern(0);
      PrepareFindObject(Pattern,false,FindObj);
      if FindObj.Execute then begin
        Value:=IntToStr(fCounterVal)+'.';
        curWordApp.Selection.Text:=Value;
      end;
      for j:=1 to fVariableCount do begin
        Pattern:=VarToPattern(j);
        PrepareFindObject(Pattern,false,FindObj);
        if FindObj.Execute then begin
          cField:=fDataset.FieldByName(fVariableNames[bvField,j]);
          if cField.IsNull then begin
            Value:='';
            goto ValueDefinedM;
          end;
          if Length(fVariableNames[bvFormat,j])<>0 then begin
            Value:=Trim(SysUtils.Format(fVariableNames[bvFormat,j],[cField.AsFloat]));
            goto ValueDefinedM;
          end;
          Value:=cField.AsString;
ValueDefinedM:
          curWordApp.Selection.Text:=Value;
        end;
      end;
      fDataset.Next;
      wProgress.DoProgress(1);
      OldKeyValue:=c;
      c:=fDataset.FieldByName(fKeyName).AsInteger;
      Inc(fCounterVal);
    end;
  end;
end;

// Подготовить секцию для вывода записей
// Размножает секцию, чтобы она повторялась столько раз, сколько записей в НД

procedure TDataBand.Prepare;
begin
  if fDataSet=nil then Exit;
  if curDocument.Bookmarks.Exists(fRowBookmark) then begin
    curDocument.Bookmarks.Item(fRowBookmark).Range.Copy;
    fDataset.First;
    while not fDataset.Eof do begin
      if fTableNumber<>0 then begin
        curDocument.Tables.Item(fTableNumber).Select;
        curWordApp.Selection.Collapse(wdCollapseEnd);
      end else
        curDocument.Bookmarks.Item(fRowBookmark).Range.Select;
      curWordApp.Selection.Paste;
      fDataset.Next;
    end;
    fDataset.First;
    curDocument.Bookmarks.Item(fRowBookmark).Range.Select;
    curWordApp.Selection.Cut;
  end;
end;

// Подключить НД к секции

procedure TDataBand.AssignDataSet(aDataSet: PDataSet);
var
  tn: integer;
begin
  if not (aDataset^ is TDataset) then begin
    raise Exception.Create('Неправильный указатель на TDataset');
    Exit;
  end;
  fDataSet:=aDataSet;
  tn:=getTableNumberByBookmark;
  if tn>0 then fTableNumber:=tn;
end;

//==============================================================================
// TWordReport
//==============================================================================

constructor TWordReport.Create(AOwner:TComponent);
begin
  inherited Create(AOwner);
  fReportVars:=TParams.Create;
  fShowResult:=true;
end;

// Закрыть активный документ

procedure TWordReport.CloseDocument;
var
  i: integer;
begin
  for i:=1 to fBandCount do fDataBands[i].Free;
  fBandCount:=0;
  fReportVars.Clear;
  curDocument.Close;
  wProgress.DoProgress(10);
end;

// Завершить Word, если он невидим или показать - если видим.
// Закрыть и освободить окно прогресса

procedure TWordReport.Quit;
begin
  if fShowResult then begin
    curWordApp.Selection.Collapse(wdCollapseEnd);
    curWordApp.Visible:=true;
  end else
    curWordApp.Quit;
  wProgress.DoProgress(40);
  wProgress.Close;
end;

destructor TWordReport.Destroy;
begin
  if Assigned(fReportVars) then fReportVars.Free;
  inherited Destroy;
end;

// Отрыть шаблон, указанный в TemplateDocFileName в MS Word

procedure TWordReport.ReadMaket;
var
  Error,verS: string;
  version: real;
  cpos: integer;
begin
  Error:='';
  try
    curWordApp:=CreateOleObject('Word.Application');
    verS:=curWordApp.Version;
  except
    Error:='MS Word недоступен';
  end;
  Val(verS,version,cpos);
  if (cpos=0) and (version>=9) then begin
    fBandCount:=0;
    curWordApp.Visible:=false;
  end else Error:=Error+#10#13+'Требуется MS Word 2000 и выше';
  try
    curDocument:=curWordApp.Documents.Add(fTemplateDocFileName);
  except
    Error:='Не удалось открыть шаблон: "'+fTemplateDocFileName+'"'+#10#13#10#13+Error;
  end;
  if Length(Error)=0 then ReadVarNames else raise Exception.Create(Error);
end;

// Установить документ шаблона

procedure TWordReport.SetTemplateFileName(Value: string);
begin
  if fTemplateDocFileName<>Value then fTemplateDocFileName:=Value;
end;

// Подсчитать сложность отчета
// = КоличествоСвобПеременных
// +КоличествоЗаписейВоВсехНД
// +10 на закрытие документа
// +10 на сохранение документа
// +40 на завершение MS Word

function TWordReport.GetReportDifficulty: integer;
var
  i,s: integer;
  SkipBands: set of Byte;
begin
  s:=fReportVars.Count;
  SkipBands:=[];
  for i:=1 to fBandCount do begin
    if i in SkipBands then Continue;
    s:=s+fDataBands[i].fDataSet^.RecordCount;
    if fDataBands[i].NextBand<>nil then
      SkipBands:=SkipBands+[GetBandNumber(fDataBands[i].fNextBand^.fName)];
  end;
  s:=s+60;
  Result:=s;
end;

// Прочитать имена секции, счетчика и переменной
// из объявления переменной секции в шаблоне

procedure TWordReport.ReadBandNames(Pattern:string; var Band,Counter,Variable: string);
var
  i1,i2,c: integer;
  s: string;
begin
  s:=PatternToName(Pattern);
  i1:=1; i2:=Pos('(',s); c:=i2-i1;
  Band:=MidStr(s,i1,c);
  i1:=i2+1; i2:=Pos(')',s); c:=i2-i1;
  Counter:=MidStr(s,i1,c);
  i1:=i2+2; i2:=Length(s); c:=i2-i1+1;
  Variable:=MidStr(s,i1,c);
end;

// Создать новую переменную секции

procedure TWordReport.NewBandVar(aBookmark: string; aBandName: string; aCounterName: string; aVariableName: string);
var
  cBand: TDataBand;
  cParam: TParam;
begin
  cBand:=Bands[aBandName];
  if cBand<>nil then begin
    if cBand.FindVariable(aVariableName)<0 then with cBand do begin
      Inc(fVariableCount);
      fVariableNames[bvVariable,fVariableCount]:=aVariableName;
    end;
  end else begin
    Inc(fBandCount);
    fDataBands[fBandCount]:=TDataBand.Create(aBookmark,aBandName,aCounterName);
    with fDataBands[fBandCount] do begin
      Inc(fVariableCount);
      fVariableNames[bvVariable,fVariableCount]:=aVariableName;
      cParam:=fReportVars.FindParam(aCounterName);
      if cParam<>nil then fReportVars.RemoveParam(cParam);
    end;
  end;
end;

// Установить значение переменной вне секций

procedure TWordReport.SetValue(VariableName:string; Value:Variant);
begin
  if fReportVars.FindParam(VariableName)<>nil then
    fReportVars.ParamByName(VariableName).Value:=Value
  else raise Exception.Create('Переменная "'+VariableName+'" отсутствует в документе');
end;

// Получить секцию по ее имени

function TWordReport.GetBandByName(BandName:string): TDataBand;
var
  i: integer;
begin
  Result:=nil;
  for i:=1 to fBandCount do if fDataBands[i].Name=BandName then begin
    Result:=fDataBands[i];
    Break;
  end;
end;

// Получить номер секции по ее имени

function TWordReport.GetBandNumber(BandName: string): integer;
var
  i:integer;
begin
  Result:=-1;
  for i:=1 to fBandCount do if fDataBands[i].Name=BandName then begin
    Result:=i;
    Break;
  end;
end;

// Начать отсчет прогресса сборки

procedure TWordReport.StartProgress;
begin
  with wProgress.ProgressBar1 do begin
    Min:=0;
    Max:=GetReportDifficulty;
    Position:=Min;
  end;
  wProgress.Show;
end;

// Прочитать имена из шаблона

procedure TWordReport.ReadVarNames;
var
  ParamName: string;
  cParam: TParam;
  cBmName,cBandName,cCounterName,cVarName: string;
  FindObj: Variant;
begin
  PrepareFindObject('\#[\(]{1}[A-z.0-9]@[\)]{1}',true,FindObj);
  while FindObj.Execute do begin
    ParamName:=PatternToName(curWordApp.Selection.Text);
    if fReportVars.FindParam(ParamName)=nil then begin
      cParam:=fReportVars.CreateParam(ftString,ParamName,ptUnknown);
      fReportVars.AddParam(cParam);
    end;
  end;

  curDocument.Range.Select;
  curWordApp.Selection.Collapse(wdCollapseStart);

  PrepareFindObject('\#[\(]{1}[A-z0-9]@[\(]{1}[A-z0-9]@[\)]{1}.[A-z0-9]@[)]{1}',true,FindObj);
  while FindObj.Execute do begin
    if curWordApp.Selection.BookmarkID>0 then begin
      cBmName:=curDocument.Bookmarks.Item(curWordApp.Selection.BookmarkID).Name;
      if (LowerCase(LeftStr(cBmName,4))<>'data') or (Pos(cBmName[5],'12345678')=0) then
        raise Exception.Create('Имя закладки секции не соответствует спецификации');
    end else
      raise Exception.Create('Переменная "'+curWordApp.Selection.Text+'" объявлена вне закладки');
    ReadBandNames(curWordApp.Selection.Text,cBandName,cCounterName,cVarName);
    NewBandVar(cBmName,cBandName,cCounterName,cVarName);
  end;
end;

// Отправить данные, привязанные пользователем, в документ

procedure TWordReport.BindVariables;
var
  FindObj: Variant;
  ParamPattern: string;
  i: integer;
  SkipBands: set of Byte;
begin
  StartProgress;
  for i:=0 to fReportVars.Count-1 do begin
    ParamPattern:=NameToPattern(fReportVars.Items[i].Name);
    PrepareFindObject(ParamPattern,false,FindObj);
    while FindObj.Execute do
      curWordApp.Selection.Text:=fReportVars.Items[i].AsString;
    wProgress.DoProgress(1);
  end;

  SkipBands:=[];
  for i:=1 to fBandCount do
    if (fDataBands[i].fNextBand<>nil) then begin
      // Если секция в группе, то выводить ее вместе со связанной секцией
      if (i in SkipBands) then Continue;
      CreateMultiBandMaket(fDataBands[i],fDataBands[i].fNextBand^);
      while not fDataBands[i].fDataSet.Eof do begin
        fDataBands[i].PokeData;
        fDataBands[i].fNextBand^.PokeData;
      end;
      SkipBands:=SkipBands+[GetBandNumber(fDataBands[i].fNextBand^.fName)];
    end else
      // Иначе, просто подготовить ...
      fDataBands[i].Prepare;

  for i:=1 to fBandCount do
    // ... и вывести ее обыкновенно
    if fDataBands[i].fNextBand=nil then fDataBands[i].PokeData;
end;

// Сохранить документ в файл

procedure TWordReport.SaveToFile(FileName: string);
begin
  curDocument.SaveAs(FileName);
  wProgress.DoProgress(10);
end;

// Объединить две секции в группу
// так чтобы они чередуясь выводили записи из одного и того же НД.

// Например: Объединить секции с именами 'R' и 'A', использовав поле 'GRID'
// в качестве ключа. При этом: если GRID=0 - вывести запись в секции R, а если
// GRID=1 - в секции A.
// JoinBands('GRID','R',0,'A',1);

procedure TWordReport.JoinBands(BandKeyField:string; BandName1:string; KeyValue1:integer; BandName2:string; KeyValue2:integer);
var
  i,j: integer;
begin
  i:=GetBandNumber(BandName1);
  j:=GetBandNumber(BandName2);
  if i<0 then raise Exception.Create('Секция "'+BandName1+'" не найдена в указанном шаблоне');
  if j<0 then raise Exception.Create('Секция "'+BandName2+'" не найдена в указанном шаблоне');
  if (i<0) or (j<0) then Exit;
  fDataBands[i].fNextBand:=@fDataBands[j];
  fDataBands[j].fNextBand:=@fDataBands[i];
  fDataBands[i].fKeyName:=BandKeyField;
  fDataBands[j].fKeyName:=BandKeyField;
  fDataBands[i].fKeyValue:=KeyValue1;
  fDataBands[j].fKeyValue:=KeyValue2;
end;

// Существует ли в шаблоне секция с именем BandName

function TWordReport.BandExists(BandName: string): boolean;
begin
  Result:=not (GetBandNumber(BandName)<0);
end;

// Подготовить вывод группы секций в документ

procedure TWordReport.CreateMultiBandMaket(Band1,Band2: TDataBand);
var
  cDataset: TDataset;
  cKey,tn: integer;
begin
  if Band1.fDataSet<>Band2.fDataSet then Exit;
  cDataset:=Band1.fDataSet^;
  tn:=Band1.fTableNumber;
  cDataset.First;
  while not cDataset.Eof do begin
    curDocument.Tables.Item(tn).Select;
    curWordApp.Selection.Collapse(wdCollapseEnd);
    cKey:=cDataset.FieldByName(Band1.fKeyName).AsInteger;
    if cKey=Band1.fKeyValue then
      curDocument.Bookmarks.Item(Band1.fRowBookmark).Range.Copy;
    if cKey=Band2.fKeyValue then
      curDocument.Bookmarks.Item(Band2.fRowBookmark).Range.Copy;
    curWordApp.Selection.Paste;
    cDataset.Next;
  end;
  curDocument.Bookmarks.Item(Band1.fRowBookmark).Range.Select;
  curWordApp.Selection.Cut;
  curDocument.Bookmarks.Item(Band2.fRowBookmark).Range.Select;
  curWordApp.Selection.Cut;
  cDataset.First;
end;

// Собрать отчет
// Выполняет всю последовательность действий
// (кроме привязки данных, которую должен осуществить пользователь в обработчике OnReadMaket)
// для формирования одного отчета из шаблона
// TemplateDocFileName в готовый документ ResultDocFileName

procedure TWordReport.Build;
begin
  ReadMaket;
  wProgress:=TProgressForm.Create(Owner);
  if Assigned(fOnReadMaket) then OnReadMaket(Self);
  BindVariables;
  SaveToFile(fResultDocFileName);
  if not fShowResult then CloseDocument;
end;

// Регистрация компонента в палитре Delphi

procedure Register;
begin
  RegisterComponents('WordReport',[TWordReport]);
end;

end.
