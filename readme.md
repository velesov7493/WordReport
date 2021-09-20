## Собранные версии компонента с примерами использования ##
![Архив на Яндекс.Диск](https://disk.yandex.ru/d/16QkRyjFLWDl5w)
### Файлы в архиве ###
Файлы для Delphi 7

- ОписаниеРазработки.txt - этот файл.
- WordReport70.zip - исполняемые файлы и ресурсы компонента.
- WR_ExampleD7.zip - пример применения компонента.
- WordReport70src.zip - исходные коды компонента.

Файлы для Delphi XE3
		
- ОписаниеРазработки.txt - этот файл.
- WordReport170.zip - исполняемые файлы и ресурсы компонента.
- WordReport170src.zip - исходные коды компонента.
- WR_ExampleDXE3.zip - пример применения компонента.
## Назначение ##
Компонент предназначен для автоматизации создания отчетов через MS Word. 
Как исходный шаблон, так и готовый отчет представляют собой обычные документы Word, что обеспечивает пользователю самыме богатые возможности редактирования, предпросмотр и печать без каких-либо дополнительных средств.
## Программные требования ##
- Borland Delphi или Embarcadero RAD Studio XE3.
- Microsoft Word 2000 и выше.
## Установка ##
1. Извлечь файлы из WordReport(Version).zip в директорию с установленной Delphi (например, WordReport70.zip в "C:\Program Files (x86)\Borland\Delphi7").
2. Запустить Delphi.
3. Выбрать пункт меню Component >> Install Packages...
4. Нажать кнопку Add... и выбрать файл WordReport(Version).bpl в Delphi\Bin (например, WordReport70.bpl в "C:\Program Files (x86)\Borland\Delphi7\Bin\WordReport70.bpl")
5. Компонент готов к работе. Его можно найти на вкладке WordReport.
## Инструкция ##
### Правила создания шаблонов ###
Шаблон в нашем случае - это документ MS Word (именно документ - т.е. файл *.doc, а не *.dot !), составленный по определенным правилам.
	
- Секция - это диапазон шаблона, который должен повторяться в результирующем документе столько раз, сколько требуется для вывода всех записей привязанного к секции набора данных.

- Каждая секция должна быть отмечена закладкой с именем DataN, где N - целое число от 1 до 8. 
Повторяться будет ТОЛЬКО то, что в диапазоне закладки, поэтому закладкой лучше отмечать всю строку документа целиком. Если секция используется 
для повторения строки таблицы, то отмечать закладкой следует также всю строку документа, в которой находится эта строка таблицы. 	
	
- Существует три категории переменных шаблона:
		
1. переменные вне секций
			
	Синтаксис объявления: #(ИмяСвободнойПеременной)
	Способ определения значения: напрямую, методом SetValue
		
2. переменные секций
			
	Синтаксис объявления: #(ИмяСекции(ИмяСчетчика).ИмяПеременной)
	Способ определения значения: из текущей записи привязанного поля набора данных.
		
3. счетчики записей секций
		
	Синтаксис объявления: #(ИмяСчетчика)
	Объявление действительно только внутри секции. 
	Заменяется на текущий номер записи привязанного набора данных при отсутствии групп секций
	или на номер записи в неразвывной последовательности при группировке секций.
	
- ИмяСвободнойПеременной - ненулевая последовательность латинских букв, цифр и точек  (только букв, цифр и точек, никаких других знаков!). Регистр букв не важен.
- ИмяПеременной - ненулевая последовательность латинских букв и цифр. Регистр букв не важен.
- ИмяСчетчика - ненулевая последовательность латинских букв и цифр. Регистр букв не важен.
- ИмяСекции - ненулевая последовательность латинских букв и цифр. Регистр букв не важен.
	
- Максимальное количество переменных в секции - 16.
- Максимальное количество секций в документе - 8.
- Максимальное количество переменных вне секций - 2^31 - 1, то есть верхняя граница 32-битного целого типа.
### Описание функционала компонента ###
#### Свойства времени разработки ####
		
Имя документа, содержащего шаблон
`TemplateDocFileName: string`
Имя документа, в котором следует сохранить готовый отчет
`ResultDocFileName: string`
Показать MS Word c готовым отчетом при вызове Quit
`ShowResult: boolean`
		
#### Свойства времени выполнения ####
		
Массив секций документа. 
Доступ по имени секции.
Если секция не найдена - возвращает nil.
`Bands[Name:string]: TDataBand`
		
Количество секций документа.
`BandCount: integer`		
#### Методы ####
		
Существует ли в шаблоне секция с именем BandName.
`function BandExists(BandName: string): boolean;`
		
Сохранить документ в файле FileName.
`procedure SaveToFile(FileName: string);`
		
Выйти. Завершает процесс MS Word или показывает готовый документ.
`procedure Quit;`
		
Сформировать отчет. Выполняет все операции для формирования одного отчета, сохраняет и закрывает полученный документ.
Однако, сам Word остается запущенным для формирования следующих отчетов.
`procedure Build;`
		
Установить значение свободной переменной.
`procedure SetValue(VariableName:string; Value:Variant);`
		
Связать две соседние секции вместе
так чтобы они чередуясь выводили записи из одного и того же набора данных (НД).
Однако, для этого требуется указать целочисленное поле в НД,
значение которого и будет определять, какую именно секцию использовать для вывода текущей записи. 
`procedure JoinBands(BandKeyField:string; BandName1:string; KeyValue1:integer; BandName2:string; KeyValue2:integer);`		
#### События ####
		
Событие наступает после прочтения структуры документа-шаблона,
т.е. тогда, когда имена всех секций (bands) и переменных уже определены, но
этим переменным еще не установлены значения или поля набора данных (НД).
`OnReadMaket:TNotifyEvent;`
		
#### Объект секции документа (TDataBand) ####
#### Свойства времени выполнения ####
		
Имя секции документа.
`Name: string`

Имя поля для переменной VarName.
`Field[VarName:string]: string` 
		
Формат вывода переменной.
`Format[VarName:string]: string`
		
#### Методы ####
		
Подключить набор данных к секции. 
Параметром служит указатель на TDataset, т.е. вызов производится так:
AssignDataSet(@IBQuery1) или AssignDataSet(@ADOTable1).
`procedure AssignDataSet(aDataSet: PDataSet);`
		
Установить переменной VariableName из секции поле FieldName набора данных, подключенного к этой секции.
Если поле содержит действительные числа, то лучше использовать маску формата для их отображения,
например %10.2f (полный список форматов см. в описании функции SysUtils.Format).
Во всех других случаях значение параметра формат можно оставить пустой строкой, т.к. это ни на что не повлияет
`procedure SetField(VariableName,FieldName,Format:string);`
### Обработка события OnReadMaket ###
Так как каждый шаблон уникален и содержит свой неповторимый набор секций и переменных, то и связывание этих секций и переменных со своими
значениями также процесс уникальный, а потому не может быть до конца автоматизирован внутри самого компонента. Эту задачу предстоит решить пользователю компонента.
		
Специально для этой цели и было создано данное событие.
В обработчике этого события НЕОБХОДИМО произвести связывание: 

+ имен переменных с их значениями (TWordReport.SetValue), 
+ секций с их заполненными наборами данных (TDataBand.AssignDataset),
+ переменных в секции с их полями (TDataBand.SetField),
+ а также, если это нужно - формирование групп секций (TWordReport.JoinBands)
		
Подробнее об этом в примере приложения.
## Иные версии Delphi ##
Если нужен компонент для другой версии Delphi - используйте исходные коды, чтобы собрать пакет на своей Delphi, а затем скопируйте его:

1. 32-битную release версию bpl - в поддиректорию bin
2. 64-битную release bpl - в bin64
3. 32-битную release версию dcu вместе с файлами WordReport.dcr и wrtprogress.dfm в lib\win32\release.
4. 64-битную release версию dcu вместе с файлами WordReport.dcr и wrtprogress.dfm в lib\win64\release.
5. Отладочные DCU, если они нужны, копируются в lib\win32\debug и lib\win64\debug уже без файлов ресурсов.