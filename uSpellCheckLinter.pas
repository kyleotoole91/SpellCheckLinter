unit uSpellCheckLinter; //designed for code files .pas and .dfm. Only words within string literals are spell checked

interface

uses
  System.SysUtils, System.IOUtils, System.Types, System.UITypes, System.Classes, System.Variants, Generics.Collections, DateUtils, uConstants,
  System.SyncObjs, System.RegularExpressions;

type
  TSpellCheckLinter = class(TObject)
  strict private
    fThreadCount: integer;
    fCSErrorLog,
    fCSThreadCount: TCriticalSection;
    fIngoreFilePath: string;
    fIgnoreCodeFile: TStringList;
    fWordCheckCount: uInt64;
    fIgnorePathContaining: TStringList;
    fIgnoreLines: TStringList;
    fIgnoreWords: TStringList;
    fIgnoreFiles: TStringList;
    fLanguageFile: TStringList;
    fErrors: TStringList;
    fStartTime: TDateTime;
    fEndTime: TDateTime;
    fFileCount: integer;
    fQuoteSym: string;
    fFileExtFilter: string;
    fSourcePath: string;
    fLanguageFilename: string;
    fWordsDict: TDictionary<string, string>;
    fRecursive: boolean;
    fErrorWords: TStringList;
    fIgnoreContainsLines: TStringList;
    fProvideSuggestions: boolean;
    fFilenameList: TStringList;
    procedure SetIngoreFilePath(const Value: string);
    procedure Clear;
    procedure LoadIgnoreFiles;
    procedure LoadLanguageDictionary;
    procedure SetLanguageFilename(const Value: string);
    procedure SpellCheckFile(const AFilename: string);
    function IngoredPathContaining(const AFilename: string): boolean;
    procedure WaitForThreads;
    procedure ScanForFiles;
  public
    constructor Create;
    destructor Destroy; override;
    procedure Run;
    function AddToIgnoreFile: boolean;
    procedure IncWordCheckCount;
    procedure DecrementThreadCount;
    procedure IncrementThreadCount;
    function ExistsInIgnoreCodeFile(const AText: string): boolean;
    procedure AddError(const AError: string; const AFilename: string; const ALineNumber: integer; const AUnTrimmedLine: string;
                       const ASuggestions: string=''; const AAltSuggestions: string=''); overload;
    procedure AddError(const AError, AFilename: string); overload;
    property LanguageFilename: string read fLanguageFilename write SetLanguageFilename;
    property SourcePath: string read fSourcePath write fSourcePath;
    property QuoteSym: string read fQuoteSym write fQuoteSym;
    property IgnoreWords: TStringList read fIgnoreWords;
    property IgnoreFiles: TStringList read fIgnoreFiles;
    property IgnoreLines: TStringList read fIgnoreLines;
    property FileCount: integer read fFileCount;
    property WordCheckCount: uInt64 read fWordCheckCount;
    property StartTime: TDateTime read fStartTime;
    property EndTime: TDateTime read fEndTime;
    property Errors: TStringList read fErrors;
    property Recursive: boolean read fRecursive write fRecursive;
    property FileExtFilter: string read fFileExtFilter write fFileExtFilter;
    property ErrorsWords: TStringList read fErrorWords;
    property IgnoreContainsLines: TStringList read fIgnoreContainsLines;
    property IngoreFilePath: string read fIngoreFilePath write SetIngoreFilePath;
    property ProvideSuggestions: boolean read fProvideSuggestions write fProvideSuggestions;
    property WordsDict: TDictionary<string, string> read fWordsDict;
    property FilenameList: TStringList read fFilenameList;
  end;

implementation

uses
  StrUtils, Masks, uSpellCheckFile;

{ TSpellCheckLinter }

constructor TSpellCheckLinter.Create;
begin
  inherited;
  fCSErrorLog := TCriticalSection.Create;
  fCSThreadCount := TCriticalSection.Create;
  fWordsDict := TDictionary<string, string>.Create;
  fProvideSuggestions := true;
  fIngoreFilePath := cDefaultIgnorePath;
  fIgnorePathContaining := TStringList.Create;
  fIgnoreContainsLines := TStringList.Create;
  fErrorWords := TStringList.Create;
  fLanguageFile := TStringList.Create;
  fErrors := TStringList.Create;
  fIgnoreLines := TStringList.Create;
  fIgnoreWords := TStringList.Create;
  fIgnoreFiles := TStringList.Create;
  fIgnoreCodeFile := TStringList.Create;
  fFilenameList := TStringList.Create;
  fQuoteSym := cDefaultQuote;
  fRecursive := true;
  Clear;
end;

destructor TSpellCheckLinter.Destroy;
begin
  try
    fFilenameList.DisposeOf;
    fIgnoreCodeFile.DisposeOf;
    fIgnorePathContaining.DisposeOf;
    fIgnoreContainsLines.DisposeOf;
    fErrorWords.DisposeOf;
    fIgnoreLines.DisposeOf;
    fIgnoreWords.DisposeOf;
    fIgnoreFiles.DisposeOf;
    fWordsDict.DisposeOf;
    fLanguageFile.DisposeOf;
    fErrors.DisposeOf;
    fCSErrorLog.DisposeOf;
    fCSThreadCount.DisposeOf;
  finally
    inherited;
  end;
end;

procedure TSpellCheckLinter.Clear;
begin
  fFileCount := 0;
  fWordCheckCount := 0;
  fIgnoreLines.Clear;
  fIgnoreWords.Clear;
  fIgnoreFiles.Clear;
  fWordsDict.Clear;
  fLanguageFile.Clear;
  fErrors.Clear;
  fSourcePath := cDefaultSourcePath;
  fFileExtFilter := cDefaultExtFilter;
  fLanguageFilename := cDefaultlanguageName;
  fIgnoreContainsLines.Clear;
  fIgnorePathContaining.Clear;
end;

procedure TSpellCheckLinter.Run;
var
  i: integer;
  filename: string;
begin
  fStartTime := Now;
  try
    try
      fThreadCount := 0;
      fErrors.Clear;
      fErrorWords.Clear;
      LoadIgnoreFiles;
      LoadLanguageDictionary;
      if FileExists(fSourcePath) then
        SpellCheckFile(fSourcePath)
      else if Trim(fSourcePath) <> '' then
        ScanForFiles;
      for i:=0 to fFilenameList.Count-1 do begin
        filename := fFilenameList.Strings[i];
        if (fIgnoreFiles.IndexOf(ExtractFileName(filename)) = -1) and
           (not IngoredPathContaining(filename)) then begin
          Inc(fFileCount);
          SpellCheckFile(filename);
        end;
      end;
      WaitForThreads;
    except
      on e: exception do
        fErrors.Add('Exception raised: '+e.ClassName+' '+e.Message);
    end;
  finally
    fEndTime := Now;
  end;
end;

procedure TSpellCheckLinter.ScanForFiles;
var
  multiExts: boolean;
  filename: string;
  filenameArr,
  maskArr: TStringDynArray;
  filterPredicate: TDirectory.TFilterPredicate;
  searchOption: TSearchOption;
begin
  fFilenameList.Clear;
  maskArr := SplitString(fFileExtFilter.Replace('"', ''), '|');
  multiExts := fFileExtFilter.Contains( '|');
  filterPredicate := function(const APath: string; const ASearchRec: TSearchRec): boolean
                     var
                       mask: string;
                     begin
                       result := false;
                       for mask in maskArr do begin
                         result := MatchesMask(ASearchRec.Name, mask);
                         if result then
                           Break;
                       end;
                     end;
  if fRecursive then
    searchOption := TSearchOption.soAllDirectories
  else
    searchOption := TSearchOption.soTopDirectoryOnly;
  if multiExts then
    filenameArr := TDirectory.GetFiles(fSourcePath, searchOption, filterPredicate)
  else
    filenameArr := TDirectory.GetFiles(fSourcePath, fFileExtFilter, searchOption);
  for filename in filenameArr do
    fFilenameList.Add(filename);
end;

procedure TSpellCheckLinter.LoadLanguageDictionary;
var
  i: integer;
  key: string;
  value: string;
  sepIndex: integer;
begin
  fWordsDict.Clear;
  fLanguageFile.LoadFromFile(fLanguageFilename);
  for i:=0 to fLanguageFile.Count-1 do begin
    key := fLanguageFile.Strings[i];
    sepIndex := key.IndexOf('/');
    if sepIndex = -1 then
      fWordsDict.AddOrSetValue(fLanguageFile.Strings[i], '')
    else begin
      key := Copy(fLanguageFile.Strings[i], 0, pos('/', fLanguageFile.Strings[i])-1);
      value := Copy(fLanguageFile.Strings[i], pos('/', fLanguageFile.Strings[i])+1, fLanguageFile.Strings[i].Length);
      fWordsDict.AddOrSetValue(key, value);
    end;
  end;
end;

procedure TSpellCheckLinter.LoadIgnoreFiles;
begin
  if FileExists(fIngoreFilePath+cIgnoreContainsName) then
    fIgnoreContainsLines.LoadFromFile(fIngoreFilePath+cIgnoreContainsName);
  if FileExists(fIngoreFilePath+cIgnoreCodeName) then
    fIgnoreCodeFile.LoadFromFile(fIngoreFilePath+cIgnoreCodeName);
  if FileExists(fIngoreFilePath+cIgnorePathsName) then
    fIgnorePathContaining.LoadFromFile(fIngoreFilePath+cIgnorePathsName);
  if FileExists(fIngoreFilePath+cIgnoreFilesName) then
    fIgnoreFiles.LoadFromFile(fIngoreFilePath+cIgnoreFilesName);
  if FileExists(fIngoreFilePath+cIgnoreWordsName) then
    fIgnoreWords.LoadFromFile(fIngoreFilePath+cIgnoreWordsName);
  if FileExists(fIngoreFilePath+cIgnoreLinesName) then
    fIgnoreLines.LoadFromFile(fIngoreFilePath+cIgnoreLinesName);
end;

function TSpellCheckLinter.IngoredPathContaining(const AFilename: string): boolean;
var
  a: integer;
  path: string;
begin
  result := false;
  path := ExtractFilePath(AFilename);
  for a := 0 to fIgnorePathContaining.Count-1 do begin
    result := path.Contains(fIgnorePathContaining.Strings[a]);
    if result then
      Break;
  end;
end;

function TSpellCheckLinter.ExistsInIgnoreCodeFile(const AText: string): boolean;
var
  a: integer;
begin
  result := false;
  for a := 0 to fIgnoreCodeFile.Count-1 do begin
    result := AText.Contains(fIgnoreCodeFile.Strings[a]);
    if result then
      Break;
  end;
end;

procedure TSpellCheckLinter.AddError(const AError: string; const AFilename: string; const ALineNumber: integer; const AUnTrimmedLine: string;
                               const ASuggestions: string=''; const AAltSuggestions: string='');
begin
  fCSErrorLog.Enter;
  try
    fErrors.Add('Error: '+AError);
    if ASuggestions <> '' then
      fErrors.Add('Suggestions: '+ASuggestions);
    if AAltSuggestions <> '' then
      fErrors.Add('Alternatives: '+AAltSuggestions);
    fErrors.Add('File name: '+AFilename);
    fErrors.Add('Line number: '+IntToStr(ALineNumber));
    fErrors.Add('Line text: '+AUnTrimmedLine);
    fErrors.Add(' ');
    fErrorWords.Add(AError);
  finally
    fCSErrorLog.Leave;
  end;
end;

function TSpellCheckLinter.AddToIgnoreFile: boolean;
var
  sl: TStringList;
  a: integer;
begin
  result := true;
  try
    sl := TStringList.Create;
    if FileExists('.\'+cIgnoreWordsName) then
      sl.LoadFromFile(cIgnoreWordsName);
    sl.Add(fErrorWords.Text);
    for a := 0 to fErrorWords.Count-1 do begin
      if sl.IndexOf(fErrorWords.Strings[a]) = -1 then
        sl.Add(fErrorWords.Strings[a]);
    end;
    sl.SaveToFile('.\'+cIgnoreWordsName);
  except
    on e: exception do begin
      result := false;
      fErrors.Add(e.Classname+' '+e.Message);
    end;
  end;
end;

procedure TSpellCheckLinter.SetIngoreFilePath(const Value: string);
begin
  fIngoreFilePath := IncludeTrailingPathDelimiter(Value);
end;

procedure TSpellCheckLinter.SetLanguageFilename(const Value: string);
begin
  fLanguageFilename := Value;
  LoadLanguageDictionary;
end;

procedure TSpellCheckLinter.SpellCheckFile(const AFilename: string);
var
  spellCheckThread: TSpellCheckThread;
begin
  while fThreadCount >= cMaxThreads do
    Sleep(cSleepTime);
  spellCheckThread := TSpellCheckThread.Create(Self);
  spellCheckThread.SpellCheckFile.Filename := AFilename;
  spellCheckThread.SpellCheckFile.ProvideSuggestions := fProvideSuggestions;
  spellCheckThread.Start;
end;

procedure TSpellCheckLinter.WaitForThreads;
begin
  while fThreadCount > 0 do
    Sleep(cSleepTime);
end;

procedure TSpellCheckLinter.DecrementThreadCount;
begin
  fCSThreadCount.Enter;
  try
    Inc(fThreadCount, -1);
  finally
    fCSThreadCount.Leave;
  end;
end;

procedure TSpellCheckLinter.IncrementThreadCount;
begin
  fCSThreadCount.Enter;
  try
    Inc(fThreadCount);
  finally
    fCSThreadCount.Leave;
  end;
end;

procedure TSpellCheckLinter.IncWordCheckCount;
begin
  fCSErrorLog.Enter;
  try
    Inc(fWordCheckCount);
  finally
    fCSErrorLog.Leave;
  end;
end;

procedure TSpellCheckLinter.AddError(const AError, AFilename: string);
begin
  fCSErrorLog.Enter;
  try
    fErrors.Add('Error checking file '+AFilename+': '+AError);
  finally
    fCSErrorLog.Leave;
  end;
end;

end.


