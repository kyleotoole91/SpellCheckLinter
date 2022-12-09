unit uSpellCheck; //designed for code files .pas and .dfm. Only words within string literals are spell checked


interface

uses
  System.SysUtils, System.IOUtils, System.Types, System.UITypes, System.Classes, System.Variants, Generics.Collections, DateUtils, uConstants,
  System.SyncObjs, System.RegularExpressions;

type
  TSpellCheck = class(TObject)
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
    procedure SetIngoreFilePath(const Value: string);
    procedure Clear;
    procedure LoadIgnoreFiles;
    procedure LoadLanguageDictionary;
    procedure SetLanguageFilename(const Value: string);
    procedure SpellCheckFile(const AFilename: string);
    function IngoredPathContaining(const AFilename: string): boolean;
    procedure WaitForThreads;
  public
    constructor Create;
    destructor Destroy; override;
    function Run: boolean;
    function AddToIgnoreFile: boolean;
    procedure IncWordCheckCount;
    function InIgnoreCodeFile(const AText: string): boolean;
    procedure AddError(const AError: string; const AFilename: string; const ALineNumber: integer; const AUnTrimmedLine: string;
                       const ASuggestions: string=''; const AAltSuggestions: string=''); overload;
    procedure AddError(const AError, AFilename: string); overload;
    procedure DecrementThreadCount;
    procedure IncrementThreadCount;
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
  end;

  TSpellCheckFile = class(TObject)
  strict private
    fOwner: TSpellCheck;
    fProvideSuggestions: boolean;
    fAltSuggestions,
    fSuggestions: TStringList;
    fPossibleTruncation: boolean;
    fTextBeforeQuote: string;
    fUnTrimmedLine: string;
    fIgnoreNextFirstWord: boolean;
    fSourceFile: TStringList;
    fSkippedLineCount: integer;
    fMultiCommentSym: string;
    fIsDFM: boolean;
    fFilename: string;
    fCamelCaseWords: TStringList;
    fLineWords: TStringList;
    procedure CleanLineWords;
    function EditDistance(const AFromString, AToString: string): integer;
    procedure AddSuggestion(const ASuggestion: string; const AOpCount: integer; const AInsert: boolean=false);
    procedure BuildSuggestions(const AMispelledWord: string);
    procedure BuildCamelCaseWords(const AArr: TArray<string>);
    function IsGuid(const ALine: string): boolean;
    function SanitizeWord(const AWord: string; const ARemovePeriods: boolean=true): string;
    function RemoveStartEndQuotes(const AStr: string): string;
    function RemoveEscappedQuotes(const AStr: string): string;
    function PosOfNextQuote(const AString, AChar: string): integer;
    function NeedsSpellCheck(const AValue: string): boolean;
    function IsCommentLine(const ALine: string): boolean;
    function IsSQL(const ALine: string): boolean;
    function SpelledCorrectly(const AWord: string; const ATryLower: boolean=true): boolean;
    function IsNumeric(const AString: string): boolean;
    procedure LimitItems(const AStringList: TStringList);
  public
    constructor Create(const AOwner: TSpellCheck); reintroduce;
    destructor Destroy; override;
    function Run: boolean;
    property Filename: string read fFilename write fFilename;
    property ProvideSuggestions: boolean read fProvideSuggestions write fProvideSuggestions;
    property Owner: TSpellCheck read fOwner;
  end;

  TSpellCheckThread = class(TThread)
  strict private
    fSpellCheckFile: TSpellCheckFile;
  protected
    procedure Execute; override;
  public
    constructor Create(const AOwner: TSpellCheck); reintroduce;
    destructor Destroy; override;
    property SpellCheckFile: TSpellCheckFile read fSpellCheckFile;
  end;

implementation

uses
  System.Character, Math, ActiveX;

{ TSpellCheck }

constructor TSpellCheck.Create;
begin
  inherited;
  fProvideSuggestions := true;
  fIngoreFilePath := cDefaultSourcePath;
  fIgnorePathContaining := TStringList.Create;
  fIgnoreContainsLines := TStringList.Create;
  fErrorWords := TStringList.Create;
  fWordsDict := TDictionary<string, string>.Create;
  fLanguageFile := TStringList.Create;
  fErrors := TStringList.Create;
  fIgnoreLines := TStringList.Create;
  fIgnoreWords := TStringList.Create;
  fIgnoreFiles := TStringList.Create;
  fIgnoreCodeFile := TStringList.Create;
  fQuoteSym := cDefaultQuote;
  fRecursive := true;
  fCSErrorLog := TCriticalSection.Create;
  fCSThreadCount := TCriticalSection.Create;
  Clear;
end;

procedure TSpellCheck.DecrementThreadCount;
begin
  fCSThreadCount.Enter;
  try
    Inc(fThreadCount, -1);
  finally
    fCSThreadCount.Leave;
  end;
end;

procedure TSpellCheck.IncrementThreadCount;
begin
  fCSThreadCount.Enter;
  try
    Inc(fThreadCount);
  finally
    fCSThreadCount.Leave;
  end;
end;

destructor TSpellCheck.Destroy;
begin
  try
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

procedure TSpellCheck.Clear;
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

function TSpellCheck.Run: boolean;
var
  filenames: TStringDynArray;
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
      else begin
        if Trim(fSourcePath) = '' then
          fSourcePath := '.\';
        if fRecursive then
          filenames := TDirectory.GetFiles(fSourcePath, fFileExtFilter, TSearchOption.soAllDirectories)
        else
          filenames := TDirectory.GetFiles(fSourcePath, fFileExtFilter, TSearchOption.soTopDirectoryOnly);
      end;
      for filename in filenames do begin
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
    result := (fErrorWords.Count = 0) and
              (fErrors.Count = 0);
  end;
end;

procedure TSpellCheck.LoadLanguageDictionary;
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

procedure TSpellCheck.LoadIgnoreFiles;
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

procedure TSpellCheck.IncWordCheckCount;
begin
  fCSErrorLog.Enter;
  try
    Inc(fWordCheckCount);
  finally
    fCSErrorLog.Leave;
  end;
end;

function TSpellCheck.IngoredPathContaining(const AFilename: string): boolean;
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

function TSpellCheck.InIgnoreCodeFile(const AText: string): boolean;
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

procedure TSpellCheck.AddError(const AError: string; const AFilename: string; const ALineNumber: integer; const AUnTrimmedLine: string;
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

procedure TSpellCheck.AddError(const AError, AFilename: string);
begin
  fCSErrorLog.Enter;
  try
    fErrors.Add('Error checking file '+AFilename+': '+AError);
  finally
    fCSErrorLog.Leave;
  end;
end;

function TSpellCheck.AddToIgnoreFile: boolean;
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

procedure TSpellCheck.SetIngoreFilePath(const Value: string);
begin
  fIngoreFilePath := IncludeTrailingPathDelimiter(Value);
end;

procedure TSpellCheck.SetLanguageFilename(const Value: string);
begin
  fLanguageFilename := Value;
  LoadLanguageDictionary;
end;

procedure TSpellCheck.SpellCheckFile(const AFilename: string);
var
  spellCheckThread: TSpellCheckThread;
begin
  while fThreadCount >= cMaxThreads do
    Sleep(10);
  spellCheckThread := TSpellCheckThread.Create(Self);
  spellCheckThread.SpellCheckFile.Filename := AFilename;
  spellCheckThread.SpellCheckFile.ProvideSuggestions := fProvideSuggestions;
  spellCheckThread.Start;
end;

procedure TSpellCheck.WaitForThreads;
begin
  while fThreadCount > 0 do
    Sleep(10);
end;

{ TSpellCheckFile }

constructor TSpellCheckFile.Create(const AOwner: TSpellCheck);
begin
  inherited Create;
  fOwner := AOwner;
  fProvideSuggestions := true;
  fAltSuggestions := TStringList.Create;
  fSuggestions := TStringList.Create;
  fLineWords := TStringList.Create;
  fLineWords.Delimiter := ' ';
  fLineWords.StrictDelimiter := true;
  fCamelCaseWords := TStringList.Create;
  fSourceFile := TStringList.Create;
end;

destructor TSpellCheckFile.Destroy;
begin
  fAltSuggestions.DisposeOf;
  fSuggestions.DisposeOf;
  fLineWords.DisposeOf;
  fCamelCaseWords.DisposeOf;
  fSourceFile.DisposeOf;
  inherited;
end;

function TSpellCheckFile.Run: boolean;
const
  cBreakout=1000;
var
  itrCount: integer;
  i, j: integer;
  theStr: string;
  lineStr: string;
  theWord: string;
  fileExt: string;
  lastWordChecked: string;
  nextFirstWord: string;
  ok, okWithNextWord: boolean;
  function ContainsIgnore: boolean;
  var
    a: integer;
  begin
    result := false;
    for a := 0 to fOwner.IgnoreContainsLines.Count-1 do begin
      result := lineStr.Contains(fOwner.IgnoreContainsLines.Strings[a]);
      if result then
        Break;
    end;
  end;
  function FirstWordNextLine: string;
  var
    str: string;
    sl: TStringList;
  begin
    result := '';
    sl := TStringList.Create;
    try
      sl.Delimiter := ' ';
      if i+1 <= fSourceFile.Count-1 then begin
        str := fSourceFile.Strings[i+1];
        str := Copy(str, Pos(fOwner.QuoteSym, str)+1, str.Length);
        str := Copy(str, 1, PosOfNextQuote(str, fOwner.QuoteSym)-1);
        sl.CommaText := Str;
        if sl.Count >= 1 then
          result := Trim(sl.Strings[0]);
      end;
    finally
      sl.DisposeOf;
    end;
  end;
  function CheckCamelCaseWords(AWord: string; const ALogError: boolean=true): boolean;
  var
    k: integer;
    hasError: boolean;
  begin
    hasError := false;
    BuildCamelCaseWords(TRegEx.Split(AWord, cRegExCamelPascalCaseSpliter));
    for k := 0 to fCamelCaseWords.Count-1 do begin //for each camel case word
      AWord := RemoveEscappedQuotes(fCamelCaseWords.Strings[k]);
      if not SpelledCorrectly(AWord) then begin
        hasError := true;
        if ALogError then
          fOwner.AddError(AWord, fFilename, i+1, fUnTrimmedLine, fSuggestions.CommaText, fAltSuggestions.CommaText);
      end;
    end;
    result := not hasError;
  end;
begin
  fSkippedLineCount := 0;
  result := true;
  fMultiCommentSym := '';
  fIsDFM := ExtractFileExt(fFilename) = '.dfm';
  fSourceFile.LoadFromFile(fFilename);
  fIgnoreNextFirstWord := false;
  for i:=0 to fSourceFile.Count-1 do begin
    try
      fUnTrimmedLine := fSourceFile.Strings[i];
      lineStr := Trim(fSourceFile.Strings[i]);
      if Pos(fOwner.QuoteSym, lineStr) >= 1 then
        fTextBeforeQuote := Copy(lineStr, 1, Pos(fOwner.QuoteSym, lineStr)-1)
      else
        fTextBeforeQuote := lineStr;
      if (lineStr <> '') and
         (not IsCommentLine(lineStr)) and
         (not ContainsIgnore) and
         (fOwner.IgnoreLines.IndexOf(Trim(lineStr)) = -1) then begin
        itrCount := 0;
        while Pos(fOwner.QuoteSym, lineStr) >= 1 do begin //while the line has string literals
          Inc(itrCount);
          lineStr := Copy(lineStr, Pos(fOwner.QuoteSym, lineStr)+1, lineStr.Length);
          if itrCount >= cBreakout then
            raise Exception.Create('An infinite loop has been detected in '+fFilename+' on line '+IntToStr(i+1)+'. '+
                                   'Please check the syntax of file. ');
          theStr := Copy(lineStr, 1, PosOfNextQuote(lineStr, fOwner.QuoteSym)-1);
          fPossibleTruncation := fIsDFM and (theStr.Length = cMaxStringLengthDFM);
          if Pos(theStr, lineStr)+theStr.Length <> 0 then //remove the string from the line
            lineStr := Copy(lineStr, Pos(theStr, lineStr)+theStr.Length+1, lineStr.Length)
          else
            lineStr := Copy(lineStr, 3, lineStr.Length); //if empty string, remove the next two chars
          if (not Trim(theStr).StartsWith('<')) and //ignore html
             (ExtractFilePath(Trim(theStr)) = '') or (Pos( ':', theStr) > 2) then begin //ignore file paths
            theStr := theStr.Replace('&', ''); //underlining menu items
            if NeedsSpellCheck(theStr) then begin
              if IsGuid(theStr) then
                fLineWords.DelimitedText := ''
              else begin
                theStr := SanitizeWord(theStr, false);
                if theStr.EndsWith(fOwner.QuoteSym) then //strings ending with double quotes was causing the last char to get removed when setting delimited text
                  theStr := Copy(theStr, 1, theStr.Length-1)+' '+fOwner.QuoteSym;
                fLineWords.DelimitedText := theStr;
              end;
              CleanLineWords;
              theWord := '';
              for j := 0 to fLineWords.Count-1 do begin //for each word in the line
                if (fIgnoreNextFirstWord) and
                   (j = 0) then
                  fIgnoreNextFirstWord := false
                else begin
                  theWord := SanitizeWord(fLineWords.Strings[j], false);
                  if Trim(theWord) <> '' then begin
                    lastWordChecked := theWord;
                    if theWord.Contains('''') then begin //the regex will split on this. You don't normally see apostrophies in camelCase or PascalCase words anyway
                      lineStr := lineStr.Remove(pos(fLineWords.Strings[j],lineStr)-1, theWord.Length);
                      theWord := RemoveEscappedQuotes(theWord);
                      if not SpelledCorrectly(theWord) then begin
                        result := false;
                        fOwner.AddError(theWord, fFilename, i+1, fUnTrimmedLine, fSuggestions.CommaText, fAltSuggestions.CommaText);
                      end;
                    end else begin
                      fileExt := ExtractFileExt(theWord); //ignore exts
                      if fileExt.Length = 1 then
                        fileExt := '';
                      if (fileExt = '') then begin
                        nextFirstWord := FirstWordNextLine;
                        if (fPossibleTruncation) and
                           (j = fLineWords.Count-1) and
                           (nextFirstWord <> '')then begin
                          okWithNextWord := CheckCamelCaseWords(theWord+nextFirstWord, false);
                          if not okWithNextWord then begin
                            ok := CheckCamelCaseWords(theWord, false);
                            if not ok then begin
                              result := false;
                              fOwner.AddError(theWord+nextFirstWord+cWordSeparator+theWord+cWordSeparator+nextFirstWord,
                                              fFilename, i+1, fUnTrimmedLine, fSuggestions.CommaText, fAltSuggestions.CommaText);
                            end;
                          end else
                            fIgnoreNextFirstWord := true;
                        end else
                          CheckCamelCaseWords(theWord);
                      end;
                    end;
                  end;
                end;
              end;
            end;
          end else
            Break;
        end;
      end;
    except
      on e: exception do begin
        result := false;
        fOwner.AddError('Exception raised: '+e.ClassName+' '+e.Message, fFilename);
      end;
    end;
  end;
end;

procedure TSpellCheckFile.BuildCamelCaseWords(const AArr: TArray<string>); //eg msmWeb, don't check msm
var
  a,
  firstWordIdx,
  wordCount: integer;
  isCamelCase: boolean;
begin
  fCamelCaseWords.Clear;
  wordCount := 0;
  firstWordIdx := -1;
  isCamelCase := false;
  for a := 0 to Length(AArr)-1 do begin
    if Trim(AArr[a]) <> '' then begin
      Inc(wordCount);
      if wordCount = 1 then
        firstWordIdx := a;
      isCamelCase := wordCount >= 2;
      if isCamelCase then
        Break;
    end;
  end;
  if (isCamelCase) and
     (AArr[firstWordIdx].Length < cIgnoreCamelPrefixLength) then //eg ignore msm in msmWeb
    AArr[firstWordIdx] := '';
  for a := 0 to Length(AArr)-1 do begin
    if Trim(AArr[a]) <> '' then
      fCamelCaseWords.Add(AArr[a]);
  end;
end;

function TSpellCheckFile.IsCommentLine(const ALine: string): boolean;
begin
  if fMultiCommentSym = '' then begin
    if fTextBeforeQuote.Contains(cSpellCheckOff) then
      fMultiCommentSym := cSpellCheckOn
    else if fTextBeforeQuote.Contains('{') then
      fMultiCommentSym := '}'
    else if fTextBeforeQuote.Contains('(*') then
      fMultiCommentSym := '*)'
    else if fTextBeforeQuote.Contains('<html>') then
      fMultiCommentSym := '</html>'
    else if IsSQL(ALine) or
            fOwner.InIgnoreCodeFile(fTextBeforeQuote) then begin
      if fIsDFM then
        fMultiCommentSym := cSkipLineEndStringDFM
      else
        fMultiCommentSym := cSkipLineEndString;
    end;
  end else begin
    Inc(fSkippedLineCount);
    if not fIsDFM then begin
      if (fUnTrimmedLine.StartsWith(fMultiCommentSym)) then
        fMultiCommentSym := ''
      else if ((fMultiCommentSym = cSkipLineEndString) or (fMultiCommentSym = '</html>')) and
              (fSkippedLineCount > cSkippedLineCountLimit) then //break out if limit is reached, eg </ html> instead of </html>
        fMultiCommentSym := '';
    end;
  end;
  result := (ALine.Contains('//')) or
            (fMultiCommentSym <> '') or
            (ALine.StartsWith('function')) or
            (ALine.StartsWith('procedure')) or
            (ALine.StartsWith('  function')) or
            (ALine.StartsWith('  procedure'));
  if (fMultiCommentSym <> '') then begin
    if (((fIsDFM) and ((ALine.StartsWith(fMultiCommentSym)))) or
        ((not fIsDFM) and (ALine.Contains(fMultiCommentSym)))) then
      fMultiCommentSym := '';
  end;
  if fMultiCommentSym = '' then
    fSkippedLineCount := 0;
end;

function TSpellCheckFile.PosOfNextQuote(const AString, AChar: string): integer;
var
  str: string;
  i, occurrence: integer;
  posOffset: integer;
begin
  result := -1;
  occurrence := 0;
  posOffset := 1;
  str := AString.Replace(AChar+AChar, '  '); //replace escapped quotes with empty spaces so they don't count
  for i:=1 to str.Length do begin
    if (str[i] = AChar) then begin
      if (str[i+1] = AChar) then
        Inc(posOffset)
      else begin
        Inc(occurrence);
        if occurrence = posOffset then begin
          result := i;
          Break;
        end;
      end;
    end;
  end;
end;

function TSpellCheckFile.RemoveEscappedQuotes(const AStr: string): string;
begin
  result := Trim(TRegEx.Replace(AStr, cRegExRemoveNumbers, ''));
  if fOwner.QuoteSym = '''' then //remove Delphi escaped quotes and replace with one single quote
    result := result.Replace('''''', '''');
end;

function TSpellCheckFile.SpelledCorrectly(const AWord: string; const ATryLower: boolean=true): boolean;
var
  word: string;
begin
  word := Trim(RemoveStartEndQuotes(AWord));
  result := word = '';
  if not result then begin
    word := Trim(TRegEx.Replace(word, cRegExKeepLettersAndQuotes, ' '));
    if word.EndsWith('.') then
      word := Copy(word, 1, word.Length-1);
    result := (word.Length < cMinCheckLength) or //Don't check short words
              (UpperCase(word) = word) or //IGNORE UPPER CASE TEXT
              (IsNumeric(word)) or //Don't check numbers
              (fOwner.IgnoreWords.IndexOf(word) >= 0); //Don't ignore file for the word
    if not result then begin
      result := (fOwner.WordsDict.ContainsKey(word)) or
                (ATryLower and fOwner.WordsDict.ContainsKey(LowerCase(word)));
      fOwner.IncWordCheckCount;
      if (not result) and
         (fProvideSuggestions) then
        BuildSuggestions(AWord);
    end;
  end;
end;

function TSpellCheckFile.RemoveStartEndQuotes(const AStr: string): string;
begin
  result := AStr;
  if (result.Length > 0) and
     (result[1] = fOwner.QuoteSym) then
    result := Copy(result, 2, result.Length);
  if (result.Length > 0) and
     (result[result.Length] = fOwner.QuoteSym) then
    result := Copy(result, 1, result.Length-1);
end;

function TSpellCheckFile.IsNumeric(const AString: string): boolean;
var
  a: integer;
begin
  result := false;
  {$WARNINGS OFF}
  for a := 1 to AString.Length do begin
    result := System.Character.IsNumber(AString, a);
    if not result then
      Break;
  end;
  {$WARNINGS ON}
end;

function TSpellCheckFile.IsSQL(const ALine: string): boolean;
begin
  result := ALine.Contains('SELECT') or
            ALine.Contains('TABLE') or
            ALine.Contains('INSERT') or
            ALine.Contains('UPDATE') or
            ALine.Contains('DELETE') or
            ALine.Contains('GROUP BY') or
            ALine.Contains('ORDER BY') or
            ALine.Contains(' AND ') or
            ALine.Contains(' WHERE ') or
            ALine.Contains('LEFT OUTER JOIN') or
            ALine.Contains('INNER JOIN');
end;

procedure TSpellCheckFile.CleanLineWords;
var
  a, k, lastPos: integer;
begin
  lastPos := -1;
  for a := 0 to fLineWords.Count-1 do begin
    if ((a > lastPos) or (lastPos=-1)) and
       ((fLineWords.Strings[a].StartsWith(''''))) then begin //remove delphi escape quotes
      for k:=a to fLineWords.Count-1 do begin
        if fLineWords.Strings[k].EndsWith('''') then begin
          fLineWords.Strings[a] := Copy(fLineWords.Strings[a], 2, fLineWords.Strings[a].Length-1);
          fLineWords.Strings[k] := Copy(fLineWords.Strings[k], 1, fLineWords.Strings[k].Length-2);
          lastPos := k;
          Break;
        end;
      end;
    end;
  end;
end;

function TSpellCheckFile.SanitizeWord(const AWord: string; const ARemovePeriods: boolean=true): string;
begin
  result := Trim(TRegEx.Replace(AWord, cRegExKeepLettersAndQuotes, ' '));
  result := Trim(RemoveStartEndQuotes(result));
  if ARemovePeriods then
    result := result.Replace('.', ' ');
end;

function TSpellCheckFile.NeedsSpellCheck(const AValue: string): boolean;
begin
  result := not (AValue.StartsWith('#') or
                 AValue.StartsWith('//') or
                 AValue.StartsWith('/') or
                 AValue.StartsWith('\') or
                 AValue.StartsWith('\\'));
end;

function TSpellCheckFile.IsGuid(const ALine: string): boolean;
begin
  result := (ALine.Length >= cGuidLen) and
            (ALine.Chars[0] = '{') and
            (ALine.Chars[ALine.Length-1] = '}') and
            (ALine.Contains('-'));
end;

procedure TSpellCheckFile.BuildSuggestions(const AMispelledWord: string);
const
  cNormalCase=0;
  cLowerCase=1;
var
  suggestedCount,
  operationCount: integer;
  splitWords: TStringList;
  procedure FindSuggestions(AWord: string; const ACaseType: integer=cNormalCase);
  var
    a, k: integer;
    word,
    dictionaryWord: string;
    procedure TypoSearchAlteratives;
    var
      a, idx: integer;
      typoFix, tmp: string;
    begin
      for a := 1 to Length(AMispelledWord) do begin
        if a >= 2 then begin
          typoFix := AMispelledWord;
          tmp := typoFix[a];
          typoFix[a] := typoFix[a-1];
          typoFix[a-1] := tmp[1];
          idx := fAltSuggestions.IndexOf(typoFix);
          if idx >= 0 then begin
            fSuggestions.Insert(0, typoFix);
            fAltSuggestions.Delete(idx);
          end else begin
            idx := fSuggestions.IndexOf(typoFix);
            if idx >= 0 then begin
              fSuggestions.Delete(idx);
              fSuggestions.Insert(0, typoFix);
            end;
          end;
        end;
      end;
    end;
    procedure ExtractMostLikelyAlts;
    var
      a: integer;
      commonChar: string;
    begin
      commonChar := '';
      for a := 0 to fSuggestions.Count-1 do begin
        if (commonChar = '') or
           (commonChar = fSuggestions.Strings[a][1]) then
          commonChar := fSuggestions.Strings[a][1]
        else begin
          commonChar := '';
          Break;
        end;
      end;
      if (commonChar <> '') and
         (LowerCase(commonChar) = LowerCase(AMispelledWord[1])) then begin
        for a := fAltSuggestions.Count-1 downto 0 do begin
          if LowerCase(fAltSuggestions.Strings[a][1]) <> LowerCase(commonChar) then
            fAltSuggestions.Delete(a)
        end;
      end;
    end;
    function NeedsCheck: boolean;
    begin
      result := (Abs(dictionaryWord.Length-AMispelledWord.Length) <= cSuggestLengthOffset) and
                ((ACaseType <> cLowerCase) or
                 (LowerCase(dictionaryWord) = LowerCase(dictionaryWord)));
    end;
  begin
    splitWords.Delimiter := cWordSeparator;
    splitWords.StrictDelimiter := true;
    splitWords.CommaText := AWord;
    for a := 0 to splitWords.Count-1 do begin
      word := LowerCase(splitWords.Strings[a]);
      suggestedCount := 0;
      for dictionaryWord in fOwner.WordsDict.Keys do begin
        if NeedsCheck then begin
          Inc(suggestedCount);
          operationCount := EditDistance(word, LowerCase(dictionaryWord));
          AddSuggestion(dictionaryWord, operationCount);
        end;
      end;
      suggestedCount := 0;
      for k := 0 to fOwner.IgnoreWords.Count-1 do begin
        if NeedsCheck then begin
          Inc(suggestedCount);
          dictionaryWord := fOwner.IgnoreWords.Strings[k];
          operationCount := EditDistance(word, LowerCase(dictionaryWord));
          AddSuggestion(dictionaryWord, operationCount);
        end;
      end;
      ExtractMostLikelyAlts;
      TypoSearchAlteratives;
    end;
  end;
begin
  splitWords := TStringList.Create;
  fSuggestions.Clear;
  fAltSuggestions.Clear;
  try
    FindSuggestions(AMispelledWord);
    LimitItems(fSuggestions);
    LimitItems(fAltSuggestions);
  finally
    splitWords.DisposeOf;
  end;
end;

procedure TSpellCheckFile.LimitItems(const AStringList: TStringList);
var
  a: integer;
begin
  if AStringList.Count > cMaxSuggestions then begin
    for a := AStringList.Count-1 downto cMaxSuggestions-1 do
      AStringList.Delete(a);
    AStringList.Add('...');
  end;
end;

procedure TSpellCheckFile.AddSuggestion(const ASuggestion: string; const AOpCount: integer; const AInsert: boolean=false);
  procedure AddItem;
  begin
    if not AInsert then
      fSuggestions.AddObject(ASuggestion, pointer(AOpCount))
    else
      fSuggestions.InsertObject(0, ASuggestion, pointer(AOpCount));
  end;
begin
  if fSuggestions.IndexOf(ASuggestion) = -1 then begin
    if fSuggestions.Count = 0 then
      AddItem
    else begin
      if AOpCount < NativeInt(fSuggestions.Objects[0]) then begin //store top scoring words
        fAltSuggestions.SetStrings(fSuggestions);
        fSuggestions.Clear;
        AddItem;
      end else if AOpCount = NativeInt(fSuggestions.Objects[0]) then
        AddItem;
    end;
  end;
end;

function TSpellCheckFile.EditDistance(const AFromString, AToString: string): integer;
var
  matrix: array of array of integer;
  i, k, cost: integer;
begin
  SetLength(matrix, Length(AFromString)+1);
  for i := Low(matrix) to High(matrix) do
    SetLength(matrix[i], Length(AToString)+1);
  for i := Low(matrix) to High(matrix) do begin
    matrix[i, 0] := i;
    for k := Low(matrix[i]) to High(matrix[i]) do
      matrix[0, k] := k;
  end;
  for i := Low(matrix)+1 to High(matrix) do begin
    for k := Low(matrix[i])+1 to High(matrix[i]) do begin
      if AFromString[i] = AToString[k] then
        cost := 0
      else
        cost := 1;
      matrix[i, k] := Min(Min(matrix[i-1, k] + 1,
                              matrix[i,   k-1] + 1),
                              matrix[i-1, k-1] + cost);
    end;
  end;
  result := matrix[Length(AFromString), Length(AToString)];
end;

{ TSpellCheckThread }

constructor TSpellCheckThread.Create(const AOwner: TSpellCheck);
begin
  inherited Create(true);
  CoInitialize(nil);
  try
    if Assigned(AOwner) then
      AOwner.IncrementThreadCount;
    fSpellCheckFile := TSpellCheckFile.Create(AOwner);
    FreeOnTerminate := true;
  finally
    CoUnInitialize;
  end;
end;

destructor TSpellCheckThread.Destroy;
begin
  fSpellCheckFile.DisposeOf;
  inherited;
end;

procedure TSpellCheckThread.Execute;
begin
  inherited;
  try
    fSpellCheckFile.Run;
  finally
    fSpellCheckFile.Owner.DecrementThreadCount;
  end;
end;

end.


