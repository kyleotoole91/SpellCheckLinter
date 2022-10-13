unit uSpellCheck; //designed for code files .pas and .dfm. Only words within string literals are spell checked

interface

uses
  System.SysUtils, System.IOUtils, System.Types, System.UITypes, System.Classes, System.Variants, Generics.Collections, DateUtils, uConstants, System.RegularExpressions;

type
  TSpellCheck = class(TObject)
  strict private
    fIgnoreNextFirstWord: boolean;
    fCamelCaseWords: TStringList;
    fPossibleTruncation: boolean;
    fIsDFM: boolean;
    fIngoreFilePath: string;
    fUntrimmedLine: string;
    fIgnoreCodeFile: TStringList;
    fTextBeforeQuote: string;
    fSkippedLineCount: integer;
    fWordCheckCount: uInt64;
    fIgnorePathContaining: TStringList;
    fIgnoreLines: TStringList;
    fIgnoreWords: TStringList;
    fIgnoreFiles: TStringList;
    fLineWords: TStringList;
    fLanguageFile: TStringList;
    fSourceFile: TStringList;
    fErrors: TStringList;
    fStartTime: TDateTime;
    fEndTime: TDateTime;
    fFileCount: integer;
    fQuoteSym: string;
    fFileExtFilter: string;
    fSourcePath: string;
    fLanguageFilename: string;
    fIgnorelinesSymbol: string;
    fWordsDict: TDictionary<string, string>;
    fRecursive: boolean;
    fErrorWords: TStringList;
    fIgnoreContainsLines: TStringList;
    procedure SetIngoreFilePath(const Value: string);
    procedure BuildCamelCaseWords(const AArr: TArray<string>);
    procedure CleanLineWords;
    procedure Clear;
    procedure LoadIgnoreFiles;
    procedure LoadLanguageDictionary;
    procedure SetLanguageFilename(const Value: string);
    function NeedsSpellCheck(const AValue: string): boolean;
    function IsNumeric(const AString: string): boolean;
    function IsIgnoreLine(const ALine: string): boolean;
    function IsGuid(const ALine: string): boolean;
    function PosOfNextQuote(const AString: string; const AChar: string): integer;
    function SpellCheckFile(const AFilename: string): boolean;
    function SanitizeWord(const AWord: string; const ARemovePeriods: boolean=true): string;
    function RemoveStartEndQuotes(const AStr: string): string;
    function RemoveEscappedQuotes(const AStr: string): string;
    function InIgnoreCodeFile(const AText: string): boolean;
    function IngoredPathContaining(const AFilename: string): boolean;
  public
    constructor Create;
    destructor Destroy; override;
    function Run: boolean;
    function AddToIgnoreFile: boolean;
    function SpelledCorrectly(const AWord: string; const ATryLower: boolean=true): boolean;
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
    property ErrorsWords: TStringList read fErrorWords write fErrorWords;
    property IgnoreContainsLines: TStringList read fIgnoreContainsLines;
    property IngoreFilePath: string read fIngoreFilePath write SetIngoreFilePath;
  end;

implementation

uses
  System.Character;

{ TSpellCheck }

constructor TSpellCheck.Create;
begin
  inherited;
  fCamelCaseWords := TStringList.Create;
  fIngoreFilePath := cDefaultSourcePath;
  fSkippedLineCount := 0;
  fIgnorePathContaining := TStringList.Create;
  fIgnoreContainsLines := TStringList.Create;
  fErrorWords := TStringList.Create;
  fWordsDict := TDictionary<string, string>.Create;
  fLanguageFile := TStringList.Create;
  fSourceFile := TStringList.Create;
  fErrors := TStringList.Create;
  fIgnoreLines := TStringList.Create;
  fIgnoreWords := TStringList.Create;
  fIgnoreFiles := TStringList.Create;
  fIgnoreCodeFile := TStringList.Create;
  fLineWords := TStringList.Create;
  fLineWords.Delimiter := ' ';
  fLineWords.StrictDelimiter := true;
  fQuoteSym := cDefaultQuote;
  fRecursive := true;
  Clear;
end;

destructor TSpellCheck.Destroy;
begin
  try
    fCamelCaseWords.DisposeOf;
    fIgnoreCodeFile.DisposeOf;
    fIgnorePathContaining.DisposeOf;
    fIgnoreContainsLines.DisposeOf;
    fErrorWords.DisposeOf;
    fIgnoreLines.DisposeOf;
    fIgnoreWords.DisposeOf;
    fIgnoreFiles.DisposeOf;
    fLineWords.DisposeOf;
    fWordsDict.DisposeOf;
    fLanguageFile.DisposeOf;
    fSourceFile.DisposeOf;
    fErrors.DisposeOf;
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
  fLineWords.Clear;
  fWordsDict.Clear;
  fLanguageFile.Clear;
  fSourceFile.Clear;
  fErrors.Clear;
  fSourcePath := cDefaultSourcePath;
  fFileExtFilter := cDefaultExtFilter;
  fLanguageFilename := cDefaultlanguageName;
  fIgnoreContainsLines.Clear;
  fIgnorePathContaining.Clear;
  fTextBeforeQuote := '';
end;

function TSpellCheck.Run: boolean;
var
  filenames: TStringDynArray;
  filename: string;
begin
  result := true;
  fStartTime := Now;
  try
    try
      fErrorWords.Clear;
      LoadIgnoreFiles;
      LoadLanguageDictionary;
      if FileExists(fSourcePath) then
        result := SpellCheckFile(fSourcePath)
      else begin
        if Trim(fSourcePath) = '' then
          fSourcePath := '.\';
        if fRecursive then
          filenames := TDirectory.GetFiles(fSourcePath, fFileExtFilter, TSearchOption.soAllDirectories)
        else
          filenames := TDirectory.GetFiles(fSourcePath, fFileExtFilter, TSearchOption.soTopDirectoryOnly);
      end;
      for filename in filenames do begin
        if fIgnoreFiles.IndexOf(ExtractFileName(filename)) = -1 then begin
          if not IngoredPathContaining(filename) then begin
            Inc(fFileCount);
            if not SpellCheckFile(filename) then
              result := false;
          end;
        end;
      end;
    except
      on e: exception do begin
        result := false;
        fErrors.Add('Exception raised: '+e.ClassName+' '+e.Message);
      end;
    end;
  finally
    fEndTime := Now;
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

function TSpellCheck.NeedsSpellCheck(const AValue: string): boolean;
begin
  result := not (AValue.StartsWith('#') or
                 AValue.StartsWith('//') or
                 AValue.StartsWith('/') or
                 AValue.StartsWith('\') or
                 AValue.StartsWith('\\'));
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

function TSpellCheck.IsIgnoreLine(const ALine: string): boolean;
begin
  if fIgnorelinesSymbol= '' then begin
    if fTextBeforeQuote.Contains(cSpellCheckOff) then
      fIgnorelinesSymbol := cSpellCheckOn
    else if fTextBeforeQuote.Contains('{') then
      fIgnorelinesSymbol := '}'
    else if fTextBeforeQuote.Contains('(*') then
      fIgnorelinesSymbol := '*)'
    else if fTextBeforeQuote.Contains('<html>') then
      fIgnorelinesSymbol := '</html>'
    else if ALine.Contains('SELECT') or
            ALine.Contains('TABLE') or
            ALine.Contains('INSERT') or
            ALine.Contains('UPDATE') or
            ALine.Contains('DELETE') or
            ALine.Contains('GROUP BY') or
            ALine.Contains('ORDER BY') or
            ALine.Contains(' AND ') or
            ALine.Contains(' WHERE ') or
            ALine.Contains('LEFT OUTER JOIN') or
            ALine.Contains('INNER JOIN') or
            InIgnoreCodeFile(fTextBeforeQuote) then begin
      if fIsDFM then
        fIgnorelinesSymbol := cSkipLineEndStringDFM
      else
        fIgnorelinesSymbol := cSkipLineEndString;
    end else if fTextBeforeQuote.StartsWith('function GetIcon(') or
            fTextBeforeQuote.StartsWith('function ChartTypeText(') or
            fTextBeforeQuote.StartsWith('function PieChartTypeText(') or
            fTextBeforeQuote.StartsWith('procedure TwcPieChart.AfterLoadDFMValues;') then
      fIgnorelinesSymbol := cSkipLineEndOfFunctionBlock;
  end else begin
    Inc(fSkippedLineCount);
    if not fIsDFM then begin
      if (fIgnorelinesSymbol = cSkipLineEndOfFunctionBlock) and
         (fUntrimmedLine.StartsWith(fIgnorelinesSymbol)) then
        fIgnorelinesSymbol := ''
      else if ((fIgnorelinesSymbol = cSkipLineEndString) or (fIgnorelinesSymbol = '</html>')) and
              (fSkippedLineCount > cSkippedLineCountLimit) then //break out if limit is reached, eg </ html> instead of </html>
        fIgnorelinesSymbol := '';
    end;
  end;
  result := (ALine.Contains('//')) or
            (fIgnorelinesSymbol <> '') or
            (ALine.StartsWith('function')) or
            (ALine.StartsWith('procedure')) or
            (ALine.StartsWith('  function')) or
            (ALine.StartsWith('  procedure'));
  if (fIgnorelinesSymbol <> '') then begin
    if (((fIsDFM) and ((ALine.StartsWith(fIgnorelinesSymbol)))) or
        ((not fIsDFM) and (ALine.Contains(fIgnorelinesSymbol)))) then
      fIgnorelinesSymbol := '';
  end;
  if fIgnorelinesSymbol = '' then
    fSkippedLineCount := 0;
end;

function TSpellCheck.IsGuid(const ALine: string): boolean;
begin
  result := (ALine.Length >= cGuidLen) and
            (ALine.Chars[0] = '{') and
            (ALine.Chars[ALine.Length-1] = '}') and
            (ALine.Contains('-'));
end;

function TSpellCheck.SpellCheckFile(const AFilename: string): boolean;
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
  procedure AddError(const AError: string);
  begin
    result := false;
    fErrors.Add('Error: '+AError);
    fErrors.Add('File name: '+AFilename);
    fErrors.Add('Line number: '+IntToStr(i+1));
    fErrors.Add('Line text: '+fUntrimmedLine);
    fErrors.Add(' ');
    fErrorWords.Add(theWord);
  end;
  function ContainsIgnore: boolean;
  var
    a: integer;
  begin
    result := (not IsIgnoreLine(lineStr)) and
              (fIgnoreLines.IndexOf(Trim(lineStr)) = -1);
    if not result then begin
      for a := 0 to fIgnoreContainsLines.Count-1 do begin
        result := lineStr.Contains(fIgnoreContainsLines.Strings[a]);
        if result then
          Break;
      end;
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
        str := Copy(str, Pos(fQuoteSym, str)+1, str.Length);
        str := Copy(str, 1, PosOfNextQuote(str, fQuoteSym)-1);
        sl.CommaText := Str;
        if sl.Count >= 1 then
          result := Trim(sl.Strings[0]);
      end;
    finally
      sl.Free;
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
          AddError(AWord);
      end;
    end;
    result := not hasError;
  end;
begin
  result := true;
  fIgnorelinesSymbol := '';
  fIsDFM := ExtractFileExt(AFilename) = '.dfm';
  fSourceFile.LoadFromFile(AFilename);
  fIgnoreNextFirstWord := false;
  for i:=0 to fSourceFile.Count-1 do begin
    try
      fUntrimmedLine := fSourceFile.Strings[i];
      lineStr := Trim(fSourceFile.Strings[i]);
      if Pos(fQuoteSym, lineStr) >= 1 then
        fTextBeforeQuote := Copy(lineStr, 1, Pos(fQuoteSym, lineStr)-1)
      else
        fTextBeforeQuote := lineStr;
      if (lineStr <> '') and
         (not ContainsIgnore) then begin
        itrCount := 0;
        while Pos(fQuoteSym, lineStr) >= 1 do begin //while the line has string literals
          Inc(itrCount);
          lineStr := Copy(lineStr, Pos(fQuoteSym, lineStr)+1, lineStr.Length);
          if itrCount >= cBreakout then
            raise Exception.Create('An infinite loop has been detected in '+AFilename+' on line '+IntToStr(i+1)+'. '+
                                   'Please check the syntax of file. ');
          theStr := Copy(lineStr, 1, PosOfNextQuote(lineStr, fQuoteSym)-1);
          fPossibleTruncation := fIsDFM and (theStr.Length = cMaxStringLengthDFM);
          if Pos(theStr, lineStr)+theStr.Length <> 0 then //remove the string from the line
            lineStr := Copy(lineStr, Pos(theStr, lineStr)+theStr.Length+1, lineStr.Length)
          else
            lineStr := Copy(lineStr, 3, lineStr.Length); //if empty string, remove the next two chars
          if (not Trim(theStr).StartsWith('<')) and //ignore html
             (ExtractFilePath(Trim(theStr)) = '') then begin //ignore file paths
            theStr := theStr.Replace('&', ''); //underlining menu items
            if NeedsSpellCheck(lineStr) then begin
              if IsGuid(theStr) then
                fLineWords.DelimitedText := ''
              else begin
                theStr := SanitizeWord(theStr, false);
                if theStr.EndsWith(fQuoteSym) then //strings ending with double quotes was causing the last char to get removed when setting delimited text
                  theStr := Copy(theStr, 1, theStr.Length-1)+' '+fQuoteSym;
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
                      if not SpelledCorrectly(theWord) then
                        AddError(theWord);
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
                            if not ok then 
                              AddError(theWord+'/'+theWord+nextFirstWord+'/'+nextFirstWord);
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
        fErrors.Add('Exception raised: '+e.ClassName+' '+e.Message);
      end;
    end;
  end;
end;

function TSpellCheck.SpelledCorrectly(const AWord: string; const ATryLower: boolean=true): boolean;
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
              (fIgnoreWords.IndexOf(word) >= 0); //Don't ignore file for the word
    if not result then begin
      result := (fWordsDict.ContainsKey(word)) or
                (ATryLower and fWordsDict.ContainsKey(LowerCase(word)));
      Inc(fWordCheckCount);
    end;
  end;
end;

function TSpellCheck.PosOfNextQuote(const AString, AChar: string): integer;
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

function TSpellCheck.SanitizeWord(const AWord: string; const ARemovePeriods: boolean=true): string;
begin
  result := Trim(TRegEx.Replace(AWord, cRegExKeepLettersAndQuotes, ' '));
  result := Trim(RemoveStartEndQuotes(result));
  if ARemovePeriods then
    result := result.Replace('.', ' ');
end;

function TSpellCheck.RemoveEscappedQuotes(const AStr: string): string;
begin
  result := Trim(TRegEx.Replace(AStr, cRegExRemoveNumbers, ''));
  if fQuoteSym = '''' then //remove Delphi escaped quotes and replace with one single quote
    result := result.Replace('''''', '''');
end;

procedure TSpellCheck.BuildCamelCaseWords(const AArr: TArray<string>); //eg msmWeb, don't check msm
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

function TSpellCheck.RemoveStartEndQuotes(const AStr: string): string;
begin
  result := AStr;
  if (result.Length > 0) and
     (result[1] = fQuoteSym) then
    result := Copy(result, 2, result.Length);
  if (result.Length > 0) and
     (result[result.Length] = fQuoteSym) then
    result := Copy(result, 1, result.Length-1);
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

procedure TSpellCheck.CleanLineWords;
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

procedure TSpellCheck.SetIngoreFilePath(const Value: string);
begin
  fIngoreFilePath := IncludeTrailingPathDelimiter(Value);
end;

procedure TSpellCheck.SetLanguageFilename(const Value: string);
begin
  fLanguageFilename := Value;
  LoadLanguageDictionary;
end;

function TSpellCheck.IsNumeric(const AString: string): boolean;
var
  a: integer;
begin
  result := false;
  {$WARNINGS OFF}
  for a := 1 to AString.Length do begin
    result := System.Character.IsNumber(AString, a);
    if not result then
      break;
  end;
  {$WARNINGS ON}
end;

end.


