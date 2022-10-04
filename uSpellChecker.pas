unit uSpellChecker; //designed for code files. Only words within string literals are spell checked

interface

uses
  System.SysUtils, System.IOUtils, System.Types, System.UITypes, System.Classes, System.Variants, Generics.Collections, DateUtils, uConstants, System.RegularExpressions;

type
  TSpellChecker = class(TObject)
  strict private
    fWordCheckCount: uInt64;
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
    fMultiCommentSym: string;
    fWordsDict: TDictionary<string, string>;
    fRecursive: boolean;
    fErrorWords: TStringList;
    fIgnoreContainsLines: TStringList;
    procedure CleanLineWords;
    procedure Clear;
    procedure LoadIgnoreFiles;
    procedure LoadLanguageDictionary;
    procedure SetLanguageFilename(const Value: string);
    function RemoveEmptyStringQuotes(const ALine: string): string;
    function NeedsSpellCheck(const AValue: string): boolean;
    function IsNumeric(const AString: string): boolean;
    function IsCommentLine(const ALine: string): boolean;
    function IsGuid(const ALine: string): boolean;
    function PosOfNextQuote(const AString: string; const AChar: string): integer;
    function OccurrenceCount(const AString: string; const AChar: string): integer;
    function SpellCheckFile(const AFilename: string): boolean;
    function SanitizeWord(const AWord: string; const ARemovePeriods: boolean=true): string;
    function RemoveStartEndQuotes(const AStr: string): string;
    function RemoveEscappedQuotes(const AStr: string): string;
  public
    constructor Create;
    destructor Destroy; override;
    function Run: boolean;
    function AddToIgnoreFile: boolean;
    function SpelledCorrectly(AWord: string; const ATryLower: boolean=true): boolean;
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
  end;

implementation

uses
  System.Character;

const
  cSQLEndStr=';';

{ TSpellChecker }

constructor TSpellChecker.Create;
begin
  inherited;
  fIgnoreContainsLines := TStringList.Create;
  fErrorWords := TStringList.Create;
  fWordsDict := TDictionary<string, string>.Create;
  fLanguageFile := TStringList.Create;
  fSourceFile := TStringList.Create;
  fErrors := TStringList.Create;
  fIgnoreLines := TStringList.Create;
  fIgnoreWords := TStringList.Create;
  fIgnoreFiles := TStringList.Create;
  fLineWords := TStringList.Create;
  fLineWords.Delimiter := ' ';
  fQuoteSym := cDefaultQuote;
  fRecursive := true;
  Clear;
end;

destructor TSpellChecker.Destroy;
begin
  try
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

function TSpellChecker.IsNumeric(const AString: string): boolean;
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

function TSpellChecker.AddToIgnoreFile: boolean;
var
  sl: TStringList;
  a: integer;
begin
  result := true;
  try
    sl := TStringList.Create;
    if FileExists('.\'+cIgnoreWords) then
      sl.LoadFromFile(cIgnoreWords);
    sl.Add(fErrorWords.Text);
    for a := 0 to fErrorWords.Count-1 do begin
      if sl.IndexOf(fErrorWords.Strings[a]) = -1 then
        sl.Add(fErrorWords.Strings[a]);
    end;
    sl.SaveToFile('.\'+cIgnoreWords);
  except
    on e: exception do begin
      result := false;
      fErrors.Add(e.Classname+' '+e.Message);
    end;
  end;
end;

procedure TSpellChecker.CleanLineWords;
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

procedure TSpellChecker.Clear;
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
  fLanguageFilename := cDefaultlanguagePath;
  fIgnoreContainsLines.Clear;
end;

function TSpellChecker.Run: boolean;
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
          Inc(fFileCount);
          result := SpellCheckFile(filename) and result;
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

procedure TSpellChecker.LoadLanguageDictionary;
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

function TSpellChecker.NeedsSpellCheck(const AValue: string): boolean;
begin
  result := not (AValue.StartsWith('#') or
                 AValue.StartsWith('//') or
                 AValue.StartsWith('/') or
                 AValue.StartsWith('\') or
                 AValue.StartsWith('\\'));
end;

function TSpellChecker.IsCommentLine(const ALine: string): boolean;
begin
  if ALine.StartsWith('{') then
    fMultiCommentSym := '}'
  else if ALine.StartsWith('(*') then
    fMultiCommentSym := '*)'
  else if ALine.Contains('SELECT') or
          ALine.Contains('INSERT') or
          ALine.Contains('UPDATE') or
          ALine.Contains('DELETE') then
    fMultiCommentSym := cSQLEndStr;
  result := (ALine.StartsWith('//')) or
            (fMultiCommentSym <> '');
  if (fMultiCommentSym <> '') and
     (ALine.Contains(fMultiCommentSym)) then //continue to return false until a closing tag is found, resume checking on the next line
    fMultiCommentSym := '';
end;

function TSpellChecker.IsGuid(const ALine: string): boolean;
begin
  result := (ALine.Chars[0] = '{') and
            (ALine.Chars[ALine.Length-1] = '}');
  if result then
    result := (ALine.Contains('-')) and
              (ALine.Length >= cGuidLen)
end;

procedure TSpellChecker.LoadIgnoreFiles;
begin
  fIgnoreContainsLines.Add('dh.');//sql
  fIgnoreContainsLines.Add('sql');
  fIgnoreContainsLines.Add('Sql');
  fIgnoreContainsLines.Add('SQL');
  fIgnoreContainsLines.Add('FieldByName(');
  fIgnoreContainsLines.Add('WriteElement('); //xml
  fIgnoreContainsLines.Add('FormatDateTime(');
  fIgnoreContainsLines.Add('ChildNodes[');
  fIgnoreContainsLines.Add('jString[');  //super object
  fIgnoreContainsLines.Add('jInteger[');
  fIgnoreContainsLines.Add('jArray[');
  fIgnoreContainsLines.Add('jFloat[');
  fIgnoreContainsLines.Add('jDouble[');
  fIgnoreContainsLines.Add('.Read'); //ini files
  fIgnoreContainsLines.Add('.Write');
  fIgnoreContainsLines.Add('.read');
  fIgnoreContainsLines.Add('.write');
  fIgnoreContainsLines.Add('<table');
  if FileExists(cIgnoreFiles) then
    fIgnoreFiles.LoadFromFile(cIgnoreFiles);
  if FileExists(cIgnoreWords) then
    fIgnoreWords.LoadFromFile(cIgnoreWords);
  if FileExists(cIgnoreLines) then
    fIgnoreLines.LoadFromFile(cIgnoreLines);
end;

function TSpellChecker.SpelledCorrectly(AWord: string; const ATryLower: boolean=true): boolean;
begin
  result := Trim(AWord) = '';
  if not result then begin
    AWord := Trim(TRegEx.Replace(AWord, cRegExKeepLettersAndQuotes, ' '));
    result := (AWord.Length < cMinCheckLength) or //Don't check short words
              (UpperCase(AWord) = AWord) or //IGNORE UPPER CASE TEXT
              (IsNumeric(AWord)) or //Don't check numbers
              (fIgnoreWords.IndexOf(AWord) >= 0); //Don't ignore file for the word
    if not result then begin
      result := (fWordsDict.ContainsKey(AWord)) or
                (ATryLower and fWordsDict.ContainsKey(LowerCase(AWord)));
      Inc(fWordCheckCount);
    end;
  end;
end;

function TSpellChecker.SpellCheckFile(const AFilename: string): boolean;
const
  cBreakout=1000;
var
  itrCount: integer;
  i, j, k: integer;
  theStr: string;
  lineStr: string;
  quotePos: integer;
  theWord: string;
  camelCaseWords: TArray<string>;
  fileExt: string;
  lastWordChecked: string;
  procedure AddError;
  begin
    result := false;
    fErrors.Add(AFilename+' line '+IntToStr(i+1)+': '+theWord);
    fErrorWords.Add(theWord);
  end;
  function ContainsIgnore: boolean;
  var
    a: integer;
  begin
    result := false;
    for a := 0 to fIgnoreContainsLines.Count-1 do begin
      result := lineStr.Contains(fIgnoreContainsLines.Strings[a]);
      if result then
        Break;
    end;
  end;
begin
  result := true;
  fSourceFile.LoadFromFile(AFilename);
  for i:=0 to fSourceFile.Count-1 do begin
    try
      lineStr := RemoveEmptyStringQuotes(fSourceFile.Strings[i]);
      if (lineStr <> '') and
         (not IsCommentLine(lineStr)) and
         (not ContainsIgnore) and
         (fIgnoreLines.IndexOf(Trim(lineStr)) = -1) then begin
        itrCount := 0;
        while OccurrenceCount(lineStr.Replace('''''', ''), fQuoteSym) > 1 do begin //while the line has string literals
          Inc(itrCount);
          if itrCount >= cBreakout then
            raise Exception.Create('An infinite loop has been detected in '+AFilename+' on line '+IntToStr(i+1)+'. '+
                                   'Please check the syntax of file. ');
          theStr := RemoveStartEndQuotes(Copy(lineStr, pos(fQuoteSym, lineStr), lineStr.Length));
          if not (theStr.Contains('/') or theStr.Contains('\')) then begin
            quotePos := PosOfNextQuote(theStr, fQuoteSym);
            if quotePos >= 0 then
              theStr := RemoveStartEndQuotes(Copy(theStr, 0, quotePos+1));
            if (Trim(theStr) <> '') and
               (Pos(theStr,lineStr) > 0) then
              lineStr := Trim(Copy(lineStr, pos(theStr, lineStr)+theStr.Length+1, lineStr.Length));
            if NeedsSpellCheck(lineStr) then begin
              theStr := SanitizeWord(theStr, false);
              if Trim(theStr.Replace(fQuoteSym, '')) = '' then //break, if the string only contains quotes and spaces
                Break
              else if IsGuid(theStr) then
                fLineWords.DelimitedText := ''
              else
                fLineWords.DelimitedText := theStr;
              CleanLineWords;
              theWord := '';
              for j := 0 to fLineWords.Count-1 do begin //for each word in the line
                theWord := SanitizeWord(fLineWords.Strings[j], false);
                if Trim(theWord) <> '' then begin
                  lastWordChecked := theWord;
                  if theWord.Contains('''') then begin //the regex will split on this. You don't normally see apostrophies in camelCase or PascalCase words anyway
                    lineStr := lineStr.Remove(pos(fLineWords.Strings[j],lineStr)-1, theWord.Length);
                    if not SpelledCorrectly(RemoveEscappedQuotes(theWord)) then
                      AddError;
                  end else begin
                    fileExt := ExtractFileExt(theWord);
                    if (fileExt <> '') and //remove file extensions
                       (fileExt.Length = cFileExtLen) and
                       (fileExt.Length <> theWord.Length) then
                      theWord := theWord.Replace(ExtractFileExt(theWord), '');
                    camelCaseWords := TRegEx.Split(theWord, cRegExCamelPascalCaseSpliter);
                    for k := 0 to Length(camelCaseWords)-1 do begin //for each camel case word
                      theWord := camelCaseWords[k];
                      if (not SpelledCorrectly(theWord)) then
                        AddError;
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

function TSpellChecker.OccurrenceCount(const AString: string; const AChar: string): integer;
var
  i: integer;
begin
  result := 0;
  for i:=1 to AString.Length do begin
    if (AString[i] = AChar) and
       (AString[i+1] <> AChar) then //detect escaped delphi quote eg: 'couldn''t'
      Inc(result);
  end;
end;

function TSpellChecker.PosOfNextQuote(const AString, AChar: string): integer;
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

function TSpellChecker.RemoveEmptyStringQuotes(const ALine: string): string;
begin
  result := Trim(ALine);
  result := result.Replace(''''''''',', '')
                  .Replace(''''''''')', '')
                  .Replace(' = '''''''' ', '')
                  .Replace(' = '''''''' ', '')
                  .Replace('='''''''' ', '')
end;

function TSpellChecker.RemoveEscappedQuotes(const AStr: string): string;
begin
  result := Trim(AStr);
  if fQuoteSym = '''' then //remove Delphi escaped quotes and replace with one single quote
    result := AStr.Replace('''''', '''');
end;

function TSpellChecker.RemoveStartEndQuotes(const AStr: string): string;
begin
  result := AStr;
  if (result.Length > 0) and
     (result[1] = fQuoteSym) then
    result := Copy(result, 2, result.Length);
  if (result.Length > 0) and
     (result[result.Length] = fQuoteSym) then
    result := Copy(result, 1, result.Length-1);
end;

function TSpellChecker.SanitizeWord(const AWord: string; const ARemovePeriods: boolean=true): string;
begin
  result := Trim(TRegEx.Replace(AWord, cRegExKeepLettersAndQuotes, ' '));
  result := Trim(RemoveStartEndQuotes(result));
  if ARemovePeriods then
    result := result.Replace('.', ' ');
end;

procedure TSpellChecker.SetLanguageFilename(const Value: string);
begin
  fLanguageFilename := Value;
  LoadLanguageDictionary;
end;

end.


