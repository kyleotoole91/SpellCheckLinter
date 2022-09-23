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
    procedure Clear;
    procedure LoadIgnoreFiles;
    procedure LoadLanguageDictionary;
    procedure SetLanguageFilename(const Value: string);
    function NeedsSpellCheck(const AValue: string): boolean;
    function IsNumeric(const AString: string): boolean;
    function IsCommentLine(const ALine: string): boolean;
    function IsGuid(const ALine: string): boolean;
    function MultiCommentCloseTag: string;
    function PosOfOccurence(const AString: string; const AChar: string; AOccurencePos: integer=1): integer;
    function OccurrenceCount(const AString: string; const AChar: string): integer;
    function SpellCheckFile(const AFilename: string): boolean;
    function SanitizeWord(const AWord: string; const ARemovePeriods: boolean=true): string;
    function RemoveStartEndQuotes(const AStr: string): string;
  public
    constructor Create;
    destructor Destroy; override;
    function Run: boolean;
    function SpelledCorrectly(AWord: string): boolean;
    function AddToIgnoreFile: boolean; 
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
  end;

implementation

uses
  System.Character;

{ TSpellChecker }

constructor TSpellChecker.Create;
begin
  inherited;
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
        if fIgnoreFiles.IndexOf(filename) = -1 then begin
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
    fMultiCommentSym := '{'
  else if ALine.StartsWith('/*') then
    fMultiCommentSym := '{';
  result := ALine.StartsWith('//') or (fMultiCommentSym <> '');
  if ALine.Contains(MultiCommentCloseTag) then //continue to return false until a closing tag is found, resume checking on the next line
    fMultiCommentSym := '';
end;

function TSpellChecker.IsGuid(const ALine: string): boolean;
begin
  result := (ALine.Chars[0] = '{') and
            (ALine.Chars[ALine.Length-1] = '}');
  if result then
    result := ALine.Contains('-') and
              (ALine.Length >= cGuidLen)
end;

function TSpellChecker.MultiCommentCloseTag: string;
begin
  if fMultiCommentSym = '{' then
    result := '}'
  else
    result := '';
end;

procedure TSpellChecker.LoadIgnoreFiles;
begin
  if FileExists(cIgnoreFiles) then
    fIgnoreFiles.LoadFromFile(cIgnoreFiles);
  if FileExists(cIgnoreWords) then
    fIgnoreWords.LoadFromFile(cIgnoreWords);
  if FileExists(cIgnoreLines) then
    fIgnoreLines.LoadFromFile(cIgnoreLines);
end;

function TSpellChecker.SpelledCorrectly(AWord: string): boolean;
begin  
  AWord := Trim(TRegEx.Replace(AWord, cRegExKeepLettersAndQuotes, ' '));
  result := (AWord.Length < cMinCheckLength) or
            (UpperCase(AWord) = AWord) or
            (IsNumeric(AWord)) or
            (fIgnoreWords.IndexOf(AWord) >= 0);
  if not result then begin
    result := fWordsDict.ContainsKey(LowerCase(AWord));
    Inc(fWordCheckCount);
  end;
end;

function TSpellChecker.SpellCheckFile(const AFilename: string): boolean;
var
  i, j, k: integer;
  theStr: string;
  lineStr: string;
  quotePos: integer;
  theWord: string;
  camelCaseWords: TArray<string>;
  fileExt: string;
const
  cEmptyStr=' ''''';
  cEmptyStr2=' '''',';
begin
  result := true;
  fSourceFile.LoadFromFile(AFilename);
  for i := 0 to fSourceFile.Count-1 do begin
    lineStr := Trim(fSourceFile.Strings[i]);
    if (lineStr <> '') and
       (not IsCommentLine(lineStr)) and
       (fIgnoreLines.IndexOf(Trim(lineStr)) = -1) then begin
      lineStr := lineStr.Replace(cEmptyStr2, '')
                        .Replace(cEmptyStr, ''); //remove emtpy strings, without this everything after an empty string ('' or '',) is ignored
      while OccurrenceCount(lineStr.Replace('''''', ''), fQuoteSym) > 1 do begin
        theStr := RemoveStartEndQuotes(Copy(lineStr, pos(fQuoteSym, lineStr), lineStr.Length));
        quotePos := PosOfOccurence(theStr, fQuoteSym, 1);
        if (quotePos >= 0) then
          theStr := RemoveStartEndQuotes(Copy(theStr, 0, quotePos));
        lineStr := lineStr.Replace(fQuoteSym + theStr + fQuoteSym, '');
        if NeedsSpellCheck(lineStr) then begin
          theStr := SanitizeWord(theStr, false);
          if Trim(theStr.Replace(fQuoteSym, '')) = '' then //break, if the string only containes quotes and spaces
            Break
          else if IsGuid(theStr) then
            fLineWords.DelimitedText := ''
          else
            fLineWords.DelimitedText := theStr;
          theWord := '';
          for j := 0 to fLineWords.Count-1 do begin
            theWord := SanitizeWord(fLineWords.Strings[j], false);
            if Trim(theWord) <> '' then begin
              if theWord.Contains('''') then begin //the regex will split on this. You don't normally see apostrophies in camelCase or PascalCase words
                if not SpelledCorrectly(theWord) then begin
                  result := false;
                  fErrors.Add(AFilename+' (Line '+IntToStr(i+1)+')'+': '+theWord);
                  fErrorWords.Add(theWord);
                end;
              end else begin
                fileExt := ExtractFileExt(theWord);
                if (fileExt <> '') and //remove file extensions
                   ((fileExt.Length = cFileExtLen) or (fileExt.Length = cFileExtLen+1)) then
                  theWord := theWord.Replace(ExtractFileExt(theWord), '');
                camelCaseWords := TRegEx.Split(theWord, cRegExCamelPascalCaseSpliter);
                for k := 0 to Length(camelCaseWords)-1 do begin
                  theWord := camelCaseWords[k];
                  if (Trim(theWord) <> '') and
                     (not SpelledCorrectly(Trim(theWord))) then begin
                    result := false;
                    fErrors.Add(AFilename+' (Line '+IntToStr(i+1)+')'+': '+theWord);
                    fErrorWords.Add(theWord);
                  end;
                end;
              end;
            end;
          end;
        end;
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

function TSpellChecker.PosOfOccurence(const AString, AChar: string; AOccurencePos: integer): integer;
var
  i, occur: integer;
begin
  result := -1;
  occur := 0;
  for i:=1 to AString.Length do begin
    if (AString[i] = AChar) then begin
      if (AString[i+1] = AChar) then //detect escaped delphi quote eg: 'couldn''t'
        Inc(AOccurencePos)
      else begin
        Inc(occur);
        if occur = AOccurencePos then begin
          result := i;
          Break;
        end;
      end;
    end;
  end;
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
  if fQuoteSym = '''' then //remove Delphi escaped quotes and replace with one single quote
    result := result.Replace('''''', '''');
  if ARemovePeriods then
    result := result.Replace('.', ' ');
end;

procedure TSpellChecker.SetLanguageFilename(const Value: string);
begin
  fLanguageFilename := Value;
  LoadLanguageDictionary;
end;

end.
