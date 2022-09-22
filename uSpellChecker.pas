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
    fWordsDict: TDictionary<string, string>;
    fRecursive: boolean;
    procedure Clear;
    procedure LoadIgnoreFiles;
    procedure LoadLanguageDictionary;
    procedure SetLanguageFilename(const Value: string);
    function PosOfOccurence(const AString: string; const AChar: string; AOccurencePos: integer=1): integer;
    function OccurrenceCount(const AString: string; const AChar: string): integer;
    function SpellCheckFile(const AFilename: string): boolean;
    function SanitizeWord(const AWord: string): string;
    function RemoveStartEndQuotes(const AStr: string): string;
  public
    constructor Create;
    destructor Destroy; override;
    function Run: boolean;
    function SpelledCorrectly(const AWord: string): boolean;
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
  end;

implementation

const
  cRegExCamelPascalCaseSpliter='([A-Z]+|[A-Z]?[a-z]+)(?=[A-Z]|\b)';

{ TSpellChecker }

constructor TSpellChecker.Create;
begin
  inherited;    
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

procedure TSpellChecker.LoadIgnoreFiles;
begin
  if FileExists(cIgnoreFiles) then
    fIgnoreFiles.LoadFromFile(cIgnoreFiles);
  if FileExists(cIgnoreWords) then
    fIgnoreWords.LoadFromFile(cIgnoreWords);
  if FileExists(cIgnoreLines) then
    fIgnoreLines.LoadFromFile(cIgnoreLines);
end;

function TSpellChecker.SpelledCorrectly(const AWord: string): boolean;
begin
  result := (Trim(AWord) = '') or
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
begin
  result := true;
  fSourceFile.LoadFromFile(AFilename);
  for i := 0 to fSourceFile.Count-1 do begin
    lineStr := fSourceFile.Strings[i];
    if (Trim(lineStr) <> '') and
       (fIgnoreLines.IndexOf(Trim(lineStr)) = -1) then begin
      while OccurrenceCount(lineStr, fQuoteSym) > 1 do begin
        theStr := RemoveStartEndQuotes(Copy(lineStr, pos(fQuoteSym, lineStr), lineStr.Length));
        quotePos := PosOfOccurence(theStr, fQuoteSym, 1);
        if quotePos >= 0 then
          theStr := RemoveStartEndQuotes(Copy(theStr, 0, quotePos));
        lineStr := lineStr.Replace(fQuoteSym + theStr + fQuoteSym, '');
        if (not theStr.Contains('\')) and //ignore path strings
           (not theStr.Contains('/')) then begin
          theStr := SanitizeWord(theStr);
          if Trim(theStr.Replace(fQuoteSym, '')) = '' then //break, if the string only containes quotes and spaces
            Break
          else
            fLineWords.DelimitedText := SanitizeWord(theStr);
          theWord := '';
          for j := 0 to fLineWords.Count-1 do begin
            theWord := SanitizeWord(fLineWords.Strings[j]);
            if Trim(theWord) <> '' then begin
              if theWord.Contains('''') then begin //the rexex will split on this. You don't normally see 's in camelCase or PascalCase words
                if not SpelledCorrectly(theWord) then begin
                  result := false;
                  fErrors.Add(AFilename+' (Line '+IntToStr(i+1)+')'+': '+theWord);
                end;
              end else begin
                camelCaseWords := TRegEx.Split(theWord, cRegExCamelPascalCaseSpliter);
                for k := 0 to Length(camelCaseWords)-1 do begin
                  theWord := camelCaseWords[k];
                  if (Trim(theWord) <> '') and
                     (not SpelledCorrectly(theWord)) then begin
                    result := false;
                    fErrors.Add(AFilename+' (Line '+IntToStr(i+1)+')'+': '+theWord);
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
      if (AString[i+1] = AChar) then begin //detect escaped delphi quote eg: 'couldn''t'
        Inc(AOccurencePos);
        Continue; //move onto next the occurence, since you cannot increment the loop var
      end;
      Inc(occur);
      if occur = AOccurencePos then begin
        result := i;
        Break;
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

function TSpellChecker.SanitizeWord(const AWord: string): string;
begin
  result := AWord.Replace(':', ' ')
                 .Replace('.', ' ')
                 .Replace(',', ' ')
                 .Replace('*', ' ')
                 .Replace('-', ' ')
                 .Replace('/', ' ')
                 .Replace(')', ' ')
                 .Replace('(', ' ')
                 .Replace(']', ' ')
                 .Replace('[', ' ')
                 .Replace(';', ' ')
                 .Replace('=', ' ')
                 .Replace('&', ' ')
                 .Replace('?', ' ')
                 .Replace('>', ' ')
                 .Replace('<', ' ')
                 .Replace(';', ' ')
                 .Replace('_', ' ');
  result := Trim(RemoveStartEndQuotes(result));
  if fQuoteSym = '''' then //remove Delphi escaped quotes and replace with one single quote
    result := result.Replace('''''', '''');
end;

procedure TSpellChecker.SetLanguageFilename(const Value: string);
begin
  fLanguageFilename := Value;
  LoadLanguageDictionary;
end;

end.
