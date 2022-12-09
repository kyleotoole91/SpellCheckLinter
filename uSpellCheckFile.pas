unit uSpellCheckFile;

interface

uses
  System.SysUtils, System.Classes, uConstants,
  System.RegularExpressions, uSpellcheckLinter;

type
  TSpellCheckFile = class(TObject)
  strict private
    fOwner: TSpellCheckLinter;
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
    fLineNum: integer;
    fCamelCaseWords: TStringList;
    fLineWords: TStringList;
    fLineStr: string;
    fTheStr: string;
    fTheWord: string;
    fQuoteSym: string;
    procedure CleanLineWords;
    procedure PrepareLineString;
    procedure BuildLineWords(AString: string);
    function EditDistance(const AFromString, AToString: string): integer;
    procedure AddSuggestion(const ASuggestion: string; const AOpCount: integer; const AInsert: boolean=false);
    procedure BuildSuggestions(const AMispelledWord: string);
    procedure BuildCamelCaseWords(const AArr: TArray<string>);
    procedure LimitItems(const AStringList: TStringList);
    procedure InitNewFile;
    procedure InitNewLine(const ALineIdx: integer);
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
    function ContainsIgnore(const ALine: string): boolean;
    function CheckCamelCaseWords(AWord: string; const ALogError: boolean=true): boolean;
    function FirstWordNextLine: string;
    function IsFileExtension(const AStr: string): boolean;
    procedure SpellCheckWithoutRegex(const ALineWordIndex: integer);
    procedure SpellCheckWord(const AIsLast: boolean);
    function IgnoreLine: boolean;
    function IgnoreString: boolean;
    procedure CheckEachWordInString;
  public
    constructor Create(const AOwner: TSpellCheckLinter); reintroduce;
    destructor Destroy; override;
    procedure Run;
    property Filename: string read fFilename write fFilename;
    property ProvideSuggestions: boolean read fProvideSuggestions write fProvideSuggestions;
    property Owner: TSpellCheckLinter read fOwner;
  end;

  TSpellCheckThread = class(TThread)
  strict private
    fSpellCheckFile: TSpellCheckFile;
  protected
    procedure Execute; override;
  public
    constructor Create(const AOwner: TSpellCheckLinter); reintroduce;
    destructor Destroy; override;
    property SpellCheckFile: TSpellCheckFile read fSpellCheckFile;
  end;

implementation

uses
  System.Character, Math, ActiveX;

{ TSpellCheckFile }

constructor TSpellCheckFile.Create(const AOwner: TSpellCheckLinter);
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

function TSpellCheckFile.ContainsIgnore(const ALine: string): boolean;
var
  a: integer;
begin
  result := false;
  if Assigned(fOwner) then begin
    for a := 0 to fOwner.IgnoreContainsLines.Count-1 do begin
      result := ALine.Contains(fOwner.IgnoreContainsLines.Strings[a]);
      if result then
        Break;
    end;
  end;
end;

function TSpellCheckFile.CheckCamelCaseWords(AWord: string; const ALogError: boolean=true): boolean;
var
  k: integer;
  hasError: boolean;
begin
  hasError := false;
  BuildCamelCaseWords(TRegEx.Split(AWord, cRegExCamelPascalCaseSpliter));
  for k := 0 to fCamelCaseWords.Count-1 do begin
    AWord := RemoveEscappedQuotes(fCamelCaseWords.Strings[k]);
    if not SpelledCorrectly(AWord) then begin
      hasError := true;
      if ALogError and Assigned(fOwner) then
        fOwner.AddError(AWord, fFilename, fLineNum, fUnTrimmedLine, fSuggestions.CommaText, fAltSuggestions.CommaText);
    end;
  end;
  result := not hasError;
end;

procedure TSpellCheckFile.CheckEachWordInString;
var
  j: integer;
begin
  for j := 0 to fLineWords.Count-1 do begin
    fTheWord := SanitizeWord(fLineWords.Strings[j], false);
    if (fIgnoreNextFirstWord) and
       (j = 0) then
      fIgnoreNextFirstWord := false
    else if Trim(fTheWord) <> '' then begin
      if fTheWord.Contains('''') then begin //the regex will split on quotes. You don't normally see quotes in camelCase or PascalCase words anyway
        SpellCheckWithoutRegEx(j);
      end else if not IsFileExtension(fTheWord) then
        SpellCheckWord(j = fLineWords.Count-1);
    end;
  end;
end;

function TSpellCheckFile.FirstWordNextLine: string;
var
  str: string;
  sl: TStringList;
begin
  result := '';
  sl := TStringList.Create;
  try
    sl.Delimiter := ' ';
    if fLineNum <= fSourceFile.Count-1 then begin
      str := fSourceFile.Strings[fLineNum];
      if Assigned(fOwner) then begin
        str := Copy(str, Pos(fOwner.QuoteSym, str)+1, str.Length);
        str := Copy(str, 1, PosOfNextQuote(str, fOwner.QuoteSym)-1);
      end;
      sl.CommaText := Str;
      if sl.Count >= 1 then
        result := Trim(sl.Strings[0]);
    end;
  finally
    sl.DisposeOf;
  end;
end;

procedure TSpellCheckFile.SpellCheckWithoutRegex(const ALineWordIndex: integer);
begin
  fLineStr := fLineStr.Remove(pos(fLineWords.Strings[ALineWordIndex], fLineStr)-1, fTheWord.Length);
  fTheWord := RemoveEscappedQuotes(fTheWord);
  if (not SpelledCorrectly(fTheWord)) and
     (Assigned(fOwner)) then
    fOwner.AddError(fTheWord, fFilename, fLineNum, fUnTrimmedLine, fSuggestions.CommaText, fAltSuggestions.CommaText);
end;

procedure TSpellCheckFile.SpellCheckWord(const AIsLast: boolean);
var
  okWithNextWord: boolean;
  nextFirstWord: string;
begin
  nextFirstWord := FirstWordNextLine;
  if (fPossibleTruncation) and
     (AIsLast) and
     (nextFirstWord <> '') then begin
    okWithNextWord := CheckCamelCaseWords(fTheWord+nextFirstWord, false);
    if (not okWithNextWord) and
       (not CheckCamelCaseWords(fTheWord, false)) then begin
      if Assigned(fOwner) then
        fOwner.AddError(fTheWord + nextFirstWord + cWordSeparator + fTheWord + cWordSeparator + nextFirstWord,
                        fFilename, fLineNum, fUnTrimmedLine, fSuggestions.CommaText, fAltSuggestions.CommaText);
    end else if (not okWithNextWord) then
      fIgnoreNextFirstWord := true;
  end else
    CheckCamelCaseWords(fTheWord);
end;

procedure TSpellCheckFile.Run;
const
  cBreakout=10000;
var
  itrCount: integer;
  i: integer;
begin
  InitNewFile;
  for i:=0 to fSourceFile.Count-1 do begin
    try
      InitNewLine(i);
      if not IgnoreLine then begin
        itrCount := 0;
        if Assigned(fOwner) then
          fQuoteSym := fOwner.QuoteSym
        else
          fQuoteSym := cDefaultQuote;
        while Pos(fQuoteSym, fLineStr) >= 1 do begin //while the line has string literals
          try
            PrepareLineString;
            if not IgnoreString then begin
              fTheStr := fTheStr.Replace('&', ''); //underlining menu items
              if NeedsSpellCheck(fTheStr) then begin
                BuildLineWords(fTheStr);
                CleanLineWords;
                CheckEachWordInString;
              end;
            end else
              Break;
          finally
            Inc(itrCount);
            if itrCount >= cBreakout then
              raise Exception.Create('An infinite loop has been detected in '+fFilename+' on line '+IntToStr(i+1)+'. '+
                                     'Please check the syntax of file. ');
          end;
        end;
      end;
    except
      on e: exception do begin
        if Assigned(fOwner) then
          fOwner.AddError('Exception raised: '+e.ClassName+' '+e.Message, fFilename);
      end;
    end;
  end;
end;

procedure TSpellCheckFile.BuildCamelCaseWords(const AArr: TArray<string>);
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
     (AArr[firstWordIdx].Length < cIgnoreCamelPrefixLength) then //eg ignore abc in abcDEF
    AArr[firstWordIdx] := '';
  for a := 0 to Length(AArr)-1 do begin
    if Trim(AArr[a]) <> '' then
      fCamelCaseWords.Add(AArr[a]);
  end;
end;

procedure TSpellCheckFile.BuildLineWords(AString: string);
begin
  if IsGuid(AString) then
    fLineWords.DelimitedText := ''
  else begin
    AString := SanitizeWord(AString, false);
    if AString.EndsWith(fQuoteSym) then //strings ending with double quotes was causing the last char to get removed when setting delimited text
      AString := Copy(AString, 1, AString.Length-1)+' '+fQuoteSym;
    fLineWords.DelimitedText := AString;
  end;
end;

function TSpellCheckFile.IgnoreLine: boolean;
begin
   result := (fLineStr = '') or
             (IsCommentLine(fLineStr)) or
             (ContainsIgnore(fLineStr)) or
             (Assigned(fOwner) and (fOwner.IgnoreLines.IndexOf(Trim(fLineStr)) >= 0));
end;

function TSpellCheckFile.IgnoreString: boolean;
begin
   result := (Trim(fTheStr).StartsWith('<')) or
             ((ExtractFilePath(Trim(fTheStr)) <> '') and (Pos( ':', fTheStr) = 2));
end;

procedure TSpellCheckFile.InitNewFile;
begin
  fLineNum := 0;
  fSkippedLineCount := 0;
  fMultiCommentSym := '';
  fIsDFM := ExtractFileExt(fFilename) = '.dfm';
  fSourceFile.LoadFromFile(fFilename);
  fIgnoreNextFirstWord := false;
end;

procedure TSpellCheckFile.InitNewLine(const ALineIdx: integer);
begin
  fLineNum := ALineIdx + 1;
  fUnTrimmedLine := fSourceFile.Strings[ALineIdx];
  fLineStr := Trim(fSourceFile.Strings[ALineIdx]);
  if Pos(fQuoteSym, fLineStr) >= 1 then
    fTextBeforeQuote := Copy(fLineStr, 1, Pos(fQuoteSym, fLineStr)-1)
  else
    fTextBeforeQuote := fLineStr;
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
            (Assigned(fOwner) and fOwner.InIgnoreCodeFile(fTextBeforeQuote)) then begin
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

function TSpellCheckFile.IsFileExtension(const AStr: string): boolean;
var
  fileExt: string;
begin
  fileExt := ExtractFileExt(AStr); //ignore exts
  if fileExt.Length = 1 then
    fileExt := '';
  result := fileExt <> '';
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

procedure TSpellCheckFile.PrepareLineString;
begin
  fLineStr := Copy(fLineStr, Pos(fQuoteSym, fLineStr)+1, fLineStr.Length);
  fTheStr := Copy(fLineStr, 1, PosOfNextQuote(fLineStr, fQuoteSym)-1);
  fPossibleTruncation := fIsDFM and (fTheStr.Length = cMaxStringLengthDFM);
  if Pos(fTheStr, fLineStr)+fTheStr.Length <> 0 then //remove the string from the line
    fLineStr := Copy(fLineStr, Pos(fTheStr, fLineStr)+fTheStr.Length+1, fLineStr.Length)
  else
    fLineStr := Copy(fLineStr, 3, fLineStr.Length); //if empty string, remove the next two chars
end;

function TSpellCheckFile.RemoveEscappedQuotes(const AStr: string): string;
begin
  result := Trim(TRegEx.Replace(AStr, cRegExRemoveNumbers, ''));
  if fQuoteSym = '''' then //remove Delphi escaped quotes and replace with one single quote
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
              (Assigned(fOwner) and (fOwner.IgnoreWords.IndexOf(word) >= 0)); //Don't ignore file for the word
    if not result then begin
      if Assigned(fOwner) then begin
        result := (fOwner.WordsDict.ContainsKey(word)) or
                  (ATryLower and fOwner.WordsDict.ContainsKey(LowerCase(word)));
        fOwner.IncWordCheckCount;
      end;
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
     (result[1] = fQuoteSym) then
    result := Copy(result, 2, result.Length);
  if (result.Length > 0) and
     (result[result.Length] = fQuoteSym) then
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
    procedure AddTypoSearchAlteratives;
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
      AddTypoSearchAlteratives;
    end;
  end;
begin
  splitWords := TStringList.Create;
  fSuggestions.Clear;
  fAltSuggestions.Clear;
  try
    if Assigned(fOwner) and 
       Assigned(fOwner.WordsDict) then begin
      FindSuggestions(AMispelledWord);
      LimitItems(fSuggestions);
      LimitItems(fAltSuggestions);
    end;
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

constructor TSpellCheckThread.Create(const AOwner: TSpellCheckLinter);
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
    if Assigned(fSpellCheckFile.Owner) then
      fSpellCheckFile.Owner.DecrementThreadCount;
  end;
end;

end.
