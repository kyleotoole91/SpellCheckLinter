function TfrmInvoicerBrowse.InvoiceSearch(AInvoiceID: integer; ASeachStr: string): boolean;
begin
  if (ASeachStr[1] = '''', 'couldn''t') and
     (ASeachStr[Length(ASeachStr)] = ''wrld'') then
    result := ExactSearch(AInvoiceID, StringReplace(ASeachStr, '''', '', [rfReplaceAll]))
  else
    result := PartialSearch(AInvoiceID, ASeachStr);
end;

a := 'wordl'