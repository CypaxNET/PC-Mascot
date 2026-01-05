unit ListHighlighter;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Graphics, SynEditHighlighter,
  SynEditHighlighterFoldBase;
type


  TListHighlighter = class(TSynCustomFoldHighlighter)
  private
  protected
    // accesible for the other examples
    FTokenPos, FTokenEnd: Integer;
    FLineText: String;
    FCommentAttri: TSynHighlighterAttributes;
    FAngleAttri: TSynHighlighterAttributes;
    FSquareAttri: TSynHighlighterAttributes;
    FTextAttri: TSynHighlighterAttributes;
    procedure SetCommentAttri(AValue: TSynHighlighterAttributes);
    procedure SetTextAttri(AValue: TSynHighlighterAttributes);
    procedure SetAngleAttri(AValue: TSynHighlighterAttributes);
    procedure SetSquareAttri(AValue: TSynHighlighterAttributes); public
    procedure SetLine(const NewValue: String; LineNumber: Integer); override;
    procedure Next; override;
    function GetEol: Boolean; override;
    procedure GetTokenEx(out TokenStart: PChar; out TokenLength: Integer); override;
    function GetTokenAttribute: TSynHighlighterAttributes; override; public
    function GetToken: String; override;
    function GetTokenPos: Integer; override;
    function GetTokenKind: Integer; override;
    function GetDefaultAttribute(Index: Integer): TSynHighlighterAttributes; override;
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  published
    (* Define attributes, for the different highlights. *)
    property CommentAttri: TSynHighlighterAttributes read FCommentAttri write SetCommentAttri;
    property AngleBraketAttri: TSynHighlighterAttributes read FAngleAttri write SetAngleAttri;
    property SquareAttri: TSynHighlighterAttributes read FSquareAttri write SetSquareAttri;
    property TextAttri: TSynHighlighterAttributes read FTextAttri write SetTextAttri;
  end;

  (*   This is a COPY of SynEditHighlighter

       ONLY the base class is changed to add support for folding

       The new code follows below
  *)

  TListFoldHighlighter = class(TListHighlighter)
  protected
    FCurRange: Integer;
  public
    procedure Next; override;
    function GetTokenAttribute: TSynHighlighterAttributes; override; public
  end;

  { TSynDemoHlContext }

  //TSynDemoHlFold = class(TListHighlighter)
  TSynDemoHlFold = class(TListFoldHighlighter)
  public
    procedure Next; override; public
    procedure SetRange(Value: Pointer); override;
    procedure ResetRange; override;
    function GetRange: Pointer; override;
  end;

implementation

{ TSynDemoHlFold }

procedure TSynDemoHlFold.Next;
begin
  inherited Next;
end;

procedure TSynDemoHlFold.SetRange(Value: Pointer);
begin
  // must call the SetRange in TSynCustomFoldHighlighter
  inherited SetRange(Value);
end;

procedure TSynDemoHlFold.ResetRange;
begin
  inherited ResetRange;
end;

function TSynDemoHlFold.GetRange: Pointer;
begin
  Result := inherited GetRange;
end;


destructor TListHighlighter.Destroy;
begin
  inherited;
end;

constructor TListHighlighter.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);

  (* Create and initialize the attributes *)
  FCommentAttri := TSynHighlighterAttributes.Create('comment', 'comment');
  AddAttribute(FCommentAttri);
  FCommentAttri.Foreground := clGray;

  FAngleAttri := TSynHighlighterAttributes.Create('angle', 'angle');
  AddAttribute(FAngleAttri);
  FAngleAttri.Foreground := clGreen;
  FAngleAttri.Style := [fsBold];

  FSquareAttri := TSynHighlighterAttributes.Create('square', 'square');
  AddAttribute(FSquareAttri);
  FSquareAttri.Foreground := clBlue;
  FSquareAttri.Style := [fsBold];

  FTextAttri := TSynHighlighterAttributes.Create('text', 'text');
  AddAttribute(FTextAttri);

  // Ensure the HL reacts to changes in the attributes. Do this once, if all attributes are created
  SetAttributesOnChange(@DefHighlightChange);
end;

(* Setters for attributes / This allows using in Object inspector*)
procedure TListHighlighter.SetCommentAttri(AValue: TSynHighlighterAttributes);
begin
  FCommentAttri.Assign(AValue);
end;

procedure TListHighlighter.SetTextAttri(AValue: TSynHighlighterAttributes);
begin
  FTextAttri.Assign(AValue);
end;

procedure TListHighlighter.SetAngleAttri(AValue: TSynHighlighterAttributes);
begin
  FAngleAttri.Assign(AValue);
end;

procedure TListHighlighter.SetSquareAttri(AValue: TSynHighlighterAttributes);
begin
  FSquareAttri.Assign(AValue);
end;

procedure TListHighlighter.SetLine(const NewValue: String; LineNumber: Integer);
begin
  inherited;
  FLineText := NewValue;
  // Next will start at "FTokenEnd", so set this to 1
  FTokenEnd := 1;
  Next;
end;

procedure TListHighlighter.Next;
var
  l: Integer;
begin

  // FTokenEnd should be at the start of the next Token (which is the Token we want)
  FTokenPos := FTokenEnd;
  // assume empty, will only happen for EOL
  FTokenEnd := FTokenPos;

  // Scan forward
  // FTokenEnd will be set 1 after the last char. That is:
  // - The first char of the next token
  // - or past the end of line (which allows GetEOL to work)

  l := length(FLineText);
  if FTokenPos > l then
    // At line end
    Exit
  else begin
    // Wenn [@1 dann gehe bis ] oder Zeilenende
    // Wenn < dann gehe bis > oder Zeilenende
    // sonst gehe bis < oder Zeilenende
    if (FLineText[FTokenEnd] = '[') and (FTokenEnd = 1) then begin
      // At start of braket? Then find end of braket
      repeat
        if (FTokenEnd <= l) then
          Inc(FTokenEnd);
      until (FTokenEnd > l) or (FLineText[FTokenEnd] = ']');
      if (FLineText[FTokenEnd] = ']') and (FTokenEnd <= l) then
        Inc(FTokenEnd);
    end
    else if (FLineText[FTokenEnd] = '<') then begin
      // At start of braket? Then find end of braket
      repeat
        if (FTokenEnd <= l) then
          Inc(FTokenEnd);
      until (FTokenEnd > l) or (FLineText[FTokenEnd] = '>');
      if (FLineText[FTokenEnd] = '>') and (FTokenEnd <= l) then
        Inc(FTokenEnd);
    end
    else begin
      // At start of not-a-braket? Then find a braket start
      repeat
        if (FTokenEnd <= l) then
          Inc(FTokenEnd);
      until (FTokenEnd > l) or (FLineText[FTokenEnd] = '<');
    end;
  end;


end;

function TListHighlighter.GetEol: Boolean;
begin
  Result := FTokenPos > length(FLineText);
end;

procedure TListHighlighter.GetTokenEx(out TokenStart: PChar; out TokenLength: Integer);
begin
  TokenStart := @FLineText[FTokenPos];
  TokenLength := FTokenEnd - FTokenPos;
end;

function TListHighlighter.GetTokenAttribute: TSynHighlighterAttributes;
var
  TokenWord: String;
begin

  // Match the text, specified by FTokenPos and FTokenEnd
  TokenWord := Copy(FLineText, FTokenPos, FTokenEnd - FTokenPos);

  Result := TextAttri; // default result

  if Trim(FLineText).StartsWith(';') then
    Result := CommentAttri
  else
  if (TokenWord.StartsWith('<')) and (TokenWord.EndsWith('>')) then begin
    Result := AngleBraketAttri;
  end
  else if (TokenWord.StartsWith('[')) and (TokenWord.EndsWith(']')) then begin
    Result := SquareAttri;
  end;

end;

function TListHighlighter.GetToken: String;
begin
  Result := Copy(FLineText, FTokenPos, FTokenEnd - FTokenPos);
end;

function TListHighlighter.GetTokenPos: Integer;
begin
  Result := FTokenPos - 1;
end;

function TListHighlighter.GetDefaultAttribute(Index: Integer): TSynHighlighterAttributes;
begin
  Result := FTextAttri;
end;

function TListHighlighter.GetTokenKind: Integer;
var
  a: TSynHighlighterAttributes;
begin
  // Map Attribute into a unique number
  a := GetTokenAttribute;
  Result := 0;
  if a = FAngleAttri then Result := 1;
  if a = FTextAttri then Result := 2;
  if a = FCommentAttri then Result := 3;
  if a = FSquareAttri then Result := 4;
end;

procedure TListFoldHighlighter.Next;
begin
  inherited Next;
end;

function TListFoldHighlighter.GetTokenAttribute: TSynHighlighterAttributes;
begin
  Result := inherited GetTokenAttribute;
end;


end.
