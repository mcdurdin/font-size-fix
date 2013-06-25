unit frmFontsizefix;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  TFontSize = (fs100, fs125);

  TForm1 = class(TForm)
    Button1: TButton;
    Memo1: TMemo;
    Button2: TButton;
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    function GetFontFolder: string;
    procedure Log(s: string);
    procedure SetFontSize(FFontSize: TFontSize);
    procedure LogRegistryChange(key, value, valuetype, old, new: string);
    function StringToHexData(s: string; var h: array of byte): Integer;
    function HexDataToString(h: array of byte; sz: Integer): string;
    procedure GetDPI(key: HKEY; path: string);
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

uses
  Registry,
  ShlObj;

{$R *.dfm}

type
  TRegistrySetting = record
    Name: string;
    ValueInteger: Integer;
    ValueString: string;
    ValueHex: string;
  end;

const
  //HKLM\SYSTEM\CurrentControlSet\Hardware Profiles\CurrentSoftware\Fonts\LogPixels
  //HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\FontDPI\LogPixels
  //to 96 (60 Hex)
  //[HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics]
  WindowMetricsRegistrySettings: array[TFontSize,0..21] of TRegistrySetting = (
   (
    (Name: 'AppliedDPI'; ValueInteger: $60),
    (Name: 'BorderWidth'; ValueString: '-15'),
    (Name: 'CaptionFont'; ValueHex:'f5,ff,ff,ff,00,00,00,00,00,00,00,00,00,00,00,00,bc,02,00,00,'+
      '00,00,00,01,00,00,00,00,54,00,61,00,68,00,6f,00,6d,00,61,00,00,00,00,00,00,'+
      '00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,'+
      '00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00'),
    (Name: 'CaptionHeight'; ValueString: '-270'),
    (Name: 'CaptionWidth'; ValueString: '-270'),
    (Name: 'IconFont'; ValueHex:'f5,ff,ff,ff,00,00,00,00,00,00,00,00,00,00,00,00,90,01,00,00,00,'+
      '00,00,01,00,00,00,00,54,00,61,00,68,00,6f,00,6d,00,61,00,00,00,45,6d,1a,f7,'+
      'd0,11,9e,a7,00,80,5f,71,47,72,00,00,00,00,4c,00,00,00,8b,00,00,00,8b,00,00,'+
      '00,01,00,00,00,f5,ff,ff,ff,00,00,00,00,00,00,00,00,00,00,00,00'),
    (Name: 'IconSpacing'; ValueString: '-1125'),
    (Name: 'IconTitleWrap'; ValueString: '1'),
    (Name: 'IconVerticalspacing'; ValueString: '-1125'),
    (Name: 'MenuFont'; ValueHex:'f5,ff,ff,ff,00,00,00,00,00,00,00,00,00,00,00,00,90,01,00,00,00,'+
      '00,00,01,00,00,00,00,54,00,61,00,68,00,6f,00,6d,00,61,00,00,00,00,00,00,00,'+
      '00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,'+
      '00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00'),
    (Name: 'MenuHeight'; ValueString: '-270'),
    (Name: 'MenuWidth'; ValueString: '-270'),
    (Name: 'MessageFont'; ValueHex:'f5,ff,ff,ff,00,00,00,00,00,00,00,00,00,00,00,00,90,01,00,00,'+
      '00,00,00,01,00,00,00,00,54,00,61,00,68,00,6f,00,6d,00,61,00,00,00,00,00,00,'+
      '00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,'+
      '00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00'),
    (Name: 'ScrollHeight'; ValueString: '-240'),
    (Name: 'ScrollWidth'; ValueString: '-240'),
    (Name: 'Shell Icon BPP'; ValueString: '16'),
    (Name: 'SmCaptionFont'; ValueHex:'f5,ff,ff,ff,00,00,00,00,00,00,00,00,00,00,00,00,bc,02,00,'+
      '00,00,00,00,01,00,00,00,00,54,00,61,00,68,00,6f,00,6d,00,61,00,00,00,00,00,'+
      '00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,'+
      '00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00'),
    (Name: 'SmCaptionHeight'; ValueString: '-225'),
    (Name: 'SmCaptionWidth'; ValueString: '-180'),
    (Name: 'StatusFont'; ValueHex:'f5,ff,ff,ff,00,00,00,00,00,00,00,00,00,00,00,00,90,01,00,00,'+
      '00,00,00,01,00,00,00,00,54,00,61,00,68,00,6f,00,6d,00,61,00,00,00,00,00,00,'+
      '00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,'+
      '00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00'),
    (Name: 'Shell Icon Size'; ValueString: '32'),
    (Name: 'MinAnimate'; ValueString: '0')
   ),
   (
    (Name: 'AppliedDPI'; ValueInteger: $78),
    (Name: 'BorderWidth'; ValueString: '-15'),
    (Name: 'CaptionFont'; ValueHex:'f2,ff,ff,ff,00,00,00,00,00,00,00,00,00,00,00,00,bc,02,00,00,'+
      '00,00,00,01,00,00,00,00,54,00,72,00,65,00,62,00,75,00,63,00,68,00,65,00,74,'+
      '00,20,00,4d,00,53,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,'+
      '00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00'),
    (Name: 'CaptionHeight'; ValueString: '-375'),
    (Name: 'CaptionWidth'; ValueString: '-270'),
    (Name: 'IconFont'; ValueHex:'f2,ff,ff,ff,00,00,00,00,00,00,00,00,00,00,00,00,90,01,00,00,00,'+
      '00,00,01,00,00,00,00,54,00,61,00,68,00,6f,00,6d,00,61,00,00,00,00,00,00,00,'+
      '00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,'+
      '00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00'),
    (Name: 'IconSpacing'; ValueString: '-1125'),
    (Name: 'IconTitleWrap'; ValueString: '1'),
    (Name: 'IconVerticalspacing'; ValueString: '-1125'),
    (Name: 'MenuFont'; ValueHex:'f2,ff,ff,ff,00,00,00,00,00,00,00,00,00,00,00,00,90,01,00,00,00,'+
      '00,00,01,00,00,00,00,54,00,61,00,68,00,6f,00,6d,00,61,00,00,00,00,00,00,00,'+
      '00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,'+
      '00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00'),
    (Name: 'MenuHeight'; ValueString: '-285'),
    (Name: 'MenuWidth'; ValueString: '-270'),
    (Name: 'MessageFont'; ValueHex:'f2,ff,ff,ff,00,00,00,00,00,00,00,00,00,00,00,00,90,01,00,00,'+
      '00,00,00,01,00,00,00,00,54,00,61,00,68,00,6f,00,6d,00,61,00,00,00,00,00,00,'+
      '00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,'+
      '00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00'),
    (Name: 'ScrollHeight'; ValueString: '-255'),
    (Name: 'ScrollWidth'; ValueString: '-255'),
    (Name: 'Shell Icon BPP'; ValueString: '16'),
    (Name: 'SmCaptionFont'; ValueHex:'f2,ff,ff,ff,00,00,00,00,00,00,00,00,00,00,00,00,bc,02,00,'+
      '00,00,00,00,01,00,00,00,00,54,00,61,00,68,00,6f,00,6d,00,61,00,00,00,00,00,'+
      '00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,'+
      '00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00'),
    (Name: 'SmCaptionHeight'; ValueString: '-255'),
    (Name: 'SmCaptionWidth'; ValueString: '-255'),
    (Name: 'StatusFont'; ValueHex:'f2,ff,ff,ff,00,00,00,00,00,00,00,00,00,00,00,00,90,01,00,00,'+
      '00,00,00,01,00,00,00,00,54,00,61,00,68,00,6f,00,6d,00,61,00,00,00,00,00,00,'+
      '00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,'+
      '00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00'),
    (Name: 'Shell Icon Size'; ValueString: '32'),
    (Name: 'MinAnimate'; ValueString: '1')
   )
  );


function TForm1.GetFontFolder: string;
var
  buf: array[0..260] of char;
begin
  if SHGetFolderPath(Handle, CSIDL_FONTS, 0, SHGFP_TYPE_CURRENT, buf) = S_OK
    then Result := IncludeTrailingPathDelimiter(buf)
    else Result := '';
end;

procedure TForm1.Log(s: string);
begin
  memo1.Lines.Add(s);
end;

procedure TForm1.LogRegistryChange(key, value, valuetype, old, new: string);
begin
  Log(Format('Updating registry %s:%s [%s] from "%s" to "%s"', [key,value,valuetype,old,new]));
end;

procedure TForm1.Button1Click(Sender: TObject);
begin
  SetFontSize(fs100);
end;


function TForm1.StringToHexData(s: string; var h: array of byte): Integer;
var
  i: Integer;
begin
  s := s + ',';
  for i := 0 to Length(s) div 3 - 1 do
    h[i] := StrToInt('$'+Copy(s,i*3+1,2));
  Result := Length(s) div 3;
end;

function TForm1.HexDataToString(h: array of byte; sz: Integer): string;
var
  i: Integer;
begin
  Result := '';
  for i := 0 to sz - 1 do
  begin
    Result := Result + IntToHex(h[i], 2);
    if i < sz - 1 then Result := Result + ',';
  end;
end;

procedure TForm1.SetFontSize(FFontSize: TFontSize);
const
  //oldfontnames: array[0..2] of string = ('sseriff.fon', 'seriff.fon', 'courf.fon');
  //newfontnames: array[0..2] of string = ('sserife.fon', 'serife.fon', 'coure.fon');
  FontSizeName: array[TFontSize] of string = ('100', '125');
  FontSizeValue: array[TFontSize] of integer = (96, 120);
  BitmapFontNames: array[0..2] of string = ('MS Sans Serif 8,10,12,14,18,24', 'MS Serif 8,10,12,14,18,24', 'Courier 10,12,15');
  BitmapFileNames: array[TFontSize,0..24] of string =
    (('SSERIFE.FON','SSERIFEE.FON','SSERIFEG.FON','SSERIFER.FON','SSERIFET.FON','SSEE1255.FON','SSEE1256.FON','SSEE1257.FON','SSEE874.FON',
      'SERIFE.FON','SERIFEE.FON','SERIFEG.FON','SERIFER.FON','SERIFET.FON','SERE1255.FON','SERE1256.FON','SERE1257.FON',
      'COURE.FON','COUREE.FON','COUREG.FON','COURER.FON','COURET.FON','COUE1255.FON','COUE1256.FON','COUE1257.FON'),

     ('SSERIFF.FON','SSERIFFE.FON','SSERIFFG.FON','SSERIFFR.FON','SSERIFFT.FON','SSEF1255.FON','SSEF1256.FON','SSEF1257.FON','SSEF874.FON',
      'SERIFF.FON','SERIFFE.FON','SERIFFG.FON','SERIFFR.FON','SERIFFT.FON','SERF1255.FON','SERF1256.FON','SERF1257.FON',
      'COURF.FON','COURFE.FON','COURFG.FON','COURFR.FON','COURFT.FON','COUF1255.FON','COUF1256.FON','COUF1257.FON')
    );

    
    function FindBitmapFontIndex(Name: string): Integer;
    var
      j: Integer;
    begin
      for j := Low(BitmapFontNames) to High(BitmapFontNames) do
        if SameText(Copy(Name, 1, Length(BitmapFontNames[j])), BitmapFontNames[j]) then
        begin
          Result := j;
          Exit;
        end;
      Result := -1;
    end;

    function FindBitmapFileIndex(Name: string): Integer;
    var
      j: TFontSize;
      k: Integer;
    begin
      for j := Low(BitmapFileNames) to High(BitmapFileNames) do
        for k := Low(BitmapFileNames[j]) to High(BitmapFileNames[j]) do
          if SameText(Copy(Name, 1, Length(BitmapFileNames[j][k])), BitmapFileNames[j][k]) then
          begin
            Result := k;
            Exit;
          end;
      Result := -1;
    end;

var
  FLogPixels, i: Integer;
  v: Cardinal;
  n: TFontSize;
  FBitmapFileIndex, FBitmapFontIndex: Integer;
  FFontDir, FFontFileName: string;
  FFontNames: TStrings;
  rs: TRegistrySetting;
  HexData: array[0..255] of byte;
  HexDataSize: Integer;
begin
  Screen.Cursor := crHourglass;
  try
    FFontDir := GetFontFolder;
    if FFontDir = '' then
    begin
      Log('Could not locate font folder.');
      Exit;
    end;

    FFontNames := TStringList.Create;
    with TRegistry.Create do
    try
      RootKey := HKEY_LOCAL_MACHINE;
      if OpenKey('SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts', False) then
      begin
        GetValueNames(FFontNames);
        for i := 0 to FFontNames.Count - 1 do
        begin
          FBitmapFontIndex := FindBitmapFontIndex(FFontNames[i]);
          if FBitmapFontIndex < 0 then Continue;

          FFontFileName := ReadString(FFontNames[i]);
          FBitmapFileIndex := FindBitmapFileIndex(FFontFileName);
          if FBitmapFileIndex < 0 then Continue;

          // Here's a font to set
          Log('Switching font '+FFontNames[i]+' to '+FontSizeName[FFontSize]+'% version');
          for n := Low(BitmapFileNames) to High(BitmapFileNames) do
            if not RemoveFontResource(PChar(FFontDir + BitmapFileNames[n][FBitmapFileIndex])) then
              Log('Failed to remove font resource '+BitmapFileNames[n][FBitmapFileIndex]+' (this is probably not an issue): '+SysErrorMessage(GetLastError));

          LogRegistryChange(CurrentPath, FFontNames[i], 'REG_SZ', FFontFileName, BitmapFileNames[FFontSize][FBitmapFileIndex]);
          WriteString(FFontNames[i], BitmapFileNames[FFontSize][FBitmapFileIndex]);

          if AddFontResource(PChar(FFontDir + BitmapFileNames[FFontSize][FBitmapFileIndex])) = 0 then
            Log('Failed to add font resource '+BitmapFileNames[FFontSize][FBitmapFileIndex]+': '+SysErrorMessage(GetLastError));
        end;
      end;
    finally
      Free;
      FFontNames.Free;
    end;

    with TRegistry.Create do
    try
      RootKey := HKEY_LOCAL_MACHINE;
      if OpenKey('SYSTEM\CurrentControlSet\Hardware Profiles\CurrentSoftware\Fonts', True) then
      begin
        if ValueExists('LogPixels') then FLogPixels := ReadInteger('LogPixels') else FLogPixels := 0;
        LogRegistryChange(CurrentPath, 'LogPixels', 'REG_DWORD', IntToStr(FLogPixels), IntToStr(FontSizeValue[FFontSize]));
        WriteInteger('LogPixels', FontSizeValue[FFontSize]);
      end;

      if OpenKey('\SOFTWARE\Microsoft\Windows NT\CurrentVersion\FontDPI\LogPixels', True) then
      begin
        if ValueExists('LogPixels') then FLogPixels := ReadInteger('LogPixels') else FLogPixels := 0;
        LogRegistryChange(CurrentPath, 'LogPixels', 'REG_DWORD', IntToStr(FLogPixels), IntToStr(FontSizeValue[FFontSize]));
        WriteInteger('LogPixels', FontSizeValue[FFontSize]);
      end;
    finally
      Free;
    end;

    with TRegistry.Create do
    try
      if OpenKey('\Control Panel\Desktop', True) then
      begin
        if ValueExists('LogPixels') then FLogPixels := ReadInteger('LogPixels') else FLogPixels := 0;
        LogRegistryChange(CurrentPath, 'LogPixels', 'REG_DWORD', IntToStr(FLogPixels), IntToStr(FontSizeValue[FFontSize]));
        WriteInteger('LogPixels', FontSizeValue[FFontSize]);
      end;

      if OpenKey('\Control Panel\Desktop\WindowMetrics', True) then
      begin
        for i := Low(WindowMetricsRegistrySettings[FFontSize]) to High(WindowMetricsRegistrySettings[FFontSize]) do
        begin
          rs := WindowMetricsRegistrySettings[FFontSize][i];
          if rs.ValueHex <> '' then
          begin
            HexDataSize := SizeOf(HexData);
            HexDataSize := ReadBinaryData(rs.Name, HexData, HexDataSize);
            LogRegistryChange(CurrentPath, rs.Name, 'REG_BINARY', HexDataToString(HexData, HexDataSize), rs.ValueHex);
            HexDataSize := StringToHexData(rs.ValueHex, HexData);
            WriteBinaryData(rs.Name, HexData, HexDataSize);
          end
          else if rs.ValueString <> '' then
            WriteString(rs.Name, rs.ValueString)
          else
            WriteInteger(rs.Name, rs.ValueInteger);
        end;
      end;
    finally
      Free;
    end;

    Log('Sending font and setting change notifications');

    SendMessageTimeout(HWND_BROADCAST, WM_FONTCHANGE, 0, 0, SMTO_NORMAL, 5000, @v);
    SendMessageTimeout(HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0, SMTO_NORMAL, 5000, @v);

    Log('You will need to log off and log in again for font size changes to take effect.  Hopefully, you should not have to restart your computer.');

    Invalidate;
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TForm1.GetDPI(key: HKEY; path: string);
var
  FLogPixels: Integer;
  s: string;
begin
  with TRegistry.Create do
  try
    RootKey := key;
    if OpenKeyReadOnly(path) and ValueExists('LogPixels') then
    begin
      FLogPixels := ReadInteger('LogPixels');
      if key = HKEY_CURRENT_USER then s := 'HKCU' else s := 'HKLM';
      Log('Current font DPI setting for '+s+path+' is '+IntToStr(FLogPixels));
    end;
  finally
    Free;
  end;
end;
procedure TForm1.FormCreate(Sender: TObject);
begin
  Log('Font folder is: '+GetFontFolder);
  GetDPI(HKEY_CURRENT_USER,'\Control Panel\Desktop');
  GetDPI(HKEY_LOCAL_MACHINE,'\SYSTEM\CurrentControlSet\Hardware Profiles\CurrentSoftware\Fonts');
  GetDPI(HKEY_LOCAL_MACHINE,'\SOFTWARE\Microsoft\Windows NT\CurrentVersion\FontDPI\LogPixels');
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
  SetFontSize(fs125);
end;

end.


