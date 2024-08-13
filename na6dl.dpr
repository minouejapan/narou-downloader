(*
  小説家になろう小説ダウンローダー

  3.0 2024/08/13  文字列が空かどうかのチェックが抜けている箇所がありメモリアクセス違反が発生することがあった不具合を修正した
  2.8 2024/06/28  Githubにあげるためソースコードを整理した
                  小説家になろうサーバーへの負荷を減らすためDL1話毎のインターバルを0.2→0.4秒にした
  2.7 2024/06/14  本文中の<>を余計なHTMLタグとして除去してしまう不具合を修正した
                  &#????;にエンコードされた文字のデコード処理をkakuyomudlと同じにした
  2.6 2024/05/14  各話見出しへの青空文庫タグ挿入時に改行コードが入る場合があった不具合を修正した
  2.5 2024/05/06  ルビ装飾タグの代替え表示用カッコに《》が指定されている場合青空文庫タグへの変換が誤作動する不具合を修正した
  2.4 2024/04/30  ルビ装飾用のタグが変更されて正しく青空文庫形式ルビタグに変換されなくなっていた不具合を修正した
  2.3 2024/03/27  作品情報の「最新部分掲載日」が「最新掲載日」に変更されていて進捗情報の連載中と
                  中断を判定できなかった不具合を修正した
                  na6dl.exeと同じフォルダ内に中断を判定するための日数を記したna6dl.cfgファイルが
                  あればその日数を基に連載中か中断かを判定するようにした
  2.2 2024/03/10  連載状況取得処理をIndyライブラリに変更しておらず情報を取得出来ていなかった不具合を修正した
  2.1 2024/03/09  R18系のDLに対応するためIdHTTPにクッキーを保存する処理を追加した
  2.0 2024/03/08  ダウンロード出来なくなった問題に対応するためHTML取得をWinINETからIndyライブラリに
                  変更した。但し、現時点R18系のDLは出来ない
  1.9 2024/01/24  トップページの目次が100話単位で複数ページにまたがるようになった仕様変更に対応した
  1.8 2023/11/26  作品の進捗状況を取得出来なくなっていた不具合を修正した
  1.7 2023/09/02  DL開始ページ指定に対応した際の副作用で短編がダウンロード出来なくなっていた不具合を修正した
  1.6 2023/08/29  タイトルと作者名の《》を青空文庫タグにエンコードする処理を追加した
                  章・話タイトルへの》青空文庫タグエンコード、HTML特殊文字デコード処理を追加した
  1.5 2023/08/23  タイトル名に含まれるエンコード記号をデコードしていなかった不具合を修正した
  1.4 2023/08/19  ルビ指定HTMLタグに《》が使われている場合があり青空文庫タグへの変換が誤作動していた
                  不具合を修正した
  1.3 2023/08/07  DL開始ページを指定した場合のNaro2mobiに送るDLページ数が1少なかった不具合を修正した
                  短編の場合Naro2mobiに作品情報を送信していなかった不具合を修正した
  1.2 2023/07/23  オプション引数確認処理を変更し、DL開始ページ指定オプション-sを追加した
  1.1 2023/07/10  章・話タイトルに半角スペースが含まれているとタイトルが途切れる不具合を修正した
  1.0 2023/06/28  Naro2mobi内臓のダウンロード処理部分を単独のダウンローダーとして切り出した
*)
program na6dl;

{$APPTYPE CONSOLE}

{$R *.res}

{$R *.dres}

uses
  System.SysUtils,
  System.Classes,
  System.DateUtils,
  System.RegularExpressions,
  Windows,
  // Indy10が必要
  IdHTTP,
  IdCookieManager,
  IdSSLOpenSSL,     // openssl-1.0.2が必要
  IdGlobal,
  WinAPI.Messages,
  IdURI;

const
  // データ抽出用の識別タグ
  STITLEB  = '<h1 id="workTitle"><a href=';     // 小説表題
  STITLEE  = '</a>';
  SAUTHERB = '<span id="workAuthor-activityName"><a href=';   // 作者
  SAUTHERE = '</a>';
  SAUTHERG = '<i class="icon-official" title="Official"></i>';
  SHEADERB = 'js-work-introduction">'; // 前書き
  SHEADERE = '</p>';
  SCOVERB  = '<div id="coverImage"><img src="';
  SCOVERE  = '"';

  SSTRURLB = '<li class="widget-toc-episode">';  // 各話リンクURL
  SSTRURLM = '<a href="';
  SSTRURLE = '" ';
  SSTTLB   = '<span class="widget-toc-episode-titleLabel js-vertical-composition-item">';
  SSTTLE   = '</span>';

  SNEXTCT  = '<a href=\".*?\" class=\"novelview_pager-next\">次へ<\/a>';  // 目次の次ページ

  SCAPTB   = '<p class="chapterTitle level1 js-vertical-composition-item"><span>';
  SCAPTB2   = '<p class="chapterTitle level2 js-vertical-composition-item"><span>';
  SCAPTE   = '</span>';
  SEPISB   = '<p class="widget-episodeTitle js-vertical-composition-item">';
  SEPISE   = '</p>';
  SBODY1   = '<div id="novel_p" class="novel_view">';
  SBODY2   = '<div id="novel_honbun" class="novel_view">';
  SBODY3   = '<div id="novel_a" class="novel_view">';
  SBODYB   = '<p id="p1"';
  SBODYM   = '>';
  SBODYE   = '</div>';

  SERRSTR  = '<div class="dots-indicator" id="LoadingEpisode">';
  SPICTB   = '<img src="';
  SPICTM   = '';
  SPICTE   = '" /></figure>';
  SHREFB   = '<a href="';
  SHREFE   = '">';
  SURLB    = 'https://';

  SHEAD    = '<h3 class="heading-level2">目次</h3>';

  // 青空文庫形式
  AO_RBI = '｜';							// ルビのかかり始め(必ずある訳ではない)
  AO_RBL = '《';              // ルビ始め
  AO_RBR = '》';              // ルビ終わり
  AO_TGI = '［＃';            // 青空文庫書式設定開始
  AO_TGO = '］';              //        〃       終了
  AO_CPI = '［＃「';          // 見出しの開始
  AO_CPT = '」は大見出し］';	// 章
  AO_SEC = '」は中見出し］';  // 話
  AO_PRT = '」は小見出し］';

  AO_CPB = '［＃大見出し］';        // 2022/12/28 こちらのタグに変更
  AO_CPE = '［＃大見出し終わり］';
  AO_SEB = '［＃中見出し］';
  AO_SEE = '［＃中見出し終わり］';
  AO_PRB = '［＃小見出し］';
  AO_PRE = '［＃小見出し終わり］';

  AO_DAI = '［＃ここから';		// ブロックの字下げ開始
  AO_DAO = '［＃ここで字下げ終わり］';
  AO_DAN = '字下げ］';
  AO_PGB = '［＃改丁］';			// 改丁と会ページはページ送りなのか見開き分の
  AO_PB2 = '［＃改ページ］';	// 送りかの違いがあるがどちらもページ送りとする
  AO_SM1 = '」に傍点］';			// ルビ傍点
  AO_SM2 = '」に丸傍点］';		// ルビ傍点 どちらもsesami_dotで扱う
  AO_EMB = '［＃丸傍点］';        // 横転開始
  AO_EME = '［＃丸傍点終わり］';  // 傍点終わり
  AO_KKL = '［＃ここから罫囲み］' ;     // 本来は罫囲み範囲の指定だが、前書きや後書き等を
  AO_KKR = '［＃ここで罫囲み終わり］';  // 一段小さい文字で表記するために使用する
  AO_END = '底本：';          // ページフッダ開始（必ずあるとは限らない）
  AO_PIB = '［＃リンクの図（';          // 画像埋め込み
  AO_PIE = '）入る］';        // 画像埋め込み終わり
  AO_LIB = '［＃リンク（';          // 画像埋め込み
  AO_LIE = '）入る］';        // 画像埋め込み終わり
  AO_CVB = '［＃表紙の図（';  // 表紙画像指定
  AO_CVE = '）入る］';        // 終わり
  AO_RAB = '［＃右寄せ］';
  AO_RAE = '［＃右寄せ終わり］';
  AO_HR  = '［＃水平線］';    // 水平線<hr />

  CRLF   = #$0D#$0A;

// ユーザメッセージID
  WM_DLINFO  = WM_USER + 30;

var
  IdHTTP: TIdHTTP;
  IdSSL: TIdSSLIOHandlerSocketOpenSSL;
  Cookies: TIdCookieManager;
  URI: TIdURI;
  PageList,     // URL, Chapter, Section, Date
  TextPage,
  LogFile: TStringList;
  Capter, URL, Path, FileName, TextLine, PBody, StartPage: string;
  hWnd: THandle;
  CDS: TCopyDataStruct;
  isOver18: boolean;
  StartN,
  ConteP,
  LimitDay: integer;

// Indyを用いたHTMLファイルのダウンロード
function LoadHTMLbyIndy(URLadr: string): string;
var
  IdURL: string;
  rbuff: TMemoryStream;
  tbuff: TStringList;
  ret: Boolean;
label
  Terminate;
begin
	Result := '';
  ret := False;

  rbuff := TMemoryStream.Create;
  tbuff := TStringList.Create;
  try
    try
      IdHTTP.Head(URLadr);
    except
      //取得に失敗した場合、再度取得を試みる
      try
        IdHTTP.Head(URLadr);
      except
        ret := True;
      end;
    end;
    if IdHTTP.ResponseCode = 302 then
    begin
      //リダイレクト後のURLで再度Headメソッドを実行して情報取得
      IdHTTP.Head(IdHTTP.Response.Location);
    end;
    if not ret then
    begin
      IdURL := IdHTTP.URL.URI;
      try
        IdHTTP.Get(IdURL, rbuff);
      except
        //取得に失敗した場合、再度取得を試みる
        try
          IdHTTP.Get(IdURL, rbuff);
        except
          Result := '';
        end;
      end;
    end;
    IdHTTP.Disconnect;
    rbuff.Position := 0;
    tbuff.LoadFromStream(rbuff, TEncoding.UTF8);
    Result := tbuff.Text;
  finally
    tbuff.Free;
    rbuff.Free;
  end;
end;

// Delphi XE2ではPos関数に検索開始位置を指定出来ないための代替え
function PosN(SubStr, Str: string; StartPos: integer): integer;
var
  tmp: string;
  p: integer;
begin
  tmp := Copy(Str, StartPos, Length(Str));
  p := Pos(SubStr, tmp);
  if p > 0 then
    Result := p + StartPos - 1  // 1ベーススタートのため1を引く
  else
    Result := 0;
end;

// 全角空白も除去するTrim
function TrimJ(src: string): string;
var
  tmp: string;
begin
  Result := src;
  if Length(src) > 0 then
  begin
    tmp := Trim(src);
    if Length(tmp) > 0 then
      while tmp[1] = '　' do
        Delete(tmp, 1, 1);
    Result := tmp;
  end;
end;

// タイトル名をファイル名として使用出来るかどうかチェックし、使用不可文字が
// あれば修正する('-'に置き換える)
// フォルダ名の最後が'.'の場合、フォルダ作成時に"."が無視されてフォルダ名が
// 見つからないことになるため'.'も'-'で置き換える(2019/12/20)
function PathFilter(PassName: string): string;
var
	i, l: integer;
  path: string;
  tmp: AnsiString;
  ch: char;
begin
  // ファイル名を一旦ShiftJISに変換して再度Unicode化することでShiftJISで使用
  // 出来ない文字を除去する
  tmp := AnsiString(PassName);
	path := string(tmp);
  l :=  Length(path);
  for i := 1 to l do
  begin
  	ch := Char(path[i]);
    if Pos(ch, '\/;:*?"<>|. '+#$09) > 0 then
      path[i] := '-';
  end;
  Result := path;
end;

// id=行番号タグを除去する
function LTagFilter(SrcText: string): string;
var
  line: string;
begin
  line := SrcText;
  // <p id=タグの終端を除去して<br />に置き換える
  line := StringReplace(line, '</p>', '<br />', [rfReplaceAll]);
  // <p id=xxx>を削除する
  line := TRegEx.Replace(line, '<p id=.*?>', '');
  Result := line;
end;

// 本文の青空文庫ルビタグ文字を代替文字に変換する
function ChangeAozoraTag(Base: string): string;
var
  tmp: string;
begin
  // ルビの代替え文字に《》が使われていれば最初に変換しておく
  tmp := StringReplace(Base, '<rp>《</rp><rt>', '</rb><rp>(</rp><rt>',  [rfReplaceAll]);
  tmp := StringReplace(tmp,  '</rt><rp>》</rp></ruby>', '</rt><rp>)</rp></ruby>',  [rfReplaceAll]);

  tmp := StringReplace(tmp,  '《', '※［＃始め二重山括弧、1-1-52］',  [rfReplaceAll]);
  tmp := StringReplace(tmp,  '》', '※［＃終わり二重山括弧、1-1-53］',  [rfReplaceAll]);
  tmp := StringReplace(tmp,  '｜', '※［＃縦線、1-1-35］',   [rfReplaceAll]);
  Result := tmp;
end;

// HTML特殊文字の処理
// 1)エスケープ文字列 → 実際の文字
// 2)&#x????; → 通常の文字
function Restore2RealChar(Base: string): string;
var
  tmp, cd: string;
  w: integer;
  ch: Char;
  m: TMatch;
begin
  // エスケープされた文字
  tmp := StringReplace(Base, '&lt;',      '<',  [rfReplaceAll]);
  tmp := StringReplace(tmp,  '&gt;',      '>',  [rfReplaceAll]);
  tmp := StringReplace(tmp,  '&quot;',    '"',  [rfReplaceAll]);
  tmp := StringReplace(tmp,  '&nbsp;',    ' ',  [rfReplaceAll]);
  tmp := StringReplace(tmp,  '&yen;',     '\',  [rfReplaceAll]);
  tmp := StringReplace(tmp,  '&brvbar;',  '|',  [rfReplaceAll]);
  tmp := StringReplace(tmp,  '&copy;',    '©',  [rfReplaceAll]);
  tmp := StringReplace(tmp,  '&amp;',     '&',  [rfReplaceAll]);
  // &#????;にエンコードされた文字をデコードする(2023/3/19)
  // 正規表現による処理に変更した(2024/3/9)
  m := TRegEx.Match(tmp, '&#.*?;');
  while m.Index > 1 do
  begin
    Delete(tmp, m.Index, m.Length);
    cd := m.Value;
    Delete(cd, 1, 2);           // &#を削除する
    Delete(cd, Length(cd), 1);  // 最後の;を削除する
    if cd[1] = 'x' then         // 先頭が16進数を表すxであればDelphiの16進数接頭文字$に変更する
      cd[1] := '$';
    try
      w := StrToInt(cd);
      ch := Char(w);
    except
      ch := '？';
    end;
    Insert(ch, tmp, m.Index);
    m := TRegEx.Match(tmp, '&#.*?;');
  end;

  Result := tmp;
end;

// テキストファイル用フィルタ
function GetText(htmltext: string): string;
var
  line: string;
begin
	Result := '';
  line := htmltext;
  // HTMLエスケープシーケンスを実際の文字に戻す
  line := Restore2RealChar(line);
  // 青空文庫タグを変換する
  line := ChangeAozoraTag(line);
  // 青空文庫形式で保存する際のルビの変換
  line := StringReplace(line,  '<ruby><rb>',              AO_RBI, [rfReplaceAll]);
  line := TRegEx.Replace(line, '</rb><rp>.</rp><rt>',     AO_RBL);
  line := TRegEx.Replace(line, '</rt><rp>.</rp></ruby>',  AO_RBR);

  line := StringReplace(line,  '<ruby>',                  AO_RBI, [rfReplaceAll]);
  line := TRegEx.Replace(line, '<rp>.</rp><rt>',          AO_RBL);
  line := TRegEx.Replace(line, '</rt><rp>.</rp></ruby>',  AO_RBR);

  // 埋め込み画像を変換する
  line := TRegEx.Replace(line, '<a href=".*?"><img src="', AO_PIB);
  line := TRegEx.Replace(line, '" alt=".*?/></a>', AO_PIE);
  // 埋め込みリンクを変換する
  line := StringReplace(line, '<a href="',                AO_LIB, [rfReplaceAll]);
  line := StringReplace(line, '">挿絵</a>',               AO_LIE, [rfReplaceAll]);
  // ダウンロード出来なかった画像を変換する
  line := StringReplace(line, '">画像をDL出来ませんでした</a>', '(DL Error)' + AO_LIE, [rfReplaceAll]);
  // 装飾用HTMLタグの除去
  line := StringReplace(line, '<br />',                   '',     [rfReplaceAll]);
  // そのの他のタグを削除する
  //line := TRegEx.Replace(line, '<.*?>', '');
  // HTMLタグ以外も除去してしまうため先頭文字が半角英字の場合だけ削除するように変更(2024/6/14)
  line := TRegEx.Replace(line, '<[a-x]+.*?>', '');
  Result := line;
end;

// ****************************************************************************
//
// １話ページの構文を解析して本文を取り出す
//
// ****************************************************************************
function ParsePage(Line: string): string;
var
	ps, pe, nps, npe, nas, nae: integer;
  chapter, section, body,
  view, preamble, afterword, ptxt: string;
begin
	Result := '';
	chapter := ''; section := ''; body := ''; ptxt := '';
  view := ''; preamble := ''; afterword := '';
  // 文章
  nps := Pos(SBODY1, Line);
  ps  := Pos(SBODY2, Line);
  nas := Pos(SBODY3, Line);

  if nps > 0 then
  begin
	  preamble 	:= Copy(Line, nps + 37, Length(Line));
    npe 		 	:= Pos('</div>', preamble);
    preamble 	:= Copy(preamble, 1, npe - 1);
    // <p id="行番号">タグを除去する
    preamble := LTagFilter(preamble);
  end;
  if nas > 0 then
  begin
	  afterword := Copy(Line, nas + 37, Length(Line));
    nae 			:= Pos('</div>', afterword);
    afterword := Copy(afterword, 1, nae - 1);
    // <p id="行番号">タグを除去する
    afterword := LTagFilter(afterword);
  end;
  if ps > 0 then
  begin
    // 本文
    Delete(Line, 1, ps + 41);
    pe			:= Pos('</div', Line);
    body 		:= Copy(Line, 1, pe - 1);
    // <p id="行番号">タグを除去する
    body := LTagFilter(body);
  end else begin
    PBody := PBody + '本文を取得できませんでした.' + #13#10 + AO_PB2 + #13#10;
		Exit;
  end;
  if preamble <> '' then
 	  ptxt := ptxt + #13#10 + AO_HR + AO_KKL + GetText(preamble) + #13#10 + AO_KKR + AO_HR + #13#10;
  body := GetText(body);
  ptxt := ptxt + body;
  if afterword <> '' then
 	  ptxt := ptxt + #13#10 + AO_HR + AO_KKL + GetText(afterword) + #13#10 + AO_KKR + AO_HR + #13#10;
  PBody := PBody + ptxt + AO_PB2 + #13#10;
end;

// ****************************************************************************
//
// 小説情報にアクセスして小説が完結、連載中、中断のいづれかを取得する
//
// ****************************************************************************
function GetNovelStatus(NiURL: string): string;
var
  str, stat: string;
  ps, pe, dd: integer;
  upd, tdy: TDateTime;
  function FormatDate(DateStr: string): string;  // 2019年 09月22日 22時26分を2019/09/22に変える
  begin
    Result := Copy(DateStr, 1, 4) + '/' + Copy(DateStr, 7, 2) + '/' + Copy(DateStr, 10, 2);
  end;
begin
  Result := '';
  str := LoadHTMLbyIndy(NiURL);
  if Length(str) =0 then
    Exit;
  // 小説が短編かどうかチェックする
  if Pos('<span id="noveltype">短編</span>', str) > 0 then
  begin
    Result := '【短編】';     // 短編のシンボルテキストを返す
    StartN := 0;
    Exit;
  end;
  // 小説の状態が完結済かどうかをチェックする
  ps := Pos('<span id="noveltype_notend">完結済', str) + Pos('<span id="noveltype">完結済', str);
  if ps > 0 then
  begin
    Result := '【完結】';     // 完結済のシンボルテキストを返す
    Exit;
  end;
  // 完結済みでない場合は最新更新日時を確認して連載中なのか中断かを判断する
  //ps := Pos('<th>最新部分掲載日</th>', str);
  ps := Pos('<th>最新掲載日</th>', str);
  if ps > 0  then
  begin
    Delete(str, 1, ps + 15);
    ps    := pos('<td>', str);
    if ps > 0 then
    begin
      Delete(str, 1, ps + 3);
      pe    := pos('</td>', str);
      stat  := Copy(str, 1, pe - 1);  // 最新部分掲載日を取得する
      stat  := FormatDate(stat);
      upd   := StrToDateTime(stat);
      tdy   := Today;
      dd    := DaysBetween(tdy, upd); // 最新掲載日から現在までの経過日数を取得する
      if dd > LimitDay then
        Result := '【中断】' // 60日以上更新されていなければ中断のシンボルテキストを返す
      else
        Result := '【連載中】';   // 連載中のシンボルテキストを返す
    end;
  end;
end;

//
//  各話を取得する
//
procedure LoadEachPage;
var
  i, n, cnt, sc, st: integer;
  pinfo: TStringList;
  line: string;
  CSBI: TConsoleScreenBufferInfo;
  CCI: TConsoleCursorInfo;
  hCOutput: THandle;
begin
  pinfo := TStringList.Create;
  try
    hCOutput := GetStdHandle(STD_OUTPUT_HANDLE);
    GetConsoleScreenBufferInfo(hCOutput, CSBI);
    GetConsoleCursorInfo(hCOutput, CCI);
    cnt := PageList.Count;
    sc  := cnt - StartN;
    n   := 0;
    Write('各話を取得中 [  0/' + Format('%3d', [cnt]) + ']');
    CCI.bVisible := False;
    SetConsoleCursorInfo(hCoutput, CCI);
    if StartN > 0 then
      st := StartN - 1
    else
      st := 0;
    for i := st to PageList.Count - 1 do
    begin
      Inc(n);
      SetConsoleCursorPosition(hCOutput, CSBI.dwCursorPosition);
      Write('各話を取得中 [' + Format('%3d', [i + 1]) + '/' + Format('%3d', [cnt]) + '(' + Format('%d', [(n * 100) div sc]) + '%)]');
      if hWnd <> 0 then
        SendMessage(hWnd, WM_DLINFO, n, 1);
      pinfo.CommaText := PageList.Strings[i];
      line := LoadHTMLbyIndy(pinfo.Strings[0]);
      if line <> '' then
      begin
        if pinfo.Strings[1] <> '' then
          PBody := PBody + AO_CPB + ChangeAozoraTag(Restore2RealChar(TrimJ(pinfo.Strings[1]))) + AO_CPE + #13#10;
        if pinfo.Strings[2] <> '' then
          PBody := PBody + AO_SEB + ChangeAozoraTag(Restore2RealChar(TrimJ(pinfo.Strings[2]))) + AO_SEE + #13#10;
        if pinfo.Strings[3] <> '' then
          PBody := PBody + AO_RAB + AO_KKL + pinfo.Strings[3] + #13#10 + AO_KKR + AO_RAE + #13#10;
        ParsePage(line);
      end;
      // サーバー側に負担をかけないため0.4秒のインターバルを入れる
      Sleep(400);
    end;
  finally
    pinfo.Free;
  end;
  CCI.bVisible := True;
  SetConsoleCursorInfo(hCoutput, CCI);
  Writeln('');
end;

//
//  トップページの解析
//
function ParseChapter(Line: string): boolean;
var
	ps, pe, cs, ce, page: integer;
  str, sub, nstat: string;
  title, auth, authurl, synop, chapter, section, dtime, sendstr: string;
  conhdl: THandle;
begin
  Write('小説情報を取得中 ' + URL + ' ... ');
  Result := False;
  synop := '';
  page := 1;
	// タイトル
  ps := Pos('<title>', Line);
  pe := Pos('</title>', Line);
  if (ps > 0) and (pe >ps) then
  begin
  	ps 	 		  := ps + 7;
    title		  := TrimJ(Restore2Realchar(Copy(Line, ps, pe - ps)));
    Delete(Line, 1, pe + 8);
  end else
  	Exit;
  // 小説情報URL
  ps := Pos('<li><a href="', Line);
  pe := Pos('">作品情報</a></li>', Line);
  if (ps > 0) and (pe >ps) then
  begin
  	ps 					:= ps + 13;
    str 				:= Copy(Line, ps, pe - ps);
    Delete(Line, 1, pe + 15);
    nstat := GetNovelStatus(str);
    // すでにタイトルに進捗状況が付いていないことをチェックして（タイトル先頭に【完結】が付いて
    // いる場合があるため）、なければ進捗状況を追加する
    if (nstat = '【完結】') and (Pos('完結', title) > 0) then
      nstat := ''
    else if (nstat = '【短編】') and (Pos('短編', title) > 0) then
      nstat := '';
    // タイトル名に進捗状況を付加する
    title := nstat + title;
    // 保存するファイル名を準備する
    if FileName = '' then
    begin
      FileName := PathFilter(title);
      if StartPage <> '' then
        FileName := FileName + '(' + StartPage + ')';
      FileName := ExtractFilePath(ParamStr(0)) + Copy(FileName, 1, 32) + '.txt';
    end;
  end;
  // 作者
  //  パターン1: <div class="novel_writername">作者：<a href="https://mypage.syosetu.com/XXXXXX/">作者名</a></div>
  //  パターン2: <div class="novel_writername">作者：作者名</div>
  ps := Pos('作者：<a href="', Line);
  if ps > 1 then
  begin
    authurl := Copy(Line, ps + Length('作者：<a href="'), 40);
    ps := Pos('">', authurl);
    if ps > 1 then
      Delete(authurl, ps, 40);
  end else begin
    authurl := '';
  end;
  ps := Pos('="novel_writername"', Line);
  if ps > 0 then
  begin
  	Delete(Line, 1, ps + 19);
    ps := 1;
  end;
  pe := Pos('</div>', Line);
  if (ps > 0) and (pe > ps) then
  begin
  	auth := Restore2Realchar(Copy(Line, ps, pe - ps));
    ps := Pos(#$0D, auth);
    while ps > 0 do
    begin
      Delete(auth, ps, 1);
      ps := Pos(#$0D, auth);
    end;
    ps := Pos(#$0A, auth);
    while ps > 0 do
    begin
      Delete(auth, ps, 1);
      ps := Pos(#$0A, auth);
    end;
    ps   := Pos('作者：', auth);
    if ps > 0 then
    	Delete(auth, 1, ps + 2)
    else
    	Exit;
    pe   := Pos('</a>', auth);
    if pe > 0 then
    begin
    	Delete(auth, pe, 8);
      ps := Pos('>', auth);
      if ps > 0 then
      	Delete(auth, 1, ps)
      else
      	Exit;
    end;
    // タイトル、作者名を保存
    PBody := title + #13#10 + auth + #13#10 + AO_PB2 + #13#10;
    LogFile.Add('小説URL : ' + URL);
    LogFile.Add('タイトル:' + title);
    LogFile.Add('作者　　:' + auth);
    if authurl <> '' then
      LogFile.Add('作者URL : ' + authurl);
	end else
  	Exit;

  // 説明
  ps := Pos('="novel_ex">', Line);
  if ps > 0 then
  begin
  	Delete(Line, 1, ps + 11);
    ps := 1;
  end;
  pe := Pos('</div>', Line);
  if (ps > 0) and (pe > ps) then
  begin
    str := Copy(Line, ps, pe - ps);

    synop := GetText(str);
    // あらすじページを保存
    if synop <> '' then
    begin
      LogFile.Add('あらすじ');
      LogFile.Add(synop);
      LogFile.Add('');
      LogFile.Add(DateToStr(Today));
      if nstat <> '【短編】' then
        PBody := PBody + AO_KKL + URL + #13#10 + synop + #13#10 + AO_KKR + #13#10 + AO_PB2 + #13#10;
    end;

    ps := Pos('<br />', str);
    while ps > 0 do
    begin
      Delete(str, ps, 6);
      ps := Pos('<br />', str);
    end;
    Delete(Line, 1, pe);
  end;
  title := ChangeAozoraTag(title);
  auth := ChangeAozoraTag(auth);
  // 短編の場合
  if nstat = '【短編】' then
  begin
    Writeln('短編を取得中');
    PBody := PBody + AO_SEB + title + AO_SEE + #13#10;  // 短編の場合は表題をタイトル名にする
    ParsePage(Line);
  end else begin
    // 各話タイトルとリンク
    //   最初に余分な情報を削除する
    ps := Pos('<div class="index_box">', Line);
    if ps > 1 then
      Delete(Line, 1, ps + Length('<div class="index_box">'));

    while True do
    begin
  	  chapter := '';
  	  cs := Pos('="chapter_title">', Line);
  	  ps := Pos('="subtitle">', Line);
  	  if ps = 0 then
  		  ps := Pos('="period_subtitle">', Line);
  	  pe := Pos('</a>', Line);
      // Chapterをチェック
      if (cs > 0) and (ps > 0) and (cs < ps) then
      begin
        Delete(Line, 1, cs + 16);
        ce := Pos('</div>', Line);
        chapter := Copy(Line, 1, ce - 1);
        Delete(Line, 1, ce + 6);
  		  ps := Pos('="subtitle">', Line);
  		  if ps = 0 then
  			  ps := Pos('="period_subtitle">', Line);
  		  pe := Pos('</a>', Line);
      end;
      if (ps > 0) and (pe > ps) then
  	  begin
  		  Delete(Line, 1, ps);
    	  ps := Pos('<a href="', Line);
    	  pe := Pos('</a>', Line);
    	  if (ps > 0) and (pe > ps) then
    	  begin
      	  str := Copy(Line, ps + 9, pe - ps - 9);
      	  Delete(Line, 1, pe + 4);
      	  ps 	:= Pos('">', str);
      	  if ps > 0 then
      	  begin
        	  // Section
        	  sub := IntToStr(page);
            Inc(page);
        	  section := Copy(str, ps + 2, Length(str));
            // 作成日時
            str := ''; dtime := '';
            ps := Pos('="long_update">', Line);
            if ps > 0 then
            begin
          	  pe := Pos('</dt>', Line);
              if pe > 0 then
              begin
            	  str := Copy(Line, ps + 15, pe - ps - 15);
                pe := Pos('<', str);
                if pe > 0 then
                begin
                  dtime := Copy(str, 1, pe - 3);
                  Delete(str, 1, pe - 3);
                  ps := Pos('<span title="', str);
                  if ps > 0 then
                  begin
                	  Delete(str, 1, ps + 12);
                	  pe := Pos('>', str);
                    dtime := dtime + ' (' + Copy(str, 1, pe - 2) + ')';
                  end;
                end else
              	  dtime := StringReplace(str, #13#10, '', [rfReplaceAll]);
              end;
            end;
            PageList.Add(URL + sub + '/,"' + chapter + '","' + section + '",'{ + dtime});
      	  end;
    	  end;
      // 次の目次ページがあるかどうかチェックしてある場合は次ページの目次を取得する
  	  end else if TRegEx.Match(Line, SNEXTCT).Index > 1 then
      begin
        Line := LoadHTMLbyIndy(URL + '?p=' + IntToStr(ConteP));
        if Line <> '' then
        begin
          ps := Pos('<div class="index_box">', Line);
          if ps > 1 then
          begin
            Delete(Line, 1, ps + Length('<div class="index_box">'));
            Inc(ConteP);
          end else begin
            Writeln('目次の' + IntToStr(ConteP) +'ページを読み込めませんでした.');
            Break;
          end;
        end else begin
          Writeln('目次の' + IntToStr(ConteP) +'ページを読み込めませんでした.');
          Break;
        end;
      end else begin
    	  Break;
      end;
    end;
    Writeln(IntToStr(PageList.Count) + ' 話の情報を取得しました.');
  end;
  // リストが取得出来ないかDL開始ページ以降のページがなければエラーにする
  if (nstat <> '【短編】') and ((PageList.Count = 0) or ((PageList.Count - StartN) <= 0)) then
  begin
    if (PageList.Count - StartN) <= 0 then
      Writeln('DL開始ページがありません.');
    Exit;
  end;
  // Naro2mobiから呼び出された場合は進捗状況をSendする
  if hWnd <> 0 then
  begin
    conhdl := GetStdHandle(STD_OUTPUT_HANDLE);
    sendstr := title + ',' + auth;
    Cds.dwData := PageList.Count - StartN + 1;
    Cds.cbData := (Length(sendstr) + 1) * SizeOf(Char);
    Cds.lpData := Pointer(sendstr);
    SendMessage(hWnd, WM_COPYDATA, conhdl, LPARAM(Addr(Cds)));
  end;
  Result := True;
end;

function GetVersionInfo(const AFileName:string): string;
var
  InfoSize:DWORD;
  SFI:string;
  Buf,Trans,Value:Pointer;
begin
  Result := '';
  if AFileName = '' then Exit;
  InfoSize := GetFileVersionInfoSize(PChar(AFileName),InfoSize);
  if InfoSize <> 0 then
  begin
    GetMem(Buf,InfoSize);
    try
      if GetFileVersionInfo(PChar(AFileName),0,InfoSize,Buf) then
      begin
        if VerQueryValue(Buf,'\VarFileInfo\Translation',Trans,InfoSize) then
        begin
          SFI := Format('\StringFileInfo\%4.4x%4.4x\FileVersion',
                 [LOWORD(DWORD(Trans^)),HIWORD(DWORD(Trans^))]);
          if VerQueryValue(Buf,PChar(SFI),Value,InfoSize) then
            Result := PChar(Value)
          else Result := 'UnKnown';
        end;
      end;
    finally
      FreeMem(Buf);
    end;
  end;
end;

// OpenSSLが使用出来るかどうかチェックする
function CheckOpenSSL: Boolean;
var
  hnd: THandle;
begin
  Result := True;
  hnd := LoadLibrary('libeay32.dll');
  if hnd = 0 then
    Result := False
  else
    FreeLibrary(hnd);
  hnd := LoadLibrary('ssleay32.dll');
  if hnd = 0 then
    Result := False
  else
    FreeLibrary(hnd);
end;

var
  i: integer;
  op, df, dy: string;
  asource: TStringStream;
  t: TextFile;

begin
  // OpenSSLライブラリをチェック
  if not CheckOpenSSL then
  begin
    Writeln('');
    Writeln('na6dlを使用するためのOpenSSLライブラリが見つかりません.');
    Writeln('以下のサイトからopenssl-1.0.2q-i386-win32.zipをダウンロードしてlibeay32.dllとssleay32.dllをna6dl.exeがあるフォルダにコピーして下さい.');
    Writeln('https://github.com/IndySockets/OpenSSL-Binaries');
    ExitCode := 2;
    Exit;
  end;
  // OpenSSLのバージョンをチェック
  if (Pos('1.0.2', GetVersionInfo('libeay32.dll')) = 0)
    or (Pos('1.0.2', GetVersionInfo('ssleay32.dll')) = 0) then
  begin
    Writeln('');
    Writeln('OpenSSLライブラリのバージョンが違います.');
    Writeln('以下のサイトからopenssl-1.0.2q-i386-win32.zipをダウンロードしてlibeay32.dllとssleay32.dllをna6dl.exeがあるフォルダにコピーして下さい.');
    Writeln('https://github.com/IndySockets/OpenSSL-Binaries');
    ExitCode := 2;
    Exit;
  end;

  if ParamCount = 0 then
  begin
    Writeln('');
    Writeln('na6dl ver2.8 2024/6/27 (c) INOUE, masahiro.');
    Writeln('  使用方法');
    Writeln('  na6dl [-sDL開始ページ番号] 小説トップページのURL [保存するファイル名(省略するとタイトル名で保存します)]');
    Exit;
  end;

  ExitCode  := 0;
  hWnd      := 0;
  StartN    := 0; // 開始ページ番号(0スタート)
  ConteP    := 2; // 目次の2ページ目
  FileName  := '';
  StartPage := '';

  Path := ExtractFilePath(ParamStr(0));
  // オプション引数取得
  for i := 0 to ParamCount - 1 do
  begin
    op := ParamStr(i + 1);
    // Naro2mobiのWindowsハンドル
    if Pos('-h', op) = 1 then
    begin
      Delete(op, 1, 2);
      try
        hWnd := StrToInt(op);
      except
        Writeln('Error: Invalid Naro2mobi Handle.');
        ExitCode := -1;
        Exit;
      end;
    // DL開始ページ番号
    end else if Pos('-s', op) = 1 then
    begin
      Delete(op, 1, 2);
      StartPage := op;
      try
        StartN := StrToInt(op);
      except
        Writeln('Error: Invalid Start Page Number.');
        ExitCode := -1;
        Exit;
      end;
    // 作品URL
    end else if Pos('https:', op) = 1 then
    begin
      URL := op;
      if URL[Length(URL)] <> '/' then
        URL := URL + '/';
    // それ以外であれば保存ファイル名
    end else begin
      FileName := op;
      if UpperCase(ExtractFileExt(op)) <> '.TXT' then
        FileName := FileName + '.txt';
    end;
  end;

  if (Pos('https://ncode.syosetu.com/n', URL) = 0) and (Pos('https://novel18.syosetu.com/n', URL) = 0) then
  begin
    Writeln('小説のURLが違います.');
    ExitCode := -1;
    Exit;
  end;

  df := Path + 'na6dl.cfg';
  LimitDay := 60;
  if FileExists(df) then
  begin
    AssignFile(t, df);
    Reset(t);
    Readln(t, dy);
    try
      LimitDay := StrToInt(dy);
    except
      LimitDay := 60;
    end;
    CloseFile(t);
  end;

  isOver18 := Pos('https://novel18.syosetu.com/n', URL) > 0;

  IdHTTP := TIdHTTP.Create(nil);
  IdSSL := TIdSSLIOHandlerSocketOpenSSL.Create;
  Cookies := TIdCookieManager.Create(nil);
  IdSSL.IPVersion := Id_IPv4;
  IdSSL.MaxLineLength := 32768;
  IdSSL.SSLOptions.Method := sslvSSLv23;
  IdSSL.SSLOptions.SSLVersions := [sslvSSLv2,sslvTLSv1];
  IdHTTP.HandleRedirects := True;
  IdHTTP.AllowCookies := True;
  IdHTTP.IOHandler := IdSSL;
  // IdHTTPインスタンスにover18=yesのキャッシュを設定する
  if isOver18 then
  begin
    IdHTTP.CookieManager := TIdCookieManager.Create(IdHTTP);
    URI := TIdURI.Create('https://novel18.syosetu.com/');
    try
      IdHTTP.CookieManager.AddServerCookie('over18=yes', URI);
    finally
      URI.Free;
    end;
    asource := TStringStream.Create;
    try
      IdHTTP.Post('https://novel18.syosetu.com/', asource);
    finally
      asource.Free;
    end;
  end;
  Capter := '';
  TextLine := LoadHTMLbyIndy(URL);
  if TextLine <> '' then
  begin
    PageList := TStringList.Create;
    TextPage := TStringList.Create;
    LogFile  := TStringList.Create;
    try
      if ParseChapter(TextLine) then          // 小説の目次情報を取得
      begin
        // 短編でなければ各話を取得する
        if PageList.Count >= StartN then
          LoadEachPage;                       // 小説各話情報を取得
        try
          TextPage.Text := PBody;
          TextPage.SaveToFile(Filename, TEncoding.UTF8);
          LogFile.SaveToFile(ChangeFileExt(FileName, '.log'), TEncoding.UTF8);
          Writeln(ExtractFileName(Filename) + ' に保存しました.');
        except
          ExitCode := -1;
          Writeln('ファイルの保存に失敗しました.');
        end;
      end else begin
        Writeln(URL + 'から作品情報を取得できませんでした.');
        ExitCode := -1;
      end;
    finally
      LogFile.Free;
      PageList.Free;
      TextPage.Free;
    end;
  end else begin
    Writeln(URL + 'からHTMLソースを取得できませんでした.');
    ExitCode := -1;
  end;
    IdSSL.Free;
    IdHTTP.Free;
    Cookies.Free;
end.

