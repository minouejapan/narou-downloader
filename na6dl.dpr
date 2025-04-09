(*
  小説家になろう小説ダウンローダー

  4.1 2025/04/10  IndyのUser-AgentにMicrosoft Edgeを設定するようにした
  4.0 2025/03/05  作品情報ページのHTML構成が変更されたため連載状況取得部分を修正した
  3.91     03/03  ファイル名に使用できない文字から漏れていた"<>の処理を追加した
  3.9 2025/02/23  作者URLがない場合の作者名がうまく取得出来なかった不具合を修正した
  3.8 2025/02/13  短編の作者名がおかしかった不具合を修正した
  3.7 2025/02/08  日付変換でEConvertErrorが発生する場合がある不具合に対処した
                  &#????;にエンコードされた文字のデコード処理の不具合を修正した
  3.6 2025/01/30  作品情報が読み込めなくなったためファイルを保存出来なくなった不具合を修正した
  3.5 2024/12/10  作者URLが設定されていない場合、作者名をうまく取得出来なかった不具合を修正した
  3.4 2024/11/25  短編作品のタイトルに"短編"が入っているとダウンロード出来なかった不具合を修正した
  3.3 2024/11/24  -sオプションで最終エピソードを指定した場合ダウンロード出来なかった不具合を修正した
  3.2 2024/09/20  なろうのHTML構成が変更されてDL出来なくなったため修正した
  3.1 2024/09/04  ダウンロード出来ない作品があったため文字列が空かどうかのチェック方法を修正した
	3.0	2024/07/18	Lazarus/Delphiどちらでもビルド出来るようにした
                  文字列が空かどうかのチェックが抜けている箇所がありメモリアクセス違反が発生することがあった不具合を修正した
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

{$IFDEF FPC}
  {$MODE Delphi}
  {$codepage utf8}
{$ENDIF}

{$R *.res}
{$R 'verinfo.res' 'verinfo.rc'}
{$R *.dres}

uses
  LazUTF8wrap,
  System.SysUtils,
  System.Classes,
  System.DateUtils,
  WinAPI.Messages,
  Windows,
  regexpr,
  IdHTTP,
  IdCookieManager,
  IdSSLOpenSSL,
  IdGlobal,
  IdURI;

const
  // データ抽出用の識別タグ
  // 識別タグのほとんどはソース内に直接埋め込んだため削除した
  SNEXTCT  = '<a href=\".*?\" class=\"c-pager__item c-pager__item--next\">次へ<\/a>';  // 目次の次ページ
  SBODY1   = '<div class="js-novel-text p-novel__text p-novel__text--preface".*?>';
  SBODY2   = '<div class="js-novel-text p-novel__text".*?>';
  SBODY3   = '<div class="js-novel-text p-novel__text p-novel__text--afterword".*?>';
  SBODYB   = '<p id="p1"';
  SBODYM   = '>';
  SBODYE   = '</div>';


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
  DLErr: string;


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
      while UTF8Copy(tmp, 1, 1) = '　' do
        Delete(tmp, 1, 1);
    Result := tmp;
  end;
end;

// タイトル名にファイル名として使用出来ない文字を'-'に置換する
// Lazarus(FPC)とDelphiで文字コード変換方法が異なるためコンパイル環境で
// 変換処理を切り替える
function PathFilter(PassName: string): string;
var
  path: string;
  tmp: AnsiString;
begin
  // ファイル名を一旦ShiftJISに変換して再度Unicode化することでShiftJISで使用
  // 出来ない文字を除去する
{$IFDEF FPC}
  tmp  := UTF8ToWinCP(PassName);
  path := WinCPToUTF8(tmp);      // これでUTF-8依存文字は??に置き換わる
{$ELSE}
  tmp  := AnsiString(PassName);
	path := string(tmp);
{$ENDIF}
  // ファイル名として使用できない文字を'-'に置換する
  path := ReplaceRegExpr('[\\/:;\*\?\+,."<>|\.\t ]', path, '-');

  Result := path;
end;

// id=行番号タグを除去する
function LTagFilter(SrcText: string): string;
var
  line: string;
begin
  line := SrcText;
  // <p id=タグの終端を除去して<br />に置き換える
  line := UTF8StringReplace(line, '</p>', '<br />', [rfReplaceAll]);
  // <p id=xxx>を削除する
  line := ReplaceRegExpr('<p id=.*?>', line, '');
  Result := line;
end;

// 本文の青空文庫ルビタグ文字を代替文字に変換する
function ChangeAozoraTag(Base: string): string;
var
  tmp: string;
begin
  // ルビの代替え文字に《》が使われていれば最初に変換しておく
  tmp := UTF8StringReplace(Base, '<rp>《</rp><rt>', '</rb><rp>(</rp><rt>',  [rfReplaceAll]);
  tmp := UTF8StringReplace(tmp,  '</rt><rp>》</rp></ruby>', '</rt><rp>)</rp></ruby>',  [rfReplaceAll]);

  tmp := UTF8StringReplace(tmp,  '《', '※［＃始め二重山括弧、1-1-52］',  [rfReplaceAll]);
  tmp := UTF8StringReplace(tmp,  '》', '※［＃終わり二重山括弧、1-1-53］',  [rfReplaceAll]);
  tmp := UTF8StringReplace(tmp,  '｜', '※［＃縦線、1-1-35］',   [rfReplaceAll]);
  Result := tmp;
end;

// HTML特殊文字の処理
// 1)エスケープ文字列 → 実際の文字
// 2)&#x????; → 通常の文字
function Restore2RealChar(Base: string): string;
var
  tmp, cd: string;
  w, mp, ml: integer;
  ch: Char;
  r: TRegExpr;
begin
  // エスケープされた文字
  tmp := UTF8StringReplace(Base, '&lt;',      '<',  [rfReplaceAll]);
  tmp := UTF8StringReplace(tmp,  '&gt;',      '>',  [rfReplaceAll]);
  tmp := UTF8StringReplace(tmp,  '&quot;',    '"',  [rfReplaceAll]);
  tmp := UTF8StringReplace(tmp,  '&nbsp;',    ' ',  [rfReplaceAll]);
  tmp := UTF8StringReplace(tmp,  '&yen;',     '\',  [rfReplaceAll]);
  tmp := UTF8StringReplace(tmp,  '&brvbar;',  '|',  [rfReplaceAll]);
  tmp := UTF8StringReplace(tmp,  '&copy;',    '©',  [rfReplaceAll]);
  tmp := UTF8StringReplace(tmp,  '&amp;',     '&',  [rfReplaceAll]);
  // &#????;にエンコードされた文字をデコードする(2023/3/19)
  // 正規表現による処理に変更した(2024/3/9)
  r := TRegExpr.Create;
  try
    r.Expression  := '&#.*?;';
    r.InputString := tmp;
    if r.Exec then
    begin
      repeat
        cd := r.Match[0];
        mp := r.MatchPos[0];
        ml := r.MatchLen[0];
        UTF8Delete(tmp, mp, ml);
        UTF8Delete(cd, 1, 2);           // &#を削除する
        UTF8Delete(cd, UTF8Length(cd), 1);  // 最後の;を削除する
        if cd[1] = 'x' then         // 先頭が16進数を表すxであればDelphiの16進数接頭文字$に変更する
          cd[1] := '$';
        try
          w := StrToInt(cd);
          ch := Char(w);
        except
          ch := '?';
        end;
        UTF8Insert(ch, tmp, mp);
        r.InputString := tmp;
      until not r.Exec;
    end;
  finally
    r.Free;
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
  line := UTF8StringReplace(line,  '<ruby><rb>',              AO_RBI, [rfReplaceAll]);
  line := ReplaceRegExpr('</rb><rp>.</rp><rt>', line,     AO_RBL);
  line := ReplaceRegExpr('</rt><rp>.</rp></ruby>', line,  AO_RBR);

  line := UTF8StringReplace(line,  '<ruby>',                  AO_RBI, [rfReplaceAll]);
  line := ReplaceRegExpr('<rp>.</rp><rt>', line,          AO_RBL);
  line := ReplaceRegExpr('</rt><rp>.</rp></ruby>', line,  AO_RBR);

  // 埋め込み画像を変換する
  line := ReplaceRegExpr('<a href=".*?"><img src="', line, AO_PIB);
  line := ReplaceRegExpr('" alt=".*?/></a>', line, AO_PIE);
  // 埋め込みリンクを変換する
  line := UTF8StringReplace(line, '<a href="',                AO_LIB, [rfReplaceAll]);
  line := UTF8StringReplace(line, '">挿絵</a>',               AO_LIE, [rfReplaceAll]);
  // ダウンロード出来なかった画像を変換する
  line := UTF8StringReplace(line, '">画像をDL出来ませんでした</a>', '(DL Error)' + AO_LIE, [rfReplaceAll]);
  // 装飾用HTMLタグの除去
  line := UTF8StringReplace(line, '<br />',                   '',     [rfReplaceAll]);
  // そのの他のタグを削除する
  // HTMLタグ以外も除去してしまうため先頭文字が半角英字の場合だけ削除するように変更(2024/6/14)
  line := ReplaceRegExpr('<[a-x]+.*?>', line, '');
  Result := line;
end;

// ****************************************************************************
//
// １話ページの構文を解析して本文を取り出す
//
// ****************************************************************************
function ParsePage(Line: string): string;
var
	ps, pe, nps, npe, nas, nae, npslen, pslen, naslen: integer;
  body, preamble, afterword, ptxt: string;
  r: TRegExpr;
begin
	Result := '';
	body := '';
  ptxt := '';
  preamble := '';
  afterword := '';
  // 文章
  r := TRegExpr.Create;
  try
    r.InputString := Line;
    r.Expression := SBODY1;
    r.Exec;
    nps := r.MatchPos[0];
    npslen := r.MatchLen[0];
    r.Expression := SBODY2;
    r.Exec;
    ps := r.MatchPos[0];
    pslen := r.MatchLen[0];
    r.Expression := SBODY3;
    r.Exec;
    nas := r.MatchPos[0];
    naslen := r.MatchLen[0];

    if nps > 0 then
    begin
      preamble 	:= UTF8Copy(Line, nps + npslen, UTF8Length(Line));
      npe 		 	:= UTF8Pos('</div>', preamble);
      preamble 	:= UTF8Copy(preamble, 1, npe - 1);
      // <p id="行番号">タグを除去する
      preamble := LTagFilter(preamble);
    end;
    if nas > 0 then
    begin
      afterword := UTF8Copy(Line, nas + naslen, UTF8Length(Line));
      nae 			:= UTF8Pos('</div>', afterword);
      afterword := UTF8Copy(afterword, 1, nae - 1);
      // <p id="行番号">タグを除去する
      afterword := LTagFilter(afterword);
    end;
    if ps > 0 then
    begin
      // 本文
      UTF8Delete(Line, 1, ps + pslen);
      pe			:= UTF8Pos('</div', Line);
      body 		:= UTF8Copy(Line, 1, pe - 1);
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
  finally
    r.Free;
  end;
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
    Result := UTF8Copy(DateStr, 1, 4) + '/' + UTF8Copy(DateStr, 7, 2) + '/' + UTF8Copy(DateStr, 10, 2);
  end;
begin
  Result := '';
  str := LoadHTMLbyIndy(NiURL);
  if UTF8Length(str) =0 then
    Exit;
  if UTF8Pos('">短編</span>', str) > 0 then
  begin
    Result := '【短編】';     // 短編のシンボルテキストを返す
    StartN := 0;
    Exit;
  end;
  // 小説の状態が完結済かどうかをチェックする
  ps := UTF8Pos('">完結済</span>', str);
  if ps > 0 then
  begin
    Result := '【完結】';     // 完結済のシンボルテキストを返す
    Exit;
  end;
  // 完結済みでない場合は最新更新日時を確認して連載中なのか中断かを判断する
  ps := UTF8Pos('<dt class="p-infotop-data__title">最新掲載日</dt>', str);
  if ps > 0  then
  begin
    UTF8Delete(str, 1, ps + Length('<dt class="p-infotop-data__title">最新掲載日</dt>') - 1);
    ps    := UTF8Pos('<dd class="p-infotop-data__value">', str);
    if ps > 0 then
    begin
      UTF8Delete(str, 1, ps + Length('<dd class="p-infotop-data__value">') - 1);
      pe    := UTF8Pos('</dd>', str);
      stat  := UTF8Copy(str, 1, pe - 1);  // 最新部分掲載日を取得する
      stat  := FormatDate(stat);
      // EConvertError回避
      try
        upd   := StrToDateTime(stat);
      except
        upd := Yesterday;
      end;
      tdy   := Today;
      dd    := DaysBetween(tdy, upd); // 最新掲載日から現在までの経過日数を取得する
      if dd > LimitDay then
        Result := '【中断】'   // 60日以上更新されていなければ中断のシンボルテキストを返す
      else
        Result := '【連載中】';
    end;
  end;
end;

//
//  各話を取得する
//
procedure LoadEachPage;
var
  i, n, cnt, sc, st, cc: integer;
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
    cc  := 0;
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
      if StartN = 0 then
        Write('各話を取得中 [' + Format('%3d', [i + 1]) + '/' + Format('%3d', [cnt]) + '(' + Format('%d', [(n * 100) div sc]) + '%)]')
      else
        Write('各話を取得中 [' + Format('%3d', [i + 1]) + '/' + Format('%3d', [cnt]) +']');
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
        Inc(cc);
      end;
      // サーバー側に負担をかけないため0.4秒のインターバルを入れる
      Sleep(400);
    end;
  finally
    pinfo.Free;
  end;
  CCI.bVisible := True;
  SetConsoleCursorInfo(hCoutput, CCI);
  Writeln(' ... ' + IntToStr(n) + ' 個のエピソードを取得しました.');
  if cc < sc then
  begin
    Writeln('!!! ' + IntToStr(sc - cc) + ' 個の取得に失敗しいました.');
    DLErr := '［DL失敗］';  // ダウンロードが失敗した場合はファイル名の頭にマークを付ける
  end;
end;

//
//  トップページの解析
//
function ParseChapter(Line: string): boolean;
var
	ps, pe, cs, ce, page: integer;
  str, sub, nstat, sn, dt: string;
  title, auth, authurl, synop, chapter, section, sendstr: string;
{$IFDEF FPC}
  ws: WideString;
{$ENDIF}
  conhdl: THandle;
  r: TRegExpr;
begin
  Write('小説情報を取得中 ' + URL + ' ... ');
  Result := False;
  synop := '';
  page := 1;
	// タイトル
  ps := UTF8Pos('<title>', Line);
  pe := UTF8Pos('</title>', Line);
  if (ps > 0) and (pe >ps) then
  begin
  	ps 	 		  := ps + 7;
    title		  := TrimJ(Restore2Realchar(UTF8Copy(Line, ps, pe - ps)));
    UTF8Delete(Line, 1, pe + 8);
  end else
  	Exit;
  // 小説情報URL
  ps := UTF8Pos('<a class="c-menu__item c-menu__item--headnav" href="', Line);
  pe := UTF8Pos('">作品情報</a>', Line);
  if (ps > 0) and (pe >ps) then
  begin
  	ps 					:= ps + Utf8Length('<a class="c-menu__item c-menu__item--headnav" href="');
    str 				:= UTF8Copy(Line, ps, pe - ps);
    UTF8Delete(Line, 1, pe + 15);
    nstat := GetNovelStatus(str);
    // すでにタイトルに進捗状況が付いていないことをチェックして（タイトル先頭に【完結】が付いて
    // いる場合があるため）、なければ進捗状況を追加する
    if (nstat = '【完結】') and (UTF8Pos('完結', title) > 0) then
      nstat := '';
    //else if (nstat = '【短編】') and (UTF8Pos('短編', title) > 0) then
    //  nstat := '';
    // タイトル名に進捗状況を付加する
    title := nstat + title;
  end;
  // 保存するファイル名を準備する
  if FileName = '' then
  begin
    FileName := PathFilter(title);
    if StartPage <> '' then
      sn := '(' + StartPage + ')'
    else
      sn := '';
    FileName := ExtractFilePath(ParamStr(0)) + UTF8Copy(FileName, 1, 32) + sn + '.txt';
  end;
  // 作者
  authurl := '';
  auth := '';
  r := TRegExpr.Create;
  try
    r.InputString := Line;
    r.Expression  := '<div class="p-novel__author">作者：.*?</div>';
    if r.Exec then
    begin
      auth := r.Match[0];
      auth := ReplaceRegExpr('<div class="p-novel__author">作者：', auth, '');
      auth := ReplaceRegExpr('</div>', auth, '');
      r.InputString := auth;
      r.Expression  := '<a href=".*?">';
      // 作者URLがある
      if r.Exec then
      begin
        authurl := r.Match[0];
        auth    := ReplaceRegExpr(authurl, auth, '');
        auth    := ReplaceRegExpr('</a>', auth, '');
        authurl := ReplaceRegExpr('<a href="', authurl, '');
        authurl := ReplaceRegExpr('">', authurl, '');
      end;
      auth := ReplaceRegExpr('\r\n', auth, '');
    end;
  finally
    r.Free;
  end;
  ps := UTF8Pos('</a>', Line);
  if ps > 0 then
  begin
  	//auth := Restore2Realchar(UTF8Copy(Line, 1, ps - 1));
    // タイトル、作者名を保存
    PBody := title + #13#10 + auth + #13#10 + AO_PB2 + #13#10;
    LogFile.Add('小説URL :' + URL);
    LogFile.Add('タイトル:' + title);
    LogFile.Add('作者　　:' + auth);
    if authurl <> '' then
      LogFile.Add('作者URL : ' + authurl);
	end else
  	Exit;

  // 説明
  ps := UTF8Pos('="p-novel__summary">', Line);
  if ps > 0 then
  begin
  	UTF8Delete(Line, 1, ps + UTF8Length('="p-novel__summary>"') - 1);
    ps := 1;
  end;
  pe := UTF8Pos('</div>', Line);
  if (ps > 0) and (pe > ps) then
  begin
    str := UTF8Copy(Line, ps, pe - ps);

    synop := GetText(str);
    // あらすじページを保存
    if synop <> '' then
    begin
      LogFile.Add('あらすじ');
      LogFile.Add(synop);
      LogFile.Add('');
      // EConvertError回避
      try
        DateTimeToString(dt, 'yyyy/MM/dd', Today);
        LogFile.Add(dt);
      except
        ;
      end;
      //LogFile.Add(DateToStr(Today));
      if nstat <> '【短編】' then
        PBody := PBody + AO_KKL + URL + #13#10 + synop + #13#10 + AO_KKR + #13#10 + AO_PB2 + #13#10;
    end;

    ps := UTF8Pos('<br />', str);
    while ps > 0 do
    begin
      UTF8Delete(str, ps, 6);
      ps := UTF8Pos('<br />', str);
    end;
    UTF8Delete(Line, 1, pe);
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
    ps := UTF8Pos('<div class="p-eplist">', Line);
    if ps > 1 then
      UTF8Delete(Line, 1, ps + UTF8Length('<div class="p-eplist">'));

    while True do
    begin
  	  chapter := '';
  	  cs := UTF8Pos('="p-eplist__chapter-title">', Line);
		  ps := UTF8Pos('="p-eplist__subtitle">', Line);
      // Chapterがある場合
      if (cs > 0) and (ps > 0) and (cs < ps) then
      begin
        UTF8Delete(Line, 1, cs + Length('="p-eplist__chapter-title">') - 1);
        ce := UTF8Pos('</div>', Line);
        chapter := UTF8Copy(Line, 1, ce - 1);
        UTF8Delete(Line, 1, ce + 6);
        // ページURLを抽出する
        ps := UTF8Pos('<a href="', Line);
        if ps > 0 then
        begin
          UTF8Delete(Line , 1, ps + Length('<a href="') - 1);
          pe := UTF8Pos(' class="p-eplist__subtitle">', Line);
          if pe > 0 then
          begin
            str := UTF8Copy(Line, 1, pe - 1);
            UTF8Delete(Line, 1, pe + UTF8Length(' class="p-eplist__subtitle">') - 1);
            pe := UTF8Pos('</a>', Line);
            if pe > 0 then
            begin
              section := TrimJ(UTF8Copy(Line, 1, pe - 1));
            end else begin
              Writeln(#13#10'目次情報を取得出来ませんでした.');
              Exit;
            end;
          end;
          sub := IntToStr(page);
          Inc(page);
          PageList.Add(URL + sub + '/,"' + chapter + '","' + section + '",'{ + dtime});
    	  end;
      // Sectionだけの場合
      end else if ps > 0 then
      begin
        ps := UTF8Pos('<a href="', Line);
        if ps > 0 then
        begin
          UTF8Delete(Line , 1, ps + Length('<a href="') - 1);
          pe := UTF8Pos(' class="p-eplist__subtitle">', Line);
          if pe > 0 then
          begin
            str := UTF8Copy(Line, 1, pe - 1);
            UTF8Delete(Line, 1, pe + UTF8Length(' class="p-eplist__subtitle">') - 1);
            pe := UTF8Pos('</a>', Line);
            if pe > 0 then
            begin
              section := TrimJ(UTF8Copy(Line, 1, pe - 1));
            end else begin
              Writeln(#13#10'目次情報を取得出来ませんでした.');
              Exit;
            end;
          end;
          sub := IntToStr(page);
          Inc(page);
          PageList.Add(URL + sub + '/,"","' + section + '",'{ + dtime});
        end;
      // もう目次はないが次の目次ページがある場合は次ページの目次を取得する
  	  end else if ExecRegExpr(SNEXTCT, Line) then
      begin
        Line := LoadHTMLbyIndy(URL + '?p=' + IntToStr(ConteP));
        if Line <> '' then
        begin
          ps := UTF8Pos('<div class="p-eplist">', Line);
          if ps > 1 then
          begin
            UTF8Delete(Line, 1, ps + UTF8Length('<div class="p-eplist">'));
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
  if (nstat <> '【短編】') and ((PageList.Count = 0) or ((PageList.Count - StartN) < 0)) then
  begin
    if (PageList.Count - StartN) < 0 then
      Writeln('DL開始ページがありません.');
    Exit;
  end;
  // Naro2mobiから呼び出された場合は進捗状況をSendする
  if hWnd <> 0 then
  begin
    conhdl := GetStdHandle(STD_OUTPUT_HANDLE);
    sendstr := title + ',' + auth;
    Cds.dwData := PageList.Count - StartN + 1;
  {$IFDEF FPC}
    ws := UTF8ToUTF16(sendstr);
    Cds.cbData := ByteLength(ws) + 2;
    Cds.lpData := PWideChar(ws);
  {$ELSE}
    Cds.cbData := (UTF8Length(sendstr) + 1) * SizeOf(Char);
    Cds.lpData := Pointer(sendstr);
  {$ENDIF}
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
    Writeln('以下のサイトからopenssl-1.0.2u-x64_86-win64.zipをダウンロードしてlibeay32.dllとssleay32.dllをna6dl.exeがあるフォルダにコピーして下さい.');
    Writeln('https://github.com/IndySockets/OpenSSL-Binaries');
    ExitCode := 2;
    Exit;
  end;
  // OpenSSLのバージョンをチェック
  if (UTF8Pos('1.0.2', GetVersionInfo('libeay32.dll')) = 0)
    or (UTF8Pos('1.0.2', GetVersionInfo('ssleay32.dll')) = 0) then
  begin
    Writeln('');
    Writeln('OpenSSLライブラリのバージョンが違います.');
    Writeln('以下のサイトからopenssl-1.0.2u-x64_86-win64.zipをダウンロードしてlibeay32.dllとssleay32.dllをna6dl.exeがあるフォルダにコピーして下さい.');
    Writeln('https://github.com/IndySockets/OpenSSL-Binaries');
    ExitCode := 2;
    Exit;
  end;

  if ParamCount = 0 then
  begin
    Writeln('');
    Writeln('na6dl ver4.1 2025/4/10 (c) INOUE, masahiro.');
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
    if UTF8Pos('-h', op) = 1 then
    begin
      UTF8Delete(op, 1, 2);
      try
        hWnd := StrToInt(op);
      except
        Writeln('Error: Invalid Naro2mobi Handle.');
        ExitCode := -1;
        Exit;
      end;
    // DL開始ページ番号
    end else if UTF8Pos('-s', op) = 1 then
    begin
      UTF8Delete(op, 1, 2);
      StartPage := op;
      try
        StartN := StrToInt(op);
      except
        Writeln('Error: Invalid Start Page Number.');
        ExitCode := -1;
        Exit;
      end;
    // 作品URL
    end else if UTF8Pos('https:', op) = 1 then
    begin
      URL := op;
      if UTF8Copy(URL, UTF8Length(URL), 1) <> '/' then
        URL := URL + '/';
    // それ以外であれば保存ファイル名
    end else begin
      FileName := op;
      if UTF8UpperCase(ExtractFileExt(op)) <> '.TXT' then
        FileName := FileName + '.txt';
    end;
  end;

  if (UTF8Pos('https://ncode.syosetu.com/n', URL) = 0) and (UTF8Pos('https://novel18.syosetu.com/n', URL) = 0) then
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

  isOver18 := UTF8Pos('https://novel18.syosetu.com/n', URL) > 0;

  IdHTTP := TIdHTTP.Create(nil);
  IdSSL := TIdSSLIOHandlerSocketOpenSSL.Create;
  Cookies := TIdCookieManager.Create(nil);
  try
    IdSSL.IPVersion := Id_IPv4;
    IdSSL.MaxLineLength := 32768;
    IdSSL.SSLOptions.Method := sslvSSLv23;
    IdSSL.SSLOptions.SSLVersions := [sslvSSLv2,sslvTLSv1];
    IdHTTP.HandleRedirects := True;
    IdHTTP.AllowCookies := True;
    IdHTTP.IOHandler := IdSSL;
    // user-agentをMicrosoft Edgeにする
    IdHTTP.Request.UserAgent := 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/79.0.3945.117 Safari/537.36 Edg/79.0.309.65';
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
    DLErr := '';
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
            TextPage.WriteBOM := True;
            LogFile.WriteBOM := True;
            TextPage.SaveToFile(DLErr + Filename, TEncoding.UTF8);
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
  finally
    IdSSL.Free;
    IdHTTP.Free;
    Cookies.Free;
  end;
end.

