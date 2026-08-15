# アプリケーション別リファレンス

[HANDOFF.md](HANDOFF.md) の補足資料。11画面それぞれの役割・使用しているGASアクション・データの流れ・このセッションで見つけて直した不具合とその理由をまとめる。
「同じ調査を繰り返さない」ことが目的なので、憶測は書かず、コード上で確認できた事実と、実際に修正したコミットの理由だけを書く。

## 目次
- [共通基盤](#共通基盤)
- [1. portal.html](#1-portalhtml-ポータル)
- [2. index.html](#2-indexhtml-症例登録)
- [3. search.html](#3-searchhtml-事案検索コンソール)
- [4. management.html](#4-managementhtml-統括業務管理)
- [5. report.html](#5-reporthtml-業務日報)
- [6. checklist.html](#6-checklisthtml-統合チェックリスト)
- [7. debriefing.html](#7-debriefinghtml-デブリーフィング)
- [8. debrief_search.html](#8-debrief_searchhtml-教育事案検索)
- [9. library.html](#9-libraryhtml-資料集)
- [10. admin.html](#10-adminhtml-システム管理コンソール)
- [11. summary.html](#11-summaryhtml-jsamsダッシュボード)
- [Code.gs 全アクション一覧](#codegs-全アクション一覧)
- [スプレッドシートのシート名一覧](#スプレッドシートのシート名一覧実際にgetsheetflexibleが探している名前)

---

## 共通基盤

全11画面が共有している設計・実装パターン。新しい画面を追加する、または既存画面を大きく改修するときは、ここからの逸脱がないか確認する。

### 認証ガード（画面冒頭の定型コード）
```js
if (!localStorage.getItem('aw109_auth_token')) { location.href = 'portal.html'; }
// 🌟 勤務シフト想定：ログインから８時間経過で自動ログアウト
(function() { ... aw109_auth_loginTime を見て8h超なら強制ログアウト ... })();
const PIN_CODE = localStorage.getItem('aw109_auth_token') || "";
```
`PIN_CODE`という変数名は歴史的経緯（PIN認証時代の名残）で、実体は認証トークン文字列。GASへのリクエストの`password`フィールドにそのまま載せて送る。admin.htmlだけは`PIN_ADMIN`という別名を使う（詳細は10番参照）。

### エラーログ送信
```js
sendErrorLog(`Msg: ${e.message || e}\nStack: ${e.stack || e.toString()}`);
```
**必ずこの形式にすること。** `sendErrorLog(e.stack || e.toString())`だけだと、iOS Safariでは`e.stack`にエラーメッセージ本文が含まれないため、実際の失敗理由が永久に分からなくなる（2026-08-15、全11画面38箇所で発覚・修正済み）。`window.onerror`ハンドラも`Source: ${source}`（発生元URL）を含めること。

### タイムゾーン・日付処理
GAS側は`Asia/Tokyo`前提。スプレッドシート上の日付表記(`YYYY/M/D`)とHTML5 `<input type="date">`の表記(`YYYY-MM-DD`)の揺れは`normalizeDate`系の関数で吸収している（library.html/report.html/management.htmlに実装）。

### 時間入力
2026-08-15に全画面で`<input type="time">`（ネイティブのスクロール式ピッカー）を廃止し、数値キーパッド入力＋自動整形に統一した。
```js
function formatTimeInput(el) { /* "1430" → "14:30" に自動整形。oninput属性で呼ぶ */ }
function clampTimeInput(el)  { /* HH:0-23, MM:0-59 に丸める。onblur属性で呼ぶ */ }
```
値は常に厳密な`"HH:MM"`文字列（ゼロ埋め）。`new Date(\`1970-01-01T${v}:00\`)`のようにDateへパースするコードが複数画面にあるため、**この文字列フォーマットを崩す変更をしないこと。**

### ヘッダーのレイアウト
2026-08-15にヘッダーの「タイトルや操作ボタンが狭い画面で見切れる／不自然に折り返る」問題を修正した。方針：
- 左右のボタン群は`flex-shrink:0; white-space:nowrap;`で絶対に折り返さない
- 優先度の低いボタン（更新・Q&A・印刷など）は400px以下で文字ラベルを`<span class="btn-label">`で囲み、そのspanを`display:none`にしてアイコンのみにする（`◀ ポータル`のような主要導線ボタンは常に文字を残す）
- 中央のタイトルは`flex:1; min-width:0;`にし、それでも入らない場合だけ`overflow:hidden; text-overflow:ellipsis;`で省略記号にする（最終手段。折り返しでレイアウトが崩れるほうが省略よりも悪いという判断）
- search.htmlだけは`header`自体に`flex-wrap:wrap`を付けて、ヘッダー全体が2段に折り返る別方式を採用済み（既存の動作を変えていない）

### PWA / Service Worker
`sw.js`は以前どのページからも`register()`されておらず、オフラインキャッシュが一度も有効化されていなかった（2026-08-15発見・修正）。現在は`portal.html`（start_url）のみが登録している。スコープはサイト全体に及ぶので他画面での登録は不要。

---

## 1. portal.html（ポータル）

全画面の玄関口。ログインもここで行う。

- **GASアクション**: `fetch_init`（起動時。ニュース・アラート・マニュアル・Q&A・マスタを一括取得）、`error_log`、`submit_question`
- **認証**: `auth_login`/`auth_register`は別画面ではなくこのファイル内の`doLogin()`/`doRegister()`が直接呼ぶ（Code.gs側の`action`名も同じ）
- **主要機能**:
  - `fetchInit(isSilent)`: キャッシュ即時表示→バックグラウンドで最新化。ヘッダーの🔄更新ボタンからも`fetchInit(false)`で手動起動できる（2026-08-15追加、以前は無かった）
  - 「カルテ未完了の事案がありますN件」アラートは`search.html?auto_filter=uncompleted`へリンク（詳細はsearch.html側の既知問題を参照）
  - 取扱説明書モーダル：`システム_取扱説明書`シートの内容を表示。**このシート名の取り違えで長期間空表示になっていたバグを2026-08-15に修正済み**（HANDOFF.md参照）
- **既知の設計**:
  - `account-admin-badge`はisAdmin===trueの時だけ表示。バッジは1種類のみ、区分けなし
  - アカウントボタンの🔐アイコンはログイン済みなら誰にでも付く（管理者かどうかとは無関係）

## 2. index.html（症例登録）

現場でのメイン入力画面。最もフィールドが多く複雑。

- **GASアクション**: `fetch_init`, `submit`/`update_record`（新規/編集で出し分け, `mode`変数由来）, `update_status`（カルテ完了処理）, `delete_record`, `fetch_recent_cases`, `add_master`, `send_email`, `submit_question`, `error_log`
- **データモデル**: `getCollectData()`が最終的な送信payloadを組み立てる。処置・ライン確保系は「実施フラグを持つ列」と「実施フラグを持たずサイズ/回数等のサブ項目だけで記録される処置」が混在している（例: 気管挿管は単独の◯フラグ列が無く、`気管挿管サイズ`等が入っていることで実施を判定する）。**この構造がdebriefing.html側の処置タグ取り込みロジックと密結合している（7番参照）。新しい処置チェックボックスを追加したら必ずdebriefing.htmlのimport処理も見直すこと。**
- **既知の設計・修正履歴**:
  - チェックボックスIDと列名が食い違っている箇所がある（例: UI上「ターニケット」の実体列は`止血帯`、UI上「バックボード固定」相当の実体列は`バックボード`）。debriefing.html側で別名変換テーブルを持って吸収している
  - `chk-ls-shock`/`chk-ls-defib`のようなID不一致がかつて本番障害を起こした（2026年前半、修正済み）
  - Safari特有の`querySelector`セレクタインジェクション（自由入力文字列をそのままCSSセレクタに埋め込むとクラッシュ）が過去2箇所で本番障害の原因になった。自由入力値との比較は`document.querySelectorAll(...).forEach`で`.value === val`比較する方式に統一済み。**新しいコードでテンプレートリテラルによる動的セレクタ生成をしないこと**

## 3. search.html（事案検索コンソール）

過去の登録済み事案を検索・修正・カルテ完了処理する画面。

- **GASアクション**: `fetch_init`, `fetch_all`, `submit_question`, `error_log`
- **キャッシュ戦略**: マスタは6時間キャッシュ(`CACHE_EXPIRE_MS`)、検索結果データは90秒キャッシュ(`DATA_CACHE_EXPIRE_MS`, sessionStorage)。電波の弱い現場での毎回全件通信によるタイムアウトを避けるため
- **既知の設計・修正履歴**:
  - 10分無操作でトークンを完全に破棄しポータルへ強制ログアウトするタイマー機能あり（`startTimer()`）。admin.htmlの同機能とは異なり正しく実装されている
  - PIN認証時代の独立ログインオーバーレイが認証システム刷新後もPINと平文比較する形で残っており「何を入力してもログインできない」障害を起こした。現在は削除済みで、ポータルでのトークン認証のみに一本化
  - `?auto_filter=uncompleted`（ポータルの「カルテ未完了」アラートから遷移してくる）を処理する経路が、初期化ロジックの別の分岐（前回検索状態の復元）と競合し、**同じページ内で`fetchData()`が2回同時に走る競合状態があった**（2026-08-15修正）。`aw109_search_state`がsessionStorageに残っている状態＋`auto_filter`付き遷移、の組み合わせで再現する
  - モバイル幅で`main`が`box-sizing:border-box`を欠いておりpadding分だけ横スクロールが出る不具合があった（修正済み）

## 4. management.html（統括・業務管理）

月間カレンダーで出動状況・待機理由を一覧表示。Tailwind CSS + Chart.js使用。

- **GASアクション**: `fetch_all_reports`, `error_log`
- **既知の設計**: 月を切り替えたときだけ`fetchData()`が呼ばれ、同じ月のまま裏でデータが更新されても手動で再取得する手段が無かった。2026-08-15にヘッダーへ🔄更新ボタンを追加した
- **注意**: ヘッダー中央のタイトル領域に`#sync-timestamp`という補助バッジがあり、狭い画面では隠すようにしている（`display:none !important`）

## 5. report.html（業務日報）

浜町日誌。当直帯のタイムラインを自動生成し印刷する。SunCalcで日の出没を計算。

- **GASアクション**: `fetch_init`, `fetch_daily_report`, `save_daily_report`, `error_log`
- **既知の設計・修正履歴**:
  - 「🚁 症例情報を再取り込み」ボタン（`syncFromDb(true)`）がフォーム内に既にあり、これがマスタ更新ボタンの役割を果たしている（ヘッダーには無い。取扱説明書にも「事案データの再同期と上書き保護」として説明あり：始業点検メモ等の手入力データは再同期しても保護される）
  - 自由入力の理由文字列を`document.querySelector(`input[name="reason"][value="${r.trim()}"]`)`のようにテンプレートリテラルでセレクタ化しておりSafariでクラッシュしていた。チェックボックス総当たりの`.value === val`比較に修正済み
  - 「担当医師」表記が実際は「フライトドクター」の意味で使われていた表記ゆれを修正済み

## 6. checklist.html（統合チェックリスト）

ME機器・ナースバック等の日常点検。ペーパーレス化。

- **GASアクション**: `fetch_init`, `fetch_checklist`, `fetch_checklist_status`, `fetch_checklist_history`, `submit_checklist`, `delete_checklist_record`, `submit_question`, `error_log`
- **既知の設計**: 点検種別（始業点検/終業点検/ME機器/ナースバック/ヘリ内/ヘリバック/待機物品/定期点検）ごとに`TYPE_TO_SHEET_MAP`（Code.gs）でシート接尾辞に変換され、`点検マスタ_${接尾辞}`/`DB_点検_${接尾辞}`という専用シートに読み書きする（詳細は本資料末尾の表を参照）
- **修正履歴**: `toHalfWidth`関数がこのファイルにだけ存在しなかった（2026-08-15、時間入力の数値化対応時に追加）

## 7. debriefing.html（デブリーフィング）

多職種での事案振り返り。**症例登録(index.html)で入力されたデータをインポートして評価タグを付ける画面**であり、独自に処置を一から入力する画面ではない。

- **GASアクション**: `fetch_init`, `fetch_recent_cases`（症例検索）, `fetch_debriefings`, `submit_debriefing`, `update_review_status`, `add_master`, `submit_question`, `error_log`
- **データの流れ（重要）**: `selectCase()`が症例登録側の生データ(`currentCase`)から評価用タグ(`sysData.tags.globalTreatment`)を自動生成する。この生成ロジックが**index.htmlの列名変更に追従できておらず、2026-08-15に大きく手直しした**：
  - スキップすべき列（算定用の`管理料_◯◯`列、サイズ/回数/タイプ等のサブ項目列、動的な`酸素N開始`等の時間列）が網羅されておらず、生の列名がそのまま処置タグとして表示されてしまっていた
  - 単独の実施フラグ列を持たない処置（気管挿管・NPPV・除細動・胸腔ドレナージ・心嚢穿刺・胃管挿入・末梢静脈路・骨髄路・酸素投与・人工呼吸器管理・胸骨圧迫）は、サブ項目の存在から実施を推定して正しい処置名でタグ化する`PROC_ALIASES`テーブルを新設
  - 症例登録側の列名とデブリーフィング側の既存タグ名が食い違う項目（`バックボード`→`バックボード固定`、`止血帯`→`ターニケット`）は`COL_TO_TAG`で変換
  - **今後index.htmlに処置チェックボックスを追加・変更したら、必ず`selectCase()`内の`skipKeys`/`PROC_ALIASES`/`COL_TO_TAG`を見直すこと。忘れると同じ「意味不明な列名タグが混入する」不具合が再発する**
- **その他の修正履歴**: 実施日時・キーワード・疾患分類・小児母性多数傷病フラグ（`"◯"`判定漏れ）のインポート漏れ、使用薬剤リストが汎用タグ取り込みループに巻き込まれる不具合、読み込み速度改善（並列fetch化）

## 8. debrief_search.html（教育・事案検索）

過去のデブリーフィング内容を横断検索する、教育目的のアーカイブ画面。

- **GASアクション**: `fetch_init`, `fetch_all`, `fetch_debriefings`, `submit_question`, `error_log`（3つの通信を`Promise.all`で並列実行）
- **既知の設計・修正履歴**:
  - `DB_デブリーフィング`の「要請番号」列には症例のUUID(sysId)が入っているため、`id`/`displayId`の両方で紐付け検索する必要がある
  - `タイムライン評価JSON`/`多職種フィードバックJSON`列が配列でない値（`"{}"`等）を含む行が1件でもあると、`.map()`が例外を投げて**forEachループ全体が停止し検索結果が丸ごと表示されなくなる**クラッシュがあった（2026-08-15修正。`Array.isArray()`で防御）
  - `jianList`/`debData`が不正な応答で`undefined`になった場合の`forEach`クラッシュ対策で`|| []`ガードを追加済み

## 9. library.html（資料集）

外部資料（PDFプロトコル集等）のカテゴリ検索・閲覧。

- **GASアクション**: `fetch_init`, `fetch_library`, `update_library_record`, `submit_question`, `error_log`
- **既知の設計**: ヘッダー右端の「管理」ボタン（`toggleAdmin()`）は絵文字なしのテキストのみのボタン。admin.htmlとは別の、この画面内だけの簡易管理モード

## 10. admin.html（システム管理コンソール）

お知らせ・取扱説明書・Q&A・各種マスタのCRUD管理画面。**このセッションで最も深刻なバグが見つかった画面。**

- **GASアクション**: `fetch_init`, `add_master`, `manage_news`, `manage_manual`, `manage_qa_full`, `error_log`
- **認証の特殊性**: 他画面と違い、ログイントークンが無い（または管理者でない）場合はアクセスコード入力画面が出る。`PIN_ADMIN`は「トークンがあり管理者ならそのトークン、なければ緊急用の`"9999"`」という二重の意味を持つ変数。**`9999`は意図的に画面上へヒントを出していない**（GitHub公開リポジトリのため、表示すると秘密の意味がなくなる）
- **重大バグ（修正済み）**: 「表記統一」コミット(`b1ff600`)で削除確認メッセージの改行を`\n`ではなく生の改行文字で書いてしまい、JS文字列リテラルとして構文エラーになっていた。**この1文字のミスでscriptタグ全体が読み込み失敗し、`checkLogin`関数を含む全機能がそのコミット以降ずっと動作していなかった**（2026-08-15発見・修正）。今後同種の複数行メッセージを書くときは`\n`を使うこと、生の改行を埋め込まないこと
- **もう一つのバグ（修正済み）**: 10分無操作ロックタイマーが`location.reload()`するだけで、管理者トークンでログイン中はリロード後も自動再ログインしてしまい実質ロックが機能していなかった。search.htmlと同じ「トークンを消してポータルへ強制ログアウト」に統一済み
- **`manage_manual`アクションのシート名バグ（修正済み）**: `fetch_init`の読み取り側だけでなく、この書き込み側(`manage_manual`)も`マスタ_取扱説明書`という誤った名前でシートを探しており、取扱説明書の追加・編集・削除が常に失敗していた（2026-08-15発見・修正。詳細はGit履歴`d717f5b`）

## 11. summary.html（JSAMSダッシュボード）

多次元統計・BIダッシュボード。Chart.jsでPivotヒートマップ・レーダーチャート。

- **GASアクション**: `fetch_all`, `fetch_debriefings`, `error_log`
- **既知の設計**: タイトルが11画面中もっとも長い（「JSAMS レジストリ・ダッシュボード」）。狭い画面での見切れ対策として右側ボタンのアイコン化・フォント縮小を最も強くかけている（`font-size:12px`まで）。ヘッダーにマスタ更新ボタンが無かったため2026-08-15に追加

---

## Code.gs 全アクション一覧

`action`パラメータで振り分けられる全ハンドラ（2026-08-15時点、Code.gs内`if (action === "...")`を全て列挙）。新しいアクションを追加したら、`doPost`冒頭の`allowed`ホワイトリスト配列にも追加しないと実行されない。

```
add_master, auth_login, auth_register,
delete_checklist_record, delete_record, error_log,
fetch_all, fetch_all_reports, fetch_checklist, fetch_checklist_history,
fetch_checklist_status, fetch_daily_report, fetch_debriefings, fetch_init,
fetch_library, fetch_recent_cases,
manage_manual, manage_news, manage_qa_full,
save_daily_report, send_email, submit, submit_checklist, submit_debriefing,
submit_question, update_library_record, update_record, update_review_status,
update_status
```

## スプレッドシートのシート名一覧（実際にgetSheetFlexibleが探している名前）

**シート名を変更・追加するときは、この一覧とCode.gs内の全`getSheetFlexible`呼び出しを両方確認すること。** 1箇所直しても別のアクションが古い名前を探したままになっているケースが2件見つかっている（取扱説明書のread/write両方）。

| 用途 | 探索候補（先頭が優先） |
|---|---|
| 事案データベース | `DB_事案`, `事案データベース` |
| 基本設定マスタ（ユーザー・認証情報含む） | `マスタ_基本設定`, `マスタデータ` |
| 日報 | `DB_日報`, `浜町日誌` |
| 処置項目マスタ | `マスタ_処置項目` |
| お知らせ | `DB_お知らせ`, `お知らせデータベース` |
| Q&A | `DB_QA`, `Q&Aデータベース` |
| 取扱説明書 | `システム_取扱説明書`, `マスタ_取扱説明書`, `取扱説明書_現場`, `取扱説明書_検索` ← 実際のシート名は先頭の`システム_取扱説明書`。他は歴史的な誤った候補で、後方互換のために残っているだけ |
| デブリーフィング | `DB_デブリーフィング`, `デブリーフィングデータベース` |
| 点検マスタ/DB | `getSheetFlexible`ではなく`resolveSheetName(type, isMaster)`が`点検マスタ_${base}`/`DB_点検_${base}`を動的生成。`base`は`TYPE_TO_SHEET_MAP`（Code.gs 61行目付近）で点検種別→シート名接尾辞に変換される: `始業点検`/`終業点検`→`日常`、`ME機器`→`ME`、`ナースバック`→`ナースバック`、`ヘリ内`→`ヘリ内`、`ヘリバック`→`ヘリバック`、`待機物品`→`待機物品`、`定期点検`→`定期`。新しい点検種別を追加するときは`TYPE_TO_SHEET_MAP`とスプレッドシート側のシート新設を両方行うこと |

シート一覧の正本はスプレッドシート内の「シート管理」シートに一覧があるので、疑わしいときはコード上の推測ではなく `read_file_content` でスプレッドシートを直接読んで確認すること。
