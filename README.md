# footnote-api

.docx 内の注マーカーを Word ネイティブの脚注に変換する、小さな FastAPI サービス。

本文に `【【1】】` と書いておき、.docx と注釈テキストの一覧を送ると、各マーカーが本物の脚注に置き換わった .docx が返る。

同じ処理を Word の自動操作（pywin32）で行う、Windows 専用の CLI も同梱している。

## 特徴

- `【【N】】` マーカーを Word の脚注参照に置換する
- `word/footnotes.xml` を新規生成し、パッケージに登録する
- マーカー前後の run 書式を引き継ぐ
- 使用した注番号と欠番をレスポンスヘッダーで返す
- 対応する注釈がないマーカーは、本文にそのまま残す
- クライアント IP でアクセスを制限する（「アクセス制御」を参照）
- Vercel 上で単一の Python 関数として動作する

## 必要環境

- Python 3.14 以上
- [uv](https://docs.astral.sh/uv/)
- （デプロイ時のみ）Vercel CLI。`pnpm dlx vercel` 経由で実行する
- （CLI のみ）Microsoft Word と Excel がインストールされた Windows

## セットアップ

```cmd
uv sync
```

## ローカル実行

`.env` を作成し、以下のようにlocalhostの IP アドレスを許可する。許可リストにない IP からのリクエストは 403 で拒否されるため、空の `.env` ではすべてのリクエストが失敗する。

```
ALLOWED_IP_ADDRESSES=::1,127.0.0.1
SERVE_STATIC=1
```

| 名前 | 説明 |
| --- | --- |
| `ALLOWED_IP_ADDRESSES` | 許可するクライアント IP のカンマ区切りリスト（必須） |
| `SERVE_STATIC` | 設定すると `public/` 内のファイルを `/` で配信する（任意） |


```cmd
uv run --env-file .env fastapi dev index.py
```

`http://127.0.0.1:8000/docs` を開くと、Swagger UI から API を試せる。


## デプロイ

デプロイ前に、Vercel のプロジェクト設定で `ALLOWED_IP_ADDRESSES` を設定する。

```cmd
pnpm dlx vercel deploy
```

Vercel は `index.py` 内の `app` を FastAPI のエントリポイントとして検出する。
Vercel 上ではリクエストボディが約 4.5 MB に制限される。

## API

### `POST /api/footnotes`

Content-Type: `multipart/form-data`

| フィールド | 型 | 説明 |
| --- | --- | --- |
| `file` | file | 変換元の .docx ファイル |
| `footnotes` | string | `[番号, テキスト]` ペアの JSON 配列 |

`footnotes` の例:

```json
[["1", "1つ目の注釈"], ["2", "2つ目の注釈"]]
```

リクエスト例:

```cmd
curl -X POST http://127.0.0.1:8000/api/footnotes ^
  -F "file=@input.docx" ^
  -F "footnotes=[[\"1\",\"1つ目の注釈\"],[\"2\",\"2つ目の注釈\"]]" ^
  -o output.docx -D -
```

成功時のレスポンス（200）:

- ボディ: 変換後の .docx
- `X-Footnote-Count`: 挿入した脚注の数
- `X-Missing-Numbers`: `footnotes` 内の最大の番号までのうち、文書内で使われなかった番号（カンマ区切り。なければ空。`footnotes` に数字の番号がない場合も空）

エラー時のレスポンス:

| ステータス | 原因 |
| --- | --- |
| 400 | `footnotes` が `[番号, テキスト]` ペアの JSON 配列ではない |
| 403 | クライアント IP が `ALLOWED_IP_ADDRESSES` に含まれていない |

`footnotes` に対応するエントリがないマーカーはエラーにならず、本文にそのまま残る。

## アクセス制御

`POST /api/footnotes` は `security.py` の `require_allowed_ip` に依存する。

- 許可リストは、起動時に環境変数 `ALLOWED_IP_ADDRESSES`（カンマ区切り。各要素の前後の空白は無視）から一度だけ読み込む
- クライアント IP は `X-Forwarded-For` ヘッダーの先頭の値。ヘッダーがなければ接続元アドレスを使う
- リストにない IP からのリクエストは `403 Forbidden` で拒否する
- 環境変数が未設定または空の場合、すべてのリクエストを拒否する
- 完全一致のみ対応。CIDR 範囲やワイルドカードは使えない
- 環境変数を変更したら、再起動（Vercel では再デプロイ）が必要

注意: Vercel のような信頼できるプロキシの背後にない場合、クライアントが `X-Forwarded-For` を偽装できる。公開ネットワーク上でアプリを直接動かす場合は、このチェックに頼らないこと。

## 処理の流れ

`footnotes/converter.py` の `apply_footnotes` が、.docx を zip として開き、必要なパーツを書き換えて新しいアーカイブを作る。

```
apply_footnotes(docx_bytes, mapping)
|
+- docx を zip として開き、各メンバーを順に処理
|   +- word/footnotes.xml             -> 破棄（最後に作り直す）
|   +- word/document.xml              -> マーカーを脚注参照に置換
|   +- [Content_Types].xml            -> footnotes の Override を保証
|   +- word/_rels/document.xml.rels   -> footnotes の Relationship を保証
|   +- word/styles.xml                -> 脚注用スタイルを保証
|   +- その他                         -> そのままコピー
|
+- 元に styles.xml がなければ新規作成
+- word/footnotes.xml を生成して追加
```

### マーカーの置換

`footnotes/markers.py` で実装している。`document.xml` 内のすべての段落（`w:p`）に対して、次の処理を行う。

1. 段落直下の単純なテキスト run を集める（「前提とする XML 構造」を参照）
2. 各 run のテキストを連結し、run ごとのオフセット範囲を記録する
3. 正規表現で `【【N】】` を探す
4. マーカー番号に対応する注釈を探す。なければマーカーを残し、その直後から探索を続ける
5. あれば、マーカーにかかる run を取り除き、次の 3 つを順に挿入する
   マーカー前のテキスト、脚注参照 run、マーカー後のテキスト。
   前後のテキストには、元の run の `w:rPr` を引き継ぐ
6. 段落内に置換できるマーカーがなくなるまで繰り返す

脚注 ID は出現順（1, 2, 3...）に採番する。
本文中のマーカー番号とは独立した値である。

## 前提とする XML 構造

変換処理は、以下の WordprocessingML 構造（ECMA-376）を前提にしている。
これと異なる構造の文書は、一部しか変換されない、または変換されない場合がある。

### 1. 本文側の脚注参照（`word/document.xml`）

```xml
<w:r>
  <w:rPr>
    <w:rStyle w:val="FootnoteReference"/>
  </w:rPr>
  <w:footnoteReference w:id="1"/>
</w:r>
```

- `w:id` は、`footnotes.xml` 内の `w:footnote` の ID と一致させる
- 参照の番号は Word が自動で振るため、ID の値は画面に表示されない
- 上付き表示は `FootnoteReference` スタイルが担当する

### 2. 脚注本体（`word/footnotes.xml`）

```xml
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:type="separator" w:id="-1">
    <w:p><w:r><w:separator/></w:r></w:p>
  </w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0">
    <w:p><w:r><w:continuationSeparator/></w:r></w:p>
  </w:footnote>
  <w:footnote w:id="1">
    <w:p>
      <w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>
      <w:r>
        <w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>
        <w:footnoteRef/>
      </w:r>
      <w:r><w:t xml:space="preserve"> 注釈テキスト</w:t></w:r>
    </w:p>
  </w:footnote>
</w:footnotes>
```

- ID -1 と 0 は Word 必須の区切り用エントリ。通常の脚注は 1 から始まる
- `w:footnoteRef` が、脚注エリア側の番号を描画する
- `xml:space="preserve"` は、注釈テキスト先頭の空白を保持するために必要
- 元ファイルに `footnotes.xml` があっても、破棄して置き換える
- ルート要素の名前空間宣言は `document.xml` と同じものを使う

### 3. パッケージへの登録

`[Content_Types].xml`:

```xml
<Override PartName="/word/footnotes.xml"
  ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
```

`word/_rels/document.xml.rels`:

```xml
<Relationship Id="rIdFootnotes"
  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
  Target="footnotes.xml"/>
```

`rIdFootnotes` が既に使われている場合は、重複しなくなるまで末尾に `X` を足す。

`word/styles.xml` には、上記で使う 2 つのスタイルの定義が必要:

```xml
<w:style w:type="paragraph" w:styleId="FootnoteText">...</w:style>
<w:style w:type="character" w:styleId="FootnoteReference">
  <w:rPr><w:vertAlign w:val="superscript"/></w:rPr>
</w:style>
```

3 つの登録のうち 1 つでも欠けると、Word が「ファイルが破損している」と報告したり、脚注が表示されなかったりする。
`footnotes/package.py` の `ensure_*` 系の関数は、不足しているものだけを追加し、既存のエントリには手を付けない。

### 4. 元文書の run に関する前提

- マーカーの検出対象は、`w:p` の直下にあり、`w:t` のみを含む `w:r` に限る
- マーカーが複数の run に分割されていてもよい。オフセットを各 run に対応付けて処理する
- `document.xml` の名前空間は接頭辞付き（`w`、`r` など）で宣言されている前提。既定名前空間（接頭辞なし）は `footnotes.xml` に引き継がない

## 制限事項

- ハイパーリンク内のマーカーは検出されない
- タブ・改行・画像を含む run にまたがるマーカーは検出されない
- 処理対象は `word/document.xml` のみ。ヘッダー、フッター、別パーツのテキストボックスは対象外
- 元ファイルにあった既存の脚注は削除される
- 注釈はプレーンテキストのみ。注釈内の書式指定は未対応

## ディレクトリ構成

```
footnote-api/
+- index.py              FastAPI アプリと HTTP 層
+- security.py           IP 許可リストによるアクセス制御（ALLOWED_IP_ADDRESSES）
+- footnotes/            docx 変換ロジック
|   +- __init__.py       公開 API（apply_footnotes, find_unused_numbers）
|   +- converter.py      zip の処理と全体の流れ
|   +- markers.py        マーカーの検索と run の置換
|   +- numbers.py        欠番の計算
|   +- package.py        Content Types・Relationship・スタイルの保証
|   +- parts.py          footnotes.xml の要素の生成
|   +- xml_utils.py      XML の解析とシリアライズ
|   +- constants.py      名前空間とマーカーのパターン
+- cli/
|   +- main.py           Word と Excel を操作する一括処理 CLI
+- public/               静的ファイル（SERVE_STATIC 設定時のみ配信）
+- .env                  ローカル用の環境変数ファイル（git 管理外）
+- pyproject.toml
+- uv.lock
+- README.md
```

## CLI

pywin32 で Word と Excel を起動し、Word の脚注機能で注マーカーを置換する。
API とは異なり、ディレクトリ内の .docx ファイルを（サブディレクトリも含めて）一括処理できる。

```cmd
uv run --group cli cli\main.py （対象ディレクトリのパス）
```

パスを省略した場合は `data` ディレクトリを対象にする。

入力:

- `名前.docx` ごとに、同じ場所に同名の `名前.xlsx` を置く
- 最初のワークシートに脚注を定義する。1 列目が注番号、2 列目が注釈テキスト

出力（元ファイルと同じ場所）:

| ファイル | 説明 |
| --- | --- |
| `名前_脚注付.docx` | 変換後の文書 |
| `名前_脚注付_欠番.txt` | 欠番とその注釈テキスト。欠番がある場合のみ出力する |
| `report.csv` | 対象ディレクトリ直下に出力。ファイル名と追加した脚注数 |

挙動:

- 対応する .xlsx がない .docx は、スキップする
- 出力済みの `_脚注付` ファイルと、Word の一時ファイル（`~$` で始まるもの）は処理しない
- `X-Missing-Numbers` と同様に、Excel 内の最大の番号までのうち、本文で使われなかった番号を欠番とする。Excel に数字の番号がない場合は欠番なしとして扱う
- API と同様に、Excel に対応する注釈がないマーカーは本文にそのまま残す。残したマーカーの番号は、コンソールに警告として表示する
- 処理中は変更履歴の記録を無効にする