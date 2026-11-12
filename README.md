# footnote-api

.docx 内の注マーカーを Word ネイティブの脚注に変換する、小さな FastAPI サービス。

本文に `【【1】】` と書いておき、.docx と注釈テキストの一覧を送ると、各マーカーが本物の脚注に置き換わった .docx が返る。

## 特徴

- `【【N】】` マーカーを Word の脚注参照に置換する
- `word/footnotes.xml` を新規生成し、パッケージに登録する
- マーカー前後の run 書式を引き継ぐ
- 使用した注番号と欠番をレスポンスヘッダーで返す
- Vercel 上で単一の Python 関数として動作する

## 必要環境

- Python 3.12 以上
- [uv](https://docs.astral.sh/uv/)
- （デプロイ時のみ）Vercel CLI。`pnpm dlx vercel` 経由で実行する

## セットアップ

```cmd
uv sync
```

## ローカル実行

```cmd
uv run --env-file .env fastapi dev index.py
```

`http://127.0.0.1:8000/docs` を開くと、Swagger UI から API を試せる。
このサービスに環境変数は不要なので、`.env` は空ファイルで構わない。

## デプロイ

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
- `X-Missing-Numbers`: 使用した注番号の欠番（カンマ区切り。なければ空）

エラー時のレスポンス:

| ステータス | 原因 |
| --- | --- |
| 400 | `footnotes` が `[番号, テキスト]` ペアの JSON 配列ではない |
| 422 | 文書内のマーカーに対応するエントリが `footnotes` にない |

## 処理の流れ

`footnotes.py` の `apply_footnotes` が、.docx を zip として開き、必要なパーツを書き換えて新しいアーカイブを作る。

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

`document.xml` 内のすべての段落（`w:p`）に対して、次の処理を行う。

1. 段落直下の単純なテキスト run を集める（「前提とする XML 構造」を参照）
2. 各 run のテキストを連結し、run ごとのオフセット範囲を記録する
3. 正規表現で `【【N】】` を探す
4. マーカーにかかる run を取り除き、次の 3 つを順に挿入する
   マーカー前のテキスト、脚注参照 run、マーカー後のテキスト。
   前後のテキストには、元の run の `w:rPr` を引き継ぐ
5. 段落内にマーカーがなくなるまで繰り返す

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

### 3. パッケージへの登録

`[Content_Types].xml`:

```xml
<Override PartName="/word/footnotes.xml"
  ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
```

`word/_rels/document.xml.rels`:

```xml
<Relationship Id="rIdN"
  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
  Target="footnotes.xml"/>
```

`word/styles.xml` には、上記で使う 2 つのスタイルの定義が必要:

```xml
<w:style w:type="paragraph" w:styleId="FootnoteText">...</w:style>
<w:style w:type="character" w:styleId="FootnoteReference">
  <w:rPr><w:vertAlign w:val="superscript"/></w:rPr>
</w:style>
```

3 つの登録のうち 1 つでも欠けると、Word が「ファイルが破損している」と報告したり、脚注が表示されなかったりする。
`ensure_*` 系の関数は、不足しているものだけを追加し、既存のエントリには手を付けない。

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
+- index.py       FastAPI アプリと HTTP 層
+- footnotes.py   docx 変換ロジック
+- .env           ローカル用の環境変数ファイル（git 管理外。空でよい）
+- pyproject.toml
+- uv.lock
+- README.md
```