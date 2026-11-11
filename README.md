# adoc-pptx-compiler

AsciiDoc 文書を、編集可能な PowerPoint ファイル（.pptx）にコンパイルするツールです。
姉妹ツールの md-pptx-injector が「テンプレートへの流し込み」なのに対し、こちらは
AsciiDoc の構造を解釈してスライドを描く「コンパイル」方式です。

現在のバージョン: **0.9.3.1**

## 特徴

- 出力は通常のシェイプ（テキストボックス・表・図形）のみで構成され、マスター／
  レイアウト／テーマに依存しません。生成後の .pptx を PowerPoint 等で直接編集できます。
- テキストを計測して改ページするため、内容がスライドの外に描かれて黙って
  消失することはありません。入り切らない節は自動で続きページになります。
- 対応要素: タイトルページ、段落（太字・斜体・等幅・リンク）、箇条書き
  （ネスト・順序付き）、表、画像、コードブロック、注意書き（NOTE 等）、脚注。
- 未対応の要素や見つからない画像は、レポートと標準出力で明示します。
- `doctype: book` の文書は縦配置、それ以外は横配置が既定です（`--orientation` で変更可）。
- 目次ページを自動生成できます（後述）。

## 仕組み

```
AsciiDoc ──(Asciidoctor)──▶ DocBook XML ──▶ 内部モデル ──▶ 写像・レイアウト ──▶ .pptx
```

設計の詳細は [docs/DESIGN.md](docs/DESIGN.md) を参照してください。

## 必要なもの

- Python 3.11 以上
- python-pptx、Pillow（`pip install -r requirements.txt`）
- AsciiDoc のフロントエンド（どちらか）
  - **Asciidoctor**（推奨。asciidoctor-pdf と同じ実行環境で動きます）
  - asciidoc-py（`pip install asciidoc`）

## 使い方

```
python adoc-pptx-compiler.py samples/sample.adoc sample.pptx
```

| オプション | 説明 |
|---|---|
| `--design FILE` | デザイン設定（TOML）で色・フォント・サイズを上書き |
| `--orientation {auto,landscape,portrait}` | スライドの向き（既定: auto＝book は縦、それ以外は横） |
| `--toc {auto,always,never}` | 目次ページ（既定: auto＝`:toc:` があるか book のとき） |
| `--body-size PT` | 本文のフォントサイズ（デザイン設定より優先） |
| `--line-chars N` | 本文1行の文字数（全角換算）で本文の柱の幅を決める |
| `--frontend {auto,asciidoctor,asciidoc-py}` | フロントエンドの選択（既定: auto） |
| `--overflow {continue,error}` | 節が1ページを超えたときの動作（既定: continue＝続きページ） |
| `--keep-ir FILE` | 中間表現（DocBook XML）を保存 |
| `--report FILE` | 検証レポート（JSON）を保存 |
| `--version` | バージョン表示 |

入力に DocBook XML（.xml）を直接渡すこともできます（フロントエンド不要）。

## 目次

原稿の先頭に `:toc:` がある場合（book の場合は常に）、タイトルページの直後に
目次ページが入り、節と小見出しの一覧にスライド番号が付きます。
見出しは `:toc-title:` 属性（既定は「目次」）です。

## 見た目の設定

`--design` に TOML ファイルを渡すと、色・フォント・文字サイズ・ページ寸法を
変更できます。書けるキーは [samples/design.toml](samples/design.toml) が
一覧を兼ねています（同梱の値は既定値そのもの）。未知のキーは無視され、
レポートの `design_unknown_keys` に記録されます。

## 運用での調整

変換は決定論的なので、見え方の微調整は「数値を変えて再変換」で行えます。
人が触るのは次の2つで十分なことが多いです。

- 文字が大きい／小さい → `--body-size`（例: `--body-size 16`）
- 1行を短くして余白を作りたい → `--line-chars`（例: `--line-chars 22`。
  本文の柱が「22文字 × 本文サイズ」の幅になります。表や画像は対象外）

恒常的に使う値は `--design` の TOML（`[sizes] body`、`[page] line_chars`）に
書いておけます。環境によってフォントが差し替わり行の高さがずれる場合は、
`[page] line_factor`（既定 1.2）を少し上げると重なり方向のずれを吸収できます。

## パスの解決規則

- 入力ファイル: カレントフォルダ → 原稿のあるフォルダ → **プログラムのある
  フォルダ** の順に探します。
- 画像などの素材: 原稿のあるフォルダを基準に探します。
- 出力ファイル: カレントフォルダ基準です。
- 「プログラムのあるフォルダ」とは、スクリプト実行時は .py のあるフォルダ、
  **exe 化したときは exe のあるフォルダ**です（`app_base_dir()` が一元管理）。

## バージョン番号の規則

4つの数字 `A.B.C.D` で管理します。初期値は `0.9.0.0` です。

- バグ修正: 末端（D）を +1
- 機能追加: 下から2番目（C）を +1

## テスト

```
pip install pytest
python -m pytest tests -q
```

## ライセンス

MIT License
