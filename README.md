# adoc-pptx-compiler

AsciiDoc 文書を、編集可能な PowerPoint ファイル（.pptx）にコンパイルするツールです。
姉妹ツールの md-pptx-injector が「テンプレートへの流し込み」なのに対し、こちらは
AsciiDoc の構造を解釈してスライドを描く「コンパイル」方式です。

現在のバージョン: **0.9.1.0**

## 特徴

- 出力は通常のシェイプ（テキストボックス・表・図形）のみで構成され、マスター／
  レイアウト／テーマに依存しません。生成後の .pptx を PowerPoint 等で直接編集できます。
- テキストを計測して改ページするため、内容がスライドの外に描かれて黙って
  消失することはありません。入り切らない節は自動で続きページになります。
- 対応要素: タイトルページ、段落（太字・斜体・等幅・リンク）、箇条書き
  （ネスト・順序付き）、表、画像、コードブロック、注意書き（NOTE 等）、脚注。
- 未対応の要素や見つからない画像は、レポートと標準出力で明示します。

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
| `--frontend {auto,asciidoctor,asciidoc-py}` | フロントエンドの選択（既定: auto） |
| `--overflow {continue,error}` | 節が1ページを超えたときの動作（既定: continue＝続きページ） |
| `--keep-ir FILE` | 中間表現（DocBook XML）を保存 |
| `--report FILE` | 検証レポート（JSON）を保存 |
| `--version` | バージョン表示 |

入力に DocBook XML（.xml）を直接渡すこともできます（フロントエンド不要）。

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
