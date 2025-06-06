# PlaceCraft

![Version](https://img.shields.io/badge/version-1.0.0-blue)
![License](https://img.shields.io/badge/license-MIT-green)
![Platform](https://img.shields.io/badge/platform-InDesign%20CC%202020%2B-orange)
![Status](https://img.shields.io/badge/status-active-brightgreen)

InDesign用オブジェクト自動配置スクリプト。指定フォルダ内の画像やPDFなどをアートボードにグリッドまたはランダムに配置するツール。

## 主な機能

- 指定フォルダから画像やPDFなどのオブジェクトを一括配置
- グリッド配置またはランダム配置の選択
- ファイルタイプフィルタリング（JPEG、PNG、TIFF、PDFなど）
- 原寸配置またはアートボードフィットの選択
- リアル挿入（埋め込み）またはリンク配置
- 配置順の指定（ファイル名、サイズ、作成日、ランダム）
- 配置数制限による処理の最適化
- ランダム配置時の回転・スケール指定

## インストール方法

1. このリポジトリをクローンまたはZIPファイルとしてダウンロード
2. `PlaceCraft.jsx`ファイルをInDesignスクリプトフォルダにコピー:

**Mac**: 
```
/Users/[ユーザー名]/Library/Preferences/Adobe InDesign/[バージョン]/[言語]/Scripts/Scripts Panel/
```

Macの「ライブラリ」フォルダへのアクセス方法:
1. Finderを開く
2. 「移動」メニューをクリックする
3. Optionキー（⌥）を押しながらメニューを表示すると、「ライブラリ」が表示される
4. その後、Preferences → Adobe InDesign → [バージョン] → [言語] → Scripts → Scripts Panel へ進む

**Windows**: 
```
C:\Users\[ユーザー名]\AppData\Roaming\Adobe\InDesign\[バージョン]\[言語]\Scripts\Scripts Panel\
```

## 使用方法

1. InDesignを起動
2. ウィンドウ > ユーティリティ > スクリプトを選択
3. スクリプトパネルで`PlaceCraft.jsx`をダブルクリック
4. ダイアログが表示されたら以下の設定を行う:
   - フォルダ選択ボタンをクリックして対象フォルダを選択
   - 必要な設定（ファイルタイプ、配置スタイルなど）を指定
   - 必要に応じてプレビューボタンで配置イメージを確認
5. 「OK」をクリックして配置を実行

## オプション設定

- **フォルダ選択**: オブジェクトが格納されたフォルダを指定
- **ファイルタイプ**:
  - 全て: InDesign対応形式（JPEG, PNG, TIFF, PDFなど）
  - 特定タイプ: ドロップダウンで選択（例：JPEGのみ）
- **配置スタイル**:
  - グリッド: 整然と配置（列数1～10、余白0～100mm）
  - ランダム: ランダム配置（回転±0～180°、スケール50～200%）
- **配置順**: ファイル名、ファイルサイズ、作成日、ランダムから選択
- **配置サイズ**: 原寸またはアートボードにフィット
- **配置オブジェクト数**: 0（無制限）～9999の整数で指定
- **配置方法**: リアル挿入（埋め込み）またはリンク

## 対応ファイル形式

- JPEG (`.jpg`, `.jpeg`)
- PNG (`.png`)
- TIFF (`.tif`, `.tiff`)
- PDF (`.pdf`)
- Photoshop (`.psd`)
- Illustrator (`.ai`)
- EPS (`.eps`)

## 動作環境

- Adobe InDesign CC 2020以降
- Windows/Mac両対応

## バージョン履歴

### 1.0.0 (2025-05-30)
- 初回リリース
- 基本的なオブジェクト配置機能
- グリッド・ランダム配置スタイル
- 配置サイズ・配置方法オプション

## クレジット

開発: rinmon

## ライセンス

MIT License

Copyright (c) 2025 rinmon

本ソフトウェアおよび関連文書のファイルは自由に使用、コピー、変更、配布が可能です。商用利用の場合は作者にご連絡ください。
