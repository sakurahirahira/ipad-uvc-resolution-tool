# iPad UVC 解像度設定ツール

## 目的

iPadをUVC（USB Video Class）対応キャプチャデバイス経由で **Windows PCの外部ディスプレイとして使う** 際に、iPadの画面にぴったり合う解像度を簡単に設定するためのGUIツール。

iPadのアスペクト比は4:3系（正確にはモデルにより微妙に異なる）だが、Windowsの標準設定では16:9系の解像度しか選べないことが多く、UVC接続したiPadの表示が引き伸ばされたり黒帯が出る問題を解決する。

## 主な機能

- **ディスプレイ検出**: 接続中の全ディスプレイを一覧表示し、UVC（非プライマリ）ディスプレイを自動選択
- **iPadプリセット**: 各iPadモデルに最適な解像度をワンクリックで適用
  - iPad Pro 13" (M4): 2752x2064 / 半分 1376x1032
  - iPad Pro 12.9" / Air 13": 2732x2048 / 半分 1366x1024
  - iPad 4:3汎用: 2048x1536 / 1024x768
  - XGA 4:3: 1280x960
  - SXGA 4:3: 1400x1050
- **カスタム解像度**: 任意の幅x高さを指定して適用
- **最寄り解像度検索**: 指定した解像度に最も近い、デバイスがサポートする解像度を自動で見つけて提案
- **サポート解像度一覧**: UVCデバイスが報告する全解像度を4:3に近い順にソート表示（4:3に近いものは緑色でハイライト）
- **ダブルクリック適用**: 一覧からダブルクリックで解像度を即適用

## ファイル構成

| ファイル | 内容 |
|---|---|
| `ipad_resolution_tool.py` | Python版（tkinter GUI）。メインの実装 |
| `ipad_resolution_tool.ps1` | PowerShell版（WinForms GUI）。Python不要で動作する代替版 |
| `iPad解像度設定ツール.spec` | PyInstallerビルド設定。Python版をexe化するための定義 |
| `build/` `dist/` | PyInstallerのビルド出力 |

## 技術的なポイント

- Windows API（`EnumDisplayDevicesW`, `EnumDisplaySettingsW`, `ChangeDisplaySettingsExW`）を直接呼び出してディスプレイの解像度を変更
- 変更前に`CDS_TEST`フラグでテスト実行し、サポートされない解像度の場合は最も近い代替を提案
- Python版とPowerShell版の2実装を用意し、環境に応じて使い分け可能
