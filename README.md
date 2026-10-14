# フォルダ比較自動化ツール

WinMerge を裏で呼び出して 2 つのフォルダを比較し、差分レポートと判定結果を自動出力する PowerShell スクリプトです。
設定ファイル（JSON）に複数のフォルダ組を定義しておけば、一括で比較できます。タスクスケジューラで定期実行することも可能です。

- 現行バージョン：**1.2.1**
- 対応マニュアル：`docs/フォルダ比較自動化ツール_運用マニュアル_v1.5.docx`

---

## 主な機能

| 機能 | 内容 |
|---|---|
| 複数フォルダ組の一括比較 | 設定ファイルの `jobs` に定義したフォルダ組を順番に比較します |
| WinMerge レポート出力 | WinMerge を最小化・非対話モードで起動し、HTML／CSV のレポートを出力します |
| ハッシュによる差分判定 | ファイルサイズと SHA256 で判定するため、WinMerge の表示言語やレポート書式に左右されません |
| ファイル・フォルダ単位のエラー記録 | 読めないファイルやフォルダがあっても処理を止めず、`Error`／`FolderError` として記録します |
| 全ファイル一覧の出力 | 比較した全ファイルについて、判定・サイズ・更新日時を CSV に出力します |
| 実行ログ | 処理経過とエラーを記録します。設定ファイルのエラー時もログを残します |
| 終了コード | `0`＝差分なし／`1`＝差分あり／`2`＝エラー |
| 運用補助 | タイムアウト、二重起動防止、古いレポートの自動削除、`-WhatIf` による事前確認 |

## 動作環境

| 項目 | 要件 |
|---|---|
| OS | Windows 11 |
| PowerShell | **Windows PowerShell 5.1（推奨）**。PowerShell 7 でも動作しますが、出力 CSV が BOM なしになり、Excel で開くと文字化けします |
| WinMerge | 2.16 以降（`/noninteractive` オプションに対応した版） |
| 権限 | 比較対象フォルダの読み取り権限と、出力先およびスクリプトを置くフォルダの書き込み権限 |

## リポジトリ構成

```
.
├─ Invoke-FolderCompare.ps1        … スクリプト本体（ASCII のみで記述）
├─ compare-config.json      … 設定ファイルのサンプル
├─ docs/
│   └─ フォルダ比較自動化ツール_運用マニュアル_v1.5.docx
├─ .gitignore
├─ .gitattributes
└─ README.md
```

> **Note**
> 実運用の `compare-config.json` には、個人のフォルダパス（`C:\Users\<ユーザー名>\Box\…` など）が含まれます。リポジトリにはコミットせず、`compare-config.sample.json` をコピーして各自で作成してください。

## セットアップ

1. **ファイルを配置する**
   スクリプトと設定ファイルを、書き込みできるローカルフォルダ（例：`C:\tools`）に置きます。
   `C:\Program Files` の配下や読み取り専用の共有フォルダに置くと、設定エラー時のログが残りません。

2. **ブロックを解除する**（Web やメールから取得した場合）
   ```bat
   powershell.exe -NoProfile -Command "Unblock-File -LiteralPath C:\tools\Invoke-FolderCompare.ps1"
   ```

3. **設定ファイルを作成する**
   ```bat
   copy compare-config.sample.json C:\tools\compare-config.json
   ```
   作成したファイルを開き、`outputRoot` と `jobs` を環境に合わせて編集します。

4. **事前確認を行う**（WinMerge は起動しません）
   ```bat
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\tools\Invoke-FolderCompare.ps1 -WhatIf
   ```

## 使い方

```bat
powershell.exe -NoProfile -ExecutionPolicy Bypass -File <スクリプトのパス> [-ConfigPath <設定ファイル>] [-VerboseFileLog] [-WhatIf]
```

| 引数 | 既定値 | 説明 |
|---|---|---|
| `-ConfigPath` | スクリプトと同じフォルダの `compare-config.json` | 使用する設定ファイルのパス |
| `-VerboseFileLog` | オフ | ファイル 1 件ごとの判定結果をログに出力します（調査用） |
| `-WhatIf` | オフ | 事前確認モードです。WinMerge の起動と古いレポートの削除は行いません。ハッシュ判定・一覧・サマリ・ログは出力します |

### 実行例

```bat
rem 標準の実行
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\tools\Invoke-FolderCompare.ps1

rem 設定ファイルを指定して実行
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\tools\Invoke-FolderCompare.ps1 -ConfigPath C:\tools\config_basic-design.json

rem ファイルごとのログを出力して実行
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\tools\Invoke-FolderCompare.ps1 -VerboseFileLog
```

### 終了コード

| コード | 意味 |
|---|---|
| `0` | 全ジョブで差分なし |
| `1` | 差分あり（エラーはなし） |
| `2` | エラーあり。設定の誤り、フォルダが存在しない、WinMerge の失敗、読めないファイル・フォルダがある、`jobs` が空、など |

> **Warning**
> `hashCheck: false` の場合は差分判定を行わないため、差分があっても終了コードは `0` になります。

### タスクスケジューラで定期実行する

| 設定欄 | 入力値 |
|---|---|
| プログラム/スクリプト | `powershell.exe` |
| 引数の追加 | `-NoProfile -ExecutionPolicy Bypass -File "C:\tools\Invoke-FolderCompare.ps1"` |
| 開始（オプション） | `C:\tools` |

- WinMerge は画面を持つアプリのため、［ユーザーがログオンしているときのみ実行する］を選択してください。
- ［前回の実行結果］が `(0x1)` の場合は「差分あり」の意味です。エラーではありません。
- 二重起動防止は PC 単位で働きます。設定ファイルを分けて複数のタスクを登録する場合は、開始時刻をずらしてください。

## 設定ファイル（compare-config.json）

### サンプル

```json
{
  "winMergePath": "C:/Program Files/WinMerge/WinMergeU.exe",
  "outputRoot": "C:/tools/reports",
  "reportFormats": ["html", "csv"],
  "retentionDays": 30,
  "timeoutSec": 1800,
  "hashCheck": true,
  "jobs": [
    {
      "name": "基本設計書_v1_vs_v2",
      "left": "C:/work/docs/v1/基本設計書",
      "right": "C:/work/docs/v2/基本設計書",
      "leftDesc": "旧版(v1)",
      "rightDesc": "新版(v2)",
      "recurse": true,
      "filter": "*.docx;*.xlsx;*.pptx;*.md",
      "include": ["*.docx", "*.xlsx", "*.pptx", "*.md"],
      "exclude": ["~$*", "*.tmp", "Thumbs.db", "desktop.ini", ".git\\*"]
    }
  ]
}
```

### 全体設定

| 項目 | 必須 | 既定値 | 説明 |
|---|---|---|---|
| `winMergePath` | 任意 | 自動検出 | WinMergeU.exe のフルパス |
| `outputRoot` | 任意 | `<スクリプトのフォルダ>\reports` | 出力先。実行ごとに日時フォルダを作成します |
| `reportFormats` | 任意 | `["html", "csv"]` | WinMerge レポートの形式（html／csv／tsv／xml） |
| `retentionDays` | 任意 | `30` | 出力フォルダの保管日数。`0` を指定すると削除しません |
| `timeoutSec` | 任意 | `1800` | WinMerge 1 回あたりの最大待ち時間（秒） |
| `hashCheck` | 任意 | `true` | ハッシュによる差分判定を行うかどうか |
| `jobs` | **必須（1 件以上）** | － | 比較するフォルダ組の一覧 |

### jobs の項目

| 項目 | 必須 | 既定値 | 説明 |
|---|---|---|---|
| `name` | 必須 | － | ジョブ名。出力ファイル名に使われるため、**重複させないでください** |
| `left` / `right` | 必須 | － | 比較元／比較先のフォルダ |
| `leftDesc` / `rightDesc` | 任意 | パス | レポートに表示する名称 |
| `recurse` | 任意 | `true` | サブフォルダも比較するかどうか |
| `filter` | 任意 | 全ファイル | WinMerge に渡すフィルタ（`;` 区切り）。WinMerge レポートにのみ適用されます |
| `include` | 任意 | 全ファイル | ハッシュ判定の対象とするファイル名パターン |
| `exclude` | 任意 | なし | ハッシュ判定から除外するパターン。ファイル名と相対パスの両方で照合します |

> **Important**
> JSON では `\` を 1 つだけ書くとエラーになります。パスは `/` 区切り（推奨）か、`\\` と 2 つ重ねて書いてください。

## 出力ファイル

```
<outputRoot>\
 └─ yyyyMMdd_HHmmss\
     ├─ compare.log              … 実行ログ
     ├─ summary.csv              … 全ジョブの結果サマリ
     ├─ <ジョブ名>.html          … WinMerge レポート
     ├─ <ジョブ名>.csv           … WinMerge レポート
     └─ <ジョブ名>_files.csv     … 全ファイルの判定一覧
```

- 設定ファイルを読めない場合や、`outputRoot` のフォルダを作成できない場合、ログは `<スクリプトのフォルダ>\reports\yyyyMMdd_HHmmss\compare.log` に出力されます。
- `FATAL` で処理全体が停止した場合、出力されるのは `compare.log` のみです。

### summary.csv の Result

| 値 | 意味 |
|---|---|
| `SAME` | 全ファイルが同一 |
| `DIFF` | 相違・左のみ・右のみのファイルがある |
| `PARTIAL_ERROR` | 比較はできたが、読めないファイルまたはフォルダがある |
| `ERROR` | ジョブ全体が失敗した（フォルダが存在しないなど） |
| `REPORT_ONLY` | `hashCheck: false` のため判定していない |

### _files.csv の Status

| 値 | 意味 |
|---|---|
| `Same` | 同一（サイズと SHA256 が一致） |
| `Different` | 内容が異なる |
| `LeftOnly` / `RightOnly` | 片側にのみ存在する |
| `Error` | ファイルを読めなかった（理由は Message 列） |
| `FolderError` | フォルダを読めなかった。配下のファイルは比較されていません |

## 結果の確認手順

1. 終了コード、または `compare.log` の最終行（`Folder compare END`）を確認します。
2. `summary.csv` で、Result が `SAME` 以外の行と、Message が空でない行を確認します。
3. 該当ジョブの `_files.csv` を開き、Status が `Same` 以外のファイルを特定します。
4. 必要に応じて、HTML レポートや WinMerge で中身を確認します。

## よくあるエラー

| 症状 | 対処 |
|---|---|
| `FATAL: 認識できないエスケープ シーケンスです` | JSON のパスに `\` が 1 つだけの箇所があります。`/` 区切りに直してください |
| `FATAL: No jobs defined in config file` | `jobs` に 1 件以上のジョブを記述してください |
| `FATAL: Another compare process is already running.` | 別の実行の終了を待ってください。プロセスが残っている場合は終了させます |
| `Cannot read folder` / `FolderError` | アクセス権、またはパス長（260 文字）を確認してください |
| `File ERROR: …別のプロセスで使用されている…` | Excel などで開いているファイルを閉じてください |
| Box のファイルで `File ERROR` | フォルダを［オフラインで利用可能にする］に設定してください |
| Office ファイルが `Different` になる | 保存し直しただけでも内部情報が変わります。WinMerge で内容を確認してください |

詳しくは運用マニュアルの 9 章「トラブルシューティング」を参照してください。

## 注意事項

- スクリプト本体は英数字（ASCII）のみで記述しています。日本語のコメントを追記すると、Windows PowerShell 5.1 で文字化けによる構文エラーになります。
- 比較は相対パスが同じファイル同士で行います。ファイル名を変更したファイルは、LeftOnly と RightOnly の 2 件として表示されます。
- 隠しファイル（`Thumbs.db`、`desktop.ini` など）も比較対象になるため、必要に応じて `exclude` で除外してください。
- `filter` には除外条件を指定できません。そのため WinMerge レポートと `_files.csv` の件数が一致しないことがあります。件数は `_files.csv` を正としてください。

## 推奨する .gitignore / .gitattributes

```gitignore
# 出力・実運用設定はコミットしない
reports/
compare-config.json
*.log
```

```gitattributes
# Windows PowerShell 向けに改行コードを CRLF で固定
*.ps1  text eol=crlf
*.json text eol=crlf
*.bat  text eol=crlf
*.docx binary
```

## 変更履歴

| バージョン | 日付 | 内容 |
|---|---|---|
| 1.2.1 | 2026-10-09 | 設定ファイルのエラー時もログを記録するよう変更。`-WhatIf` 時に古いレポートを「削除した」と記録しないよう修正。`jobs` が空の場合をエラーに変更 |
| 1.2.0 | 2026-10-09 | 残ったロックからの回復に対応。読めないフォルダを `FolderError`（PARTIAL_ERROR／終了コード 2）として扱うよう変更。ログに版数を出力 |
| 1.1.0 | 2026-10-08 | ハッシュ計算を自前実装に変更。ファイル単位のエラー記録と全ファイル一覧 CSV の出力に対応 |

## ライセンス

本ソフトウェアは MIT License のもとで公開しています。
著作権表示とライセンス文を残せば、商用・非商用を問わず、自由に利用・改変・再配布できます。
本ソフトウェアは無保証です。利用によって生じた損害について、作者は責任を負いません。

## 問い合わせ

ohta@ios-net.co.jp
