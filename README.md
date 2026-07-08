# dl_commons

請求書自動DLシステム（seikyudl）関連の共通処理ライブラリ。

## パッケージ構成

```
dl_commons/
├── pyproject.toml
└── src/
    └── dl_commons/
        ├── __init__.py
        ├── core.py                  # Cmn_SpreadIO 等の共通処理本体
        └── seikyudl_config.ini      # Google認証情報・出力フォルダパス等の設定（外出し済み）
```

## セットアップ

```bash
cd dl_commons
pip install .
```

編集しながら使う場合：

```bash
pip install -e .
```

## 注意

- `seikyudl_config.ini` はこのパッケージと同じ階層（`src/dl_commons/`）に置く前提です。
  `core.py` 内の `_resolve_path()` が以下の順で探索します。
  1. `core.py` と同じディレクトリ
  2. 実行時のカレントディレクトリ
  3. カレントディレクトリの1つ上
- `seikyudl_config.ini` に実際のCLIENT_ID / CLIENT_SECRET / SPREADSHEET_KEYを入れた場合、
  そのファイルは `.gitignore` に追加し、Git管理から外してください。
  プレースホルダー版を `seikyudl_config.ini.sample` として別途コミットする運用を推奨します。
- インストール時は **名前だけを指定する `pip install dl_commons` ではなく**、
  必ずこのフォルダをローカルパス指定でインストールしてください
  （PyPI上の無関係な同名/類似名パッケージを誤って拾わないようにするため）。
