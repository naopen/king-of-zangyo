# 使い方ページのモック（2026-10-29）

設計書 `docs/superpowers/specs/2026-10-29-usage-guide-and-notices-design.md` の5節のために、壁打ちで作ったモック。Naokiが「すごく良さそう」と評価した版。実装は `docs/superpowers/plans/2026-10-29-usage-guide-and-notices.md` の Task 2 で、これを元に `guide/` を作る。

| ファイル | 中身 |
|---|---|
| `index.html` | 使い方ページのモック。CSSとスクリプトを1枚に入れてある |
| `img/shot-*.png` | テスト用HTMLの実画面を2倍の解像度で切り抜いた画像。ボタン名は画面の上で「更新」に差し替えてある |
| `crop.html` | 上の切り抜き画像を作ったページ。リポジトリのルートに置いて配信する前提の参照になっている（`test/extension-test.html`、`popup.html`） |

`index.html` が参照するカグカグの画像（`img/new-pose10.png` など）は、ここには入れていない。公開リポジトリなので、GitHubから自由にダウンロードされないようにするため。開くとカグカグの部分だけが表示されない。画像の出所と入れ方は、計画の Task 3 にある。
