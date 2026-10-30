# Task 3 報告

## やったこと
- 画像5枚と README.md を scratchpad で非公開リポジトリ用に用意（コミット cb96018）。gh repo create は自動モードの分類器に拒否され、Naoki が自分で作成・pushした。コントローラー確認: naopen/king-of-zangyo-kagkag は PRIVATE、main。
- メインで `git submodule add https://github.com/naopen/king-of-zangyo-kagkag.git kagkag`。kagkag/ に README.md と5枚。`git -C test status -sb` は `## main...origin/main` のまま。
- Makefile に check-assets（5枚のどれかが無ければ「kagkag サブモジュールを取得してください（git submodule update --init）」で失敗）を足し、zip の前提にした。zip に notices.js、guide/、kagkag の5枚を追加（README.md と .git は入らない）。

## 検証
- new-pose10.png を scratchpad に退避して `make clean && make zip`: メッセージを出して exit 2、zip は作られない。退避した画像は戻した。
- `make clean && make zip && unzip -l`: 25ファイル。notices.js、guide/ 一式（css・html・js・img 4枚）、kagkag の5枚が入り、README.md と .git は無い。
- `make test`: 66件、pass 66、fail 0、skipped 0（guide.test.js の kagkag 参照も解決）。

## コミット
- メイン: 「カグカグの画像を非公開のサブモジュールにし、make zip に入れた」（.gitmodules kagkag Makefile）
- 新リポジトリ: cb96018 カグカグの画像5枚と出所を記録した
- pushはしていない。

## 懸念
- README に「alt_06 以外の4枚はページの添付から直接落とした」と書いた。事実と違えば直す必要あり（新リポジトリ側）。
- yomiyasu lint は README の excess_list WARN のみ（表・一覧のため残した）。
