# final-fix report
1. content.js: toggleOvertimeDisplay / toggleLegendDisplay の末尾で renderNotices().catch(...) を呼ぶ。変更点の吹き出しは #king-of-zangyo-annual-section 自身に append し、座標はセクション基準（clientLeft/Top, scroll 込み）で再計算。
2. tests/guide.test.js: skip 判定を kagkag/new-pose10.png の存在確認に変更。
3. Makefile zip: -x '*.DS_Store' を追加。
Verification: make test 66 pass / 0 fail. zip の DS_Store 0件.
Browser (change bubble, 1500x800): init 1個 top 595 >= legend bottom 587 / overtime OFF 0個 / ON 1個 595 / legend OFF 1個 top 575 >= table bottom 567 / legend ON 1個 595 >= 587. Usage bubble (sync {}) 1個 + dot 1個で従来どおり。
