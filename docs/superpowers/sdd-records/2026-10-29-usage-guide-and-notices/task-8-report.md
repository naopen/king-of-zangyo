# Task 8 report

Status: DONE. Commit 2825a21 (branch feature/usage-guide-and-notices).

## What
- manifest.json version 2.6.0
- screenshot-compose.html: removed applyPlannedUi (+call) and the injected 「使い方を見る」 link (real popup.html has it, kept visible)
- guide-shots.html: removed the 「更新」 label replacement and its header comment
- STORE_LISTING.txt: v2.6.0 line added at top of JA and EN version histories (draft wording as given)
- Re-shot store-assets/screenshot.png (1280x800) and guide/img/shot-annual / annual-unfetched / popup. shot-monthly came out byte-identical, so unchanged.

## Extra change not in the brief
The real extension now draws notice bubbles (usage dot+bubble, change bubble). Headless Chrome starts clean, so the first shoot showed a change-bubble sliver under the annual legend and a usage bubble on the unfetched shot. Both export pages now inject `.kot-zangyo-notice-bubble,.kot-zangyo-notice-dot{display:none!important}` before measuring crop boxes. Images show only the real update button, entry pill and popup link.

## Image comparison
- store screenshot vs drafts/screenshot-off-on.png: same 1280x800, 108,319 differing pixels. Cause: the real popup is taller (848/2 vs 774/2 css px, the link now has real spacing), so the corner box bottom moved from y~395 to ~405 and OFF/ON shifted down ~10px. Layout otherwise the same; update button, 「?使い方」 pill and popup link look the same. I viewed the image; no stray artifacts.
- guide shot-annual vs old: same size, 2,626 px differ (real rendering of button/pill). shot-annual-unfetched 637->806 wide, shot-popup 774->848 tall (pill and link now real). guide/index.html fixes width via style on unfetched (320px) and annual (500px), so display widths hold.

## Verification
- make test: 66 tests, 66 pass, 0 fail
- make clean && make zip OK; zip manifest.json has "version": "2.6.0" (zip is untracked/ignored, not committed)
- Server (pid 71329) killed. Nothing deleted, no pip install.

## Concerns
- Guide page references shot-annual-unfetched with width 320px; aspect changed slightly (wider with pill) but only scales.
- Popup shot in guide is now taller (includes 「使い方を見る」); fine for the guide but worth a glance.
