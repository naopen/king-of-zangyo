.PHONY: zip clean test check-assets

KAGKAG_PNGS = kagkag/new-pose10.png kagkag/new-pose21.png kagkag/new-pose15.png kagkag/alt_06.png kagkag/new-pose04.png

check-assets:
	@for f in $(KAGKAG_PNGS); do \
		if [ ! -f "$$f" ]; then echo "kagkag サブモジュールを取得してください（git submodule update --init）"; exit 1; fi; \
	done

zip: check-assets
	zip -r king-of-zangyo.zip manifest.json popup.html popup.js popup.css annual-pace.js content.js notices.js icons/ guide/ $(KAGKAG_PNGS)

test:
	node --test "tests/*.test.js"

clean:
	rm -f king-of-zangyo.zip
