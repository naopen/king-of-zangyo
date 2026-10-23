.PHONY: zip clean test

zip:
	zip -r king-of-zangyo.zip manifest.json popup.html popup.js popup.css content.js icons/

test:
	node --test "tests/*.test.js"

clean:
	rm -f king-of-zangyo.zip
