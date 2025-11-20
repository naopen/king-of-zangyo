.PHONY: zip clean

zip:
	zip -r king-of-zangyo.zip manifest.json popup.html popup.js popup.css content.js icons/

clean:
	rm -f king-of-zangyo.zip
