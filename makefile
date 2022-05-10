install:
	pip install -r requirements.txt

test:
	python3 -m pytest -vv -cov=go tests.py

lint:
	echo "pending implementation"