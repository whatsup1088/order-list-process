SHELL = /bin/bash
all:
	python -m venv .venv
	source .venv/bin/activate
	pip install -r requirements.txt
build:
	pyinstaller -wF app.py -i img/icon.ico
