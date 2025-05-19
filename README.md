# Docxearch
Official name has changed a bit from original 'docsearch'. A bit of a word play with the fact that this application only searches through .docx file extensions.

Docxearch is a search tool to search for words within Word (docx) files. It looks for terms occurring within the same paragraph. It also has an option to search for words separately, meaning each word (separated by whitespace) is searched within a paragraph.

## Version Update
### Version 3.0
**Release AI feature**: Now can use ollama model to autogenerate an answer based on documents indexed and searched through FAISS. Using Korean-optimized models for ollama as well as sentence-transformer.
The new feature is usable in a new window which spawns on clicking "Ask AI".
### Version 2.4
Removed the previous feature and replaced with 'Duplicate' functionality which opens a new window while copying the search results. This way, you can create multiple running windows based on different results.
### Version 2.3
Added a feature to display last searched set of results upon rerun. This is to maintain history of results and also allow nested search by opening multiple windows.

## How to build and create a ditributable

### 1. Run pyinstaller
```
pyinstaller docxearch.spec
```
it should prompt (if you have already created a dist before) whether you will allow overwriting the previous dist. Reply 'Y' to do so.

### 2. Run Inno Setup
To create the install file, run Inno Setup and open `docxearch_installer2.iss`. Depending on the change, update the major or minor version, then press save and the play button. This will also create a new setup executable under `/Output` folder.

## How to run

### 1. Install all dependencies
```
pip install -r requirements.txt
```

### 2. If using AI feature, make sure to download Ollama and pull model
Install [Ollama](https://ollama.com/download)
```
ollama pull exaone3.5:2.4b
```

### 2. Run by executing app.py
```
python app.py
```