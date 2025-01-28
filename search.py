import os
from PyQt5.QtWidgets import QFileDialog, QTableWidgetItem
from docx import Document

def search_word_docs(self):
    # Get search term from textbox
    search_term = self.textbox.text()

    # Open directory dialog to select directory to search
    dir_path = QFileDialog.getExistingDirectory(self, 'Open Directory', '/')

    # Search for Word documents in directory
    word_docs = []
    for root, dirs, files in os.walk(dir_path):
        for file in files:
            if file.endswith('.docx'):
                word_docs.append(os.path.join(root, file))

    # Search for search term in Word documents
    results = []
    for word_doc in word_docs:
        doc = Document(word_doc)
        for paragraph in doc.paragraphs:
            if search_term in paragraph.text:
                results.append((os.path.basename(word_doc), word_doc))
                break

    # Display results in table
    self.table.setRowCount(len(results))
    for i, (file_name, path) in enumerate(results):
        self.table.setItem(i, 0, QTableWidgetItem(file_name))
        self.table.setItem(i, 1, QTableWidgetItem(path))
