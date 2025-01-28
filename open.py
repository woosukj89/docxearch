from PyQt5.QtGui import QTextCursor
from PyQt5.QtCore import Qt
from docx import Document

def open_doc(self, row, col):
    # Get file path from table
    file_path = self.table.item(row, 1).text()

    # Open Word document with search term highlighted
    doc = Document(file_path)
    for paragraph in doc.paragraphs:
        if self.search_term in paragraph.text:
            cursor = self.textedit.textCursor()
            cursor.setPosition(paragraph.runs[0].text.find(self.search_term))
            cursor.movePosition(QTextCursor.Right, mode=QTextCursor.KeepAnchor, n=len(self.search_term))
            self.textedit.setTextCursor(cursor)
            self.textedit.setFocus(Qt.OtherFocusReason)
            break
