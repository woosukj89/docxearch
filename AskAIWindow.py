import os
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QPushButton, QLineEdit, QProgressDialog, 
    QScrollArea, QHBoxLayout, QFrame, QMessageBox, QSizePolicy, QTextBrowser
)
from PyQt5.QtCore import Qt, QTimer, QObject, pyqtSignal, QThread
from PyQt5.QtGui import QTextDocument
from DocumentProcessor import DocumentProcessor
from IndexSearch import IndexSearch
from DocumentProcessor import DocumentProcessor
from OpenAIHelper import OpenAIHelper
import math

class FileProcessingWorker(QObject):
    progress = pyqtSignal(int)
    finished = pyqtSignal()
    error_occurred = pyqtSignal(str)

    def __init__(self, directory):
        super().__init__()
        self.directory = directory
        self.file_save_threshold = 50
        self.error_threshold = 5
        self.file_processed = 0
        self.processor = DocumentProcessor()

    def run(self):
        dir_path = self.directory
        if not self.__index_complete(dir_path):
            self.__remove_all_faiss_metadata(dir_path)
            self.__run(dir_path)
        # do it subsequently for all subpaths
        for name in os.listdir(dir_path):
            subpath = os.path.join(dir_path, name)
            if os.path.isdir(subpath) and not self.__index_complete(subpath):
                self.__remove_all_faiss_metadata(subpath)
                self.__run(subpath)
        
        self.finished.emit()
    
    def __index_complete(self, directory):
        docx_count = sum(1 for f in os.scandir(directory) if f.name.endswith(".docx") and f.is_file())
        faiss_count = sum(1 for f in os.scandir(directory) if f.name.endswith(".faiss") and f.is_file())
        metadata_count = sum(1 for f in os.scandir(directory) if f.name.endswith(".metadata") and f.is_file())

        return math.ceil(docx_count / self.file_save_threshold) <= faiss_count and faiss_count == metadata_count
    
    def __remove_all_faiss_metadata(self, directory):
        for entry in os.scandir(directory):
            if entry.is_file() and (entry.name.endswith(".faiss") or entry.name.endswith(".metadata")):
                os.remove(entry.path)
    
    def __run(self, dir_path):
        print(f"Running processor for path {dir_path}")
        file_count = 0
        error_count = 0
        for file in os.listdir(dir_path):
            file_path = os.path.join(dir_path, file)
            if os.path.isfile(file_path) and file.endswith(".docx"):
                try:
                    self.processor.process_file(file_path)
                    # if number of files > 50, save
                    file_count +=1
                    self.file_processed += 1
                    self.progress.emit(self.file_processed)
                    if file_count > self.file_save_threshold:
                        self.processor.save_progress(dir_path)
                        file_count = 0
                except Exception as e:
                    error_count += 1
                    if error_count > self.error_threshold:
                        self.error_occurred.emit(str(e))
                        self.finished.emit()
                        return
        if file_count:
            self.processor.save_progress(dir_path)

class AIWorker(QObject):
    finished = pyqtSignal(str, list)  # response text and search results
    error = pyqtSignal(str)

    def __init__(self, directory, query, searcher, processor, answerer):
        super().__init__()
        self.directory = directory
        self.query = query
        self.searcher = searcher
        self.processor = processor
        self.answerer = answerer

    def run(self):
        try:
            # Run search
            search_results = self.searcher.get_all_results(self.directory, self.query)
            context = "\n\n".join([
                f"[{i}]: \n\"\"\"{self.processor.get_paragraphs(result['path'], result['para_range'])}\n\"\"\"\n"
                f"{'Author: ' + result['author'] if result.get('author') else ''}"
                for i, result in enumerate(search_results)
            ])
            # response = ollama.generate(model='exaone3.5:2.4b', prompt=prompt)
            response = self.answerer.answer_with_context(self.query, context)
            self.finished.emit(response, search_results)
        except Exception as e:
            self.error.emit(str(e))

class AskAIWindow(QWidget):
    def __init__(self, directory, query):
        super().__init__()
        self.setStyleSheet("QWidget { font-size: 12pt; }")
        self.setWindowTitle("Ask AI")
        self.setGeometry(200, 200, 700, 600)
        self.directory = directory
        self.query = query
        self.layout = QVBoxLayout()
        self.setLayout(self.layout)

        # Conversation area
        self.chat_area = QVBoxLayout()
        self.chat_container = QWidget()
        self.chat_container.setLayout(self.chat_area)

        self.chat_scroll = QScrollArea()
        self.chat_scroll.setWidgetResizable(True)
        self.chat_scroll.setWidget(self.chat_container)
        self.layout.addWidget(self.chat_scroll)

        self.searcher = IndexSearch()
        self.processor = DocumentProcessor()
        self.answerer = OpenAIHelper()

        # Follow-up input bar (goes at the bottom)
        followup_layout = QHBoxLayout()
        self.followup_input = QLineEdit()
        self.followup_input.setPlaceholderText("같은 경로에 추가 질문이 있으면 물어보세요.")
        self.followup_input.returnPressed.connect(self.followup_query)
        self.followup_button = QPushButton("Ask AI")
        self.followup_button.clicked.connect(self.followup_query)

        followup_layout.addWidget(self.followup_input)
        followup_layout.addWidget(self.followup_button)

        self.followup_input.hide()
        self.followup_button.hide()

        self.layout.addLayout(followup_layout)

        QTimer.singleShot(100, self.initialize_ai)
    
    def initialize_ai(self):
        self.display_user_question(self.query)
        self.start_processing_in_thread()
    
    def count_docx_files(self, directory: str) -> int:
        count = 0
        for dirpath, _, filenames in os.walk(directory):
            count += sum(1 for f in filenames if f.lower().endswith('.docx'))
        return count
    
    def start_processing_in_thread(self):
        # Count files in a blocking way (quick), or move this to thread too if slow
        total_files = self.count_docx_files(self.directory)

        self.progress_dialog = QProgressDialog("폴더에서 문서를 처리하는 중입니다. 문서마다 20초 정도 소요됩니다. 이 작업은 폴더당 한 번만 수행합니다.", 
                                               None, 0, total_files, self)
        self.progress_dialog.setWindowModality(Qt.WindowModal)
        self.progress_dialog.setMinimumDuration(0)
        self.progress_dialog.show()

        # Set up background thread
        self.thread = QThread()
        self.worker = FileProcessingWorker(self.directory)
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(self.update_progress)
        self.worker.finished.connect(self.processing_done)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.worker.error_occurred.connect(self.show_error_message)

        self.thread.start()

    def update_progress(self, count):
        self.progress_dialog.setValue(count)
    
    def show_error_message(self, error):
        QMessageBox.critical(self, "오류", error)

    def processing_done(self):
        self.progress_dialog.close()
        self.ask_ai()
    
    def ask_ai(self):
        self.loading_label = QLabel("AI로부터 답변을 받고 있습니다")
        self.chat_area.addWidget(self.loading_label)
        self.ellipsis_state = 0

        self.loading_timer = QTimer()
        self.loading_timer.timeout.connect(self.animate_loading_text)
        self.loading_timer.start(500)

        QTimer.singleShot(100, self.run_ai_query)

    def animate_loading_text(self):
        self.ellipsis_state = (self.ellipsis_state + 1) % 4
        self.loading_label.setText("AI로부터 답변을 받고 있습니다" + "." * self.ellipsis_state)

    def run_ai_query(self):
        self.ai_thread = QThread()
        self.ai_worker = AIWorker(self.directory, self.query, self.searcher, self.processor, self.answerer)
        self.ai_worker.moveToThread(self.ai_thread)

        self.ai_thread.started.connect(self.ai_worker.run)
        self.ai_worker.finished.connect(self.handle_ai_response)
        self.ai_worker.finished.connect(self.ai_thread.quit)
        self.ai_worker.finished.connect(self.ai_worker.deleteLater)
        self.ai_thread.finished.connect(self.ai_thread.deleteLater)
        self.ai_worker.error.connect(self.show_error_message)

        self.ai_thread.start()
    
    def handle_ai_response(self, response_text, search_results):
        self.loading_timer.stop()
        self.loading_label.deleteLater()
        self.search_results = search_results
        self.display_answer(response_text)
    
    def display_user_question(self, question_text):
        container = QVBoxLayout()

        question_label = QLabel(f"<b>You asked:</b> {question_text}")
        question_label.setStyleSheet("font-size: 12pt;")

        container.addWidget(question_label)

        wrapper = QFrame()
        wrapper.setLayout(container)
        wrapper.setFrameShape(QFrame.Shape.Box)

        self.chat_area.addWidget(wrapper)
    
    def display_answer(self, response_text):
        # Only show follow-up input after first answer
        self.followup_input.show()
        self.followup_button.show()

        message_container = QVBoxLayout()
        
        answer_label = QLabel("<b>AI's answer:</b>")
        answer_label.setStyleSheet("font-size: 12pt;")
        message_container.addWidget(answer_label)

        # ① create preview doc just for sizing
        doc = QTextDocument()
        doc.setMarkdown(response_text)
        doc.setTextWidth(self.chat_scroll.viewport().width() - 40)
        doc_height = doc.size().height()

        answer_text = QTextBrowser(self)
        answer_text.setMarkdown(response_text)
        answer_text.setFrameShape(QFrame.NoFrame)
        answer_text.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        answer_text.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        answer_text.document().setTextWidth(self.chat_scroll.viewport().width() - 40)
        answer_text.setFixedHeight(int(doc_height) + 40)

        message_container.addWidget(answer_text)

        references_label = QLabel("<b>References:</b>")
        references_label.setStyleSheet("font-size: 12pt;")
        message_container.addWidget(references_label)

        for i, result in enumerate(self.search_results):
            frame = self.create_reference_frame(i, result['title'], self.processor.get_paragraphs(result['path'], result['para_range']))
            message_container.addWidget(frame)

        wrapper = QFrame()
        wrapper.setLayout(message_container)
        wrapper.setFrameShape(QFrame.Shape.Box)
        self.chat_area.addWidget(wrapper)

        # Ensure we scroll to the top of the AI's answer
        QTimer.singleShot(100, lambda: self.chat_scroll.ensureWidgetVisible(answer_text))
    
    def create_reference_frame(self, index, title, text):
        container = QVBoxLayout()
        row = QHBoxLayout()

        label = QLabel(f"[{index}] {title}")
        toggle_button = QPushButton("▼")
        toggle_button.setFixedWidth(30)
        toggle_button.setCheckable(True)

        row.addWidget(label)
        row.addStretch()
        row.addWidget(toggle_button)
        container.addLayout(row)

        # text_box = QTextEdit(text)
        # text_box.setReadOnly(True)
        # text_box.setVisible(False)

        text_box = QTextBrowser()
        text_box.setText(text)
        text_box.setVisible(False)
        text_box.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        text_box.setFrameShape(QFrame.NoFrame)
        text_box.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        # text_box.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        # Resize the height to fit the content
        doc = QTextDocument()
        doc.setPlainText(text)
        text_box.setMinimumHeight(int(doc.size().height()) + 10)

        container.addWidget(text_box)

        toggle_button.toggled.connect(lambda checked, tb=text_box: tb.setVisible(checked))

        frame = QFrame()
        frame.setLayout(container)
        frame.setFrameShape(QFrame.Shape.Box)
        return frame

    def followup_query(self):
        question = self.followup_input.text().strip()
        if not question:
            return

        self.query = question
        self.display_user_question(self.query)
        self.followup_input.clear()
        self.ask_ai()
        QTimer.singleShot(500, lambda: self.chat_scroll.verticalScrollBar().setValue(self.chat_scroll.verticalScrollBar().maximum()))
