import os
from docx import Document
from docx.text.paragraph import Paragraph
import re
from FAISSMetadataIndex import FAISSMetadataIndex

class DocumentProcessor:
    def __init__(self, min_words=500, overlap=50, file_save_threshold=50):
        self.min_words = min_words
        self.overlap = overlap
        self.debug = False
        self.file_save_threshold = file_save_threshold
        self.processed_files_count = 0
        self.process_finished = False
        self.indexer = FAISSMetadataIndex()
        self.file_count = 0
    
    def process_file(self, file_path):
        file = os.path.basename(file_path)
        if not file.endswith('.docx'):
            raise Exception(".docx file expected.")
        if self.debug: print(f"Processing file: {file_path}")
        try:
            doc = Document(file_path)
            title, author = self.try_to_get_title_author(file)
            body = doc._body._body
            ps = body.xpath(".//w:p")
            word_count = 0
            text = []
            pre_overlap_text = []
            post_overlap_text = []
            start_index = 0
            overlap_index = 0
            for i, p in enumerate(ps):
                paragraph = Paragraph(p, None)
                words = [w for w in paragraph.text.strip().split() if w]
                word_count += len(words)
                if self.debug: print(f"Adding {len(words)} words. New word count: {word_count}")
                if word_count >= self.min_words + self.overlap:
                    # save text
                    self.indexer.add_document({
                        "text": "\n    ".join(pre_overlap_text + text + post_overlap_text),
                        "metadata": {
                            "path": file_path,
                            "para_range": (start_index, i),
                            "author": author,
                            "title": title
                        }
                    })
                    # if self.debug: print(f"Adding chunk\n {pre_overlap_text + text + post_overlap_text}\n")
                    # change post to pre and reset
                    pre_overlap_text = post_overlap_text
                    post_overlap_text = []
                    text = []
                    word_count = 0
                    start_index = i - overlap_index
                    overlap_index = 0
                elif word_count >= self.min_words:
                    if paragraph.text.strip(): post_overlap_text.append(paragraph.text)
                    overlap_index += 1
                else:
                    if paragraph.text.strip(): text.append(paragraph.text)
            
            # add any remaining
            if len(text):
                self.indexer.add_document({
                    "text": "\n    ".join(pre_overlap_text + text + post_overlap_text),
                    "metadata": {
                        "path": file_path,
                        "para_range": (start_index, i),
                        "author": author,
                        "title": title
                    }
                })
            
            self.file_count += 1
            self.processed_files_count += 1

        except Exception as e:
            print(f"Error while processing document: {file}", e)
    
    def save_progress(self, directory):
        self.indexer.save(directory, directory)
    
    def get_processed(self):
        # if self.debug: print("get_processed")
        return self.processed_files_count, self.process_finished
        
    def try_to_get_title_author(self, file_name):
        match = re.match(r"^(.*)-(.*).docx$", file_name)
        if match:
            return match.group(1), match.group(2)
        else:
            return file_name, None
    
    def get_paragraphs(self, document_path, para_range):
        para_from, para_to = para_range
        doc = Document(document_path)
        body = doc._body._body
        ps = body.xpath(".//w:p")[para_from:para_to+1]
        return "\n    ".join([Paragraph(p, None).text for p in ps if Paragraph(p, None).text.strip()])
