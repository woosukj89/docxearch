import os
import faiss
import numpy as np
import pickle
from OpenAIHelper import OpenAIHelper
import tempfile
import shutil
from UniqueCounter import UniqueCounter
import asyncio
from AsyncOpenAIHelper import AsyncOpenAIHelper
from concurrent.futures import ThreadPoolExecutor
import threading

class FAISSMetadataIndex:
    def __init__(self):
        self.index = None
        self.embeddings = []
        self.metadata = []
        # self.encoder = SentenceTransformer(encoder, device='cuda')
        self.encoder = AsyncOpenAIHelper()
        self.default_index_file_name = "document"
        self.default_metadata_file_name = "document"
        self.normalize = False
        self.debug = False
        self._lock = threading.Lock()
        self._loop = asyncio.new_event_loop()
        self._executor = ThreadPoolExecutor()
        threading.Thread(target=self._loop.run_forever, daemon=True).start()

    def add_document(self, chunk):
        # self.chunks.append(chunk["text"])
        self.embeddings.append(self.encoder.encode(chunk["text"]))
        self.metadata.append(chunk["metadata"])
    
    def add_documents_bulk(self, chunks):
        with self._lock:
            if self.debug: print(f"Adding total chunks of {len(chunks)}")
            futures = [
                asyncio.run_coroutine_threadsafe(
                    self.encoder.encode(chunk["text"]), self._loop
                )
                for chunk in chunks
            ]
            for future, chunk in zip(futures, chunks):
                embedding = future.result()  # blocks until result
                self.embeddings.append(embedding)
                self.metadata.append(chunk["metadata"])

    def save(self, index_root, metadata_root):
        with self._lock:
            if self.debug: print("Save called")
            embedding = np.array(self.embeddings)
            index = faiss.IndexFlatIP(embedding.shape[1])
            # if self.normalize: faiss.normalize_L2(embedding)
            index.add(embedding)
            with tempfile.NamedTemporaryFile(delete=False, suffix=".faiss") as tmp:
                tmp_path = tmp.name
            if self.debug: print(f"Adding to index {tmp_path}")
            faiss.write_index(index, tmp_path) 
            file_num = UniqueCounter.next()
            shutil.move(tmp_path, os.path.join(index_root, self.default_index_file_name + str(file_num) + ".faiss"))
            with open(os.path.join(metadata_root, self.default_metadata_file_name + str(file_num) + ".metadata"), "wb") as f:
                pickle.dump(self.metadata, f)
            self.embeddings = []
            self.metadata = []

    def load(self, index_path, metadata_path):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".faiss") as tmp:
            tmp_path = tmp.name
        shutil.copy(index_path, tmp_path)
        index = faiss.read_index(tmp_path)
        with open(metadata_path, 'rb') as f:
            metadata = pickle.load(f)
        
        return index, metadata
