import os
import faiss
import numpy as np
import pickle
from sentence_transformers import SentenceTransformer

class FAISSMetadataIndex:
    def __init__(self, encoder="jhgan/ko-sroberta-multitask"):
        self.index = None
        self.chunks = []
        self.metadata = []
        self.encoder = SentenceTransformer(encoder)
        self.default_index_file_name = "document"
        self.default_metadata_file_name = "document"
        self.file_number = 1
        self.normalize = False
        self.debug = False

    def add_document(self, chunk):
        self.chunks.append(chunk["text"])
        self.metadata.append(chunk["metadata"])

    def save(self, index_root, metadata_root):
        if self.debug: 
            print("Indexing chunks\n\n" + "\n\n".join(self.chunks))
        embedding = self.encoder.encode(self.chunks)
        index = faiss.IndexFlatIP(embedding.shape[1])
        if self.normalize: faiss.normalize_L2(embedding)
        if self.debug: print("Adding to index")
        index.add(np.array(embedding))
        faiss.write_index(index, os.path.join(index_root, self.default_index_file_name + str(self.file_number) + ".faiss"))
        with open(os.path.join(metadata_root, self.default_metadata_file_name + str(self.file_number) + ".metadata"), "wb") as f:
            pickle.dump(self.metadata, f)
        self.chunks = []
        self.metadata = []
        self.file_number += 1

    def load(self, index_path, metadata_path):
        index = faiss.read_index(index_path)
        with open(metadata_path, 'rb') as f:
            metadata = pickle.load(f)
        
        return index, metadata
    
    def reset_file_number(self):
        self.file_number = 1
