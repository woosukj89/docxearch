import os
import numpy as np
from FAISSMetadataIndex import FAISSMetadataIndex
from OpenAIHelper import OpenAIHelper

class IndexSearch:
    def __init__(self):
        self.debug = False
        self.indexer = FAISSMetadataIndex()
        self.model = OpenAIHelper()

    def get_query_embedding(self, query):
        return self.model.encode(query)

    def get_all_results(self, directory, query, k=10):
        query_vector = self.get_query_embedding(query)
        results = []
        for root, _, files in os.walk(directory):
            for file in files:
                if file.endswith('.faiss'):
                    if self.debug: print(f"Found .faiss file: {file}")
                    base_name = os.path.splitext(os.path.basename(file))[0]
                    metadata_filename = base_name + ".metadata"
                    if not os.path.isfile(os.path.join(root, metadata_filename)):
                        raise Exception(f"metadata file {metadata_filename} does not exist.")
                    index, metadata = self.indexer.load(os.path.join(root, file), os.path.join(root, metadata_filename))
                    distances, indices = index.search(np.array([query_vector]), k)
                    for i, idx in enumerate(indices[0]):
                        meta = metadata[idx]
                        results.append({
                            "id": int(idx),
                            "path": meta["path"],
                            "para_range": meta["para_range"],
                            "author": meta["author"],
                            "title": meta["title"],
                            "score": float(distances[0][i])
                        })
        
        if self.debug: print(f"Returning results: {results}")
        return sorted(results, key=lambda x: x['score'], reverse=True)[:10]
    
    def check_index_exists(self, directory):
        for file in os.listdir(directory):
            file_path = os.path.join(directory, file)
            if os.path.isfile(file_path) and file.endswith(".faiss"):
                return True
        return False
    