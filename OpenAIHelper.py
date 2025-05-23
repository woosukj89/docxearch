import requests
import numpy as np
import time
import tiktoken

class OpenAIHelper:
    _instance = None

    def __new__(cls, *args, **kwargs):
        if cls._instance is None:
            cls._instance = super(OpenAIHelper, cls).__new__(cls)
            cls._instance._init_once(*args, **kwargs)
        return cls._instance

    def _init_once(self, api_key=None, embedding_model="text-embedding-3-small", chat_model="gpt-4o-mini"):
        from dotenv import load_dotenv
        import os
        from pathlib import Path

        # Load from .env
        load_dotenv(dotenv_path=Path(__file__).parent / ".env")
        self.api_key = api_key or os.getenv("OPENAI_API_KEY")
        if not self.api_key:
            raise ValueError("OpenAI API key is not set.")
        self.embedding_model = embedding_model
        self.chat_model = chat_model
        self.headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }
        self.debug = False

    def encode(self, text):
        if self.debug:
            print("Encoding", text)
            tokenizer = tiktoken.encoding_for_model("text-embedding-3-small")
            print("Number of tokens:", len(tokenizer.encode(text)))

        url = "https://api.openai.com/v1/embeddings"
        payload = {
            "input": text,
            "model": self.embedding_model
        }

        max_retries = 5
        backoff = 1  # start with 1 second

        for _ in range(max_retries):
            response = requests.post(url, headers=self.headers, json=payload)
            if response.status_code == 200:
                data = response.json()
                return data["data"][0]["embedding"]
            elif response.status_code == 429:
                print(f"Rate limit hit. Retrying in {backoff} seconds...")
                time.sleep(backoff)
                backoff *= 2
            else:
                response.raise_for_status()

        raise Exception("Max retries exceeded for embedding request")

    def answer_with_context(self, query, context, language="Korean"):
        prompt = f"""Answer the user's question using ONLY the documents given in the context below. 
You MUST cite your sources using numbers in square brackets after EVERY piece of information (e.g., [0], [1], [2]).
At the end, give a summary of which parts you used from each document.
Please give your answer in {language}.

Context (numbered documents):
{context}

Question: {query}

Instructions:
1. Use information ONLY from the provided documents
2. You MUST cite sources using [X] format after EVERY claim
3. Use multiple citations if information comes from multiple documents (e.g., [0][1])
4. Make sure citations are numbers that match the context documents
5. DO NOT skip citations - every piece of information needs a citation
6. DO NOT make up information - only use what's in the documents

Example format:
기도하는 방법을 배우는 것은 "주변 환경과 경험"에 크게 의존합니다 [0]. 저자는 "부모님이나 기독교 가정에서 기도하는 모습"을 통해 배우거나, 교회에서 가르치는 기도의 예를 모방하는 것이 길이라고 가르칩니다 [1]. 예를 들어, 어린 시절부터 부모님이 기도하는 것을 보고 배운다면 그 습관을 유지하는 것이 도움이 될 수 있습니다 [2][3].

[0] 주변 환경과 경헙을 통해 기도를 배우는 내용
[1] 가족을 통해 기도를 배우는 모습
[2] 기도 습관의 중요성
[3] 부모님으로부터 기도를 배운 예시

Answer:"""
        url = "https://api.openai.com/v1/chat/completions"
        payload = {
            "model": self.chat_model,
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.2
        }
        response = requests.post(url, headers=self.headers, json=payload)
        response.raise_for_status()
        data = response.json()
        return data["choices"][0]["message"]["content"].strip()
