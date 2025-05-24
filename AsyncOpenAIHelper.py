import aiohttp
import asyncio
import os
from dotenv import load_dotenv
from pathlib import Path

class AsyncOpenAIHelper:
    _instance = None
    _semaphore = asyncio.Semaphore(5)  # limit concurrent requests

    def __new__(cls, *args, **kwargs):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._init_once(*args, **kwargs)
        return cls._instance

    def _init_once(self, api_key=None, embedding_model="text-embedding-3-small"):
        load_dotenv(dotenv_path=Path(__file__).parent / ".env")
        self.api_key = api_key or os.getenv("OPENAI_API_KEY")
        if not self.api_key:
            raise ValueError("OpenAI API key is not set.")
        self.embedding_model = embedding_model
        self.headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }

    async def encode(self, text):
        url = "https://api.openai.com/v1/embeddings"
        payload = {
            "input": text,
            "model": self.embedding_model
        }

        retries = 5
        backoff = 1

        async with AsyncOpenAIHelper._semaphore:
            for attempt in range(retries):
                async with aiohttp.ClientSession() as session:
                    async with session.post(url, headers=self.headers, json=payload) as resp:
                        if resp.status == 200:
                            data = await resp.json()
                            return data["data"][0]["embedding"]
                        elif resp.status == 429:
                            await asyncio.sleep(backoff)
                            backoff *= 2
                        else:
                            resp.raise_for_status()
            raise Exception("Max retries exceeded for embedding request")
