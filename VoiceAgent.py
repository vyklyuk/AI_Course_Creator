import re
import io
import sys
from pathlib import Path
from typing import List, Tuple

from docx import Document
from openai import OpenAI

import shutil


class PresentationVoiceAgent:
    def __init__(self, client: OpenAI, model: str = "tts-1", voice: str = "nova",
                 out_dir: str = "audio_out", instructions: str = None, verbose = True):
        self.client = client
        self.model = model
        self.voice = voice
        self.verbose = verbose
        
        self.out = Path(out_dir)
        self.out.mkdir(parents=True, exist_ok=True)
        
        # Безпечна перевірка:
        resolved = self.out.resolve()
        if resolved == Path("/") or resolved == Path.home():
            raise ValueError(f"Refusing to clean unsafe directory: {resolved}")
        
        # # Видаляємо кожен елемент всередині папки
        # for item in self.out.iterdir():
        #     try:
        #         if item.is_dir():
        #             shutil.rmtree(item)
        #         else:
        #             item.unlink()
        #     except Exception as e:
        #         print(f"Some problem with clear folder {self.out}")
        #         raise
        
        self.instructions = instructions or (
            "Read the text with a tone that is clear, professional, high energy and confident, "
            "but not monotonous. Add slight warmth and emphasis so learners stay interested."
        )

    def process(self, slide_num: int, notes: str) -> Tuple[List[Path], Path]:
        
        # ділимо на абзаци
        sentences = [p.strip() for p in notes.split("\n") if p.strip()]
        
        for i, sentence in enumerate(sentences, start=1):
            resp = self.client.audio.speech.create(
                model=self.model,
                voice=self.voice,
                instructions=self.instructions,
                input=sentence
            )

            if i == 1:
                file_n = "000_Intro"
            else:
                file_n = f"Bullet_{(i-1):03d}"    
            
            filename = self.out / f"Slide_{slide_num}_{file_n}.mp3"
            filename.write_bytes(resp.content)
            if self.verbose:
                # print(sentence)
                print(f"[{i}] saved {filename}")



if __name__ == "__main__":
    client = OpenAI()  # ключ береться з OPENAI_API_KEY
    agent = PresentationVoiceAgent(client,
                                   model="tts-1",
                                   voice="nova",
                                   out_dir="audio_out_nova")

    audio_files, text_file = agent.process("content.docx")
    print(f"Створено {len(audio_files)} аудіофайлів, текст -> {text_file}")
