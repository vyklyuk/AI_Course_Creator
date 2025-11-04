# slide_prompt_agent_openai.py
import os
from dataclasses import dataclass, field
from typing import List, Optional, Dict

# If using the official OpenAI Python SDK v1+
import openai
import random
from collections import Counter
import re
from pathlib import Path

_STOPWORDS = {
    # UA
    "і","й","та","або","що","це","як","для","при","до","від","на","у","в","з","по","таким","також","є",
    # EN
    "and","or","the","a","an","of","to","for","in","on","with","into","by","from","as","is","are","be",
}

def _extract_keywords(text: str, k: int = 8):
    toks = re.findall(r"[A-Za-zА-Яа-яІіЇїЄєҐґ0-9\-]+", text.lower())
    toks = [t for t in toks if len(t) >= 3 and t not in _STOPWORDS]
    common = [w for w,_ in Counter(toks).most_common(20)]
    return common[:k]

def _make_content_cues(title: str, bullets: list[str]) -> str:
    text = f"{title} " + " ".join(bullets)
    kws = _extract_keywords(text, k=8)
    # зробимо їх читабельними фразами через коми
    return ", ".join(kws) if kws else "key topic nouns from the slide"



PROMPT_SYSTEM = """Compose concise, production-ready prompts for presentation backgrounds. Output ONLY the final prompt (no commentary).

Constraints:
• Palette/style: corporate blue/gray; clean, modern, photorealistic (maximally realistic) with subtle depth.
• Format: 16:9, 1920×1080.
• Layout: LEFT 70% (0–70% width) must be perfectly uniform light #F4F6F9 (no noise). No objects/icons/lines/shadows/textures/gradients left of 70%. Nothing may cross the 70% boundary.
• Composition: confine ALL visuals to 75–95% width as a tight, right-aligned cluster with ≥5% right padding.
• Gradient: very subtle left→right light-to-dark ONLY within 70–100% width (starts at 70%, must not spill left).
• Edges: feathered edges + soft vignette at extreme left/right (no banding or hard seams).
• People: allowed only if the user prompt requests; generic, non-identifiable; keep entirely within the right cluster.
• Prohibitions: no text, logos, watermarks, clutter.
• Content fidelity: the final prompt MUST concretely reflect the slide’s topic using the provided content cues (no unrelated themes).
"""



PROMPT_USER_TEMPLATE = """Create ONE final paragraph for a photorealistic presentation background based on the slide below. Respect SYSTEM constraints. Output ONLY the final paragraph (no commentary).

Slide title: {title}
Bullets:
{bullets}

Content cues (MUST appear as explicit visual motifs in the final paragraph): {content_cues}
People directive: {people_directive}

Critical reminders:
- LEFT 70% stays perfectly uniform #F4F6F9; nothing crosses left of 70%.
- ALL visuals confined to 75–95% width, right-aligned cluster (≥5% right padding).
- Gradient only in 70–100% (starts at 70%, never spills left); feather edges at extreme sides (no banding).
- Format: 16:9, 1920×1080. No text/logos/watermarks.
"""



@dataclass
class Slide:
    title: str
    bullets: List[str]

@dataclass
class OpenAIAgentConfig:
    model: str = "gpt-5-nano"  # pick your preferred text model
    image_model: str = "gpt-image-1"  # for optional image rendering
    people_probability: float = 0.3   # 30% імовірність додати людей
    random_seed: Optional[int] = None  # для відтворюваності (опційно)


class ImagePromptAgent:
    def __init__(self, client, cfg: OpenAIAgentConfig = OpenAIAgentConfig()):
        self.cfg = cfg
        self.client = client
    # self.client = openai.OpenAI(api_key="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")
       

    def generate_prompt(self, slide: Slide) -> str:
        bullets_formatted = "\n".join(f"- {b}" for b in slide.bullets)
        
        # user = PROMPT_USER_TEMPLATE.format(title=slide.title, bullets=bullets_formatted)
        
        # вирішуємо, чи додавати людей
        # rng = random.Random(self.cfg.random_seed) if self.cfg.random_seed is not None else random
        rng = random
        include_people = (rng.random() < self.cfg.people_probability)
    
        people_directive = (
            "People: include 1–2 generic, non-identifiable professionals; business attire; "
            "place entirely within the right cluster (75–95% width)."
            if include_people else
            "People: DO NOT include any people, faces, silhouettes, or body parts."
        )

        content_cues = _make_content_cues(slide.title, slide.bullets)
        
        user = PROMPT_USER_TEMPLATE.format(
            title=slide.title,
            bullets=bullets_formatted,
            content_cues=content_cues,
            people_directive=people_directive
        )


        
        res = self.client.responses.create(
            model=self.cfg.model,
            # temperature=self.cfg.temperature,
            input=[
                {"role": "system", "content": PROMPT_SYSTEM},
                {"role": "user", "content": user},
            ],
        )
        # Responses API text convenience:
        text = res.output_text if hasattr(res, "output_text") else (
            res.choices[0].message.content if hasattr(res, "choices") else str(res)
        )
        return text.strip()

    def render_image(
        self,
        prompt: str,
        size: str = "1920x1080",
        save_path: Optional[str] = None,
        out_dir: str = "image_out"
    ) -> Optional[str]:        
        
        """
        Optional: call the Images API to render. Requires internet + valid API key.
        Returns local path if saved.
        """
        if not hasattr(self.client, "images"):
            raise RuntimeError("Images API not available in the installed SDK.")
        img = self.client.images.generate(
            model=self.cfg.image_model,
            prompt=prompt,
            size=size,
        )
        # Save first image as PNG if data URL provided in b64_json
        if hasattr(img.data[0], "b64_json") and img.data[0].b64_json:
            import base64, io
            from PIL import Image

            # створюємо папку, якщо не існує
            out_path = Path(out_dir)
            out_path.mkdir(parents=True, exist_ok=True)
            
            png_bytes = base64.b64decode(img.data[0].b64_json)
            image = Image.open(io.BytesIO(png_bytes)).convert("RGBA")

            # Якщо користувач не передав ім'я → дефолтне
            filename = Path(save_path).name if save_path else "slide_background.png"
            save_path = out_path / filename
            # if save_path is None:
            #     save_path = "slide_background.png"
            image.save(save_path)
            return save_path
        # Or if URL is provided (older SDKs):
        if hasattr(img.data[0], "url") and img.data[0].url:
            return img.data[0].url
        return None
