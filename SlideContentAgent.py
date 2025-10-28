"""
Slides Creation Agent (Intro slide = ONLY Lecturer Notes + Visuals; NO Content)
- External OpenAI client is injected.
- Parameters per run:
  1) course_title      (str)
  2) module_title      (str)
  3) lecture_title     (str)
  4) minutes           (int)   # 8000 words/hour budget for Lecturer Notes
  5) approx_slides     (int)
  6) save_path         (Optional[str]; .docx enforced)

Requires:
  pip install openai python-docx
"""

from typing import Optional, Tuple, List
from pathlib import Path
from docx import Document
from datetime import datetime
import re

# ------------------------- Difficulty Profiles -------------------------
DIFFICULTY_PROFILES = {
    "Beginner": (
        "- Audience: newcomers; avoid jargon or define on first use.\n"
        "- Scaffolding: strong. Use short sentences, step-by-step guidance.\n"
        "- Examples: add simple, concrete examples and analogies.\n"
        "- Explanations: define acronyms and terms inline.\n"
        "- Tone: friendly, clear, encouraging.\n"
    ),
    "Intermediate": (
        "- Audience: practitioners; moderate jargon with quick clarifications.\n"
        "- Scaffolding: medium. Focus on practical tips and workflows.\n"
        "- Examples: show realistic scenarios and trade-offs.\n"
        "- Explanations: assume basics known; clarify non-obvious parts.\n"
        "- Tone: professional, concise.\n"
    ),
    "Advanced": (
        "- Audience: experts; domain terminology without basic definitions.\n"
        "- Scaffolding: light. Emphasize trade-offs, assumptions, edge cases.\n"
        "- Examples: deep dives, performance and limitation notes.\n"
        "- Explanations: focus on rigor, references, comparatives.\n"
        "- Tone: precise, analytical.\n"
    ),
    "Executive": (
        "- Audience: non-technical decision-makers.\n"
        "- Focus: business outcomes, risk, cost, ROI, timelines.\n"
        "- Scaffolding: narrative summaries, KPIs, options & implications.\n"
        "- Explanations: minimal technical detail; clear decisions needed.\n"
        "- Tone: crisp, outcome-oriented.\n"
    ),
}


# ------------------------- Single source of truth: approved transition phrases -------------------------

APPROVED_PHRASES: List[str] = [
    # "Now let's shift our focus to...",
    "Another key aspect to consider is...",
    "Building on what we just discussed, let's look at...",
    # "Next, let's examine...",
    "To understand this better, let's explore...",
    "From here, we move on to...",
    "With that in mind, let's turn to...",
    "Let's take a closer look at...",
    "Moving forward, let's focus on...",
    "At this point, it's important to highlight...",
    "Now, let's break down...",
    "To build on this idea, let's consider...",
    "Shifting gears, let's explore...",
    "This brings us to...",
    "Let's continue by looking at...",
    "Another perspective worth noting is...",
    "Keeping that in mind, let's move to...",
    "Let's now focus our attention on...",
    "With that background, let's examine...",
    # "Finally, let's turn our attention to...",
]

# ------------------------- Prompt template -------------------------

SYSTEM_TEMPLATE = """You are a professional voiceover narrator for an educational AML course.
Your job is to make the lecture engaging, approachable, and easy to follow.

STYLE RULES (mandatory):
1) Transitions: At the start of Lecturer Notes for each slide (except Intro), write a short transition paragraph (1–2 sentences) that smoothly introduces the slide as a whole. Use one of the approved transition phrases to begin. After this intro, provide bullet-aligned explanations for the Content items. Each bullet from Content must have its own corresponding explanation.
1.a) Approved transition phrases (use and vary them):
{approved_list}
1.b) Do not use the same transition phrase on two consecutive slides. If a repetition occurs, choose a different phrase from the approved list.
2) Sentence length: Keep sentences short and digestible. Target average <= 18 words per sentence in Lecturer Notes.
3) Tone: Professional but conversational; explain for understanding, not rote memorization.
4) Clarity: Split long ideas into multiple short sentences.
5) ASCII quotes and ellipses only: " " and ...

OUTPUT REQUIREMENTS (must follow):
- Plain text only (no JSON, no tables, no code blocks).
- Exactly N slides (where N is provided in the user prompt).
- Slide 1 (INTRO) is a special case:
  • Do NOT include a "Content:" section on Slide 1. If any "Content:" appears on Slide 1, the output is invalid.
  • Lecturer Notes: ONE short paragraph (2–4 sentences) that greets the audience and summarizes what the lecture is about. No bullets on this slide.
  • Wording rule: use a lecture-focused opener (e.g., "In this lecture, we will...", "Today’s lecture covers...", "Welcome..."). Do NOT say "This module introduces ...".
  • Visuals as usual.
- Slide 2 (AGENDA) is a special case:
  • Content: 3–6 grouped themes (clusters) that summarize the scope of the lecture.
    Do NOT list slide titles or slide numbers. Keep each bullet a short phrase (1–4 words).
    Max 6 bullets total.
  • Lecturer Notes: short intro paragraph + 1:1 brief explanation for each theme
    (why grouped this way / what falls under each cluster).
  • Visuals: simple roadmap with 3–6 nodes (no slide-by-slide enumeration).
- Slides 3..N: Each slide must include: Title, Content (3–5 bullets), Lecturer Notes (intro transition paragraph + bullet-aligned 1:1 explanations; no labels like "[for Content point 1]"), Visuals.
- Bullet alignment rule: The number of bullets under Content must EXACTLY match the number of explanation bullets under Lecturer Notes (excluding the intro paragraph). Any mismatch makes the output invalid.
- Total word count target for all Lecturer Notes is provided in the user prompt and must be followed (±10%).
- Mandatory slides: INTRO, AGENDA, Key Takeaways, Conclusion.
- Use English language.

TEMPLATE (follow literally):
Slide 1
Title: INTRO
Lecturer Notes:
Short welcoming paragraph (2–4 sentences) that greets the audience and previews the lecture. No bullets on this slide.
Visuals: ...
---
Slide N (for N ≥ 2)
Title: ...
Content:
- ...
- ...
- ...
Lecturer Notes:
Intro/transition paragraph here.
- ...
- ...
- ...
Visuals: ...
---"""

# ------------------------- Builders -------------------------

def calc_target_words(minutes: int) -> int:
    """8000 words/hour => 8000/60 words/min."""
    return round(minutes * (8000 / 60.0))

def build_system(difficulty_level: str = "Beginner") -> str:
    bullets = "\n".join(f"   - {p}" for p in APPROVED_PHRASES)
    base = SYSTEM_TEMPLATE.format(approved_list=bullets)
    # Append difficulty profile rules
    profile_rules = DIFFICULTY_PROFILES.get(difficulty_level, DIFFICULTY_PROFILES["Beginner"])
    base = base + "\n\nDifficulty profile:\n" + profile_rules
    
    # Hard nudge for the model to respect LN length targets
    return base + "\n\nCRITICAL: Always satisfy STRICT LENGTH & VALIDATION; when in doubt, err on the higher side by enriching explanations rather than shortening."


def build_user_prompt(
    course_title: str,
    module_title: str,
    lecture_title: str,
    minutes: int,
    approx_slides: int,
    lecture_short_desc: str | None = None,
) -> str:
    target_words = calc_target_words(minutes)
    min_words = round(target_words * 0.9)
    max_words = round(target_words * 1.1)

    # Бюджети довжини по ключових слайдах
    s1_min, s1_max = 60, 90       # INTRO
    s2_min, s2_max = 80, 120      # AGENDA
    s_kt_min, s_kt_max = 90, 130  # KEY TAKEAWAYS (передостанній)
    s_last_min, s_last_max = 90, 130  # CONCLUSION (останній)

    # Розподіл на середні слайди (3..N-2)
    remaining = max(0, target_words - (s1_min + s2_min + s_kt_min + s_last_min))
    mid_slides = max(1, approx_slides - 4)
    per_mid = max(80, round(remaining / mid_slides))
    per_bullet = max(25, round(per_mid / 4))  # орієнтир на ~4 булети

    return f"""
Create a presentation for the course:
{course_title}
Module: {module_title}
Topic: {lecture_title}

Context:
- Lecture short description (use naturally across slides for scope and examples; mention once in Slide 1 Lecturer Notes):
{(lecture_short_desc or "N/A").strip()}

STRICT LENGTH & VALIDATION:
- The output will be validated by a script that ONLY counts words inside lines following "Lecturer Notes:" on each slide.
- HARD CONSTRAINT: Total words across ALL Lecturer Notes ∈ [{min_words}, {max_words}] (target ≈ {target_words}).
- If you are near a boundary, ERR ON THE HIGHER SIDE (produce slightly more, not less).
- PER-SLIDE WORD BUDGET for Lecturer Notes:
  • Slide 1 (INTRO): {s1_min}–{s1_max} words (2–4 full sentences, no bullets).
  • Slide 2 (AGENDA): {s2_min}–{s2_max} words (intro + brief rationale per cluster).
  • Slides 3..{approx_slides-2}: ≈{per_mid} words each (±10%).
  • Slide {approx_slides-1} (KEY TAKEAWAYS): {s_kt_min}–{s_kt_max} words.
  • Slide {approx_slides} (CONCLUSION): {s_last_min}–{s_last_max} words.
- PER-BULLET on Slides 3..{approx_slides-2}: ≈{per_bullet} words per explanation (full sentences).
- DO NOT add new slides or bullets to meet length; expand explanations, examples, mini-cases, transitions.

STRUCTURE (MANDATORY ORDER):
- Exactly {approx_slides} slides.
- Slide 1 (INTRO):
  - Do NOT include a "Content:" section.
  - Lecturer Notes: one short welcoming paragraph (2–4 sentences). No bullets.
  - Visuals: short recommendation.
- Slide 2 (AGENDA):
  - Content: 3–6 grouped themes (clusters), concise phrases (1–4 words each). No enumeration of slide titles or numbers.
  - Lecturer Notes: 1–2 sentence intro + 1:1 brief rationale per theme.
  - Visuals: simple roadmap with 3–6 nodes.
- Slides 3..{approx_slides-2}:
  - Include Title, Content (3–5 concise bullets), Lecturer Notes, Visuals.
  - Bullet alignment rule: the number of Content bullets MUST equal the number of Lecturer Notes explanations (excluding the intro paragraph).
  - Lecturer Notes: short transition (1–2 sentences) + bullet-aligned explanations (≈{per_bullet} words each).
- Slide {approx_slides-1} (KEY TAKEAWAYS) — **must be the second-to-last slide**:
  - Title: **"Key Takeaways"** (exactly this wording).
  - Content: 3–5 bullets summarizing the most important points to remember.
  - Lecturer Notes: synthesize the 3–5 core insights within {s_kt_min}–{s_kt_max} words.
  - Bullet alignment rule: the number of Content bullets MUST equal the number of Lecturer Notes explanations (excluding the intro paragraph).
  - Visuals: summary icons/checklist.
- Slide {approx_slides} (CONCLUSION) — **must be the last slide**:
  - Title: **"Conclusion"** (exactly this wording).
  - Content: 3–5 bullets with final synthesis and (optionally) next steps.
  - Lecturer Notes: tie everything together, reinforce 1–2 actions/checks, keep within {s_last_min}–{s_last_max} words.
  - Bullet alignment rule: the number of Content bullets MUST equal the number of Lecturer Notes explanations (excluding the intro paragraph).
  - Visuals: wrap-up checklist or summary iconography.

STYLE & SELF-ADJUSTMENT:
- Use complete, flowing sentences (avg 12–18 words); avoid telegraphic fragments.
- If BELOW {min_words}, EXPAND (examples, rationale); if ABOVE {max_words}, CONDENSE (remove redundancy).

"""


# ------------------------- Post-processors -------------------------

def _normalize(s: str) -> str:
    return re.sub(r"\s+", " ", s.strip().lower())

def _pick_alternative(prev_phrase: str) -> str:
    for p in APPROVED_PHRASES:
        if _normalize(p) != _normalize(prev_phrase):
            return p
    return APPROVED_PHRASES[0]

def fix_transitions(presentation_text: str) -> str:
    """
    Ensures the first sentence of the Lecturer Notes intro paragraph on consecutive slides
    does not start with the same approved transition phrase. (Slides >= 2 are affected.)
    """
    parts = re.split(r'(?=^Slide\s+\d+\s*$)', presentation_text, flags=re.M)
    if len(parts) <= 1:
        return presentation_text

    prev_phrase = ""
    fixed_blocks = [parts[0]]

    for block in parts[1:]:
        # Skip Slide 1 entirely (no transition phrase enforced there)
        if re.match(r"^Slide\s*1\b", block.strip()):
            fixed_blocks.append(block)
            prev_phrase = ""
            continue

        m = re.search(r'(Lecturer Notes:\s*)(.+?)(\n- |\nVisuals: )', block, flags=re.S)
        if not m:
            fixed_blocks.append(block)
            prev_phrase = ""
            continue

        ln_header, intro_chunk, sep = m.group(1), m.group(2), m.group(3)
        intro_para = intro_chunk.strip()

        fs_match = re.match(r'^\s*(.+?\.)\s*(.*)$', intro_para, flags=re.S)
        if not fs_match:
            fixed_blocks.append(block)
            prev_phrase = ""
            continue

        first_sentence = fs_match.group(1).strip()
        remainder = fs_match.group(2).strip()

        found_phrase = None
        for phrase in APPROVED_PHRASES:
            if _normalize(first_sentence).startswith(_normalize(phrase)):
                found_phrase = phrase
                break

        if found_phrase and _normalize(found_phrase) == _normalize(prev_phrase):
            new_phrase = _pick_alternative(prev_phrase)
            tail = first_sentence[len(found_phrase):].lstrip()
            if tail and not tail.startswith((" ", ",")):
                tail = " " + tail
            first_sentence = f"{new_phrase}{tail}".strip()

        prev_phrase = found_phrase if found_phrase else ""
        new_intro = first_sentence + ((" " + remainder) if remainder else "")
        new_block = block[:m.start()] + ln_header + new_intro + sep + block[m.end():]
        fixed_blocks.append(new_block)

    return "".join(fixed_blocks)

import random

INTRO_OPENERS = [
    "Welcome to {course}. Today’s lecture covers",
    "In this lecture, we will",
    "Today, we’ll explore",
    "Let’s set the stage for",
    "Our focus today is",
]

def pick_intro_opener(course: str) -> str:
    # print("Random intro")
    return random.choice(INTRO_OPENERS).format(course=course)

def generate_intro_greeting_via_openai(
    client,
    model: str,
    course_title: str,
    module_title: str,
    lecture_title: str,
    lecture_short_desc: Optional[str] = None,
) -> str:
    """Generate a single short welcoming paragraph with a varied opener. Uses Responses API, falls back to Chat Completions."""
    opener = pick_intro_opener(course_title)
    # print(opener)
    sys = (
        "You are a professional voiceover narrator for an educational AML course. "
        "Output a single short paragraph (2–4 sentences). "
        "Begin with the provided opener verbatim, then continue naturally. "
        "Mention the course and module naturally. "
        "Use ASCII quotes only, no bullets, no headings, no 'Content:'. "
        "Do NOT use the phrase 'This module introduces'."
    )
    desc_line = f"We will focus on: {lecture_short_desc}" if lecture_short_desc else ""
    usr = (
        f"Course: {course_title}\n"
        f"Module: {module_title}\n"
        f"Lecture: {lecture_title}\n"
        f"{desc_line}\n\n"
        f"Opener to use (verbatim at the start): {opener}\n"
        "Write the Intro slide's Lecturer Notes paragraph now."
    )

    # --- Try Responses API ---
    try:
        resp = client.responses.create(
            model=model,
            input=[{"role": "system", "content": sys},
                   {"role": "user", "content": usr}]
        )
        text = getattr(resp, "output_text", None)
        if not text:
            try:
                # alternative shape
                text = resp.output[0].content[0].text
            except Exception:
                text = None
        if text:
            text = text.strip()
    except Exception:
        text = None

    # --- Fallback to Chat Completions if Responses failed ---
    if not text:
        try:
            chat = client.chat.completions.create(
                model=model,
                messages=[{"role": "system", "content": sys},
                          {"role": "user", "content": usr}]
            )
            text = chat.choices[0].message.content.strip()
        except Exception:
            text = None

    # --- Last-resort fallback: still VARIED using the opener ---
    if not text:
        text = f"{opener} '{lecture_title}' and set expectations for today’s session."

    # Soft-guard: avoid "This module introduces"
    text = re.sub(r"\bThis module introduces\b", "In this lecture, we will", text, flags=re.I).strip()
    print("Introduce generated", "*"*80)
    return text


def enforce_intro_rules(
    text: str,
    greeting_para: Optional[str] = None,
    *,
    client=None,
    model: Optional[str] = None,
    course_title: Optional[str] = None,
    module_title: Optional[str] = None,
    lecture_title: Optional[str] = None,
    lecture_short_desc: Optional[str] = None,
    force_replace: bool = False,
) -> str:
    """
    Forces Slide 1 to have NO 'Content:' and NO bullets.
    Ensures exactly one short paragraph in Lecturer Notes (with a greeting).
    If greeting_para is None and client+context provided, generates the greeting via OpenAI.
    """
    # Якщо greeting_para не подали — згенеруй через OpenAI (за наявності клієнта та контексту)
    if greeting_para is None and client and model and course_title and module_title and lecture_title:
        try:
            greeting_para = generate_intro_greeting_via_openai(
                client=client,
                model=model,
                course_title=course_title,
                module_title=module_title,
                lecture_title=lecture_title,
                lecture_short_desc=lecture_short_desc,
            )
        except Exception:
            greeting_para = "This lecture introduces the topic and sets clear expectations for what you will learn today."

    m = re.search(r"(^Slide\s*1\b.*?)(?=^Slide\s*\d+\b|^---\s*$|\Z)", text, flags=re.M | re.S)
    if not m:
        return text

    block = m.group(1)

    # 1) Прибрати будь-який Content на Intro
    block = re.sub(
        r"\nContent:\s*(?:\n(?:-.*|.+))*?(?=\nLecturer Notes:|\nVisuals:|\n---|^Slide\s*\d+|\Z)",
        "\n",
        block,
        flags=re.S | re.M
    )

    # 2) Почистити Lecturer Notes: рівно один короткий абзац, без булетів
    def _ln_clean(match: re.Match) -> str:
        prefix = match.group(1)  # "Lecturer Notes:\n"
        body   = match.group(2)

        # Прибрати булети
        lines = [ln for ln in body.splitlines() if not ln.lstrip().startswith("-")]
        text_body = "\n".join(lines).strip()

        # Беремо перший абзац
        first_par = re.split(r"\n\s*\n", text_body)[0].strip()

        # Якщо force_replace — завжди беремо greeting_para (якщо є)
        if force_replace and greeting_para:
            first_par = greeting_para.strip()
        else:
            if not first_par:
                # раніше тут був фіксований рядок "This lecture introduces the topic..."
                # зробимо варіативний fallback з opener'ом
                opener = pick_intro_opener(course_title or "this course")
        first_par = (greeting_para or f"{opener} '{lecture_title}' and outline what you will learn today.").strip()

        # Уникаємо 'This module introduces' (мʼяка заміна)
        first_par = re.sub(r"\bThis module introduces\b", "In this lecture, we will", first_par, flags=re.I)


        # Уникаємо 'This module introduces' (замінюємо на нейтральне лекційне формулювання)
        first_par = re.sub(r"\bThis module introduces\b", "In this lecture, we will", first_par, flags=re.I)
        return prefix + first_par + "\n"

    block = re.sub(
        r"(Lecturer Notes:\s*\n)(.+?)(?=\nVisuals:|\Z)",
        _ln_clean,
        block,
        flags=re.S
    )

    return text[:m.start(1)] + block + text[m.end(1):]


# ------------------------- Agent -------------------------

class SlidesAgent:
    """
    Agent that generates AML slides. Receives a pre-initialized OpenAI client instance.
    Usage:
        from openai import OpenAI
        client = OpenAI(api_key="...")
        agent = SlidesAgent(client, model="gpt-5-nano")
        text, path = agent.generate(...)
    """
    def __init__(self, client, model: str = "gpt-5-nano"):
        self.client = client
        self.model = model

    def generate(
        self,
        course_title: str,
        module_title: str,
        lecture_title: str,
        difficulty_level: str = "beginner",
        minutes: int = 10,
        approx_slides: int = 10,
        lecture_short_desc: str | None = None,
        save_path: Optional[str] = None,
        verbose: Optional[bool] = True,
    ) -> Tuple[str, str]:


        """
        Генерує текст слайдів і зберігає у DOCX.
        Якщо кількість слів у Lecturer Notes виходить за допустимі межі,
        робить повторні запити (до MAX_RETRIES).
        """
        MAX_RETRIES = 5
        TOL = 0.10  # ±10%

        def _count_ln_words(txt: str) -> int:
            text = txt
            # Шукаємо всі блоки нотаток
            blocks = re.findall(
                r"Lecturer Notes:\s*(.*?)(?=\n\s*Visuals:|\n\s*---|\n\s*Slide\s+\d+|\n\s*Title:|\Z)",
                text,
                flags=re.S
            )
            
            # Об’єднуємо їх у суцільний текст
            lecturer_text = "\n\n".join(b.strip() for b in blocks if b.strip())
            lecturer_text = re.sub(r'(?m)^- ', '', lecturer_text)
            word_count = len(lecturer_text.split())
            return word_count

        target_words = calc_target_words(minutes)
        min_words = int(target_words * (1 - TOL))
        max_words = int(target_words * (1 + TOL))

        system = build_system(difficulty_level)
        base_user = build_user_prompt(course_title, module_title, lecture_title, minutes, approx_slides, lecture_short_desc)
        if verbose:
            print(system)
            print(base_user)
        text, guidance_suffix = None, ""
        success = False

        for attempt in range(1, MAX_RETRIES):  # початковий
            print(f"Attempt {attempt}")
            user_prompt = base_user if not guidance_suffix else f"{base_user}\n\nADJUSTMENT:\n{guidance_suffix}"

            # --- Виклик API ---
            try:
                response = self.client.responses.create(
                    model=self.model,
                    input=[
                        {"role": "system", "content": system},
                        {"role": "user", "content": user_prompt},
                    ],
                )
                text = getattr(response, "output_text", None) or response.output[0].content[0].text
            except Exception:
                text = ""

            text = enforce_intro_rules(
                text,
                greeting_para=None,
                client=self.client,
                model=self.model,
                course_title=course_title,
                module_title=module_title,
                lecture_title=lecture_title,
                lecture_short_desc=lecture_short_desc,
                force_replace=True  # лишай: щоразу отримаєш новий opener через OpenAI
            )
        
            # Optional: avoid consecutive identical transition phrases on Slides >= 2
            text = fix_transitions(text)
      
            # --- Build guidance for next attempt ---
            ln_words = _count_ln_words(text)
            if min_words <= ln_words <= max_words*1.1:
                success = True
                self.last_status = f"OK: {ln_words} words in Lecturer Notes (target {min_words}–{max_words}, ≈{target_words})."
                print("✅ " + self.last_status)
                break
           
            if ln_words < min_words:
                delta = min_words - ln_words
                action = "EXPAND"
                tip = (
                    "Add more detailed, well-structured explanations in Lecturer Notes. "
                    "Prefer short flowing sentences, keep 1:1 bullet-to-note mapping. "
                    "Enrich with examples or mini-cases without adding new slides."
                )
            else:
                delta = ln_words - max_words
                action = "CONDENSE"
                tip = (
                    "Reduce verbosity in Lecturer Notes. Keep essential ideas, "
                    "remove redundancy, and preserve 1:1 bullet-to-note mapping."
                )
            
            guidance_suffix = (
                f"{action} the Lecturer Notes by about {delta} words so that the total "
                f"falls within {min_words}–{max_words} words (target ≈ {target_words}). "
                f"Maintain exactly {approx_slides} slides and follow all structure rules. {tip}"
            )

            if not success:
                try:
                    ln_words = _count_ln_words(text)
                except Exception:
                    ln_words = 0
                self.last_status = (
                    f"FAILED after {MAX_RETRIES} retries: final {ln_words} words "
                    f"(target {min_words}–{max_words}, ≈{target_words})."
                )
                print("⚠️ " + self.last_status)
        

        
        doc = Document()
        for line in text.splitlines():
            doc.add_paragraph("" if not line.strip() else line)

        doc.save(save_path)

        return text, str(save_path)

# ------------------------- Example (run directly) -------------------------
if __name__ == "__main__":
 
    agent = SlidesAgent(client, model="gpt-5-nano")

    generated_text, saved_path = agent.generate(
        course_title="AML Uncovered – Foundation Course (Beginner Level)",
        module_title="Foundations of AML & Financial Crime",
        lecture_title="How Predicate Crimes Enable Money Laundering",
        lecture_short_desc="""Explain how criminal proceeds are generated and enter the financial system. 
                            Link specific predicate crimes to placement, layering, and integration stages.
                            Importance of understanding predicate crimes for AML investigations.
                            Learning Outcome: Students can analyze the connection between criminal activity and money laundering patterns."
                            """,
        difficulty_level="Beginner",
        minutes=8,            # -> ~8000 target words (±10%) across Lecturer Notes
        approx_slides=10,      # -> exactly 10 slides
        save_path="content.docx"
    )
    print("Saved to:", saved_path)
