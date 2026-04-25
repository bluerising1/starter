#!/usr/bin/env python3
"""Generate Bhagavad Gita Instagram images chapter-by-chapter using OpenAI Images API.

Behavior:
- Generates one image per run for the next chapter in sequence.
- Remembers progress in a local state file.
- Uses a strict master style prompt for character and universe consistency.
"""

from __future__ import annotations

import argparse
import base64
import json
import os
import re
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

MASTER_STYLE_PROMPT = """MASTER CHARACTER CONSISTENCY PROMPT — BHAGAVAD GITA COMIC SERIES

Create an ultra-premium cinematic mythological comic illustration in vertical Instagram format (4:5 ratio), inspired by Marvel cinematic posters + Amar Chitra Katha storytelling + epic Indian mythology realism.

IMPORTANT RULE:
All future images must maintain EXACT SAME character faces, body structure, costume identity, jewelry style, armor design, expressions, and overall visual continuity across the entire Bhagavad Gita comic universe.

This must feel like one connected cinematic universe — not random AI-generated characters.

VISUAL STYLE:

• Hyper-realistic comic illustration
• Premium mythological cinematic quality
• Marvel comic + Mahabharata fusion
• Rich golden sunlight + royal blue + warm orange cinematic grading
• Divine atmosphere
• High-detail royal costumes
• Beautiful expressive faces
• Premium clean composition
• Light readable background
• Strong emotional storytelling
• Speech-bubble friendly composition
• NOT cartoonish
• NOT childish
• NOT dark muddy colors
• NOT fantasy random faces
• Must look elegant, powerful, and realistic

MAIN CHARACTER REFERENCE LOCK:

LORD KRISHNA:
• Divine blue skin tone
• Calm, wise, powerful expression
• Sharp beautiful eyes with peaceful authority
• Peacock feather crown
• Golden मुकुट with royal jewelry
• Yellow-golden royal dhoti
• Elegant divine ornaments
• Floral garland
• Strong but graceful physique
• Radiant aura around him
• Regal royal charioteer appearance
• Always visually divine and composed

ARJUNA:
• Strong handsome warrior prince
• Sharp jawline
• Intense expressive eyes
• Long warrior hair tied back
• Golden warrior crown
• Premium royal battle armor
• Muscular but elegant physique
• Emotional expressive face
• Heroic warrior presence
• Royal red + gold warrior costume
• Carries Gandiva bow
• Must look like the central human hero

BHISHMA:
• Tall majestic elderly warrior
• White beard
• White long hair
• Powerful divine warrior aura
• Royal silver-golden armor
• Grandfather-like authority
• Fearless battlefield presence

DRONACHARYA:
• Wise elderly guru warrior
• White beard
• Calm serious expression
• Royal sage-warrior appearance
• Traditional armor with guru presence

DURYODHANA:
• Strong royal warrior
• Proud face
• Sharp aggressive expression
• Golden crown
• Heavy royal armor
• Villain-like confidence but realistic

SETTING CONSISTENCY:

• Kurukshetra battlefield
• Massive armies
• Royal chariots
• Horses
• Elephants
• War flags
• Dust in sunlight
• Epic sunrise / golden sky
• Conch shells
• Grand mythological war atmosphere

INSTAGRAM CONTENT FORMAT:

• Clean visual storytelling
• Comic-panel style composition when needed
• Clear dialogue space for speech bubbles
• Emotional cliffhanger storytelling
• Highly engaging for Instagram audience
• Premium viral content quality

FINAL OUTPUT FEEL:

This should look like:
“Netflix presents Mahabharata Universe”

not

“Random mythology poster”

This should feel like a premium global cinematic franchise."""


@dataclass
class Chapter:
    chapter_number: int
    title: str
    scene_brief: str
    composition_notes: str
    hook_text: str


def slugify(text: str) -> str:
    cleaned = re.sub(r"[^a-zA-Z0-9]+", "-", text.strip().lower())
    return cleaned.strip("-")


def load_chapters(chapters_path: Path) -> list[Chapter]:
    data = json.loads(chapters_path.read_text(encoding="utf-8"))
    chapters: list[Chapter] = []
    for item in data:
        chapters.append(
            Chapter(
                chapter_number=int(item["chapter_number"]),
                title=item["title"],
                scene_brief=item["scene_brief"],
                composition_notes=item["composition_notes"],
                hook_text=item["hook_text"],
            )
        )
    chapters.sort(key=lambda chapter: chapter.chapter_number)
    return chapters


def load_state(state_path: Path) -> dict[str, Any]:
    if not state_path.exists():
        return {"next_chapter_index": 0, "history": []}
    return json.loads(state_path.read_text(encoding="utf-8"))


def save_state(state_path: Path, state: dict[str, Any]) -> None:
    state_path.parent.mkdir(parents=True, exist_ok=True)
    state_path.write_text(json.dumps(state, indent=2, ensure_ascii=False), encoding="utf-8")


def build_prompt(chapter: Chapter) -> str:
    return (
        f"{MASTER_STYLE_PROMPT}\n\n"
        "CHAPTER-SPECIFIC SCENE REQUEST:\n"
        f"Bhagavad Gita Chapter {chapter.chapter_number}: {chapter.title}\n"
        f"Scene summary: {chapter.scene_brief}\n"
        f"Composition direction: {chapter.composition_notes}\n"
        f"Emotional hook for post: {chapter.hook_text}\n\n"
        "OUTPUT INSTRUCTIONS:\n"
        "- Single premium Instagram-ready vertical visual (4:5).\n"
        "- Keep character identity perfectly consistent with prior chapters.\n"
        "- Keep space for future speech bubbles.\n"
        "- Use clear, readable cinematic composition with no text overlay."
    )


def generate_image(client: Any, model: str, prompt: str, size: str) -> bytes:
    try:
        response = client.images.generate(
            model=model,
            prompt=prompt,
            size=size,
        )
    except Exception as exc:
        error_text = str(exc)
        normalized = error_text.lower()
        if "billing_hard_limit_reached" in normalized or "billing hard limit" in normalized:
            raise RuntimeError(
                "OpenAI billing hard limit reached. Increase your API billing limit or wait for reset."
            ) from exc
        if "insufficient_quota" in normalized or "quota" in normalized:
            raise RuntimeError(
                "OpenAI API quota exceeded. Check plan and usage limits in your OpenAI billing dashboard."
            ) from exc
        raise RuntimeError(f"OpenAI image generation failed: {error_text}") from exc

    if not response.data:
        raise RuntimeError("Image API response had no data.")

    first = response.data[0]
    b64_data = getattr(first, "b64_json", None)
    if not b64_data:
        raise RuntimeError("Image API response did not contain b64_json image content.")

    return base64.b64decode(b64_data)


def write_metadata(metadata_path: Path, payload: dict[str, Any]) -> None:
    metadata_path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Generate Bhagavad Gita comic images chapter-by-chapter using OpenAI's latest image model."
        )
    )
    parser.add_argument(
        "--chapters-file",
        default="gita_chapters.json",
        help="Path to the chapters JSON file.",
    )
    parser.add_argument(
        "--state-file",
        default="output/gita_state.json",
        help="Path to persisted state JSON file.",
    )
    parser.add_argument(
        "--output-dir",
        default="output/gita_images",
        help="Directory where images and metadata are saved.",
    )
    parser.add_argument(
        "--model",
        default="gpt-image-1",
        help="Image model name.",
    )
    parser.add_argument(
        "--size",
        default="1024x1280",
        help="Image size in WIDTHxHEIGHT (4:5 recommended).",
    )
    parser.add_argument(
        "--force-chapter",
        type=int,
        default=None,
        help="Generate a specific chapter number without changing sequence order.",
    )
    parser.add_argument(
        "--posts-per-run",
        type=int,
        default=3,
        help="Number of chapter posts to generate in one run (default: 3).",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Do not call API; print planned chapters and prompt previews only.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    if args.posts_per_run < 1:
        raise SystemExit("--posts-per-run must be >= 1.")

    api_key = os.getenv("OPENAI_API_KEY")
    if not args.dry_run and not api_key:
        raise SystemExit("OPENAI_API_KEY is not set. Export it and run again.")

    if not args.dry_run:
        try:
            from openai import OpenAI
        except ModuleNotFoundError as exc:
            raise SystemExit(
                "openai package is not installed. Install with: pip install openai"
            ) from exc

    chapters_path = Path(args.chapters_file)
    state_path = Path(args.state_file)
    output_dir = Path(args.output_dir)

    if not chapters_path.exists():
        raise SystemExit(f"Chapters file not found: {chapters_path}")

    chapters = load_chapters(chapters_path)
    if not chapters:
        raise SystemExit("No chapters found in chapters file.")

    state = load_state(state_path)
    next_idx = int(state.get("next_chapter_index", 0))

    client = OpenAI(api_key=api_key) if not args.dry_run else None

    if args.force_chapter is not None:
        matching = [c for c in chapters if c.chapter_number == args.force_chapter]
        if not matching:
            raise SystemExit(f"Chapter {args.force_chapter} not found in chapters file.")
        selected_chapters = [matching[0]]
        update_sequence = False
    else:
        if next_idx >= len(chapters):
            raise SystemExit(
                "All chapters are already generated. Reset state file or use --force-chapter."
            )
        end_idx = min(next_idx + args.posts_per_run, len(chapters))
        selected_chapters = chapters[next_idx:end_idx]
        update_sequence = True

    output_dir.mkdir(parents=True, exist_ok=True)

    for chapter in selected_chapters:
        prompt = build_prompt(chapter)
        if args.dry_run:
            print(f"[DRY RUN] Chapter {chapter.chapter_number}: {chapter.title}")
            print(f"[DRY RUN] Prompt preview: {prompt[:400]}...\n")
            continue

        try:
            image_bytes = generate_image(
                client=client, model=args.model, prompt=prompt, size=args.size
            )
        except RuntimeError as exc:
            print(f"Image generation failed for Chapter {chapter.chapter_number}: {chapter.title}")
            print(f"Reason: {exc}")
            print("Tip: run with --dry-run to validate prompts without API usage.")
            raise SystemExit(1) from None

        timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
        chapter_slug = slugify(chapter.title)

        image_path = output_dir / f"ch{chapter.chapter_number:02d}_{chapter_slug}_{timestamp}.png"
        metadata_path = output_dir / f"ch{chapter.chapter_number:02d}_{chapter_slug}_{timestamp}.json"
        image_path.write_bytes(image_bytes)

        metadata = {
            "chapter_number": chapter.chapter_number,
            "chapter_title": chapter.title,
            "scene_brief": chapter.scene_brief,
            "composition_notes": chapter.composition_notes,
            "hook_text": chapter.hook_text,
            "model": args.model,
            "size": args.size,
            "generated_at_utc": timestamp,
            "image_path": str(image_path),
            "prompt": prompt,
        }
        write_metadata(metadata_path, metadata)

        if update_sequence and not args.dry_run:
            state.setdefault("history", []).append(
                {
                    "chapter_number": chapter.chapter_number,
                    "title": chapter.title,
                    "generated_at_utc": timestamp,
                    "image_path": str(image_path),
                    "metadata_path": str(metadata_path),
                }
            )

        print(f"Generated chapter {chapter.chapter_number}: {chapter.title}")
        print(f"Image saved to: {image_path}")
        print(f"Metadata saved to: {metadata_path}")

    if update_sequence and not args.dry_run:
        state["next_chapter_index"] = next_idx + len(selected_chapters)
        save_state(state_path, state)
        print(f"Next chapter index: {state['next_chapter_index']}")


if __name__ == "__main__":
    main()
