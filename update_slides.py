"""
Two tasks:
  1. Update "CBC PLAYER AUCTION 2026" header font to match player name size (34pt)
     on all player profile slides.
  2. Match headshots from ./headshots/ folder to player name on each slide
     and insert the photo into the photo_inner area.

Usage:
    python3 update_slides.py [--headshots-dir ./headshots]
"""
import os
import sys
import io
import re
import argparse
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ── Constants ─────────────────────────────────────────────────────────────────
PLAYER_NAME_FONT_SIZE = Pt(34)   # 431800 EMU – matches player name on slides
HEADER_TEXT           = 'CBC  PLAYER AUCTION 2026'

# Photo area (from inspection of slide template)
PHOTO_X = Inches(1.070)
PHOTO_Y = Inches(1.600)
PHOTO_W = Inches(2.700)
PHOTO_H = Inches(2.700)

# ── Name normalisation for fuzzy matching ─────────────────────────────────────
def _norm(s: str) -> str:
    """Lowercase, strip punctuation, collapse spaces."""
    s = s.lower()
    s = re.sub(r'[^a-z0-9 ]', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def find_headshot(name: str, headshots_dir: Path) -> Path | None:
    """
    Try to find an image in headshots_dir whose filename best matches `name`.
    Returns the Path or None.
    """
    if not headshots_dir.is_dir():
        return None

    image_exts = {'.jpg', '.jpeg', '.png', '.webp', '.bmp', '.gif'}
    candidates = [f for f in headshots_dir.iterdir()
                  if f.suffix.lower() in image_exts]
    if not candidates:
        return None

    norm_name = _norm(name)
    # Exact match first
    for c in candidates:
        if _norm(c.stem) == norm_name:
            return c
    # Partial: every word in the name appears in the filename
    name_words = set(norm_name.split())
    scored = []
    for c in candidates:
        fn_words = set(_norm(c.stem).split())
        overlap  = len(name_words & fn_words)
        if overlap > 0:
            # score: fraction of name words matched
            scored.append((overlap / len(name_words), c))
    if scored:
        scored.sort(key=lambda x: -x[0])
        if scored[0][0] >= 0.5:   # at least half the words match
            return scored[0][1]
    return None

# ── Headshot insertion ─────────────────────────────────────────────────────────
def insert_headshot(slide, img_path: Path):
    """
    Remove the photo_inner placeholder and insert the real photo in its place.
    """
    sp_tree = slide.shapes._spTree

    # Remove photo_inner (the placeholder with the camera emoji)
    for shape in slide.shapes:
        if shape.name == 'photo_inner':
            sp_tree.remove(shape._element)
            break

    # Add picture
    slide.shapes.add_picture(
        str(img_path),
        PHOTO_X, PHOTO_Y, PHOTO_W, PHOTO_H
    )
    print(f'    → Inserted photo: {img_path.name}')

# ── Header update ─────────────────────────────────────────────────────────────
def update_header(slide, slide_idx: int):
    """
    Find the 'CBC PLAYER AUCTION 2026' textbox and bump its font size to 34pt.
    Also expand the textbox height to fit the larger text and remove/hide the
    'PLAYER PROFILE' subtitle to make room.
    """
    header_shape  = None
    profile_shape = None

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        text = shape.text_frame.text.strip()
        if text == HEADER_TEXT:
            header_shape = shape
        elif text == 'PLAYER PROFILE':
            profile_shape = shape

    if header_shape is None:
        return  # Nothing to update

    # ── Update font size on every run in the textbox ──────────────────────────
    for para in header_shape.text_frame.paragraphs:
        for run in para.runs:
            run.font.size = PLAYER_NAME_FONT_SIZE
            run.font.bold = True

    # ── Resize the textbox to fit 34pt text ───────────────────────────────────
    # Keep same x/y but increase height from ~0.38" to ~0.60"
    header_shape.height = Inches(0.62)
    # Shift y up slightly so it doesn't overflow below the separator
    # (stays at y=0.22" naturally)

    # ── Remove 'PLAYER PROFILE' subtitle – no longer fits ────────────────────
    if profile_shape is not None:
        slide.shapes._spTree.remove(profile_shape._element)

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--src',           default='/tmp/workspace/minhajz/cbc-storage/cbc-auction-2026.pptx')
    parser.add_argument('--dest',          default='/tmp/workspace/minhajz/cbc-storage/cbc-auction-2026.pptx')
    parser.add_argument('--headshots-dir', default='/tmp/workspace/minhajz/cbc-storage/headshots')
    args = parser.parse_args()

    headshots_dir = Path(args.headshots_dir)
    if not headshots_dir.is_dir():
        print(f'[WARN] headshots dir not found: {headshots_dir}')
    else:
        imgs = [f for f in headshots_dir.iterdir()
                if f.suffix.lower() in {'.jpg', '.jpeg', '.png', '.webp', '.bmp'}]
        print(f'Found {len(imgs)} image(s) in headshots folder.')

    print(f'Opening {args.src} …')
    prs    = Presentation(args.src)
    slides = list(prs.slides)

    header_updated  = 0
    photos_inserted = 0
    photos_missing  = []

    for i, slide in enumerate(slides):
        # Only process player profile slides
        has_profile_header = any(
            (s.has_text_frame and s.text_frame.text.strip() == HEADER_TEXT)
            for s in slide.shapes
        )
        if not has_profile_header:
            continue

        # Extract player name
        player_name = None
        for shape in slide.shapes:
            if shape.name == 'TextBox 22' and shape.has_text_frame:
                player_name = shape.text_frame.text.strip()
                break

        print(f'Slide {i+1}: {player_name or "??"}')

        # 1. Update header
        update_header(slide, i)
        header_updated += 1

        # 2. Insert headshot if available
        if player_name:
            img_path = find_headshot(player_name, headshots_dir)
            if img_path:
                insert_headshot(slide, img_path)
                photos_inserted += 1
            else:
                photos_missing.append(f'Slide {i+1}: {player_name}')
                print(f'    → No headshot match found for: {player_name}')

    prs.save(args.dest)
    print(f'\n── Summary ──────────────────────────────')
    print(f'Headers updated : {header_updated}')
    print(f'Photos inserted : {photos_inserted}')
    if photos_missing:
        print(f'Missing photos  : {len(photos_missing)}')
        for m in photos_missing:
            print(f'  {m}')
    print(f'Saved → {args.dest}')

if __name__ == '__main__':
    main()
