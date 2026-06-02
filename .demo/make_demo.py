#!/usr/bin/env python3
"""Render an animated terminal demo (frames -> GIF via ffmpeg) for the README.

Scenario (combines the o365 + GitHub skills):
  A request lands in the inbox; the agent reads it, edits a Terraform repo,
  opens a PR, and replies to the email with the link.
All names/values are placeholders.
"""
from PIL import Image, ImageDraw, ImageFont
from pathlib import Path

OUT = Path(__file__).parent / "frames"
OUT.mkdir(exist_ok=True)

# ── Canvas / theme (GitHub dark) ──────────────────────────────────────────
W, H = 1000, 460
BG       = (13, 17, 23)      # #0d1117
BAR      = (22, 27, 34)      # title bar
FG       = (201, 209, 217)   # default text
GREEN    = (63, 185, 80)
BLUE     = (88, 166, 255)
PURPLE   = (188, 140, 255)
ORANGE   = (255, 166, 87)
GREY     = (110, 118, 129)
WHITE    = (240, 246, 252)
CURSOR   = (88, 166, 255)

FONT = "/System/Library/Fonts/Menlo.ttc"
f_reg  = ImageFont.truetype(FONT, 21)
f_bold = ImageFont.truetype(FONT, 21, index=1)
f_small= ImageFont.truetype(FONT, 16)

PAD_X, PAD_Y = 28, 64
LINE_H = 30

def base_frame():
    img = Image.new("RGB", (W, H), BG)
    d = ImageDraw.Draw(img)
    # title bar
    d.rectangle([0, 0, W, 44], fill=BAR)
    for i, c in enumerate([(255,95,86),(255,189,46),(39,201,63)]):
        d.ellipse([20+i*22, 16, 32+i*22, 28], fill=c)
    d.text((W//2-90, 13), "hermes — agent", font=f_small, fill=GREY)
    return img, d

# ── Script: (segments) where each segment is (text, color, font) ──────────
# Lines are revealed progressively to fake typing / streaming.
PROMPT = [("❯ ", GREEN, f_bold)]  # emoji-free below; ❯ renders fine in Menlo

USER_TEXT = ("Read the GitHub access request in my inbox, add those users to "
             "the infra repo allow-list, and open a PR.")

# Steps: (marker, marker_color, text, text_color)
STEPS = [
    ("[o365]", BLUE,   " found email  \"GitHub access request - data team\"", FG),
    ("[parse]",GREY,   " extracted 3 users: alice, bob, carol", FG),
    ("[git] ", PURPLE, " branch  add-data-team-access", FG),
    ("[edit] ",ORANGE, " access.tf  (+3 members)", FG),
    ("[gh]  ", PURPLE, " opened PR #42  \"Grant data team repo access\"", FG),
    ("[o365]", BLUE,   " replied to sender with the PR link", FG),
]
DONE = ("done in 12s", GREEN)
PRLINE = ("PR  github.com/acme/infra/pull/42", BLUE)

frames = []

def draw_state(user_chars, steps_shown, done=False, cursor=True, cursor_at=None):
    img, d = base_frame()
    y = PAD_Y
    # prompt line + typed user text
    d.text((PAD_X, y), "❯", font=f_bold, fill=GREEN)
    typed = USER_TEXT[:user_chars]
    # wrap user text at ~58 chars
    import textwrap
    wrapped = textwrap.wrap(typed, width=58) or [""]
    for i, ln in enumerate(wrapped):
        d.text((PAD_X+28, y+i*LINE_H), ln, font=f_reg, fill=WHITE)
    yend = y + max(1,len(wrapped))*LINE_H
    if cursor and cursor_at == "user":
        lastln = wrapped[-1]
        cx = PAD_X+28 + d.textlength(lastln, font=f_reg)
        cy = y + (len(wrapped)-1)*LINE_H
        d.rectangle([cx+2, cy+2, cx+12, cy+24], fill=CURSOR)
    # steps
    sy = yend + 18
    for i in range(steps_shown):
        marker, mcol, text, tcol = STEPS[i]
        d.text((PAD_X, sy), marker, font=f_bold, fill=mcol)
        d.text((PAD_X + d.textlength(marker, font=f_bold) + 6, sy), text, font=f_reg, fill=tcol)
        sy += LINE_H + 4
    if done:
        sy += 8
        d.text((PAD_X, sy), "✓", font=f_bold, fill=GREEN)
        d.text((PAD_X+26, sy), DONE[0], font=f_bold, fill=GREEN)
        d.text((PAD_X+26, sy+LINE_H+2), PRLINE[0], font=f_reg, fill=BLUE)
    return img

# ── Build frame sequence ───────────────────────────────────────────────────
# 1. type the user prompt (every ~2 chars, blink cursor)
step = 3
for n in range(0, len(USER_TEXT)+1, step):
    frames.append((draw_state(n, 0, cursor_at="user"), 40))
frames.append((draw_state(len(USER_TEXT), 0, cursor_at="user"), 350))  # pause

# 2. reveal each step with a short delay
for s in range(1, len(STEPS)+1):
    frames.append((draw_state(len(USER_TEXT), s, cursor=False), 90))
    frames.append((draw_state(len(USER_TEXT), s, cursor=False), 380))  # dwell

# 3. done line + hold
for _ in range(2):
    frames.append((draw_state(len(USER_TEXT), len(STEPS), done=True, cursor=False), 700))
frames.append((draw_state(len(USER_TEXT), len(STEPS), done=True, cursor=False), 2000))

# ── Save frames + concat list for ffmpeg ────────────────────────────────────
for i, (img, _) in enumerate(frames):
    img.save(OUT / f"f{i:04d}.png")

concat = OUT.parent / "concat.txt"
with open(concat, "w") as fh:
    for i, (_, dur) in enumerate(frames):
        fh.write(f"file 'frames/f{i:04d}.png'\nduration {dur/1000:.3f}\n")
    # repeat last frame (ffmpeg concat needs final file line)
    fh.write(f"file 'frames/f{len(frames)-1:04d}.png'\n")

print(f"frames={len(frames)} written to {OUT}")
