"""Cut a vertical Reel/Short from a longer video for @juliancriderestates.

Usage:
  python3 social/clip.py SRC START END OUT.mp4 --hook "Hook text" [--words words.json] [--no-captions] [--luxe]
    SRC        source video (any size; it's scaled/cropped to 1080x1920, 30 fps)
    START/END  seconds
    --words    faster-whisper JSON ([{s,e,t,w:[[start,end,word],...]}]) for burned-in captions
    --hook     big line at the top for the first 3 seconds
    --luxe     calm-luxury caption font (Cormorant) instead of the everyday look
    --mute     drop the original audio (for footage with private conversations; add music in Metricool)
Adds a 2.5 s end card ("Book your game plan · link in bio") in the calm black-and-white look.
"""
import argparse, json, os, subprocess, tempfile
from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
FONTS = os.path.join(HERE, 'fonts')
RED = (209, 52, 56)


def ass_time(t):
    t = max(t, 0); h = int(t // 3600); m = int(t % 3600 // 60); s = t % 60
    return f'{h}:{m:02d}:{s:05.2f}'


def build_ass(path, words, start, end, hook, luxe):
    cap_font = 'Cormorant Garamond' if luxe else 'Anton'
    cap_size = 150 if luxe else 140
    lines = [
        '[Script Info]', 'ScriptType: v4.00+', 'PlayResX: 1080', 'PlayResY: 1920', '',
        '[V4+ Styles]',
        'Format: Name, Fontname, Fontsize, PrimaryColour, SecondaryColour, OutlineColour, BackColour, Bold, Italic, Underline, StrikeOut, ScaleX, ScaleY, Spacing, Angle, BorderStyle, Outline, Shadow, Alignment, MarginL, MarginR, MarginV, Encoding',
        # Captions: white with a black outline, just below the middle (clear of face bubbles and Instagram's buttons)
        f'Style: Cap,{cap_font},{cap_size},&H00FFFFFF,&H00FFFFFF,&H00000000,&H64000000,1,{1 if luxe else 0},0,0,100,100,0,0,1,6,2,2,90,90,820,1',
        # Hook: white on a black box near the top
        f'Style: Hook,{cap_font},{116 if luxe else 106},&H00FFFFFF,&H00FFFFFF,&H00000000,&H28000000,1,0,0,0,100,100,{0 if luxe else 1},0,3,26,0,8,80,80,260,1',
        '', '[Events]', 'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text']
    if hook:
        lines.append(f'Dialogue: 1,{ass_time(0)},{ass_time(min(3.2, end - start))},Hook,,0,0,0,,{hook}')
    if words:
        ws = [w for seg in words for w in seg['w'] if w[0] >= start - 0.05 and w[1] <= end + 0.05]
        group = []
        for i, w in enumerate(ws):
            group.append(w)
            txt = ''.join(g[2] for g in group).strip()
            last = i == len(ws) - 1
            gap = (not last) and ws[i + 1][0] - w[1] > 0.6
            if len(group) >= 3 or len(txt) > 16 or txt.endswith(('.', ',', '?', '!')) or gap or last:
                t0, t1 = group[0][0] - start, (ws[i + 1][0] - start if not last and not gap else w[1] - start + 0.15)
                text = txt if luxe else txt.upper()
                lines.append(f'Dialogue: 0,{ass_time(t0)},{ass_time(t1)},Cap,,0,0,0,,{text}')
                group = []
    open(path, 'w').write('\n'.join(lines) + '\n')


def end_card(path, luxe):
    im = Image.new('RGB', (1080, 1920), (11, 11, 12)); d = ImageDraw.Draw(im)
    if luxe:
        h1 = ImageFont.truetype(os.path.join(FONTS, 'Cormorant.ttf'), 140); h1.set_variation_by_axes([500])
    else:
        h1 = ImageFont.truetype(os.path.join(FONTS, 'Anton-Regular.ttf'), 150)
    small = ImageFont.truetype(os.path.join(FONTS, 'Oswald.ttf'), 40); small.set_variation_by_axes([500])
    tiny = ImageFont.truetype(os.path.join(FONTS, 'Oswald.ttf'), 28); tiny.set_variation_by_axes([500])
    d.rectangle((90, 700, 180, 706), fill=RED)
    words = ['Book your', 'game plan'] if luxe else ['BOOK YOUR', 'GAME PLAN']
    d.text((90, 760), words[0], font=h1, fill='white'); d.text((90, 925), words[1], font=h1, fill='white')
    d.rectangle((90, 1150, 760, 1240), fill=RED)
    d.text((120, 1170), '30-MIN CALL · LINK IN BIO', font=small, fill='white')
    d.text((90, 1290), 'OR DM ME HOMES', font=small, fill=(199, 199, 204))
    d.text((90, 1640), '@JULIANCRIDERESTATES', font=small, fill='white')
    d.text((90, 1700), 'JULIÁN CRIDER · REALTOR® · BOECKLE & CO. · S.0177760', font=tiny, fill=(110, 110, 115))
    im.save(path)


def main():
    a = argparse.ArgumentParser()
    a.add_argument('src'); a.add_argument('start', type=float); a.add_argument('end', type=float); a.add_argument('out')
    a.add_argument('--hook', default=''); a.add_argument('--words'); a.add_argument('--no-captions', action='store_true')
    a.add_argument('--luxe', action='store_true')
    a.add_argument('--mute', action='store_true', help='drop the original audio (music gets added in Metricool)')
    o = a.parse_args()
    words = json.load(open(o.words)) if o.words and not o.no_captions else None
    tmp = tempfile.mkdtemp()
    ass = os.path.join(tmp, 'c.ass'); build_ass(ass, words, o.start, o.end, o.hook, o.luxe)
    card = os.path.join(tmp, 'end.png'); end_card(card, o.luxe)
    dur = o.end - o.start
    vf = (f"[0:v]scale=1080:1920:force_original_aspect_ratio=increase,crop=1080:1920,setsar=1,fps=30,"
          f"ass={ass}:fontsdir={FONTS},format=yuv420p[v0];"
          f"[1:v]scale=1080:1920,setsar=1,fps=30,format=yuv420p[v1];"
          + (f"[2:a]asplit[n0][n1];[n0]atrim=0:{dur}[a0];" if o.mute else
             f"[0:a]aresample=48000,afade=t=out:st={max(dur - 0.4, 0)}:d=0.4[a0];[2:a]anull[n1];") +
          "[n1]atrim=0:2.5,asetpts=PTS-STARTPTS[a1];"
          f"[v0][a0][v1][a1]concat=n=2:v=1:a=1[v][a]")
    subprocess.run(['ffmpeg', '-y', '-loglevel', 'error', '-ss', str(o.start), '-t', str(dur), '-i', o.src,
                    '-loop', '1', '-t', '2.5', '-i', card, '-f', 'lavfi', '-t', str(dur + 2.5), '-i', 'anullsrc=r=48000:cl=stereo',
                    '-filter_complex', vf, '-map', '[v]', '-map', '[a]', '-c:v', 'libx264', '-preset', 'medium', '-crf', '23',
                    '-c:a', 'aac', '-b:a', '160k', '-movflags', '+faststart', o.out], check=True)
    print(o.out, f'{dur + 2.5:.1f}s')


if __name__ == '__main__':
    main()
