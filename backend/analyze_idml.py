"""Temporary analysis script - delete after use"""
import zipfile
from xml.etree import ElementTree as ET

path = '/Users/dong-wankim/Downloads/독해_내지_IDML (1).idml'
ns = 'http://ns.adobe.com/AdobeInDesign/idml/1.0/packaging'

SKIP = {'Mission','Syntax','Logic Note','A','B','Chapter','CONTENTSc'}
CHAPTER_NAMES = {'첫 문장과 링킹','배경과 반전','부분 링킹','정보 분류','체화'}

def get_text(zf, sid):
    sf = f'Stories/Story_{sid}.xml'
    if sf not in zf.namelist(): return ''
    with zf.open(sf) as f:
        root = ET.parse(f).getroot()
    return ''.join(c.text for c in root.iter('Content') if c.text).strip()

def is_boilerplate(t):
    if not t: return True
    if t in SKIP or t in CHAPTER_NAMES: return True
    if t.upper().startswith('CHAPTER'): return True
    if t.isdigit(): return True
    if len(t) <= 2: return True
    if t in ('①','②','③','④','⑤'): return True
    return False

with zipfile.ZipFile(path) as zf:
    with zf.open('designmap.xml') as f:
        dm = ET.parse(f).getroot()
    spreads = [e.get('src') for e in dm.iter(f'{{{ns}}}Spread') if e.get('src')]

    for sp in spreads:
        with zf.open(sp) as f:
            root = ET.parse(f).getroot()
        frames = []
        for tf in root.iter('TextFrame'):
            sid = tf.get('ParentStory')
            t = get_text(zf, sid)
            if not is_boilerplate(t):
                frames.append(t)

        has_q1 = any(t.startswith('1.') for t in frames)
        content = [t for t in frames if not t.startswith('1.') and not t.startswith('2.')]
        name = sp.split('/')[-1]
        print(f'{name:28s} q1={str(has_q1):5s} content_frames={len(content):2d}  {[c[:35] for c in content[:3]]}')
