"""Export the exact Word hash and paragraphs; does not grant human approval."""
import argparse
import hashlib
import json
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


def freeze(path):
    raw = Path(path).read_bytes()
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read('word/document.xml'))
    ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    paragraphs = [''.join(n.text or '' for n in p.iter(ns+'t')).strip() for p in root.iter(ns+'p')]
    return {'word_sha256': hashlib.sha256(raw).hexdigest(),
            'paragraphs': [p for p in paragraphs if p], 'word_confirmed': False}


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('word')
    parser.add_argument('--output', required=True)
    args = parser.parse_args()
    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open('x', encoding='utf-8') as f:
        json.dump(freeze(args.word), f, ensure_ascii=False, indent=2)
