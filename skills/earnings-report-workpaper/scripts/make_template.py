"""Create a neutral Word draft template; no institutional branding."""
import argparse
from pathlib import Path
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt


def create(path):
    path = Path(path)
    if path.exists():
        raise ValueError('Template exists; choose a new path')
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)
    style.element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
    for alias, count in [('标题',1),('副标题',1),('报告日期',1),('投资评级',1),('评级变动',1),('事件公告',1),('摘要',8),('风险提示',1)]:
        sdt = OxmlElement('w:sdt')
        props = OxmlElement('w:sdtPr')
        a = OxmlElement('w:alias'); a.set(qn('w:val'), alias); props.append(a)
        sdt.append(props)
        content = OxmlElement('w:sdtContent')
        for _ in range(count):
            p=OxmlElement('w:p');r=OxmlElement('w:r');t=OxmlElement('w:t');t.text='文字待填充'
            r.append(t);p.append(r);content.append(p)
        sdt.append(content)
        doc.element.body.insert(len(doc.element.body)-1, sdt)
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)


if __name__ == '__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', required=True)
    create(parser.parse_args().output)
