"""Generate an entirely fictional, offline four-section test case."""
import argparse
import csv
import json
import subprocess
import sys
from pathlib import Path
from make_template import create
from freeze_word import freeze


def create_demo(base):
    base=Path(base).resolve();base.mkdir(parents=True,exist_ok=True)
    create(base/'template.docx')
    bodies=['演示公司本季收入为100万元，同比增长25%。','订阅业务收入为60万元，占总收入60%。',
            '演示AI产品付费客户达到20家。','毛利率为70%，同比提升5个百分点。']
    titles=['事件','投资要点第一段','投资要点第二段','投资要点第三段']
    originals=['Quarterly revenue was RMB 1 million, up 25% year over year.',
              'Subscription revenue was RMB 600,000, representing 60% of total revenue.',
              'The demo AI product reached 20 paying customers.',
              'Gross margin was 70%, up 5 percentage points year over year.']
    md='# 演示公司：收入增长，产品应用扩展\n\n## 事件\n\n'+bodies[0]+'\n\n## 投资要点\n\n'
    for i in range(1,4): md+='▌'+titles[i]+'\n\n'+bodies[i]+'\n\n'
    md+='## 投资建议\n\n虚构案例，仅用于工具测试，不提供投资评级。\n\n## 风险提示\n\n演示数据不代表任何真实公司。\n'
    (base/'report.md').write_text(md,encoding='utf-8')
    subprocess.run([sys.executable,str(Path(__file__).with_name('fill_report_template.py')),
        '--markdown',str(base/'report.md'),'--template',str(base/'template.docx'),
        '--output',str(base/'report.docx'),'--date','演示日期','--rating','不适用','--rating-change','不适用'],check=True)
    evidence={};sections=[]
    for i,(title,body) in enumerate(zip(titles,bodies)):
        eid=f'source_{i}'
        original=originals[i]
        (base/f'{eid}.html').write_text('<html><head><meta charset="utf-8"></head><body><h1>Fictional earnings release</h1><p>'+original+'</p><p>Synthetic test data. Not a real company.</p></body></html>',encoding='utf-8')
        evidence[eid]={'source':f'{eid}.html','label':f'{eid}.html','terms':[original],
                       'source_language':'en','source_excerpt':original,'translation_zh':body}
        sections.append({'sheet':title,'body':body,'overview_evidence':eid,'points':[{'text':body,'evidence_ids':[eid], 'review_inputs':'来源1：虚构公告中的本期数据、期间和口径。', 'review_result':'直接披露值与来源1对应；该演示不用于真实投资判断。'}]})
    cfg={'company_dir':'.','word_path':'report.docx','word_sha256':freeze(base/'report.docx')['word_sha256'],
         'word_confirmed':True,'output_xlsx':'workpaper.xlsx','process_dir':'process','evidence':evidence,'sections':sections}
    # Synthetic test only. Never copy this approval flag into a real report.
    (base/'workpaper.json').write_text(json.dumps(cfg,ensure_ascii=False,indent=2),encoding='utf-8')
    with (base/'sources.csv').open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.writer(f);w.writerow(['sentence_id','section','text','source_file','location','status'])
        for i,body in enumerate(bodies):w.writerow([f'S-{i+1}',titles[i],body,f'source_{i}.html','第1段','虚构测试'])
    return base


if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('--output-dir',required=True)
    print(create_demo(p.parse_args().output_dir))
