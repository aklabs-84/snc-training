import pathlib,re
B=pathlib.Path(__file__).parent; S=B/'src'
css=(S/'common.css').read_text(); js=(S/'common.js').read_text().replace("if (typeof module!=='undefined') module.exports=U;","")
data=(B.parent/'data.js').read_text()
LOAD=(S/'load.html').read_text()
import os
for n,title in [(1,'1_정리집계앱'),(2,'2_분석시각화앱'),(3,'3_보고서앱')]:
    if not (S/f'app{n}.html').exists(): continue
    a=(S/f'app{n}.html').read_text().replace('<!--LOAD-->',LOAD)
    i=a.rindex('<script>'); app=a[i:]; a=a[:i]
    t=re.search(r'<title>(.*?)</title>',a).group(1)
    out=f'<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{t}</title><style>{css}</style></head><body>{a}<div id="tip"></div><script>{data}</script><script>{js}</script>{app}</body></html>'
    # app script marker
    (B/f'{title}.html').write_text(out); print(title,len(out))
