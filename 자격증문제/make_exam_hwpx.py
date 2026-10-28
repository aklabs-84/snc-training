"""AX 전문컨설턴트 3급 연습 문제(4지선다 10문항) + 정답·해설을 hwpx 한 파일로 만들기.
앞쪽에 문제, 다음 쪽부터 정답. 실행: python3 make_exam_hwpx.py"""
from hwpx import HwpxDocument

QS=[ # (문제, [보기4], 정답번호, 해설)
("AX(AI Transformation, 인공지능 전환)에 대한 설명으로 가장 알맞은 것은?",
 ["종이 문서를 컴퓨터 파일로 바꾸는 일","인공지능을 업무·서비스·조직 전반에 도입해 일하는 방식과 가치를 바꾸는 일","컴퓨터 한 대에 새 프로그램을 설치하는 일","직원 수를 줄이는 일"],2,
 "AX는 AI를 일부 업무에 써 보는 데서 그치지 않고 일하는 방식과 가치를 바꾸는 것이에요."),
("생성형 AI에 대한 설명으로 알맞은 것은?",
 ["데이터를 저장하는 장치","정해진 규칙만 되풀이해서 실행하는 프로그램","글, 이미지, 코드처럼 새로운 결과물을 만들어 내는 인공지능","인터넷 속도를 높이는 기술"],3,
 "질문·지시(프롬프트)를 바탕으로 글, 이미지, 코드 등을 새로 만드는 AI예요. (ChatGPT, Gemini 등)"),
("머신러닝(기계학습)을 가장 잘 설명한 것은?",
 ["사람이 모든 규칙을 일일이 프로그래밍하는 방식","컴퓨터가 많은 데이터를 보고 스스로 규칙과 패턴을 배우는 방식","기계를 사람이 직접 조작하는 방식","컴퓨터를 오래 켜 두어도 고장 나지 않게 하는 방식"],2,
 "데이터에서 규칙과 패턴을 스스로 학습해 예측·판단하는 기술이에요. ‘일일이 프로그래밍’은 전통적인 방식이에요."),
("생성형 AI에게 원하는 일을 시키려고 입력하는 질문이나 지시문을 무엇이라고 하는가?",
 ["프로세서","프롬프트","프로토콜","프로그램"],2,
 "프롬프트(prompt)는 AI에게 하는 질문·지시문이에요. 구체적으로 쓸수록 원하는 결과를 얻기 쉬워요."),
("회사 자료를 외부의 무료 AI 서비스에 입력해야 할 때 가장 알맞은 행동은?",
 ["거래처 이름과 금액을 그대로 모두 입력한다","이름·금액 같은 민감한 정보는 가리거나 빼고 입력한다","비밀번호를 함께 입력해서 AI가 파일을 열게 한다","누가 올렸는지 알 수 없으니 신경 쓰지 않는다"],2,
 "외부 AI에 입력한 내용은 회사 밖으로 나가요. 민감한 정보는 가리거나(마스킹) 빼고, 회사 보안 규정을 먼저 확인해요."),
("기업이 AX를 추진하는 순서로 가장 알맞은 것은?",
 ["확산·정착 → 시범 적용 → 과제 발굴 → 현황 진단","시범 적용 → 현황 진단 → 확산·정착 → 과제 발굴","현황 진단 → 과제 발굴(우선순위 정하기) → 시범 적용 → 확산·정착","과제 발굴 → 확산·정착 → 현황 진단 → 시범 적용"],3,
 "진단 → 과제 발굴·우선순위 → 작게 시범 적용(PoC) → 효과 확인 후 확산·정착 순서가 일반적이에요."),
("AI에게 엑셀 자료를 분석시키기 전에 가장 먼저 해야 할 일로 알맞은 것은?",
 ["그래프 색을 예쁘게 고른다","자료의 중복, 빈칸, 표기 흔들림(예: 서울/서울시)을 정리한다","파일 이름을 영어로 바꾼다","자료를 최대한 많이 늘려서 넣는다"],2,
 "데이터가 지저분하면 AI도 틀린 결과를 내요. 중복 삭제, 날짜·숫자 모양 통일, 같은 뜻의 값 합치기가 첫 단계예요."),
("AI 도입의 효과를 확인하는 방법으로 가장 알맞은 것은?",
 ["직원들이 좋아하는 것 같다고 추측한다","도입 전과 후의 업무 시간, 오류 건수 같은 숫자를 비교한다","AI 서비스를 쓰는 사람 수가 늘었다는 것만 본다","다른 회사가 도입했다는 사실만 확인한다"],2,
 "효과는 측정할 수 있는 지표로 확인해요. 예: 월 마감 시간 8시간 → 3시간, 입력 오류 건수 감소."),
("생성형 AI가 사실이 아닌 내용을 사실처럼 그럴듯하게 만들어 내는 현상을 무엇이라고 하는가?",
 ["암호화","압축","환각(할루시네이션)","과금"],3,
 "‘환각(hallucination)’이라고 해요. 그래서 AI의 답은 사람이 사실과 숫자를 확인해야 해요."),
("AI가 만들어 준 결과(글, 숫자, 그래프)를 업무에 쓰기 전에 가장 알맞은 행동은?",
 ["그럴듯해 보이면 그대로 보고서에 쓴다","사람이 사실과 숫자를 직접 확인(검산)한 뒤 쓴다","AI에게 ‘틀리지 마’라고 한 번 더 말하면 확인하지 않아도 된다","AI가 쓴 것이므로 책임은 AI에게 있다"],2,
 "AI는 틀릴 수 있고, 업무 결과의 책임은 사람에게 있어요. 쓰기 전에 사실과 숫자를 확인해요."),
]
# 정답 번호가 한 곳에 몰리지 않게 정답 위치를 아래대로 맞춤
WANT=[2,3,1,4,3,4,1,2,3,1]
Q=[]
for (q,opts,ans,exp),w in zip(QS,WANT):
    o=opts[:];o[ans-1],o[w-1]=o[w-1],o[ans-1];Q.append((q,o,w,exp))
CIR='①②③④'
TITLE='인공지능전환(AX) 전문컨설턴트 3급 연습 문제'

doc=HwpxDocument.new()
doc.set_page_setup(paper_size='A4',margins_mm={'left':20,'right':20,'top':18,'bottom':18})
S={k:doc.ensure_run_style(**v) for k,v in {
  'title':dict(bold=True,size=18),'sub':dict(size=11),'h':dict(bold=True,size=12.5),
  'q':dict(bold=True,size=11),'o':dict(size=10.5),'n':dict(size=10),'b':dict(bold=True,size=10.5)}.items()}
first=[True]
def put(text,style,align='LEFT',before=0,after=3,left=0,pbreak=False,keep=False):
    if first[0]:
        p=doc.paragraphs[0];p.text=text;first[0]=False
        r=p.runs[0] if getattr(p,'runs',None) else None
        if r is not None:r.char_pr_id_ref=S[style]
    else:
        p=doc.add_paragraph(text,char_pr_id_ref=S[style])
    idx=len(doc.paragraphs)-1
    kw=dict(paragraph_index=idx,alignment=align,spacing_before_pt=before,spacing_after_pt=after,line_spacing_percent=140)
    kw['indent_left_mm']=left
    if pbreak:kw['page_break_before']=True
    if keep:kw['keep_with_next']=True
    doc.set_paragraph_format(**kw)
# ── 문제지
put(TITLE,'title','CENTER',0,4)
put('(문제지)','sub','CENTER',0,10)
put('응시자 이름: ____________________        점수: ________ / 100점','sub','LEFT',0,3)
put('시험 시간: 20분   ·   4지선다 10문항 (문항당 10점, 100점 만점)','n','LEFT',0,10)
put('※ 각 문항에서 알맞은 것을 하나만 고르시오.','h','LEFT',0,8)
for i,(q,opts,ans,exp) in enumerate(Q,1):
    put(f'{i}. {q}','q','LEFT',6,2,0,False,True)
    for j,o in enumerate(opts):put(f'{CIR[j]} {o}','o','LEFT',0,1,6)
# ── 정답 (새 쪽에서 시작)
put(TITLE,'title','CENTER',0,4,0,True)
put('(정답 및 해설 · 강사용)','sub','CENTER',0,10)
put('1. 정답표','h','LEFT',0,6)
t=doc.add_table(rows=2,cols=11)
cells=[['문항']+[f'{i}번' for i in range(1,11)],['정답']+[CIR[a-1] for (_,_,a,_) in Q]]
for r in range(2):
    for c in range(11):
        t.set_cell_text(r,c,cells[r][c])
put('','n','LEFT',0,4)
put('2. 해설','h','LEFT',6,6)
for i,(q,opts,ans,exp) in enumerate(Q,1):
    put(f'{i}번  정답 {CIR[ans-1]}  {opts[ans-1]}','b','LEFT',4,1,0,False,True)
    put(exp,'n','LEFT',0,2,4)
put('※ 새로 만든 연습 문제이며 실제 자격시험의 출제 범위와 다를 수 있어요. 합격선은 60점 이상을 권해요.','n','LEFT',8,0)
out='AX전문컨설턴트3급_연습문제_문제와정답.hwpx'
doc.save_to_path(out)
print('저장:',out)
r=doc.validate();print('검사:',r)
