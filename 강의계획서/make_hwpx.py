"""강의계획서 hwpx 만들기 (양식 빈칸 채우기 + 칸 번호 바로잡기).
양식 도구(스킬)의 채우기 함수를 쓰고, 그 결과에서 여러 줄로 복제된 커리큘럼 줄의 rowAddr(줄 번호)를
0부터 차례대로 다시 매겨요. 줄 번호가 겹치면 한글에서 파일이 열리지 않아요.
실행: python3 make_hwpx.py"""
import importlib.util,zipfile,re,glob,os,sys
_home=os.path.expanduser('~')
_found=(glob.glob('/root/.claude/skills/synced/*/hwpx-lecture-planner')
        +glob(_home+'/.claude/skills/**/hwpx-lecture-planner',recursive=True)
        +glob(_home+'/.claude/plugins/**/hwpx-lecture-planner',recursive=True))
if not _found:sys.exit('hwpx-lecture-planner 스킬(양식 파일)을 찾지 못했어요. 스킬 폴더 위치를 SK 변수에 직접 적어 주세요.')
SK=_found[0]
spec=importlib.util.spec_from_file_location("fill_hwpx",SK+"/scripts/fill_hwpx.py");mod=importlib.util.module_from_spec(spec);spec.loader.exec_module(mod)
TEMPLATE=SK+"/assets/template.hwpx"
OUT=os.path.join(os.path.dirname(os.path.abspath(__file__)),"강의계획서_완성본_20261028.hwpx")
DATA={
 "강좌명":"AI로 만드는 나만의 엑셀 정리·집계 도우미 (업무 보고 자동화)",
 "수업장소":"에스앤씨 본사 회의실",
 "강의기간":"2026. 11. 3.(화) 10:00 ~ 13:00 (3시간)",
 "교과목개요":"제조업 사무 업무에서 매달 되풀이되는 엑셀 데이터 정리·집계, 분석·시각화, 보고서·회의자료 작성을 무료 생성형 AI(ChatGPT·Gemini)와 인터넷 없이 열리는 간단한 도구를 활용해 효율화하는 실습 과정입니다. 참가자는 지저분한 엑셀을 정리·집계하는 도구를 직접 써 보고, 거래처명 등 민감한 자료는 가려서 AI에게 안전하게 분석을 맡기는 방법과, 원하는 결과물을 거꾸로 파헤쳐 구조적인 요청문(프롬프트)을 스스로 만드는 방법을 익힙니다.",
 "학습목표":"① 지저분한 엑셀 데이터를 도구로 정리·집계하고, 같은 정리 방법을 다음 달 파일에 다시 적용할 수 있다. ② 집계 결과에서 민감한 정보를 가려 AI에게 넘겨 분석·시각화를 하고, AI가 낸 숫자를 직접 검산할 수 있다. ③ 분석 결과를 업무 보고서와 회의자료로 정리하고 PDF로 저장할 수 있다. ④ 결과물을 부품으로 쪼개어 구조적인 요청문을 스스로 만들고, 틀린 곳을 규칙으로 고쳐 나갈 수 있다.",
 "수업방법":"강사 시연 후 1인 1대 PC 직접 실습(단계별 따라 하기 가이드와 완성본 제공). 연습용 샘플 데이터 사용, 무료 ChatGPT·Gemini 활용, 회사 자료는 이름·금액을 가려서 사용. 마무리에서 내 업무에 바로 쓸 곳 1가지를 정해 봅니다.",
 "커리큘럼":[
  {"번호":"1","주제명":"시작하기: AI에게 일 시키는 법과 안전하게 쓰는 약속","교육내용":"수업 목표와 오늘의 결과물 소개, 무료 AI 사용 준비, 회사 자료를 가려서 쓰는 약속(이름 가리기), 요청문 6칸 틀(역할·목표·재료·할 일·규칙·마지막 검사) 맛보기","시수":"0.5","담당강사":"변모세"},
  {"번호":"2","주제명":"엑셀 데이터 정리 및 자동 집계","교육내용":"지저분한 엑셀(날짜·숫자·같은 값·중복·빈칸) 정리 도구 실습, 월별·구분별 합계 내기, '내 정리 방법' 저장 후 다음 달 파일에 한 번에 적용","시수":"1","담당강사":"변모세"},
  {"번호":"3","주제명":"AI를 활용한 데이터 분석 및 시각화","교육내용":"분석 보고서(추이·변화의 출처·구성과 쏠림) 읽기, 이름을 가린 집계표로 AI에게 분석 질문하기, 질문에 맞는 그래프 고르기, AI 답의 숫자 검산","시수":"0.75","담당강사":"변모세"},
  {"번호":"4","주제명":"업무 보고서 및 회의자료 작성","교육내용":"결과물을 파헤쳐 요청문 만들기, Gemini로 보고서·회의자료 웹페이지 만들고 PDF로 저장, 내 업무에 적용할 곳 정하기","시수":"0.75","담당강사":"변모세"}]}
TMP="/tmp/_hwpx_stage.hwpx"
mod.fill_lecture_plan(TEMPLATE,TMP,DATA)
src=zipfile.ZipFile(TMP);tpl=zipfile.ZipFile(TEMPLATE)
xml=src.read('Contents/section0.xml').decode('utf8')
# 줄 번호(rowAddr)를 줄 순서대로 0,1,2,... 로 다시 매김
rows=re.findall(r'<hp:tr>.*?</hp:tr>',xml,re.S)
for i,tr in enumerate(rows):
    new=re.sub(r'(<hp:cellAddr colAddr="\d+" rowAddr=")\d+(")',lambda m:f'{m.group(1)}{i}{m.group(2)}',tr)
    xml=xml.replace(tr,new,1)
rows=re.findall(r'<hp:tr>.*?</hp:tr>',xml,re.S)
xml=re.sub(r'rowCnt="\d+"',f'rowCnt="{len(rows)}"',xml,count=1)
# 양식과 같은 순서·같은 압축 방식으로 다시 묶기 (mimetype은 맨 앞, 압축 안 함)
with zipfile.ZipFile(OUT,'w') as z:
    for info in tpl.infolist():
        data=xml.encode('utf8') if info.filename=='Contents/section0.xml' else src.read(info.filename)
        z.writestr(zipfile.ZipInfo(info.filename,date_time=info.date_time),data,compress_type=info.compress_type)
print('만들었어요:',OUT)
