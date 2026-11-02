"""수업가이드 4개 문서 → 슬라이드 덱(deck.json) 만들기. swiss 템플릿, 51장.
실행: python3 make_deck.py  →  deck.json 생성 (렌더는 README 참고)"""
import json
IM='../images/'
def scr(sec,time,title,img,pts,cap,callout=None):
    d={"type":"screen","section":sec,"title":title,"image":IM+img+".png","frame":"browser","points":pts,"caption":cap}
    if time:d["time"]=time
    if callout:d["callout"]=callout
    return d
S=[]
a=S.append
a({"type":"cover","kicker":"에스앤씨 · 스마트제조 AI 활용 교육","title":"AI로 만드는 나만의 엑셀 정리·집계 도우미","sub":"정리 → 집계 → 분석 → 보고서까지, 매달 하는 일을 짧게","icon":"rocket","chips":["3시간","13명","무료 AI 도구"]})
a({"type":"roadmap","section":"오늘의 여정","time":"3시간","title":"3시간 동안 이렇게 해요","total":"180분","blocks":[
 {"label":"시작하기 + 정리·집계","minutes":90,"icon":"clipboard","items":["시작하기 30분","정리·집계 50분 + 휴식 10분"]},
 {"label":"AI 분석·시각화","minutes":45,"icon":"search","items":["분석 보고서 읽기","가려서 AI에게 묻기"]},
 {"label":"보고서·회의자료","minutes":45,"icon":"pencil","items":["웹페이지·PDF 35분","연습 문제 10분"]}]})
a({"type":"checklist","section":"준비물","time":"5분","title":"시작 전에 준비해요","items":["앱 3개 (1·2·3번 파일)","연습용 샘플 자료","Gemini 새 대화 창","연습 문제 (정답 쪽은 나중에)"],"done":0,"callout":"다 됐으면 *시작*해요"})
# ⓪ 시작하기
a({"type":"section","no":"0","title":"시작하기","sub":"AI의 기본 말과 안전한 약속 · 30분"})
a({"type":"steps","section":"STEP 1","time":"10분","title":"AX·생성형 AI·머신러닝, 한 줄로","lead":"회사 업무 예시로 *내 말*로 정리해요","steps":[
 {"t":"AX","d":"일하는 방식과 가치를 바꾸는 것","icon":"rocket"},
 {"t":"생성형 AI","d":"글·이미지·코드를 새로 만들어요","icon":"chat"},
 {"t":"머신러닝","d":"데이터에서 규칙을 스스로 배워요","icon":"book"}],"flow":False,"callout":"자격증 연습 문제 *1·2·3번*과 이어져요"})
a({"type":"quest","section":"STEP 1","title":"내 업무 예시를 적어 봐요","quests":[
 {"t":"매달 반복하는 엑셀 업무","d":"예: 지점별 월 매출 집계","icon":"clipboard"},
 {"t":"AI가 도우면 달라질 것","d":"예: 집계 시간, 숫자 실수","icon":"clock"}],"callout":"표에 *한 줄씩* 적어 보세요"})
a({"type":"compare","section":"STEP 2","time":"10분","title":"AI는 모를 때도 *지어내요*","bad":{"label":"그냥 물어보면","text":"우리 회사 7월 매출을 알려줘","note":"그럴듯한 숫자를 지어낼 수 있어요"},"good":{"label":"규칙 한 줄을 넣으면","text":"모르면 *모른다고 해*","note":"지어내는 일이 줄어요"},"callout":"지어내는 현상을 *환각*이라고 해요 (자격증 9번)"})
a({"type":"rules","section":"STEP 2","title":"회사 자료를 넣기 전 약속","left":{"title":"이렇게 해요","items":["이름은 *가려서* 넣어요","필요 없는 열은 *빼요*","*집계표*만 넣어요"],"marker":"check"},"right":{"title":"이건 안 돼요","tone":"warn","marker":"x","items":["거래처 이름을 그대로","비밀번호를 같이","원본 파일을 통째로"]},"callout":{"text":"자격증 *5번*: 민감한 정보는 가리거나 빼요","tone":"warn"}})
a({"type":"examples","section":"STEP 3","time":"10분","title":"요청문은 *6칸*으로 써요","groups":[
 {"t":"역할","tag":"누가","icon":"user"},{"t":"목표","tag":"무엇을","icon":"target"},{"t":"재료","tag":"자료","icon":"clipboard"},
 {"t":"할 일","tag":"구체적으로","icon":"pencil"},{"t":"규칙","tag":"하지 마","icon":"alert"},{"t":"마지막 검사","tag":"스스로 확인","icon":"checkCircle"}],"hi":4,"callout":"*프롬프트*는 AI에게 하는 지시문이에요 (자격증 4번)"})
a({"type":"prompt","section":"STEP 3","title":"예시 요청문, 네 조각으로 읽어요","parts":[
 {"label":"역할","t":"너는 경영지원팀 선배야","tone":"primary"},
 {"label":"재료","t":"7월 18,461,000원 / 8월 20,647,000원","tone":"blue"},
 {"label":"규칙","t":"숫자는 재료에 있는 것만 써","tone":"accent"},
 {"label":"마지막 검사","t":"쓴 숫자가 재료에 있는지 확인해","tone":"purple"}],"callout":"*하지 마* 규칙 한 줄이 가장 효과가 커요"})
# ① 정리·집계
a({"type":"section","no":"1","title":"엑셀 정리와 자동 집계","sub":"지저분한 표를 정리하고, 합계를 내고, 저장해요 · 60분"})
a(scr("STEP 1","5분","앱①을 열고 연습 자료를 올려요","앱1_01_연습자료_올린직후",["*1_정리집계앱.html*을 열어요","*매출 (수출입)* 버튼을 눌러요","*주황색 칸*이 이상한 곳이에요"],"앱① 화면 · 39줄 읽음"))
a({"type":"examples","section":"STEP 1","title":"주황색 칸은 이런 모양이에요","groups":[
 {"t":"날짜 모양이 섞임","tag":"46206 · 2026.07.08","icon":"clock"},{"t":"보이지 않는 빈칸","tag":"이름 뒤에 빈칸","icon":"eye"},
 {"t":"원·쉼표가 붙음","tag":"792,000원","icon":"alert"},{"t":"똑같은 줄","tag":"3줄","icon":"x"}],"callout":"4가지가 *섞이면* 합계가 틀려요"})
a(scr("STEP 2","8분","이상한 곳 *59곳*을 찾았어요","앱1_02_이상한곳_목록",["빈칸 *11곳*","날짜 *39곳*","원·쉼표 *6곳*"],"똑같은 줄 3줄은 따로 지워요"))
a(scr("STEP 2","","한 번에 고치면 *3곳*만 남아요","앱1_03_추천후_이상한곳3",["*추천대로 모두 고치기*를 눌러요","날짜·숫자·빈칸이 한 번에","남은 3곳은 *똑같은 줄*"],"내가 확인하고 지워요","자격증 *7번*: AI 분석 전에 *데이터 정리*가 먼저"))
a({"type":"compare","section":"STEP 3","time":"7분","title":"똑같은 줄 3줄이 합계를 *부풀려요*","bad":{"label":"지우기 전","text":"63,244,000원","note":"똑같은 줄 3줄 포함"},"good":{"label":"지운 뒤","text":"57,641,000원","note":"차이 *5,603,000원*"},"callout":"임원이 *틀린 숫자*로 결정하면 안 돼요"})
a(scr("STEP 4","8분","다르게 쓴 같은 값을 *하나로*","앱1_06_값목록_합치기",["서울·서울시·서울 본점에 *체크*","*서울*로 바꿔요","*체크한 값 합치기*를 눌러요"],"값 목록 탭"))
a({"type":"spotlight","section":"휴식","kicker":"BREAK","title":"*10분* 쉬어요","sub":"돌아오면 직접 고치기와 합계 내기를 해요","icon":"clock"})
a(scr("STEP 5","7분","바꾸기와 *되돌리기*","앱1_07_바꾸기탭",["B칸(거래처)을 *고르고*","*Sakura Parts*를 찾으면 11곳","잘못하면 *되돌리기* 버튼"],"바꾸기 탭"))
a({"type":"steps","section":"STEP 5","title":"직접 고치고, 되돌리고, 기록 보기","lead":"실수해도 *괜찮아요*","steps":[
 {"t":"칸 고치기","d":"수식줄에 쓰고 Enter","icon":"pencil"},{"t":"노란색 확인","d":"고친 칸이 노랗게 바뀌어요","icon":"eye"},
 {"t":"되돌리기","d":"맨 위 버튼은 항상 보여요","icon":"arrowRight"},{"t":"한 일 보기","d":"한 일·저장 탭에 기록돼요","icon":"clipboard"}],"flow":True,"callout":"직접 고친 것은 *저장 안 돼요*"})
a(scr("STEP 6","8분","월별·구분별 합계와 *그래프*","앱1_08_합계_구분별",["*일자*(월별)로 묶어요","더 나눠 볼까요? → *구분*","수출 *32,884,000* / 내수 *24,757,000*"],"선택만 하면 계산돼요"))
a({"type":"examples","section":"STEP 6","title":"정답 확인: 합계 숫자","groups":[
 {"t":"7월","tag":"18,461,000","icon":"clock"},{"t":"8월","tag":"20,647,000","icon":"clock"},{"t":"9월","tag":"18,533,000","icon":"clock"},
 {"t":"전체","tag":"57,641,000","icon":"star"},{"t":"수출","tag":"32,884,000","icon":"globe"},{"t":"내수","tag":"24,757,000","icon":"home"}],"hi":3,"callout":"가로·세로 합이 같으면 *표가 맞아요*"})
a(scr("STEP 7","7분","정리 방법을 *저장*해요","앱1_09_정리방법_저장",["이름을 쓰고 *저장하기*","*2가지*와 합계 설정이 저장돼요","다른 컴퓨터는 *파일로 받기*"],"내 정리 방법"))
a({"type":"steps","section":"STEP 7","title":"다음 달 파일에 *한 번에* 적용","lead":"첫 달에 만들고, 다음 달은 *넣기만* 해요","steps":[
 {"t":"새 엑셀 준비","d":"같은 모양의 다음 달 파일","icon":"book"},{"t":"방법 고르기","d":"올리기 화면의 내 정리 방법","icon":"sliders"},
 {"t":"파일 올리기","d":"올리면 자동으로 정리돼요","icon":"play"},{"t":"합계 확인","d":"합계 설정도 그대로예요","icon":"checkCircle"}],"flow":True,"callout":"칸은 *제목*으로 찾아서 위치가 바뀌어도 돼요"})
a({"type":"compare","section":"STEP 7","title":"효과는 *숫자*로 확인해요","rows":[
 {"tag":"시간","icon":"clock","bad":{"label":"손으로","text":"25분","note":"처음부터 다시"},"good":{"label":"앱으로","text":"2분","note":"저장한 방법 적용"}},
 {"tag":"오류","icon":"alert","bad":{"label":"손으로","text":"3줄 놓침","note":"합계가 부풀어요"},"good":{"label":"앱으로","text":"0건","note":"똑같은 줄 확인"}}],"callout":"자격증 *8번*: 도입 전후의 시간·오류를 *비교*해요"})
# ② AI 분석
a({"type":"section","no":"2","title":"AI로 분석하고 보기","sub":"계산은 앱이, 해석은 AI와 함께 · 45분"})
a(scr("STEP 1","7분","앱②의 *분석 보고서*","앱2_01_분석보고서_위쪽",["정리된 *csv*를 올려요","한 줄 결론: *합계 57,641,000원*","AI 없이 *계산*해요"],"앱② 화면"))
a(scr("STEP 2","8분","변화는 *어디에서* 왔나","앱2_02_변화는어디에서",["7월에서 8월로 *+2,186,000원*","수출 *+3,909,000*","내수 *−1,723,000*"],"앱이 계산으로 알려 줘요"))
a({"type":"compare","section":"STEP 2","title":"앱이 아는 것과 *짐작*할 것","bad":{"label":"앱이 아는 것","text":"무엇이, 얼마나 변했나","note":"계산으로 확인돼요"},"good":{"label":"사람·AI가 짐작","text":"*왜* 변했나","note":"숫자만으로는 몰라요"},"callout":"그래서 *AI의 도움*을 받아요"})
a(scr("STEP 3","7분","이름을 *가려서* 보내요","앱2_03_가리기와요청문",["*거래처 가리기*에 체크","원본 말고 *집계표*만","*금액도 가리기*는 선택"],"AI 더 물어보기"))
a({"type":"examples","section":"STEP 3","title":"가명표는 *내 화면 안*에만 있어요","groups":[
 {"t":"Sakura Parts","tag":"거래처A","icon":"user"},{"t":"Nova Tech","tag":"거래처B","icon":"user"},{"t":"한빛전자","tag":"거래처C","icon":"user"},
 {"t":"대성정밀","tag":"거래처D","icon":"user"},{"t":"미래모터스","tag":"거래처E","icon":"user"}],"callout":"자격증 *5번*: 민감한 정보는 가리거나 빼요"})
a({"type":"prompt","section":"STEP 4","time":"10분","title":"앱이 만든 요청문을 읽어 봐요","parts":[
 {"label":"역할","t":"분석 선배","tone":"primary"},{"label":"재료","t":"가린 집계표 (거래처A~E)","tone":"blue"},
 {"label":"질문","t":"9월에 줄어든 이유는?","tone":"accent"},{"label":"규칙","t":"재료의 숫자만, 짐작은 추정이라고","tone":"purple"}],"callout":"6칸 틀을 *그대로* 써요"})
a({"type":"steps","section":"STEP 4","title":"Gemini에게 묻고 답을 붙여넣어요","lead":"복사 → 질문 → 붙여넣기","steps":[
 {"t":"요청문 복사","d":"앱②의 복사하기 버튼","icon":"clipboard"},{"t":"Gemini에 질문","d":"새 대화에 붙여넣고 Enter","icon":"chat"},
 {"t":"답 복사","d":"답 전체를 복사해요","icon":"pencil"},{"t":"앱에 붙여넣기","d":"가명이 원래 이름으로","icon":"checkCircle"}],"flow":True,"callout":"답은 *관찰·근거·추정*으로 나뉘어요"})
a(scr("STEP 5","8분","틀린 숫자를 앱이 *잡아내요*","앱2_05_숫자검사",["일부러 넣은 *58,000,000원*","표에 없는 숫자를 *찾아요*","정답은 *57,641,000원*"],"숫자 검사"))
a({"type":"spotlight","section":"STEP 5","kicker":"CHECK","title":"AI도 틀려요. 검산은 *내가* 해요","sub":"그럴듯하다고 그대로 쓰지 않아요","icon":"alert","callout":"자격증 *9번 환각, 10번 검산*"})
a({"type":"examples","section":"STEP 6","time":"5분","title":"질문에 맞는 *그래프*","groups":[
 {"t":"변화","tag":"꺾은선·세로 막대","icon":"target"},{"t":"비교","tag":"막대","icon":"sliders"},{"t":"순위","tag":"큰 순서 가로 막대","icon":"flag"},
 {"t":"구성","tag":"100% 쌓은 막대","icon":"globe"},{"t":"관계","tag":"점 그래프","icon":"search"}],"callout":"막대 그래프의 세로축은 *0부터*"})
# ③ 보고서·회의자료
a({"type":"section","no":"3","title":"보고서와 회의자료","sub":"웹페이지로 만들고 PDF로 저장 · 35분 + 문제 10분"})
a(scr("STEP 1","5분","앱③ 설정하기","앱3_01_설정",["*읽는 사람*과 분량을 골라요","제목·기간·강조할 점","*거래처 가리기* 확인"],"설정 화면"))
a(scr("STEP 2","8분","*내용*부터 만들어요","앱3_02_요청문_내용먼저",["*내용 먼저 만들기*를 골라요","요청문을 복사해 Gemini에","같은 대화에서 이어서 해요"],"요청문 화면"))
a({"type":"steps","section":"STEP 2~4","title":"내용 → 검산 → 웹페이지 → PDF","lead":"순서를 *지키면* 안전해요","steps":[
 {"t":"내용 만들기","d":"보고 요약 + 구성안","icon":"pencil"},{"t":"숫자 검산","d":"검산표를 앱② 숫자와 비교","icon":"checkCircle"},
 {"t":"웹페이지","d":"보고서·회의자료 요청문","icon":"globe"},{"t":"PDF 저장","d":"Ctrl+P, PDF로 저장","icon":"book"}],"flow":True})
a({"type":"spotlight","section":"STEP 2","kicker":"VERIFY","big":"0개","title":"검산표의 *근거 없음*은 0개","sub":"하나라도 있으면 그 문장을 고쳐 달라고 해요","callout":"자격증 *10번*: 사람이 사실과 숫자를 *확인*해요"})
a({"type":"rules","section":"STEP 3","time":"8분","title":"인쇄 규칙 한 줄이 PDF를 지켜요","left":{"title":"요청문에 넣어요","items":["A4 세로, 배경은 *흰색*","그래프가 쪽 사이에서 *안 잘리게*","화면 *버튼은 숨기기*"],"marker":"check"},"right":{"title":"안 넣으면","tone":"warn","marker":"x","items":["그래프가 *반으로* 잘려요","버튼이 *인쇄*돼요","잉크가 *낭비*돼요"]},"callout":{"text":"Canvas에서 인쇄가 막히면 *.html로 저장*해서 열어요","tone":"warn"}})
a(scr("STEP 3","","되돌린 웹페이지를 *받아요*","앱3_03_다듬기_웹페이지",["코드를 *붙여넣으면*","가명이 *원래 이름*으로","*.html 받기* 후 Ctrl+P"],"다듬기 화면"))
a({"type":"mixed","section":"STEP 4","time":"7분","title":"회의자료 5장은 *결론 문장* 제목","lead":"한 장에 *한 가지*만 말해요","main":{"t":"가로 PDF 5쪽","d":"방향은 가로, 한 장이 한 쪽","icon":"book"},"items":[
 {"t":"1장 합계","d":"57,641,000원","icon":"flag"},{"t":"2장 월별","d":"8월이 가장 커요","icon":"clock"},{"t":"3장 구성","d":"수출이 더 커요","icon":"globe"}],"callout":"제목만 읽어도 *이야기*가 이어져요"})
a({"type":"steps","section":"STEP 5","time":"7분","title":"결과물을 *파헤쳐* 요청문을 만들어요","lead":"부품으로 쪼개면 *문장*이 보여요","steps":[
 {"t":"영역","d":"위에서 아래 덩어리","icon":"clipboard"},{"t":"눌러서 되는 것","d":"버튼과 입력칸","icon":"play"},
 {"t":"숫자·겉모습","d":"계산식, 글자 크기","icon":"sliders"},{"t":"안 보이는 약속","d":"인터넷 없이, 오류 처리","icon":"lock"}],"flow":False,"callout":"틀린 곳은 *규칙 한 줄*로 고쳐요"})
a({"type":"compare","section":"STEP 5","title":"막연한 말을 *재 볼 수 있게*","bad":{"label":"막연한 말","text":"글씨 크게, 예쁘게","note":"AI가 마음대로 정해요"},"good":{"label":"재 볼 수 있는 말","text":"*18px 이상*, 색 *3가지*","note":"그대로 나와요"},"callout":"*하지 마* 규칙도 꼭 넣어요"})
a({"type":"steps","section":"STEP 5","title":"AX 추진 *4단계*","lead":"작게 시험하고 *넓혀요*","steps":[
 {"t":"현황 진단","d":"지금 업무와 데이터를 봐요","icon":"search"},{"t":"과제 발굴","d":"우선순위를 정해요","icon":"target"},
 {"t":"시범 적용","d":"작게 먼저 해 봐요","icon":"rocket"},{"t":"확산·정착","d":"효과가 보이면 넓혀요","icon":"users"}],"flow":True,"callout":"자격증 *6번*: 진단, 발굴, 시범, 확산"})
a({"type":"quest","section":"STEP 5","title":"내 업무 로드맵을 적어 봐요","quests":[
 {"t":"현황 진단","d":"지금 걸리는 시간·오류","icon":"search"},{"t":"과제 발굴","d":"1순위 한 가지","icon":"target"},
 {"t":"시범 적용","d":"이번 달 한 번 해 보기","icon":"rocket"},{"t":"확산·정착","d":"나눌 사람과 방법","icon":"users"}],"callout":"내일 바로 할 일 *한 가지*를 정해요"})
a({"type":"spotlight","section":"STEP 6","time":"10분","kicker":"QUIZ","big":"10문제","title":"연습 문제를 *풀어 봐요*","sub":"정답 쪽은 *다 푼 뒤*에 열어요","icon":"pencil"})
a({"type":"examples","section":"STEP 6","title":"틀린 문제는 *어디서* 다시 볼까","groups":[
 {"t":"1·2·3번","tag":"시작하기 STEP 1","icon":"book"},{"t":"4·5번","tag":"시작하기 STEP 2~3","icon":"lock"},{"t":"6번","tag":"AX 추진 4단계","icon":"flag"},
 {"t":"7·8번","tag":"정리·집계 STEP 2·7","icon":"sliders"},{"t":"9·10번","tag":"AI 분석 STEP 5","icon":"alert"}],"callout":"점수와 *틀린 이유*를 한 줄로 적어요"})
a({"type":"wrapup","kicker":"WRAP UP","title":"오늘 하루의 흐름","flow":[
 {"label":"정리·집계","icon":"clipboard"},{"label":"가려서 분석","icon":"lock"},{"label":"검산","icon":"checkCircle"},{"label":"보고서·회의자료","icon":"pencil"}],"prompt":"오늘 배운 방법으로 *내 업무* 하나를 정해요"})
# swiss 템플릿의 번호 칸(examples)은 제목(t)만 그려서, 보조 글(tag)을 제목 아래 줄로 합쳐요
for sl in S:
    if sl.get('type')=='examples':
        for g in sl['groups']:
            if 'tag' in g:g['t']=g['t']+'\n'+g.pop('tag')
        if len(sl['groups'])==4:sl['cols']=2
deck={"template":"swiss","purpose":"lesson","title":"AI로 만드는 나만의 엑셀 정리·집계 도우미","mascot":False,"slides":S}
json.dump(deck,open('deck.json','w',encoding='utf8'),ensure_ascii=False,indent=1)
print(len(S),'장')
