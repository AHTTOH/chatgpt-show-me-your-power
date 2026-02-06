const fs = require("fs");
const path = require("path");
const pptxgen = require(path.join(__dirname, "..", "vendor", "pptxgenjs"));

const OUTPUT_DIR = path.join(__dirname, "..", "output");
const OUTPUT_FILE = path.join(OUTPUT_DIR, "antonai_chatgpt_deepdive.pptx");
const OUTPUT_BASE64 = path.join(
  OUTPUT_DIR,
  "antonai_chatgpt_deepdive.pptx.b64.txt"
);

if (!fs.existsSync(OUTPUT_DIR)) {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "AntonAI";
pptx.company = "AntonAI";
pptx.subject = "ChatGPT Deep Dive";
pptx.theme = {
  headFontFace: "Arial",
  bodyFontFace: "Arial",
  lang: "ko-KR"
};

const COLORS = {
  text: "000000",
  accent: "2B6CB0",
  bg: "FFFFFF"
};

const TITLE_STYLE = {
  x: 0.6,
  y: 0.2,
  w: 12.1,
  h: 0.6,
  fontSize: 28,
  bold: true,
  color: COLORS.text
};

const BODY_STYLE = {
  x: 0.9,
  y: 1.4,
  w: 11.6,
  h: 5.4,
  fontSize: 20,
  color: COLORS.text,
  valign: "top"
};

function addHeader(slide, title) {
  slide.background = { color: COLORS.bg };
  slide.addShape(pptx.ShapeType.line, {
    x: 0.6,
    y: 0.95,
    w: 12.1,
    h: 0,
    line: { color: COLORS.accent, width: 1 }
  });
  slide.addText(title, TITLE_STYLE);
}

function addBullets(slide, lines) {
  slide.addText(lines.map((line) => `• ${line}`).join("\n"), BODY_STYLE);
}

const slides = [
  {
    title: "ChatGPT 사용법을 뼛속까지 (안똔AI)",
    subtitle: "Apps / 개발자설정 / MCP / 자동화 / 실전 데모"
  },
  {
    title: "오늘 얻어갈 것 5가지",
    bullets: [
      "ChatGPT 화면 구성 요소를 한 장으로 이해",
      "Apps와 모델의 역할 분담을 구분",
      "개발자 설정으로 생산성 올리기",
      "MCP 연결과 안전 규칙 습득",
      "실전 데모로 자동화 흐름 설계"
    ]
  },
  {
    title: "ChatGPT 화면 구조 한 장 정리",
    bullets: [
      "모델: 답변의 두뇌 (GPT 계열 선택)",
      "툴: 파일, 브라우저, 코드 실행 등 확장",
      "앱: 외부 서비스와 연결되는 UI",
      "파일: 참고 자료/근거를 직접 넣는 방식",
      "프로젝트: 작업 컨텍스트와 히스토리 묶음"
    ]
  },
  {
    title: "Apps: 앱이 하는 일 vs 모델이 하는 일",
    bullets: [
      "앱: Canva, Notion 등 UI/데이터 연결",
      "모델: 요약/분석/생성 같은 사고 작업",
      "앱은 권한/데이터를 가져오고",
      "모델은 지시문을 따라 결과를 만들기",
      "둘의 경계를 이해하면 결과가 안정적"
    ]
  },
  {
    title: "개발자 설정: 뭘 켜면 뭐가 달라질까",
    bullets: [
      "도구 허용: 모델이 외부 기능 호출",
      "파일 처리: PDF/CSV 분석 속도 개선",
      "메모리/프로젝트: 장기 컨텍스트 유지",
      "실험 기능: 베타 기능 빠르게 체험",
      "권한 설정: 안전 범위 명확히 하기"
    ]
  },
  {
    title: "MCP: 개념 + 안전하게 쓰는 룰",
    bullets: [
      "MCP = 모델이 도구를 호출하는 연결 규칙",
      "서버/클라이언트로 기능 확장",
      "권한 최소화: 꼭 필요한 도구만 연결",
      "민감 데이터는 로컬/폐쇄망 우선",
      "로그/감사로 호출 내역 확인"
    ]
  },
  {
    title: "실전 데모 1: 문서/표/리서치 3-step",
    bullets: [
      "1) 자료 수집: 파일 업로드 + 요약",
      "2) 구조화: 표/목차로 재정리",
      "3) 인사이트: 비교/결론/다음 액션",
      "템플릿화하면 재사용이 쉬움"
    ]
  },
  {
    title: "실전 데모 2: 자동화 시나리오",
    bullets: [
      "조건: 어떤 이벤트가 시작 트리거인지",
      "트리거: 스케줄/웹훅/파일 변경",
      "출력: 문서/보고서/알림으로 자동 발행",
      "실패 대비: 재시도와 에러 로그",
      "작게 시작해서 점진적으로 확장"
    ]
  },
  {
    title: "Codex란?",
    bullets: [
      "ChatGPT 채팅: 대화 중심",
      "Codex 작업공간: 프로젝트/파일 중심",
      "여러 파일 수정 + 실행 + 검증까지",
      "팀/레포 단위로 협업 가능",
      "장기 작업에 최적화된 워크플로"
    ]
  },
  {
    title: "Codex로 산출물 만들기: 끝까지",
    bullets: [
      "코드 작성 → 실행 → 결과 확인",
      "문서/슬라이드 자동 생성",
      "버전 관리와 변경 이력 유지",
      "오늘 실습: 슬라이드 덱 자동 생성",
      "반복 작업을 파이프라인화"
    ]
  },
  {
    title: "체크리스트: 실패 vs 성공 지시문",
    bullets: [
      "실패: 목적이 모호하고 예시가 없음",
      "실패: 형식/길이/제약 조건 누락",
      "성공: 목표 + 결과물 + 제약 명확",
      "성공: 입력 데이터와 맥락 제공",
      "성공: 검수 기준(체크리스트) 포함"
    ]
  },
  {
    title: "엔딩",
    bullets: [
      "다음 영상: MCP 실전 자동화 구축기",
      "구독/좋아요/댓글로 질문 남기기",
      "자료 링크 자리: QR 또는 URL",
      "오늘의 핵심: 구조 → 자동화 → 검증"
    ]
  }
];

slides.forEach((data, index) => {
  const slide = pptx.addSlide();
  addHeader(slide, data.title);

  if (index === 0 && data.subtitle) {
    slide.addText(data.subtitle, {
      x: 0.9,
      y: 2.1,
      w: 11.6,
      h: 1.2,
      fontSize: 22,
      color: COLORS.text
    });
  } else if (data.bullets) {
    addBullets(slide, data.bullets);
  }
});

pptx.writeFile({ fileName: OUTPUT_FILE }).then(() => {
  const deckBinary = fs.readFileSync(OUTPUT_FILE);
  fs.writeFileSync(OUTPUT_BASE64, deckBinary.toString("base64"));
  console.log(`Deck created: ${OUTPUT_FILE}`);
  console.log(`Base64 created: ${OUTPUT_BASE64}`);
});
