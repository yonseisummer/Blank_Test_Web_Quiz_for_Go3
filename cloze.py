import streamlit as st
from docx import Document
import nltk
from nltk import pos_tag, word_tokenize
import random
import re

# ---------- NLTK data ----------
try:
    nltk.data.find("tokenizers/punkt")
except LookupError:
    nltk.download("punkt", quiet=True)

try:
    nltk.data.find("tokenizers/punkt_tab")
except LookupError:
    nltk.download("punkt_tab", quiet=True)

try:
    nltk.data.find("taggers/averaged_perceptron_tagger")
except LookupError:
    try:
        nltk.data.find("taggers/averaged_perceptron_tagger_eng")
    except LookupError:
        nltk.download("averaged_perceptron_tagger", quiet=True)
        nltk.download("averaged_perceptron_tagger_eng", quiet=True)

# ---------- POS 그룹 ----------
POS_GROUPS = {
    "동사": {"VB", "VBD", "VBG", "VBN", "VBP", "VBZ"},
    "명사": {"NN", "NNS", "NNP", "NNPS"},
    "형용사": {"JJ", "JJR", "JJS"},
    "부사": {"RB", "RBR", "RBS"},
}

TOKEN_CANDIDATE_RE = re.compile(r"[A-Za-z0-9\uac00-\ud7a3]+")

# ---------- 수능특강 빈출 어휘 기반 보기 단어은행 ----------
# 첨부해주신 수능특강 어휘 리스트를 참고해 코드 안에 넣은 보기 후보입니다.
WORD_BANK = [
    "diversity", "enrich", "unique", "unity", "revision", "manuscript",
    "cooperation", "thoroughly", "insightful", "critical", "incorporate",
    "mention", "briefly", "complaint", "anniversary", "lifeless",
    "spotted", "compensation", "refund", "settle", "enthusiastic",
    "tiresome", "repayment", "loan", "transportation", "translate",
    "immigration", "document", "property", "ignore", "rejection",
    "ridiculous", "superior", "dizzy", "assistance", "precious",
    "possession", "comfort", "interrupt", "weapon", "whisper",
    "urgently", "trunk", "prayer", "initially", "commute",
    "cruel", "irony", "soaked", "scent", "unassuming",
    "ease", "microclimate", "scale", "legendary", "harsh",
    "approximately", "flat", "inject", "promotion", "manipulation",
    "loyalty", "astronaut", "inspire", "cause", "trap",
    "statement", "blend", "accumulate", "countless", "pitiless",
    "phenomenon", "blueprint", "dictator", "blessing", "deadly",
    "vehicle", "critically", "controversial", "assumption", "challenge",
    "gravity", "sequence", "crop", "solar", "filter",
    "concentrate", "constant", "civilization", "flourish", "sophisticated",
    "capture", "fraction", "fuel", "advancement", "engagement",
    "journalist", "relevant", "abandon", "accountable", "democratization",
    "disruption", "corporate", "shareholder", "shift", "broadcaster",
    "gatekeeper", "myth", "valid", "commitment", "accomplishment",
    "intuition", "testing", "confirm", "disprove", "reject",
    "alter", "hypothesis", "justification", "instinctively", "absolute",
    "ultimate", "awkwardness", "normal", "internal", "anxiety",
    "boost", "icebreaker", "lifeline", "insecure", "empathy",
    "flip", "compassion", "inherent", "overwhelming", "comprehend",
    "interfere", "overcome", "oversized", "identify", "break",
    "manageable", "revisit", "distribution", "portfolio", "nutrition",
    "herb", "regulatory", "misleading", "reputable", "devote",
    "facility", "embrace", "servant", "polish", "assume",
    "sweep", "activate", "acceptable", "athlete", "downshift",
    "resolve", "alien", "compatible", "distinction", "restrict",
    "democracy", "convincing", "voice", "abstract", "rubbish",
    "interpret", "inclined", "materialistic", "fascination", "behaviorism",
    "metaphor", "anthropology", "founder", "cognitive", "construction",
    "reflection", "humanistic", "emphasis", "essential", "ingredient",
    "cognition", "escape", "identification", "comparison", "alternative",
    "perception", "reverse", "signify", "populate", "imagery",
    "representation", "interweave", "occupation", "instrument", "figure",
    "dramatically", "utter", "nonsense", "innate", "draw",
    "ensure", "beverage", "capacity", "priority", "dismiss",
    "thirst", "negotiator", "justify", "component", "symposium",
    "bold", "freehand", "transition", "entail", "precision",
    "reflective", "phenomenon", "disincentive", "interconnectedness",
    "emphasize", "experimentation", "proportion", "housing", "accommodation",
    "destination", "opposite", "outlook", "income", "adoption",
    "outnumber", "foresee", "share", "engage", "occasionally",
    "commonly", "respondent", "generate", "appliance", "equipment",
    "exceed", "chemistry", "distinguished", "conduct", "substance",
    "mixture", "blossom", "vibrant", "affordable", "inner",
    "persuade", "construct", "textile", "outstanding", "portrait",
    "geology", "mentor", "interpretation", "pursue", "charge",
    "supervise", "registration", "complete", "seasonal", "session",
    "select", "primarily", "accompany", "upcoming", "donate",
    "gear", "organize", "application", "urban", "assign",
    "locate", "conclusion", "boundary", "distinct", "civilized",
    "overlook", "incidence", "obesity", "coincidence", "confront",
    "distortion", "consistently", "unintentionally", "positive", "ultimately",
    "moral", "outrage", "witness", "embody", "irrationality",
    "cooperative", "counterproductive", "objective", "potentially", "disposable",
    "reassess", "significance", "perspective", "reflect", "analysis",
    "condense", "address", "regularity", "sharpness", "arrangement",
    "variation", "odds", "sensation", "faculty", "yield",
    "coolly", "stir", "subtly", "unconsciously", "posture",
    "vocal", "perceive", "criticism", "overextension", "involve",
    "caution", "GPS", "comprehension", "knowledge", "enable",
    "outscore", "instructor", "foster", "cue", "effortful",
    "injury", "conversely", "mindset", "evidently", "enhance",
    "fundamental", "device", "digit", "calculator", "visualize",
    "obtain", "curriculum", "stress", "abstraction", "specific",
    "cluster", "define", "calculation", "indicator", "numerical",
    "seed", "crucial", "trigger", "leading", "investigate",
    "roughly", "academia", "administrator", "finalize", "combination",
    "intersect", "inhabitable", "patch", "understandably", "colonial",
    "territory", "evaluate", "contribution", "normative", "imitate",
    "associative", "self-esteem", "negative", "performance", "proximity",
    "element", "pitch", "essentially", "simultaneously", "spatial",
    "undervalue", "wetland", "provision", "diverse", "apparent",
    "indicate", "shallow", "concentration", "capability", "demanding",
    "logical", "output", "subjective", "acquire", "instrumental",
    "derive", "encounter", "viewpoint", "confusion", "episode",
    "tragedy", "grasp", "symptom", "mechanism", "psychological",
    "resource", "fatigue", "attention", "precede", "consequently",
    "determine", "inactivity", "endure", "despair", "overlap",
    "remarkably", "estimate", "barn", "tractor", "affectionate",
    "tap", "nostalgically", "remark", "fondness", "immediate",
    "verified", "multiple", "straightforward", "necessitate", "supplementary",
    "favor", "prominent", "coordinate", "indication", "outcome",
    "rational", "invisible", "unthought", "bygone", "unquestioned",
    "narrative", "portray", "repurpose", "fashion", "fabric",
    "modify", "secondhand", "sustainability", "silhouette", "trimmings",
    "refresh", "economical", "retain", "distribute", "reliability",
    "questionable", "transmit", "portable", "compact", "accessible",
    "decline", "shrink", "stable", "chronic", "simplify",
    "dominant", "opponent", "willingness", "demonstration", "aggressiveness",
    "preventive", "disarmament", "spiral", "presume", "intention",
    "hostility", "invasion", "centerpiece", "restrain", "inform",
    "care", "shelter", "confidence", "gently", "politely",
    "vendor", "merely", "reply", "patience", "disruptive",
    "amuse", "instruction", "paragraph", "notice", "newfound",
    "impression", "nervous", "struggle", "stare", "blankly",
    "retired", "brighten", "cautiously", "hesitate", "encourage",
    "encounter", "application", "nonprofit", "renewed", "purpose"
]


def is_candidate_token(tok):
    return bool(TOKEN_CANDIDATE_RE.search(tok))


def tokenize_preserve_spacing(text):
    tokens = word_tokenize(text)
    return tokens


def assemble_tokens(tokens):
    out = ""
    for i, t in enumerate(tokens):
        if i == 0:
            out += t
            continue

        if re.fullmatch(r"[^\w\s]", t):
            out += t
        else:
            out += " " + t

    return out


def make_options(correct_word):
    correct_word = str(correct_word).strip()
    correct_lower = correct_word.lower()

    pool = []
    for word in WORD_BANK:
        if word.lower() != correct_lower and word not in pool:
            pool.append(word)

    if len(pool) >= 4:
        wrong_options = random.sample(pool, 4)
    else:
        wrong_options = pool

    options = wrong_options + [correct_word]
    random.shuffle(options)

    return options


# ---------- 문제 생성용 함수 ----------
def generate_questions_from_docx(file_like, pos_choice, blank_ratio_fraction):
    src = Document(file_like)

    question_paragraphs = []
    answer_map = {}
    option_map = {}
    next_blank_num = 1

    for para in src.paragraphs:
        orig_text = para.text.strip()

        if not orig_text:
            question_paragraphs.append("")
            continue

        tokens = tokenize_preserve_spacing(orig_text)

        try:
            tagged = pos_tag(tokens)
        except Exception:
            tagged = [(t, "NN") for t in tokens]

        candidate_indices = []

        for i, (tok, tg) in enumerate(tagged):
            if is_candidate_token(tok):
                if pos_choice == "전체":
                    candidate_indices.append(i)
                else:
                    if tg in POS_GROUPS.get(pos_choice, set()):
                        candidate_indices.append(i)

        if not candidate_indices:
            candidate_indices = [
                i for i, (tok, tg) in enumerate(tagged) if is_candidate_token(tok)
            ]

        n_candidates = len(candidate_indices)
        n_blanks = max(0, int(round(n_candidates * blank_ratio_fraction)))
        n_blanks = min(n_blanks, n_candidates)

        chosen = []

        if n_blanks > 0 and n_candidates > 0:
            chosen = random.sample(candidate_indices, n_blanks)

        out_tokens = list(tokens)

        for idx in sorted(chosen):
            original_word = tokens[idx]
            underline = "_" * max(3, len(original_word))

            out_tokens[idx] = f"({next_blank_num}){underline}"
            answer_map[next_blank_num] = original_word
            option_map[next_blank_num] = make_options(original_word)

            next_blank_num += 1

        para_text = assemble_tokens(out_tokens)
        question_paragraphs.append(para_text)

    return question_paragraphs, answer_map, option_map


# ---------- 채점 함수 ----------
def grade_answers(answer_map):
    total = len(answer_map)

    if total == 0:
        return 0, 0, []

    correct_count = 0
    results = []

    for num in sorted(answer_map.keys()):
        correct = answer_map[num]
        user_key = f"answer_{num}"
        user_ans = st.session_state.get(user_key, "")

        user_norm = str(user_ans).strip().lower()
        correct_norm = str(correct).strip().lower()

        is_correct = (user_norm == correct_norm) and (user_norm != "")

        if is_correct:
            correct_count += 1

        results.append(
            {
                "num": num,
                "correct": correct,
                "user": user_ans,
                "is_correct": is_correct,
            }
        )

    return correct_count, total, results


# ---------------- Streamlit UI ----------------
st.set_page_config(page_title="Blank Test Web Quiz", layout="wide")

st.title("📘 Blank Test Web Quiz")

st.markdown(
    "업로드한 Word(.docx)에서 특정 품사만 선택하여 랜덤으로 빈칸을 생성하고, "
    "수능형 오지선다 보기로 자동 채점까지 할 수 있습니다.\n\n"
    "**문제지 전체는 항상 왼쪽 사이드바에 고정**되어 있어서, "
    "스크롤을 내려도 지문을 계속 보면서 답을 선택할 수 있습니다."
)

# 상단 정보란
col_class, col_name = st.columns(2)

with col_class:
    class_name = st.text_input("반", value="", placeholder="예: 중3A반")

with col_name:
    student_name = st.text_input("이름", value="", placeholder="예: 홍길동")

st.markdown("---")

# 설정
pos_choice = st.selectbox(
    "빈칸으로 만들 품사 선택",
    ["전체", "동사", "명사", "형용사", "부사"]
)

blank_pct = st.slider(
    "빈칸 비율 (%)",
    min_value=5,
    max_value=80,
    value=20,
    step=5
)

uploaded_file = st.file_uploader(
    "📂 Word(.docx) 파일을 드래그 앤 드롭하거나 클릭하여 업로드하세요.",
    type=["docx"]
)

# 세션 상태 초기화 버튼
if st.button("🧹 초기화(새로 시작하기)"):
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()

# 문제 생성 버튼
if uploaded_file is not None:
    if st.button("📄 문제 만들기"):
        try:
            uploaded_file.seek(0)

            questions, answer_map, option_map = generate_questions_from_docx(
                uploaded_file,
                pos_choice,
                blank_pct / 100.0
            )

            st.session_state["questions"] = questions
            st.session_state["answer_map"] = answer_map
            st.session_state["option_map"] = option_map

            # 이전 답안 초기화
            for key in list(st.session_state.keys()):
                if str(key).startswith("answer_"):
                    del st.session_state[key]

            st.success("문제가 생성되었습니다. 왼쪽 문제지를 보면서 아래에서 답을 선택하세요!")

        except Exception as e:
            st.error("문제 생성 중 오류가 발생했습니다.")
            st.exception(e)
else:
    st.info("먼저 Word(.docx) 파일을 업로드하세요.")

st.markdown("---")

# --------- 사이드바에 항상 문제지 표시 ---------
with st.sidebar:
    st.header("📝 문제지")

    if "questions" in st.session_state:
        questions = st.session_state["questions"]

        for para in questions:
            if para.strip() == "":
                st.write("")
            else:
                st.markdown(para)
    else:
        st.caption("문제지가 여기에 표시됩니다. 먼저 docx를 업로드하고 '문제 만들기'를 눌러 주세요.")

# --------- 메인 영역: 답안 선택 + 채점 ---------
if "answer_map" in st.session_state and "option_map" in st.session_state:
    answer_map = st.session_state["answer_map"]
    option_map = st.session_state["option_map"]

    if len(answer_map) == 0:
        st.warning("생성된 빈칸이 없습니다. 빈칸 비율을 올리거나 다른 품사/지문을 사용해 보세요.")

    else:
        st.subheader("✏️ 답안 선택")

        for num in sorted(answer_map.keys()):
            options = option_map.get(num, [])

            st.radio(
                label=f"{num}번",
                options=options,
                key=f"answer_{num}",
                index=None,
                horizontal=True,
            )

        if st.button("✅ 채점하기"):
            correct_count, total, results = grade_answers(answer_map)
            score_pct = (correct_count / total) * 100 if total > 0 else 0.0

            st.markdown("---")
            st.subheader("📊 채점 결과")

            st.write(f"총 {total}문항 중 **{correct_count}개** 정답입니다.")
            st.write(f"점수: **{score_pct:.1f}점 / 100점**")

            for r in results:
                num = r["num"]
                correct = r["correct"]
                user_ans = r["user"]

                if r["is_correct"]:
                    st.success(f"{num}번: 정답! 선택: **{user_ans}**")
                else:
                    if str(user_ans).strip() == "":
                        st.error(f"{num}번: 무응답. 정답은 **{correct}** 입니다.")
                    else:
                        st.error(
                            f"{num}번: 오답. 선택: `{user_ans}`, 정답: **{correct}**"
                        )

else:
    st.info("문제지를 먼저 생성해 주세요.")
