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
    nltk.data.find("taggers/averaged_perceptron_tagger_eng")
except LookupError:
    nltk.download("averaged_perceptron_tagger_eng", quiet=True)

# ---------- POS 그룹 ----------
POS_GROUPS = {
    "동사": {"VB", "VBD", "VBG", "VBN", "VBP", "VBZ"},
    "명사": {"NN", "NNS", "NNP", "NNPS"},
    "형용사": {"JJ", "JJR", "JJS"},
    "부사": {"RB", "RBR", "RBS"},
}

TOKEN_CANDIDATE_RE = re.compile(r"[A-Za-z]+")

# ---------- 품사별 수능특강 보기 단어은행 ----------
WORD_BANK_BY_POS = {
    "동사": [
        "enrich", "incorporate", "mention", "settle", "translate",
        "ignore", "interrupt", "whisper", "commute", "inject",
        "inspire", "cause", "trap", "blend", "accumulate",
        "challenge", "concentrate", "flourish", "capture", "abandon",
        "shift", "confirm", "disprove", "reject", "alter",
        "boost", "flip", "comprehend", "interfere", "overcome",
        "identify", "break", "revisit", "devote", "embrace",
        "polish", "assume", "sweep", "activate", "resolve",
        "restrict", "interpret", "escape", "reverse", "signify",
        "populate", "interweave", "draw", "ensure", "dismiss",
        "justify", "entail", "emphasize", "outnumber", "foresee",
        "share", "engage", "respond", "generate", "exceed",
        "conduct", "blossom", "persuade", "construct", "pursue",
        "charge", "supervise", "complete", "select", "accompany",
        "donate", "organize", "assign", "locate", "overlook",
        "confront", "witness", "embody", "reassess", "reflect",
        "condense", "address", "yield", "stir", "perceive",
        "involve", "enable", "outscore", "foster", "enhance",
        "visualize", "obtain", "define", "trigger", "investigate",
        "intersect", "evaluate", "imitate", "indicate", "acquire",
        "derive", "encounter", "grasp", "determine", "endure",
        "overlap", "estimate", "tap", "remark", "necessitate",
        "supplement", "favor", "coordinate", "portray", "repurpose",
        "modify", "refresh", "retain", "distribute", "transmit",
        "decline", "shrink", "simplify", "prevent", "presume",
        "restrain", "inform", "care", "shelter", "reply",
        "amuse", "notice", "struggle", "stare", "brighten",
        "hesitate", "encourage"
    ],
    "명사": [
        "diversity", "unity", "revision", "manuscript", "cooperation",
        "complaint", "anniversary", "compensation", "refund", "repayment",
        "loan", "transportation", "immigration", "document", "property",
        "rejection", "assistance", "possession", "comfort", "weapon",
        "trunk", "prayer", "irony", "scent", "ease",
        "microclimate", "scale", "promotion", "manipulation", "loyalty",
        "astronaut", "statement", "phenomenon", "blueprint", "dictator",
        "blessing", "vehicle", "assumption", "gravity", "sequence",
        "crop", "filter", "civilization", "fraction", "fuel",
        "advancement", "engagement", "journalist", "democratization", "disruption",
        "shareholder", "broadcaster", "gatekeeper", "myth", "commitment",
        "accomplishment", "intuition", "testing", "hypothesis", "justification",
        "awkwardness", "anxiety", "icebreaker", "lifeline", "empathy",
        "compassion", "distribution", "portfolio", "nutrition", "herb",
        "facility", "servant", "athlete", "distinction", "democracy",
        "voice", "abstract", "rubbish", "fascination", "behaviorism",
        "metaphor", "anthropology", "founder", "construction", "reflection",
        "emphasis", "ingredient", "cognition", "identification", "comparison",
        "alternative", "perception", "imagery", "representation", "occupation",
        "instrument", "figure", "nonsense", "beverage", "capacity",
        "priority", "thirst", "negotiator", "component", "symposium",
        "transition", "precision", "experimentation", "proportion", "housing",
        "accommodation", "destination", "outlook", "income", "adoption",
        "respondent", "appliance", "equipment", "chemistry", "substance",
        "mixture", "textile", "portrait", "geology", "mentor",
        "interpretation", "registration", "session", "gear", "application",
        "conclusion", "boundary", "incidence", "obesity", "coincidence",
        "distortion", "moral", "outrage", "irrationality", "significance",
        "perspective", "analysis", "regularity", "sharpness", "arrangement",
        "variation", "odds", "sensation", "faculty", "posture",
        "criticism", "overextension", "caution", "comprehension", "knowledge",
        "instructor", "cue", "injury", "mindset", "device",
        "digit", "calculator", "curriculum", "stress", "abstraction",
        "cluster", "calculation", "indicator", "seed", "academia",
        "administrator", "combination", "patch", "territory", "contribution",
        "self-esteem", "performance", "proximity", "element", "pitch",
        "wetland", "provision", "concentration", "capability", "demand",
        "output", "viewpoint", "confusion", "episode", "tragedy",
        "symptom", "mechanism", "resource", "fatigue", "attention",
        "inactivity", "despair", "barn", "tractor", "fondness",
        "indication", "outcome", "narrative", "fashion", "fabric",
        "sustainability", "silhouette", "trimmings", "reliability", "opponent",
        "willingness", "demonstration", "aggressiveness", "disarmament", "spiral",
        "intention", "hostility", "invasion", "centerpiece", "confidence",
        "vendor", "patience", "instruction", "paragraph", "impression",
        "purpose"
    ],
    "형용사": [
        "unique", "insightful", "critical", "brief", "lifeless",
        "spotted", "enthusiastic", "tiresome", "ridiculous", "superior",
        "dizzy", "precious", "cruel", "soaked", "unassuming",
        "legendary", "harsh", "flat", "countless", "pitiless",
        "deadly", "controversial", "solar", "constant", "sophisticated",
        "relevant", "accountable", "valid", "absolute", "ultimate",
        "normal", "internal", "insecure", "inherent", "overwhelming",
        "oversized", "manageable", "regulatory", "misleading", "reputable",
        "acceptable", "alien", "compatible", "convincing", "inclined",
        "materialistic", "cognitive", "humanistic", "essential", "dramatic",
        "utter", "innate", "bold", "freehand", "reflective",
        "interconnected", "opposite", "common", "distinguished", "vibrant",
        "affordable", "inner", "outstanding", "seasonal", "upcoming",
        "urban", "distinct", "civilized", "consistent", "unintentional",
        "positive", "cooperative", "counterproductive", "objective", "potential",
        "disposable", "vocal", "effortful", "evident", "fundamental",
        "specific", "numerical", "crucial", "leading", "rough",
        "inhabitable", "colonial", "normative", "associative", "negative",
        "spatial", "diverse", "apparent", "shallow", "demanding",
        "logical", "subjective", "instrumental", "psychological", "remarkable",
        "affectionate", "immediate", "verified", "multiple", "straightforward",
        "supplementary", "prominent", "rational", "invisible", "bygone",
        "unquestioned", "secondhand", "economical", "questionable", "portable",
        "compact", "accessible", "stable", "chronic", "dominant",
        "preventive", "newfound", "nervous", "retired", "renewed"
    ],
    "부사": [
        "thoroughly", "briefly", "urgently", "initially", "approximately",
        "critically", "instinctively", "dramatically", "occasionally", "commonly",
        "consistently", "unintentionally", "ultimately", "potentially", "coolly",
        "subtly", "unconsciously", "conversely", "evidently", "roughly",
        "understandably", "essentially", "simultaneously", "consequently", "remarkably",
        "nostalgically", "gently", "politely", "merely", "blankly",
        "cautiously"
    ]
}

# 전체 단어은행
WORD_BANK = []
for words in WORD_BANK_BY_POS.values():
    WORD_BANK.extend(words)


def is_candidate_token(tok):
    return bool(TOKEN_CANDIDATE_RE.fullmatch(tok))


def tokenize_preserve_spacing(text):
    return word_tokenize(text)


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


def get_pos_group(tag):
    for group_name, tag_set in POS_GROUPS.items():
        if tag in tag_set:
            return group_name
    return None


def make_options(correct_word, correct_pos_group):
    correct_word = str(correct_word).strip()
    correct_lower = correct_word.lower()

    # 정답과 같은 품사에서 오답 후보를 먼저 뽑음
    if correct_pos_group in WORD_BANK_BY_POS:
        pool_source = WORD_BANK_BY_POS[correct_pos_group]
    else:
        pool_source = WORD_BANK

    pool = []

    for word in pool_source:
        word_lower = str(word).strip().lower()

        if word_lower != correct_lower and word_lower not in [p.lower() for p in pool]:
            pool.append(word)

    # 혹시 같은 품사 후보가 4개 미만이면 전체 단어장에서 보충
    if len(pool) < 4:
        for word in WORD_BANK:
            word_lower = str(word).strip().lower()

            if word_lower != correct_lower and word_lower not in [p.lower() for p in pool]:
                pool.append(word)

            if len(pool) >= 4:
                break

    wrong_options = random.sample(pool, 4)
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
            tagged = pos_tag(tokens, lang="eng")
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
                i for i, (tok, tg) in enumerate(tagged)
                if is_candidate_token(tok)
            ]

        n_candidates = len(candidate_indices)
        n_blanks = max(1, int(round(n_candidates * blank_ratio_fraction))) if n_candidates > 0 else 0
        n_blanks = min(n_blanks, n_candidates)

        chosen = []

        if n_blanks > 0 and n_candidates > 0:
            chosen = random.sample(candidate_indices, n_blanks)

        out_tokens = list(tokens)

        for idx in sorted(chosen):
            original_word = tokens[idx]
            original_tag = tagged[idx][1]
            original_pos_group = get_pos_group(original_tag)

            underline = "_" * max(3, len(original_word))

            out_tokens[idx] = f"({next_blank_num}){underline}"
            answer_map[next_blank_num] = original_word
            option_map[next_blank_num] = make_options(original_word, original_pos_group)

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

        if user_ans == "선택하세요":
            user_ans = ""

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

            if len(answer_map) == 0:
                st.warning("빈칸이 생성되지 않았습니다. 다른 품사나 지문을 선택해 보세요.")
            else:
                st.success(f"문제가 생성되었습니다. 총 {len(answer_map)}개의 빈칸이 만들어졌습니다.")

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
# --------- 메인 영역: 답안 선택 + 채점 ---------
st.subheader("✏️ 답안 선택")

if "answer_map" not in st.session_state or "option_map" not in st.session_state:
    st.info("문제를 먼저 생성하면 여기에 답안 선택지가 표시됩니다.")

else:
    answer_map = st.session_state["answer_map"]
    option_map = st.session_state["option_map"]

    st.write(f"생성된 빈칸 수: **{len(answer_map)}개**")

    if len(answer_map) == 0:
        st.warning("생성된 빈칸이 없습니다. 빈칸 비율을 올리거나 다른 품사를 선택해 보세요.")

    else:
        for num in sorted(answer_map.keys()):
            options = option_map.get(num, [])

            st.markdown(f"### {num}번")

            st.selectbox(
                label=f"{num}번 답을 선택하세요",
                options=["선택하세요"] + options,
                key=f"answer_{num}",
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
                        st.error(f"{num}번: 오답. 선택: `{user_ans}`, 정답: **{correct}**")
