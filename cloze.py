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

# ---------- 품사 그룹 ----------
POS_GROUPS = {
    "동사": {"VB", "VBD", "VBG", "VBN", "VBP", "VBZ"},
    "명사": {"NN", "NNS", "NNP", "NNPS"},
    "형용사": {"JJ", "JJR", "JJS"},
    "부사": {"RB", "RBR", "RBS"},
}

TOKEN_CANDIDATE_RE = re.compile(r"[A-Za-z]+")

# ---------- 품사별 보기 단어은행 ----------
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
        "share", "engage", "generate", "exceed", "conduct",
        "blossom", "persuade", "construct", "pursue", "charge",
        "supervise", "complete", "select", "accompany", "donate",
        "organize", "assign", "locate", "overlook", "confront",
        "witness", "embody", "reassess", "reflect", "condense",
        "address", "yield", "stir", "perceive", "involve",
        "enable", "outscore", "foster", "enhance", "visualize",
        "obtain", "define", "trigger", "investigate", "intersect",
        "evaluate", "imitate", "indicate", "acquire", "derive",
        "encounter", "grasp", "determine", "endure", "overlap",
        "estimate", "tap", "remark", "necessitate", "favor",
        "coordinate", "portray", "repurpose", "modify", "refresh",
        "retain", "distribute", "transmit", "decline", "shrink",
        "simplify", "presume", "restrain", "inform", "care",
        "shelter", "reply", "amuse", "notice", "struggle",
        "stare", "brighten", "hesitate", "encourage"
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
        "advancement", "engagement", "journalist", "democratization",
        "disruption", "shareholder", "broadcaster", "gatekeeper", "myth",
        "commitment", "accomplishment", "intuition", "testing", "hypothesis",
        "justification", "awkwardness", "anxiety", "icebreaker", "lifeline",
        "empathy", "compassion", "distribution", "portfolio", "nutrition",
        "herb", "facility", "servant", "athlete", "distinction",
        "democracy", "voice", "abstract", "rubbish", "fascination",
        "behaviorism", "metaphor", "anthropology", "founder", "construction",
        "reflection", "emphasis", "ingredient", "cognition", "identification",
        "comparison", "alternative", "perception", "imagery", "representation",
        "occupation", "instrument", "figure", "nonsense", "beverage",
        "capacity", "priority", "thirst", "negotiator", "component",
        "symposium", "transition", "precision", "experimentation", "proportion",
        "housing", "accommodation", "destination", "outlook", "income",
        "adoption", "respondent", "appliance", "equipment", "chemistry",
        "substance", "mixture", "textile", "portrait", "geology",
        "mentor", "interpretation", "registration", "session", "gear",
        "application", "conclusion", "boundary", "incidence", "obesity",
        "coincidence", "distortion", "moral", "outrage", "irrationality",
        "significance", "perspective", "analysis", "regularity", "sharpness",
        "arrangement", "variation", "odds", "sensation", "faculty",
        "posture", "criticism", "overextension", "caution", "comprehension",
        "knowledge", "instructor", "cue", "injury", "mindset",
        "device", "digit", "calculator", "curriculum", "stress",
        "abstraction", "cluster", "calculation", "indicator", "seed",
        "academia", "administrator", "combination", "patch", "territory",
        "contribution", "self-esteem", "performance", "proximity", "element",
        "pitch", "wetland", "provision", "concentration", "capability",
        "output", "viewpoint", "confusion", "episode", "tragedy",
        "symptom", "mechanism", "resource", "fatigue", "attention",
        "inactivity", "despair", "barn", "tractor", "fondness",
        "indication", "outcome", "narrative", "fashion", "fabric",
        "sustainability", "silhouette", "trimmings", "reliability", "opponent",
        "willingness", "demonstration", "aggressiveness", "disarmament",
        "spiral", "intention", "hostility", "invasion", "centerpiece",
        "confidence", "vendor", "patience", "instruction", "paragraph",
        "impression", "purpose"
    ],
    "형용사": [
        "unique", "insightful", "critical", "lifeless", "spotted",
        "enthusiastic", "tiresome", "ridiculous", "superior", "dizzy",
        "precious", "cruel", "soaked", "unassuming", "legendary",
        "harsh", "flat", "countless", "pitiless", "deadly",
        "controversial", "solar", "constant", "sophisticated", "relevant",
        "accountable", "valid", "absolute", "ultimate", "normal",
        "internal", "insecure", "inherent", "overwhelming", "oversized",
        "manageable", "regulatory", "misleading", "reputable", "acceptable",
        "alien", "compatible", "convincing", "inclined", "materialistic",
        "cognitive", "humanistic", "essential", "utter", "innate",
        "bold", "freehand", "reflective", "opposite", "common",
        "distinguished", "vibrant", "affordable", "inner", "outstanding",
        "seasonal", "upcoming", "urban", "distinct", "civilized",
        "positive", "cooperative", "counterproductive", "objective", "disposable",
        "vocal", "effortful", "fundamental", "specific", "numerical",
        "crucial", "leading", "rough", "inhabitable", "colonial",
        "normative", "associative", "negative", "spatial", "diverse",
        "apparent", "shallow", "demanding", "logical", "subjective",
        "instrumental", "psychological", "remarkable", "affectionate", "immediate",
        "verified", "multiple", "straightforward", "supplementary", "prominent",
        "rational", "invisible", "bygone", "unquestioned", "secondhand",
        "economical", "questionable", "portable", "compact", "accessible",
        "stable", "chronic", "dominant", "preventive", "newfound",
        "nervous", "retired", "renewed"
    ],
    "부사": [
        "thoroughly", "briefly", "urgently", "initially", "approximately",
        "critically", "instinctively", "dramatically", "occasionally", "commonly",
        "consistently", "unintentionally", "ultimately", "potentially", "coolly",
        "subtly", "unconsciously", "conversely", "evidently", "roughly",
        "understandably", "essentially", "simultaneously", "consequently",
        "remarkably", "nostalgically", "gently", "politely", "merely",
        "blankly", "cautiously"
    ]
}

WORD_BANK = []
for word_list in WORD_BANK_BY_POS.values():
    WORD_BANK.extend(word_list)


def is_candidate_token(tok):
    return bool(TOKEN_CANDIDATE_RE.fullmatch(tok))


def tokenize_text(text):
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

    if correct_pos_group in WORD_BANK_BY_POS:
        pool_source = WORD_BANK_BY_POS[correct_pos_group]
    else:
        pool_source = WORD_BANK

    pool = []
    seen = set()

    for word in pool_source:
        word_lower = str(word).strip().lower()

        if word_lower != correct_lower and word_lower not in seen:
            seen.add(word_lower)
            pool.append(word)

    if len(pool) < 4:
        for word in WORD_BANK:
            word_lower = str(word).strip().lower()

            if word_lower != correct_lower and word_lower not in seen:
                seen.add(word_lower)
                pool.append(word)

            if len(pool) >= 4:
                break

    wrong_options = random.sample(pool, 4)
    options = wrong_options + [correct_word]
    random.shuffle(options)

    return options


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

        tokens = tokenize_text(orig_text)

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

        if n_candidates == 0:
            question_paragraphs.append(orig_text)
            continue

        n_blanks = max(1, int(round(n_candidates * blank_ratio_fraction)))
        n_blanks = min(n_blanks, n_candidates)

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

        question_paragraphs.append(assemble_tokens(out_tokens))

    return question_paragraphs, answer_map, option_map


def grade_answers(answer_map):
    total = len(answer_map)

    correct_count = 0
    results = []

    for num in sorted(answer_map.keys()):
        correct = answer_map[num]
        user_ans = st.session_state.get(f"answer_{num}", "")

        user_norm = str(user_ans).strip().lower()
        correct_norm = str(correct).strip().lower()

        is_correct = user_norm == correct_norm and user_norm != ""

        if is_correct:
            correct_count += 1

        results.append({
            "num": num,
            "correct": correct,
            "user": user_ans,
            "is_correct": is_correct,
        })

    return correct_count, total, results


# ---------------- Streamlit UI ----------------
st.set_page_config(page_title="Blank Test Web Quiz", layout="wide")

st.title("📘 Blank Test Web Quiz")

st.markdown(
    "Word(.docx) 파일을 업로드하면 왼쪽에는 문제지, 오른쪽에는 오지선다 답안 선택지가 바로 표시됩니다."
)

col_class, col_name = st.columns(2)

with col_class:
    class_name = st.text_input("반", value="", placeholder="예: 중3A반")

with col_name:
    student_name = st.text_input("이름", value="", placeholder="예: 홍길동")

st.markdown("---")

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

col_make, col_reset = st.columns(2)

with col_make:
    make_button = st.button("📄 문제 만들기")

with col_reset:
    reset_button = st.button("🧹 초기화")

if reset_button:
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()

if uploaded_file is not None and make_button:
    uploaded_file.seek(0)

    questions, answer_map, option_map = generate_questions_from_docx(
        uploaded_file,
        pos_choice,
        blank_pct / 100.0
    )

    st.session_state["questions"] = questions
    st.session_state["answer_map"] = answer_map
    st.session_state["option_map"] = option_map

    for key in list(st.session_state.keys()):
        if str(key).startswith("answer_"):
            del st.session_state[key]

    st.success(
        f"문제가 생성되었습니다. 총 {len(answer_map)}개의 빈칸이 만들어졌습니다."
    )

    st.rerun()
# ---------------- 좌측 문제지 / 우측 오지선다 ----------------
left_col, right_col = st.columns([1.25, 1])

with left_col:
    st.subheader("📝 문제지")

    if "questions" not in st.session_state:
        st.info("문제를 생성하면 여기에 문제지가 표시됩니다.")
    else:
        st.markdown(
            """
            <div style="
                border:1px solid #ddd;
                border-radius:12px;
                padding:20px;
                background-color:#ffffff;
                line-height:1.8;
                font-size:17px;
                max-height:720px;
                overflow-y:auto;
            ">
            """,
            unsafe_allow_html=True
        )

        for para in st.session_state["questions"]:
            if para.strip() == "":
                st.markdown("<br>", unsafe_allow_html=True)
            else:
                st.markdown(f"<p>{para}</p>", unsafe_allow_html=True)

        st.markdown("</div>", unsafe_allow_html=True)

with right_col:
    st.subheader("✏️ 오지선다 답안 선택")

    if "answer_map" not in st.session_state or "option_map" not in st.session_state:
        st.info("문제를 생성하면 여기에 답안 선택지가 표시됩니다.")

    else:
        answer_map = st.session_state["answer_map"]
        option_map = st.session_state["option_map"]

        if len(answer_map) == 0:
            st.warning("생성된 빈칸이 없습니다. 빈칸 비율을 올리거나 다른 품사를 선택해 보세요.")
        else:
            st.write(f"총 **{len(answer_map)}문항**")

            for num in sorted(answer_map.keys()):
                options = option_map.get(num, [])

                st.markdown(f"**{num}번**")

                st.radio(
                    label=f"{num}번 보기",
                    options=options,
                    key=f"answer_{num}",
                    horizontal=False,
                    label_visibility="collapsed"
                )

            st.markdown("---")

            if st.button("✅ 채점하기"):
                correct_count, total, results = grade_answers(answer_map)
                score_pct = (correct_count / total) * 100 if total > 0 else 0.0

                st.subheader("📊 채점 결과")
                st.write(f"총 {total}문항 중 **{correct_count}개** 정답입니다.")
                st.write(f"점수: **{score_pct:.1f}점 / 100점**")

                for r in results:
                    if r["is_correct"]:
                        st.success(f"{r['num']}번 정답: {r['user']}")
                    else:
                        st.error(
                            f"{r['num']}번 오답 / 선택: {r['user']} / 정답: {r['correct']}"
                        )
