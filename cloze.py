import streamlit as st
from docx import Document
import nltk
from nltk import word_tokenize, pos_tag
import random
import re
import html

# =========================================================
# NLTK 데이터 준비
# =========================================================
def ensure_nltk_data():
    resources = [
        ("punkt", "tokenizers/punkt"),
        ("punkt_tab", "tokenizers/punkt_tab"),
        ("averaged_perceptron_tagger_eng", "taggers/averaged_perceptron_tagger_eng"),
    ]

    for package, path in resources:
        try:
            nltk.data.find(path)
        except LookupError:
            nltk.download(package, quiet=True)

ensure_nltk_data()

# =========================================================
# 기본 설정
# =========================================================
st.set_page_config(page_title="수능형 Cloze Test Generator", layout="wide")

POS_GROUPS = {
    "동사": {"VB", "VBD", "VBG", "VBN", "VBP", "VBZ"},
    "명사": {"NN", "NNS", "NNP", "NNPS"},
    "형용사": {"JJ", "JJR", "JJS"},
    "부사": {"RB", "RBR", "RBS"},
}

# 너무 쉬운 기능어는 빈칸 대상에서 제외
STOPWORDS = {
    "the", "a", "an", "of", "to", "for", "in", "on", "at", "by", "with", "from",
    "is", "am", "are", "was", "were", "be", "been", "being",
    "and", "or", "but", "so", "yet",
    "that", "this", "these", "those", "it", "its", "they", "them", "their",
    "we", "our", "you", "your", "he", "his", "she", "her", "i", "my",
    "as", "if", "than", "then", "into", "over", "under", "about",
    "after", "before", "because", "when", "while", "where", "which", "who",
    "what", "how", "not", "no", "do", "does", "did", "can", "could",
    "will", "would", "may", "might", "must", "should", "have", "has", "had",
}

# =========================================================
# 수능특강 어휘 기반 단어은행
# 기존 코드에 들어 있던 대형 품사별 단어은행을 반영
# =========================================================
WORD_BANK_BY_POS = {'동사': ['enrich',
        'incorporate',
        'mention',
        'settle',
        'translate',
        'ignore',
        'interrupt',
        'whisper',
        'commute',
        'inject',
        'inspire',
        'cause',
        'trap',
        'blend',
        'accumulate',
        'challenge',
        'concentrate',
        'flourish',
        'capture',
        'abandon',
        'shift',
        'confirm',
        'disprove',
        'reject',
        'alter',
        'boost',
        'flip',
        'comprehend',
        'interfere',
        'overcome',
        'identify',
        'break',
        'revisit',
        'devote',
        'embrace',
        'polish',
        'assume',
        'sweep',
        'activate',
        'resolve',
        'restrict',
        'interpret',
        'escape',
        'reverse',
        'signify',
        'populate',
        'interweave',
        'draw',
        'ensure',
        'dismiss',
        'justify',
        'entail',
        'emphasize',
        'outnumber',
        'foresee',
        'share',
        'engage',
        'generate',
        'exceed',
        'conduct',
        'blossom',
        'persuade',
        'construct',
        'pursue',
        'charge',
        'supervise',
        'complete',
        'select',
        'accompany',
        'donate',
        'organize',
        'assign',
        'locate',
        'overlook',
        'confront',
        'witness',
        'embody',
        'reassess',
        'reflect',
        'condense',
        'address',
        'yield',
        'stir',
        'perceive',
        'involve',
        'enable',
        'outscore',
        'foster',
        'enhance',
        'visualize',
        'obtain',
        'define',
        'trigger',
        'investigate',
        'intersect',
        'evaluate',
        'imitate',
        'indicate',
        'acquire',
        'derive',
        'encounter',
        'grasp',
        'determine',
        'endure',
        'overlap',
        'estimate',
        'tap',
        'remark',
        'necessitate',
        'favor',
        'coordinate',
        'portray',
        'repurpose',
        'modify',
        'refresh',
        'retain',
        'distribute',
        'transmit',
        'decline',
        'shrink',
        'simplify',
        'presume',
        'restrain',
        'inform',
        'care',
        'shelter',
        'reply',
        'amuse',
        'notice',
        'struggle',
        'stare',
        'brighten',
        'hesitate',
        'encourage'],
 '명사': ['diversity',
        'unity',
        'revision',
        'manuscript',
        'cooperation',
        'complaint',
        'anniversary',
        'compensation',
        'refund',
        'repayment',
        'loan',
        'transportation',
        'immigration',
        'document',
        'property',
        'rejection',
        'assistance',
        'possession',
        'comfort',
        'weapon',
        'trunk',
        'prayer',
        'irony',
        'scent',
        'ease',
        'microclimate',
        'scale',
        'promotion',
        'manipulation',
        'loyalty',
        'astronaut',
        'statement',
        'phenomenon',
        'blueprint',
        'dictator',
        'blessing',
        'vehicle',
        'assumption',
        'gravity',
        'sequence',
        'crop',
        'filter',
        'civilization',
        'fraction',
        'fuel',
        'advancement',
        'engagement',
        'journalist',
        'democratization',
        'disruption',
        'shareholder',
        'broadcaster',
        'gatekeeper',
        'myth',
        'commitment',
        'accomplishment',
        'intuition',
        'testing',
        'hypothesis',
        'justification',
        'awkwardness',
        'anxiety',
        'icebreaker',
        'lifeline',
        'empathy',
        'compassion',
        'distribution',
        'portfolio',
        'nutrition',
        'herb',
        'facility',
        'servant',
        'athlete',
        'distinction',
        'democracy',
        'voice',
        'abstract',
        'rubbish',
        'fascination',
        'behaviorism',
        'metaphor',
        'anthropology',
        'founder',
        'construction',
        'reflection',
        'emphasis',
        'ingredient',
        'cognition',
        'identification',
        'comparison',
        'alternative',
        'perception',
        'imagery',
        'representation',
        'occupation',
        'instrument',
        'figure',
        'nonsense',
        'beverage',
        'capacity',
        'priority',
        'thirst',
        'negotiator',
        'component',
        'symposium',
        'transition',
        'precision',
        'experimentation',
        'proportion',
        'housing',
        'accommodation',
        'destination',
        'outlook',
        'income',
        'adoption',
        'respondent',
        'appliance',
        'equipment',
        'chemistry',
        'substance',
        'mixture',
        'textile',
        'portrait',
        'geology',
        'mentor',
        'interpretation',
        'registration',
        'session',
        'gear',
        'application',
        'conclusion',
        'boundary',
        'incidence',
        'obesity',
        'coincidence',
        'distortion',
        'moral',
        'outrage',
        'irrationality',
        'significance',
        'perspective',
        'analysis',
        'regularity',
        'sharpness',
        'arrangement',
        'variation',
        'odds',
        'sensation',
        'faculty',
        'posture',
        'criticism',
        'overextension',
        'caution',
        'comprehension',
        'knowledge',
        'instructor',
        'cue',
        'injury',
        'mindset',
        'device',
        'digit',
        'calculator',
        'curriculum',
        'stress',
        'abstraction',
        'cluster',
        'calculation',
        'indicator',
        'seed',
        'academia',
        'administrator',
        'combination',
        'patch',
        'territory',
        'contribution',
        'self-esteem',
        'performance',
        'proximity',
        'element',
        'pitch',
        'wetland',
        'provision',
        'concentration',
        'capability',
        'output',
        'viewpoint',
        'confusion',
        'episode',
        'tragedy',
        'symptom',
        'mechanism',
        'resource',
        'fatigue',
        'attention',
        'inactivity',
        'despair',
        'barn',
        'tractor',
        'fondness',
        'indication',
        'outcome',
        'narrative',
        'fashion',
        'fabric',
        'sustainability',
        'silhouette',
        'trimmings',
        'reliability',
        'opponent',
        'willingness',
        'demonstration',
        'aggressiveness',
        'disarmament',
        'spiral',
        'intention',
        'hostility',
        'invasion',
        'centerpiece',
        'confidence',
        'vendor',
        'patience',
        'instruction',
        'paragraph',
        'impression',
        'purpose'],
 '형용사': ['unique',
         'insightful',
         'critical',
         'lifeless',
         'spotted',
         'enthusiastic',
         'tiresome',
         'ridiculous',
         'superior',
         'dizzy',
         'precious',
         'cruel',
         'soaked',
         'unassuming',
         'legendary',
         'harsh',
         'flat',
         'countless',
         'pitiless',
         'deadly',
         'controversial',
         'solar',
         'constant',
         'sophisticated',
         'relevant',
         'accountable',
         'valid',
         'absolute',
         'ultimate',
         'normal',
         'internal',
         'insecure',
         'inherent',
         'overwhelming',
         'oversized',
         'manageable',
         'regulatory',
         'misleading',
         'reputable',
         'acceptable',
         'alien',
         'compatible',
         'convincing',
         'inclined',
         'materialistic',
         'cognitive',
         'humanistic',
         'essential',
         'utter',
         'innate',
         'bold',
         'freehand',
         'reflective',
         'opposite',
         'common',
         'distinguished',
         'vibrant',
         'affordable',
         'inner',
         'outstanding',
         'seasonal',
         'upcoming',
         'urban',
         'distinct',
         'civilized',
         'positive',
         'cooperative',
         'counterproductive',
         'objective',
         'disposable',
         'vocal',
         'effortful',
         'fundamental',
         'specific',
         'numerical',
         'crucial',
         'leading',
         'rough',
         'inhabitable',
         'colonial',
         'normative',
         'associative',
         'negative',
         'spatial',
         'diverse',
         'apparent',
         'shallow',
         'demanding',
         'logical',
         'subjective',
         'instrumental',
         'psychological',
         'remarkable',
         'affectionate',
         'immediate',
         'verified',
         'multiple',
         'straightforward',
         'supplementary',
         'prominent',
         'rational',
         'invisible',
         'bygone',
         'unquestioned',
         'secondhand',
         'economical',
         'questionable',
         'portable',
         'compact',
         'accessible',
         'stable',
         'chronic',
         'dominant',
         'preventive',
         'newfound',
         'nervous',
         'retired',
         'renewed'],
 '부사': ['thoroughly',
        'briefly',
        'urgently',
        'initially',
        'approximately',
        'critically',
        'instinctively',
        'dramatically',
        'occasionally',
        'commonly',
        'consistently',
        'unintentionally',
        'ultimately',
        'potentially',
        'coolly',
        'subtly',
        'unconsciously',
        'conversely',
        'evidently',
        'roughly',
        'understandably',
        'essentially',
        'simultaneously',
        'consequently',
        'remarkably',
        'nostalgically',
        'gently',
        'politely',
        'merely',
        'blankly',
        'cautiously']}

# 전체 단어은행
WORD_BANK = []
for word_list in WORD_BANK_BY_POS.values():
    WORD_BANK.extend(word_list)

# =========================================================
# 함수
# =========================================================
def read_docx(file_like):
    doc = Document(file_like)
    paragraphs = []

    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            paragraphs.append(text)

    return "\n".join(paragraphs)


def get_pos_group(tag):
    for group_name, tag_set in POS_GROUPS.items():
        if tag in tag_set:
            return group_name
    return None


def clean_token(token):
    return str(token).strip()


def is_candidate_word(word):
    if not re.fullmatch(r"[A-Za-z]+", word):
        return False

    if len(word) < 4:
        return False

    if word.lower() in STOPWORDS:
        return False

    return True


def assemble_tokens(tokens):
    text = " ".join(tokens)
    text = re.sub(r"\s+([,.!?;:])", r"\1", text)
    text = re.sub(r"\(\s+", "(", text)
    text = re.sub(r"\s+\)", ")", text)
    text = re.sub(r"\s+'", "'", text)
    return text


def make_options(correct_word, pos_group):
    correct_word = clean_token(correct_word)
    correct_lower = correct_word.lower()

    pool_source = WORD_BANK_BY_POS.get(pos_group, WORD_BANK)

    pool = []
    seen = set()

    for word in pool_source:
        word = clean_token(word)
        word_lower = word.lower()

        # 구동사/숙어는 radio 보기로 길어질 수 있어서 일단 단일어 위주로 사용
        if " " in word or "~" in word or "∼" in word:
            continue

        if word_lower == correct_lower:
            continue

        if word_lower in seen:
            continue

        seen.add(word_lower)
        pool.append(word)

    # 혹시 품사별 풀이 부족하면 전체 단어은행에서 보충
    if len(pool) < 4:
        for word in WORD_BANK:
            word = clean_token(word)
            word_lower = word.lower()

            if " " in word or "~" in word or "∼" in word:
                continue

            if word_lower == correct_lower:
                continue

            if word_lower in seen:
                continue

            seen.add(word_lower)
            pool.append(word)

            if len(pool) >= 4:
                break

    wrong_options = random.sample(pool, 4)
    options = wrong_options + [correct_word]
    random.shuffle(options)

    return options


def generate_quiz(text, pos_choice, blank_ratio):
    tokens = word_tokenize(text)

    try:
        tagged = pos_tag(tokens, lang="eng")
    except Exception:
        tagged = pos_tag(tokens)

    candidates = []

    for idx, (word, tag) in enumerate(tagged):
        if not is_candidate_word(word):
            continue

        group = get_pos_group(tag)

        # 같은 품사 선지를 만들기 위해 품사 분류가 안 되는 단어는 제외
        if group is None:
            continue

        if pos_choice == "전체" or group == pos_choice:
            candidates.append((idx, group))

    if not candidates:
        return text, {}, {}

    n_blanks = max(1, round(len(candidates) * blank_ratio))
    n_blanks = min(n_blanks, len(candidates))

    selected = random.sample(candidates, n_blanks)
    selected_map = {idx: group for idx, group in selected}

    answer_map = {}
    option_map = {}
    blank_no = 1

    output_tokens = []

    for idx, token in enumerate(tokens):
        if idx in selected_map:
            pos_group = selected_map[idx]
            answer_map[blank_no] = token
            option_map[blank_no] = make_options(token, pos_group)
            output_tokens.append(f"({blank_no})________")
            blank_no += 1
        else:
            output_tokens.append(token)

    question_text = assemble_tokens(output_tokens)

    return question_text, answer_map, option_map


def get_selected_word(selected_text):
    # "① acquire"에서 acquire만 추출
    if not selected_text:
        return ""

    return re.sub(r"^[①②③④⑤]\s*", "", selected_text).strip()


def reset_answers():
    for key in list(st.session_state.keys()):
        if str(key).startswith("user_answer_"):
            del st.session_state[key]


# =========================================================
# UI
# =========================================================
st.title("📘 수능형 Cloze Test Generator")

st.caption("Word(.docx) 파일을 업로드하면 왼쪽에는 문제지, 오른쪽에는 수능특강 어휘 기반 오지선다 보기가 표시됩니다.")

top1, top2 = st.columns(2)

with top1:
    class_name = st.text_input("반", value="", placeholder="예: 중3A반")

with top2:
    student_name = st.text_input("이름", value="", placeholder="예: 홍길동")

st.markdown("---")

control1, control2 = st.columns(2)

with control1:
    pos_choice = st.selectbox(
        "빈칸으로 만들 품사 선택",
        ["전체", "동사", "명사", "형용사", "부사"]
    )

with control2:
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

btn1, btn2 = st.columns(2)

with btn1:
    make_button = st.button("📄 문제 만들기", use_container_width=True)

with btn2:
    reset_button = st.button("🧹 초기화", use_container_width=True)

if reset_button:
    st.session_state.clear()
    st.rerun()

if uploaded_file is None:
    st.info("먼저 Word(.docx) 파일을 업로드하세요.")

if uploaded_file is not None and make_button:
    uploaded_file.seek(0)

    try:
        original_text = read_docx(uploaded_file)

        question_text, answer_map, option_map = generate_quiz(
            original_text,
            pos_choice,
            blank_pct / 100.0
        )

        st.session_state["question_text"] = question_text
        st.session_state["answer_map"] = answer_map
        st.session_state["option_map"] = option_map
        st.session_state["class_name"] = class_name
        st.session_state["student_name"] = student_name

        reset_answers()

        st.rerun()

    except Exception as e:
        st.error("문제 생성 중 오류가 발생했습니다.")
        st.exception(e)

st.markdown("---")

left_col, right_col = st.columns([1.25, 1])

with left_col:
    st.subheader("📝 문제지")

    if "question_text" not in st.session_state:
        st.info("문제를 생성하면 여기에 문제지가 표시됩니다.")
    else:
        meta = []
        if st.session_state.get("class_name"):
            meta.append(f"반: {st.session_state['class_name']}")
        if st.session_state.get("student_name"):
            meta.append(f"이름: {st.session_state['student_name']}")

        if meta:
            st.markdown(" / ".join(meta))

        safe_text = html.escape(st.session_state["question_text"]).replace("\n", "<br>")

        st.markdown(
            f"""
            <div style="
                border:1px solid #ddd;
                border-radius:12px;
                padding:20px;
                background-color:#ffffff;
                line-height:1.9;
                font-size:17px;
                max-height:720px;
                overflow-y:auto;
                white-space:normal;
            ">
                {safe_text}
            </div>
            """,
            unsafe_allow_html=True
        )

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
            labels = ["①", "②", "③", "④", "⑤"]

            for num in sorted(answer_map.keys()):
                options = option_map.get(num, [])

                st.markdown(f"### {num}번")

                display_options = [
                    f"{labels[i]} {option}"
                    for i, option in enumerate(options)
                ]

                st.radio(
                    label=f"{num}번 보기",
                    options=display_options,
                    key=f"user_answer_{num}",
                    label_visibility="collapsed"
                )

            st.markdown("---")

            if st.button("✅ 채점하기", use_container_width=True):
                correct_count = 0
                total = len(answer_map)

                st.subheader("📊 채점 결과")

                for num in sorted(answer_map.keys()):
                    selected_text = st.session_state.get(f"user_answer_{num}", "")
                    selected_word = get_selected_word(selected_text)
                    correct_word = answer_map[num]

                    if selected_word.strip().lower() == correct_word.strip().lower():
                        correct_count += 1
                        st.success(f"{num}번 정답")
                    else:
                        st.error(
                            f"{num}번 오답 / 선택: {selected_word} / 정답: {correct_word}"
                        )

                score_pct = (correct_count / total) * 100 if total else 0
                st.markdown("---")
                st.write(f"총 {total}문항 중 **{correct_count}개** 정답입니다.")
                st.write(f"점수: **{score_pct:.1f}점 / 100점**")
