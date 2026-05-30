
import streamlit as st
from docx import Document
import nltk
from nltk import word_tokenize, pos_tag
import random
import re

# ---------- NLTK ----------
for pkg, path in [
    ("punkt", "tokenizers/punkt"),
    ("averaged_perceptron_tagger_eng", "taggers/averaged_perceptron_tagger_eng")
]:
    try:
        nltk.data.find(path)
    except:
        nltk.download(pkg)

# ---------- 설정 ----------
st.set_page_config(page_title="수능형 Cloze Test", layout="wide")

POS_GROUPS = {
    "동사": {"VB","VBD","VBG","VBN","VBP","VBZ"},
    "명사": {"NN","NNS","NNP","NNPS"},
    "형용사": {"JJ","JJR","JJS"},
    "부사": {"RB","RBR","RBS"},
}

STOPWORDS = {
    "the","a","an","of","to","for","in","on","at","by","with","from",
    "is","am","are","was","were","be","been","being","and","or","but",
    "that","this","these","those","it","he","she","they","we","you","i",
    "as","if","than","then","into","over","under","about","after","before"
}

WORD_BANK = {
    "동사":["acquire","determine","generate","identify","encourage","enhance","obtain","analyze","improve","evaluate"],
    "명사":["knowledge","analysis","performance","resource","confidence","community","culture","attention","ability","environment"],
    "형용사":["critical","essential","remarkable","effective","complex","significant","valuable","creative","positive","reliable"],
    "부사":["carefully","effectively","significantly","gradually","clearly","actively","successfully","frequently","widely","properly"]
}

# ---------- 함수 ----------
def read_docx(file):
    doc = Document(file)
    text = []
    for p in doc.paragraphs:
        if p.text.strip():
            text.append(p.text.strip())
    return "\n".join(text)

def pos_group(tag):
    for k,v in POS_GROUPS.items():
        if tag in v:
            return k
    return None

def make_options(answer, group):
    pool = WORD_BANK.get(group, [])
    pool = [x for x in pool if x.lower()!=answer.lower()]

    while len(pool) < 4:
        pool += ["sample"]

    wrong = random.sample(pool,4)
    options = wrong + [answer]
    random.shuffle(options)
    return options

def generate(text, pos_choice, blank_ratio):
    tokens = word_tokenize(text)
    tagged = pos_tag(tokens)

    candidates = []

    for idx,(word,tag) in enumerate(tagged):

        if not re.fullmatch(r"[A-Za-z]+", word):
            continue

        if len(word) < 4:
            continue

        if word.lower() in STOPWORDS:
            continue

        group = pos_group(tag)

        if pos_choice == "전체":
            candidates.append((idx,group))
        elif group == pos_choice:
            candidates.append((idx,group))

    if not candidates:
        return text, {}, {}

    n_blank = max(1, round(len(candidates)*blank_ratio))
    selected = random.sample(candidates, min(n_blank,len(candidates)))

    selected_map = {i:g for i,g in selected}

    answer_map = {}
    option_map = {}

    blank_no = 1
    new_tokens = []

    for i,token in enumerate(tokens):

        if i in selected_map:

            answer_map[blank_no] = token
            option_map[blank_no] = make_options(
                token,
                selected_map[i]
            )

            new_tokens.append(f"({blank_no})________")
            blank_no += 1

        else:
            new_tokens.append(token)

    question = " ".join(new_tokens)
    question = re.sub(r"\s+([,.!?;:])", r"\1", question)

    return question, answer_map, option_map

# ---------- UI ----------
st.title("📘 수능형 Cloze Test Generator")

pos_choice = st.selectbox(
    "품사 선택",
    ["전체","동사","명사","형용사","부사"]
)

blank_pct = st.slider(
    "빈칸 비율 (%)",
    5,80,20,5
)

uploaded = st.file_uploader(
    "Word(.docx) 파일 업로드",
    type=["docx"]
)

c1,c2 = st.columns(2)

with c1:
    if st.button("📄 문제 만들기"):

        if uploaded:

            text = read_docx(uploaded)

            q,a,o = generate(
                text,
                pos_choice,
                blank_pct/100
            )

            st.session_state["question"] = q
            st.session_state["answers"] = a
            st.session_state["options"] = o

            for k in list(st.session_state.keys()):
                if str(k).startswith("user_"):
                    del st.session_state[k]

            st.rerun()

with c2:
    if st.button("🧹 초기화"):
        st.session_state.clear()
        st.rerun()

left,right = st.columns([1.3,1])

with left:

    st.subheader("📝 문제지")

    if "question" in st.session_state:
        st.text_area(
            "",
            st.session_state["question"],
            height=700
        )
    else:
        st.info("문제를 생성하세요.")

with right:

    st.subheader("✏️ 오지선다")

    if "options" in st.session_state:

        answers = st.session_state["answers"]
        options = st.session_state["options"]

        labels = ["①","②","③","④","⑤"]

        for num in sorted(options.keys()):

            st.markdown(f"### {num}번")

            display = [
                f"{labels[i]} {opt}"
                for i,opt in enumerate(options[num])
            ]

            st.radio(
                "",
                display,
                key=f"user_{num}"
            )

        if st.button("✅ 채점하기"):

            score = 0
            total = len(answers)

            st.markdown("---")

            for num in sorted(answers.keys()):

                selected = st.session_state.get(
                    f"user_{num}",
                    ""
                )

                selected_word = selected[2:] if len(selected)>2 else ""

                if selected_word == answers[num]:
                    score += 1
                    st.success(f"{num}번 정답")
                else:
                    st.error(
                        f"{num}번 오답 | 정답: {answers[num]}"
                    )

            st.subheader(
                f"점수: {score}/{total} ({score/total*100:.1f}점)"
            )
    else:
        st.info("문제를 생성하면 보기가 표시됩니다.")
