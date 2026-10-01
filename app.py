import os
import tempfile
from pathlib import Path

import streamlit as st
from dotenv import load_dotenv

from smeta_pipeline import PipelineError, generate_smeta

load_dotenv()

st.set_page_config(page_title="Тендер-агент — смета", layout="centered")

APP_PASSWORD = os.environ.get("APP_PASSWORD")


def _check_password() -> bool:
    if st.session_state.get("authenticated"):
        return True
    st.title("Тендер-агент")
    pwd = st.text_input("Пароль", type="password")
    if st.button("Войти"):
        if APP_PASSWORD and pwd == APP_PASSWORD:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Неверный пароль")
    return False


if not _check_password():
    st.stop()

st.title("Тендер-агент — расчёт сметы")
st.caption(
    "Загрузите ТЗ нового тендера на уборку ПОМЕЩЕНИЙ — конвейер извлечёт данные "
    "через Claude API и соберёт оценочную смету (ФОТ и материалы — по нормативам). "
    "ТЗ на уборку территории/благоустройство пока не поддерживаются (другая методика "
    "расчёта площади) — результат для них будет недостоверным без предупреждения."
)

uploaded = st.file_uploader("Загрузите ТЗ (PDF/DOCX/XLSX)", type=["pdf", "docx", "xlsx"])

col1, col2 = st.columns(2)
with col1:
    region = st.selectbox("Регион", ["Москва"], index=0)
    object_complexity = st.selectbox("Сложность объекта", ["стандартный"], index=0)
with col2:
    schedule_complexity = st.selectbox(
        "График",
        [
            "комбинированная смена (ежедневная-основная + "
            "ежедневная-поддерживающая, один сотрудник)"
        ],
        index=0,
    )
    vat_rate_pct = st.number_input(
        "Ставка НДС, %", min_value=0.0, max_value=100.0, value=20.0, step=1.0,
    )

st.caption(
    "Регион/сложность/график сейчас ограничены единственной заполненной "
    "комбинацией в базе нормативов (knowledge_base/salary_rules.json)."
)

if uploaded is not None and st.button("Рассчитать"):
    with tempfile.NamedTemporaryFile(delete=False, suffix=Path(uploaded.name).suffix) as tmp:
        tmp.write(uploaded.getbuffer())
        tmp_path = tmp.name
    try:
        with st.spinner("Извлекаем данные из ТЗ и считаем смету (может занять до минуты)..."):
            result = generate_smeta(
                tmp_path,
                region=region,
                object_complexity=object_complexity,
                schedule_complexity=schedule_complexity,
                vat_rate=vat_rate_pct / 100,
            )
    except PipelineError as e:
        st.error(str(e))
    except Exception as e:
        st.error(f"Неожиданная ошибка: {e}")
    else:
        st.session_state["last_result"] = result
    finally:
        os.unlink(tmp_path)

if "last_result" in st.session_state:
    result = st.session_state["last_result"]
    st.success(f"Смета готова: {result.object_name}")
    m1, m2, m3 = st.columns(3)
    m1.metric("Площадь, м²", f"{result.area_sqm:,.1f}")
    m1.metric("Дней уборки", result.cleaning_days)
    m2.metric("Срок, мес", result.contract_months)
    m2.metric("Численность", result.staff_count, help=result.staff_count_source)
    m3.metric(
        "ФОТ, ₽/мес",
        f"{result.fot_month:,.0f}" if result.fot_month is not None else "не найден",
    )
    for w in result.warnings:
        st.warning(w)
    with open(result.xlsx_path, "rb") as f:
        st.download_button(
            "Скачать смету (.xlsx)",
            data=f.read(),
            file_name=Path(result.xlsx_path).name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
