import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go

# Настройка страницы и стилей
st.set_page_config(
    page_title='AutoCall — Аналитика',
    layout='wide',
    initial_sidebar_state='expanded'
)

# Кастомные стили CSS
st.markdown("""
<style>
    /* Основной шрифт */
    body {
        font-family: 'Segoe UI', Arial, sans-serif;
        color: #2c3e50;
    }
    
    /* Заголовки */
    h1 {
        color: #2c3e50;
        font-weight: 600;
        font-size: 2.5rem;
        padding-bottom: 1rem;
        border-bottom: 2px solid #3498db;
        margin-bottom: 2rem;
    }
    
    h2 {
        color: #34495e;
        font-weight: 500;
        font-size: 1.8rem;
        margin-top: 2rem;
        margin-bottom: 1rem;
    }

    /* Таблицы */
    .dataframe {
        font-family: 'Segoe UI', Arial, sans-serif !important;
        font-size: 14px !important;
    }
    
    /* Метрики */
    div[data-testid="stMetricValue"] {
        font-size: 2rem !important;
        color: #2980b9 !important;
    }

    /* Подписи к метрикам */
    div[data-testid="stMetricLabel"] {
        font-size: 1rem !important;
        color: #7f8c8d !important;
    }

    /* Отступы для визуального разделения */
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }

    /* Инфо-блоки */
    div.stAlert {
        font-family: 'Segoe UI', Arial, sans-serif !important;
        border-radius: 8px !important;
        padding: 1rem !important;
    }

    /* Кнопки */
    .stButton > button {
        font-family: 'Segoe UI', Arial, sans-serif !important;
        font-size: 1rem !important;
        padding: 0.5rem 1rem !important;
        border-radius: 8px !important;
    }

    /* Селекторы */
    .stSelectbox > div > div {
        font-family: 'Segoe UI', Arial, sans-serif !important;
    }

    /* Тени для карточек */
    div[data-testid="stMetricValue"] {
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        border-radius: 8px;
        padding: 1rem;
        background: white;
    }

    /* Анимации при наведении */
    .stButton > button:hover {
        transform: translateY(-1px);
        transition: all 0.2s ease;
    }
</style>
""", unsafe_allow_html=True)

# Заголовок с центрированием
st.markdown("<h1 style='text-align: center;'>AutoCall — Аналитика по обзвону пациентов</h1>", unsafe_allow_html=True)

# Описание приложения
st.markdown("""
<div style='text-align: center; padding: 1rem; margin-bottom: 2rem; color: #7f8c8d;'>
    Интерактивный анализ данных обзвона пациентов с визуализацией KPI и CSI по отделениям
</div>
""", unsafe_allow_html=True)

uploaded = st.file_uploader('Загрузите Excel/CSV', type=['xlsx', 'xls', 'csv'])
if uploaded is not None:
    try:
        if uploaded.name.lower().endswith('.csv'):
            df = pd.read_csv(uploaded, header=None)
        else:
            df = pd.read_excel(uploaded, header=None)
        df.columns = df.iloc[3]
        df = df.iloc[4:].reset_index(drop=True)
        df = df.dropna(how='all').reset_index(drop=True)
        df.columns = [str(c).strip() for c in df.columns]
        if len(df) > 0 and str(df.iloc[-1,0]).strip().lower().startswith('всего'):
            df = df.iloc[:-1].reset_index(drop=True)
    except Exception as e:
        st.error('Ошибка чтения файла: ' + str(e))
    else:
        st.write('Размер:', df.shape)
        # Скрывающаяся секция с исходными данными
        with st.expander("📋 Исходные данные", expanded=False):
            st.dataframe(df.head(50), width='stretch')

        question_cols = df.columns[5:13]
        dept_col = df.columns[2]
        total_calls = len(df)

        def count_answers(row):
            vals = row[question_cols]
            return sum(str(v).strip().isdigit() and 1 <= int(str(v).strip()) <= 10 for v in vals)

        def csi(row):
            vals = row[question_cols]
            nums = [int(str(v).strip()) for v in vals if str(v).strip().isdigit() and 1 <= int(str(v).strip()) <= 10]
            return sum(nums) / len(nums) if nums else None

        answers_per_row = df.apply(count_answers, axis=1)
        csi_per_row = df.apply(csi, axis=1)
        df['_answers'] = answers_per_row
        df['_csi'] = csi_per_row

        all_answered = (answers_per_row == len(question_cols)).sum()
        any_answered = (answers_per_row > 0).sum()
        percent_all = all_answered / total_calls * 100 if total_calls else 0
        percent_any = any_answered / total_calls * 100 if total_calls else 0

        # Calculate overall average for question 6 (index 5 in question_cols)
        q6_col = question_cols[5]
        q6_vals = [int(str(v).strip()) for v in df[q6_col] if str(v).strip().isdigit() and 1 <= int(str(v).strip()) <= 10]
        q6_avg = np.mean(q6_vals) if q6_vals else None

        st.markdown("## 📊 Ключевые показатели")

        # KPI метрики в четыре колонки
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric(
                'Всего обзвоненных пациентов',
                f"{total_calls:,}",
                help="Общее количество пациентов в выгрузке"
            )
        with col2:
            st.metric(
                'Ответили на все вопросы',
                f"{all_answered:,}",
                delta=f"{percent_all:.1f}%",
                help="Количество и процент пациентов, ответивших на все вопросы"
            )
        with col3:
            st.metric(
                'Ответили хотя бы на один вопрос',
                f"{any_answered:,}",
                delta=f"{percent_any:.1f}%",
                help="Количество и процент пациентов, ответивших хотя бы на один вопрос"
            )
        with col4:
            st.metric(
                'Средний балл по вопросу 6',
                f"{q6_avg:.2f}" if q6_avg else "Нет данных",
                help="Оцените насколько доброжелательными были с Вами медицинские специалисты"
            )

        st.markdown("## 📈 CSI по отделениям")
        
        num_col = df.columns[0]
        dept_stats = df.groupby(dept_col).agg(
            звонков=(num_col, 'count'),
            средний_CSI=('_csi', 'mean'),
            ответили_все=(num_col, lambda x: (df.loc[x.index, '_answers'] == len(question_cols)).sum()),
            ответили_хотябы=(num_col, lambda x: (df.loc[x.index, '_answers'] > 0).sum()),
        )
        dept_stats['%_ответили_все'] = dept_stats['ответили_все'] / dept_stats['звонков'] * 100
        dept_stats['%_ответили_хотябы'] = dept_stats['ответили_хотябы'] / dept_stats['звонков'] * 100
        dept_stats['средний_CSI'] = dept_stats['средний_CSI'].round(2)

        # Форматирование таблицы с градиентной подсветкой
        styled_stats = dept_stats.reset_index().style\
            .background_gradient(subset=['средний_CSI'], cmap='RdYlGn')\
            .format({
                'средний_CSI': '{:.2f}',
                '%_ответили_все': '{:.1f}%',
                '%_ответили_хотябы': '{:.1f}%'
            })\
            .set_properties(**{
                'font-size': '14px',
                'font-family': 'Segoe UI, Arial, sans-serif',
                'text-align': 'center'
            })
        
        st.dataframe(styled_stats, width='stretch')

        st.markdown("## 📊 Сравнение отделений по вопросам")
        
        # Выбор вопросов для отображения
        selected_questions = st.multiselect(
            'Выберите вопросы для анализа:',
            options=list(enumerate(question_cols, 1)),
            default=list(enumerate(question_cols, 1)),
            format_func=lambda x: f'Вопрос {x[0]}: {x[1]}'
        )

        # Графики только для выбранных вопросов
        for i, qcol in selected_questions:
            q_stats = df.groupby(dept_col)[qcol].apply(
                lambda vals: np.mean(lst) if (lst := [int(str(v).strip()) for v in vals if str(v).strip().isdigit() and 1 <= int(str(v).strip()) <= 10]) else np.nan
            )
            
            fig = go.Figure()
            fig.add_bar(
                x=q_stats.index,
                y=q_stats.values,
                marker_color='skyblue',
                text=q_stats.values.round(2),
                textposition='auto',
                textfont=dict(size=18),
            )

            # Настройка графика
            fig.update_layout(
                height=450,
                margin=dict(t=100, b=50, l=50, r=50),
                yaxis_range=[0, 10],
                title=f'<b>Средний балл по вопросу {i}</b><br>{qcol}',
                xaxis_title='Отделение',
                yaxis_title='Средний балл',
                font=dict(size=14),
                xaxis=dict(tickfont=dict(size=14)),
                yaxis=dict(tickfont=dict(size=14)),
                title_font=dict(size=18)
            )
            
            # Вывод с минимальной конфигурацией
            st.plotly_chart(fig, use_container_width=True)

        # Экспорт данных
        st.markdown("## 💾 Экспорт данных")
        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                "⬇️ Скачать очищенные данные (CSV)",
                df.to_csv(index=False, encoding='utf-8-sig'),
                "cleaned_data.csv",
                "text/csv",
                key='download-csv'
            )
        with col2:
            st.download_button(
                "⬇️ Скачать статистику по отделениям (CSV)",
                dept_stats.reset_index().to_csv(index=False, encoding='utf-8-sig'),
                "department_stats.csv",
                "text/csv",
                key='download-stats'
            )
