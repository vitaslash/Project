# Анализ выгрузки обзвона пациентов
# Версия релиза: 1.3.0

import pandas as pd
import numpy as np
import argparse
from pathlib import Path
import sys

def clean_data(df):
    """Очистка и подготовка данных."""
    # Установка заголовков из 4-й строки
    df.columns = df.iloc[3]
    # Удаление первых 4 строк и сброс индекса
    df = df.iloc[4:].reset_index(drop=True)
    # Удаление полностью пустых строк
    df = df.dropna(how='all').reset_index(drop=True)
    # Очистка названий столбцов
    df.columns = [str(c).strip() for c in df.columns]
    # Удаление итоговой строки если она есть
    if len(df) > 0 and str(df.iloc[-1,0]).strip().lower().startswith('всего'):
        df = df.iloc[:-1].reset_index(drop=True)
    return df

def analyze_data(df):
    """Анализ данных и расчет метрик."""
    # Определение колонок с вопросами и отделениями
    question_cols = df.columns[5:13]
    dept_col = df.columns[2]
    
    # Подсчет общего количества звонков
    total_calls = len(df)
    
    # Функции для анализа
    def count_answers(row):
        """Подсчет количества ответов в строке."""
        vals = row[question_cols]
        return sum(str(v).strip().isdigit() and 1 <= int(str(v).strip()) <= 10 for v in vals)
    
    def calculate_csi(row):
        """Расчет CSI для строки."""
        vals = row[question_cols]
        nums = [int(str(v).strip()) for v in vals if str(v).strip().isdigit() and 1 <= int(str(v).strip()) <= 10]
        return sum(nums) / len(nums) if nums else None
    
    # Расчет метрик для каждой строки
    df['_answers'] = df.apply(count_answers, axis=1)
    df['_csi'] = df.apply(calculate_csi, axis=1)
    
    # Общая статистика
    all_answered = (df['_answers'] == len(question_cols)).sum()
    any_answered = (df['_answers'] > 0).sum()
    
    # Статистика по отделениям
    dept_stats = df.groupby(dept_col).agg(
        звонков=(df.columns[0], 'count'),
        средний_CSI=('_csi', 'mean'),
        ответили_все=(df.columns[0], lambda x: (df.loc[x.index, '_answers'] == len(question_cols)).sum()),
        ответили_хотябы=(df.columns[0], lambda x: (df.loc[x.index, '_answers'] > 0).sum()),
    )
    dept_stats['%_ответили_все'] = dept_stats['ответили_все'] / dept_stats['звонков'] * 100
    dept_stats['%_ответили_хотябы'] = dept_stats['ответили_хотябы'] / dept_stats['звонков'] * 100
    dept_stats['средний_CSI'] = dept_stats['средний_CSI'].round(2)
    
    # Статистика по вопросам
    question_stats = {}
    for i, qcol in enumerate(question_cols, 1):
        q_stats = df.groupby(dept_col)[qcol].apply(
            lambda vals: np.mean([
                int(str(v).strip()) 
                for v in vals 
                if str(v).strip().isdigit() and 1 <= int(str(v).strip()) <= 10
            ])
        )
        question_stats[f'Вопрос {i}'] = q_stats
    
    return {
        'total_calls': total_calls,
        'all_answered': all_answered,
        'any_answered': any_answered,
        'dept_stats': dept_stats,
        'question_stats': question_stats,
        'cleaned_data': df
    }

def save_results(results, output_dir):
    """Сохранение результатов анализа."""
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Сохранение очищенных данных
    results['cleaned_data'].to_csv(output_dir / 'cleaned_data.csv', index=False, encoding='utf-8-sig')
    
    # Сохранение статистики по отделениям
    results['dept_stats'].to_csv(output_dir / 'department_stats.csv', encoding='utf-8-sig')
    
    # Сохранение статистики по вопросам
    question_stats_df = pd.DataFrame(results['question_stats'])
    question_stats_df.to_csv(output_dir / 'question_stats.csv', encoding='utf-8-sig')
    
    # Сохранение общей статистики
    with open(output_dir / 'summary.txt', 'w', encoding='utf-8') as f:
        f.write(f"Всего звонков: {results['total_calls']}\n")
        f.write(f"Ответили на все вопросы: {results['all_answered']}\n")
        f.write(f"Процент ответивших на все вопросы: {results['all_answered']/results['total_calls']*100:.1f}%\n")
        f.write(f"Ответили хотя бы на один вопрос: {results['any_answered']}\n")
        f.write(f"Процент ответивших хотя бы на один вопрос: {results['any_answered']/results['total_calls']*100:.1f}%\n")

def main():
    parser = argparse.ArgumentParser(description='Анализ данных автообзвона пациентов')
    parser.add_argument('input_file', help='Путь к входному Excel файлу')
    parser.add_argument('-o', '--output', default='output', help='Директория для сохранения результатов')
    args = parser.parse_args()
    
    try:
        print(f"Чтение файла: {args.input_file}")
        if args.input_file.lower().endswith('.csv'):
            df = pd.read_csv(args.input_file, header=None)
        else:
            df = pd.read_excel(args.input_file, header=None)
        
        print("Очистка данных...")
        df = clean_data(df)
        
        print("Анализ данных...")
        results = analyze_data(df)
        
        print(f"Сохранение результатов в: {args.output}")
        save_results(results, args.output)
        
        print("Готово!")
        return 0
        
    except Exception as e:
        print(f"Ошибка: {str(e)}", file=sys.stderr)
        return 1

if __name__ == '__main__':
    sys.exit(main())