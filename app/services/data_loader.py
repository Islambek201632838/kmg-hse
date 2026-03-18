import os
import pandas as pd
from functools import lru_cache
from app.config import settings


def _hour_to_shift(hour) -> str:
    if pd.isna(hour):
        return "Неизвестно"
    h = int(hour)
    if 6 <= h < 14:
        return "День (06-14)"
    elif 14 <= h < 22:
        return "Вечер (14-22)"
    else:
        return "Ночь (22-06)"


@lru_cache(maxsize=1)
def load_incidents() -> pd.DataFrame:
    path = os.path.join(settings.data_dir, "incidents.xlsx")
    df = pd.read_excel(path, engine="openpyxl")
    return _clean_incidents(df)


@lru_cache(maxsize=1)
def load_korgau() -> pd.DataFrame:
    path = os.path.join(settings.data_dir, "korgau_cards.xlsx")
    df = pd.read_excel(path, engine="openpyxl")
    return _clean_korgau(df)


def _clean_incidents(df: pd.DataFrame) -> pd.DataFrame:
    # Даты
    df["_date"] = pd.to_datetime(df["Дата возникновения происшествия"], errors="coerce")
    df["_month"] = df["_date"].dt.to_period("M")
    df["_year"] = df["_date"].dt.year

    # Организация
    df["_org"] = df["Наименование организации ДЗО"].fillna("Неизвестно").astype(str).str.strip()

    # Тип: определяем по флаговым колонкам (1 = да, NaN = нет)
    def _derive_type(row) -> str:
        if row.get("Несчастный случай") == 1:
            return "Несчастный случай"
        if row.get("Дорожно-транспортное происшествие") == 1:
            return "ДТП"
        if row.get("Пожар/Возгорание") == 1:
            return "Пожар/Возгорание"
        if row.get("Инцидент") == 1:
            return "Инцидент"
        if row.get("Оказание Медицинской помощи/микротравма") == 1:
            return "Микротравма"
        return "Прочее"

    df["_type"] = df.apply(_derive_type, axis=1)

    # Описание и причины
    df["_description"] = df["Краткое описание происшествия"].fillna("").astype(str)
    df["_cause"] = df["Предварительные причины"].fillna("").astype(str)

    # Тяжесть
    df["_severity"] = df["Тяжесть травмы"].fillna("Нет данных").astype(str)

    # Локация
    df["_location"] = df["Место происшествия"].fillna("Неизвестно").astype(str)

    # Время суток → час → смена
    df["_time_raw"] = pd.to_datetime(
        df["Время возникновения происшествия"], errors="coerce"
    )
    df["_hour"] = df["_time_raw"].dt.hour
    df["_shift"] = df["_hour"].apply(_hour_to_shift)

    # День недели
    df["_weekday"] = df["_date"].dt.day_name()

    # Стаж (числовое значение из строки вида "3 года", "5 лет")
    df["_experience_years"] = (
        df["Стаж работы в организации"]
        .astype(str)
        .str.extract(r"(\d+)")
        .astype(float)
    )

    # В рабочее время
    df["_in_work_hours"] = df["В рабочее время"].fillna("Нет").astype(str)

    return df.dropna(subset=["_date"]).reset_index(drop=True)


def _clean_korgau(df: pd.DataFrame) -> pd.DataFrame:
    # Даты
    df["_date"] = pd.to_datetime(df["Дата"], errors="coerce")
    df["_month"] = df["_date"].dt.to_period("M")
    df["_year"] = df["_date"].dt.year

    # Организация
    df["_org"] = df["Организация"].fillna("Неизвестно").astype(str).str.strip()

    # Категория (может быть составная через запятую — берём первую)
    df["_category"] = (
        df["Категория наблюдения"]
        .fillna("Прочее")
        .astype(str)
        .str.split(",")
        .str[0]
        .str.strip()
    )

    # Тип наблюдения: хорошая / плохая практика / прочее
    df["_obs_type"] = df["Тип наблюдения"].fillna("Неизвестно").astype(str).str.strip()
    df["_is_violation"] = ~df["_obs_type"].isin(["Хорошая практика", "Предложение (инициатива)"])

    # Описание
    df["_description"] = df["Опишите ваше наблюдение/предложение"].fillna("").astype(str)

    # Остановка работ
    df["_work_stopped"] = df["Производилась ли остановка работ?"].fillna("Нет").astype(str)

    # Статус устранения
    resolved_col = "Было ли небезопасное условие / поведение исправлено и опасность устранена?"
    df["_resolved"] = df[resolved_col].fillna("Нет").astype(str)
    df["_is_resolved"] = df["_resolved"].str.lower().str.contains("да")

    # Час наблюдения + смена
    df["_time_raw"] = pd.to_datetime(df["Время"], errors="coerce")
    df["_hour"] = df["_time_raw"].dt.hour
    df["_shift"] = df["_hour"].apply(_hour_to_shift)

    # День недели
    df["_weekday"] = df["_date"].dt.day_name()

    # Сообщили ответственному
    df["_reported"] = df["Сообщили ли ответственному лицу?"].fillna("Нет").astype(str)

    return df.dropna(subset=["_date"]).reset_index(drop=True)
