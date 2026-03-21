from __future__ import annotations

from pathlib import Path
from typing import List, Tuple

import numpy as np
import pandas as pd


def normalize_lith_intervals(
    df: pd.DataFrame,
    bhid_col: str = "BHID",
    from_col: str = "FROM",
    to_col: str = "TO",
    lith_col: str = "LITH",
    eps: float = 1e-12,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    required_cols = [bhid_col, from_col, to_col, lith_col]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"В таблице отсутствуют обязательные столбцы: {missing}")

    work = df.copy()
    work["_row_order"] = np.arange(len(work))
    work["_excel_row"] = np.arange(2, len(work) + 2)  # номер строки Excel, если заголовок в 1-й строке

    work[from_col] = pd.to_numeric(work[from_col], errors="coerce")
    work[to_col] = pd.to_numeric(work[to_col], errors="coerce")

    invalid_mask = (
        work[bhid_col].isna()
        | work[from_col].isna()
        | work[to_col].isna()
        | (work[to_col] <= work[from_col])
    )

    skipped = work.loc[invalid_mask].copy()

    def get_reason(row) -> str:
        reasons = []
        if pd.isna(row[bhid_col]):
            reasons.append("BHID пустой")
        if pd.isna(row[from_col]):
            reasons.append("FROM пустой/не число")
        if pd.isna(row[to_col]):
            reasons.append("TO пустой/не число")
        if pd.notna(row[from_col]) and pd.notna(row[to_col]) and row[to_col] <= row[from_col]:
            reasons.append("TO <= FROM")
        return "; ".join(reasons)

    if not skipped.empty:
        skipped["SKIP_REASON"] = skipped.apply(get_reason, axis=1)

    work = work.loc[~invalid_mask].copy()

    if work.empty:
        raise ValueError("После фильтрации не осталось корректных интервалов.")

    work.sort_values([bhid_col, "_row_order"], kind="stable", inplace=True)

    output_records: List[dict] = []
    max_lith_cols = 0

    for bhid, grp in work.groupby(bhid_col, sort=False):
        grp = grp.copy()

        boundaries = np.sort(
            pd.unique(
                np.concatenate(
                    [grp[from_col].to_numpy(dtype=float), grp[to_col].to_numpy(dtype=float)]
                )
            )
        )

        if len(boundaries) < 2:
            continue

        for i in range(len(boundaries) - 1):
            seg_from = float(boundaries[i])
            seg_to = float(boundaries[i + 1])

            if seg_to - seg_from <= eps:
                continue

            active = grp.loc[
                (grp[from_col] < seg_to - eps) & (grp[to_col] > seg_from + eps)
            ].sort_values("_row_order", kind="stable")

            if active.empty:
                continue

            lith_values = active[lith_col].astype(str).tolist()

            rec = {
                bhid_col: bhid,
                from_col: seg_from,
                to_col: seg_to,
            }

            for j, lith in enumerate(lith_values, start=1):
                col_name = lith_col if j == 1 else f"{lith_col}{j}"
                rec[col_name] = lith

            max_lith_cols = max(max_lith_cols, len(lith_values))
            output_records.append(rec)

    if not output_records:
        raise ValueError("Не удалось сформировать выходные интервалы.")

    out = pd.DataFrame(output_records)

    lith_cols = [lith_col] + [f"{lith_col}{i}" for i in range(2, max_lith_cols + 1)]
    final_cols = [bhid_col, from_col, to_col] + lith_cols

    for col in final_cols:
        if col not in out.columns:
            out[col] = pd.NA

    out = out[final_cols]

    skipped = skipped.drop(columns=["_row_order"], errors="ignore")

    return out, skipped


def process_excel_file(
    input_file: str | Path,
    input_sheet: str = "Sheet1",
    output_file: str | Path | None = None,
    output_sheet: str = "Sequential_Lith",
) -> Path:
    input_file = Path(input_file)

    if not input_file.exists():
        raise FileNotFoundError(f"Файл не найден: {input_file}")

    df = pd.read_excel(input_file, sheet_name=input_sheet)
    result, skipped = normalize_lith_intervals(df)

    if output_file is None:
        output_file = input_file.with_name(f"{input_file.stem}_sequential.xlsx")
    else:
        output_file = Path(output_file)

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Original", index=False)
        result.to_excel(writer, sheet_name=output_sheet, index=False)
        if not skipped.empty:
            skipped.to_excel(writer, sheet_name="Skipped_Rows", index=False)

    print(f"Пропущено строк: {len(skipped)}")
    if not skipped.empty:
        print("Они записаны на лист Skipped_Rows")

    return output_file


if __name__ == "__main__":
    input_path = r"C:\Users\1\Documents\TEST.xlsx"
    input_sheet_name = "LITH"

    saved_file = process_excel_file(
        input_file=input_path,
        input_sheet=input_sheet_name,
        output_file=None,
        output_sheet="Sequential_Lith",
    )

    print(f"Готово. Результат сохранен в файл: {saved_file}")
