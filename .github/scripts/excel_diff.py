import pandas as pd
import os

def load_excel(path):
    return pd.read_excel(path, sheet_name=None)

def diff_excel(old, new):
    diffs = []

    for sheet in old:
        if sheet not in new:
            diffs.append(f"Sheet removed: {sheet}")
            continue

        df1 = old[sheet].fillna("")
        df2 = new[sheet].fillna("")

        # 行列サイズ合わせ
        df2 = df2.reindex_like(df1).fillna("")

        diff = df1.compare(df2)

        if not diff.empty:
            for idx, row in diff.iterrows():
                for col in diff.columns.levels[0]:
                    before = row[(col, "self")]
                    after = row[(col, "other")]
                    diffs.append({
                        "sheet": sheet,
                        "row": idx,
                        "column": col,
                        "before": before,
                        "after": after
                    })

    return diffs

def to_markdown(diffs):
    if not diffs:
        return "## Excel差分\n差分はありません"

    lines = ["## Excel差分", "", "| Sheet | Row | Column | Before | After |", "|---|---|---|---|---|"]

    for d in diffs:
        lines.append(f"| {d['sheet']} | {d['row']} | {d['column']} | {d['before']} | {d['after']} |")

    return "\n".join(lines)

if __name__ == "__main__":
    old = load_excel(os.environ["OLD_FILE"])
    new = load_excel(os.environ["NEW_FILE"])
    diffs = diff_excel(old, new)
    md = to_markdown(diffs)

    print(md)
