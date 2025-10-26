from flask import Flask, render_template, request, send_file
import pandas as pd
import io
import re

app = Flask(__name__)

EXCEL_FILE = "C:/Users/claud/OneDrive/Desktop/NBB_Lists/nbb_flask/NBBC23 Spreadsheet Lists_edited.xlsx"

# Load all sheets once at startup
xls = pd.ExcelFile(EXCEL_FILE)
data = {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names}

def search_terms(df, query, columns=None):
    if columns is None:
        columns = df.columns
    mask = pd.Series(False, index=df.index)
    for col in columns:
        mask |= df[col].astype(str).str.contains(query, case=False, na=False)
    return df[mask]

def highlight_text(text, query):
    # Handle missing values first
    if pd.isna(text):
        return "-"
    if not query:
        return str(text)
    pattern = re.compile(re.escape(query), re.IGNORECASE)
    return pattern.sub(lambda m: f"<mark>{m.group(0)}</mark>", str(text))

def search_all(query, sheet=None):
    results = {}
    if sheet and sheet in data:
        hits = search_terms(data[sheet], query)
        if not hits.empty:
            hits = hits.applymap(lambda v: highlight_text(v, query))
            results[sheet] = hits
    else:
        for s, df in data.items():
            hits = search_terms(df, query)
            if not hits.empty:
                hits = hits.applymap(lambda v: highlight_text(v, query))
                results[s] = hits
    return results

@app.route("/", methods=["GET", "POST"])
def index():
    sheets = list(data.keys())
    if request.method == "POST":
        query = request.form.get("query", "")
        sheet = request.form.get("sheet", "")
        results = search_all(query, sheet if sheet != "All" else None)
        return render_template("results.html", query=query, results=results,
                               sheets=sheets, selected=sheet)
    return render_template("index.html", sheets=sheets, selected="All")

@app.route("/clear", methods=["GET"])
def clear():
    return render_template("index.html", sheets=list(data.keys()), selected="All")

@app.route("/export", methods=["POST"])
def export():
    query = request.form.get("query", "")
    sheet = request.form.get("sheet", "")
    results = search_all(query, sheet if sheet != "All" else None)

    if not results:
        return "No results to export."

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for sheet_name, df in results.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    output.seek(0)

    return send_file(output,
                     as_attachment=True,
                     download_name=f"search_results_{query}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    app.run(debug=True)