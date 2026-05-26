
from flask import Flask, render_template
import pandas as pd

app = Flask(__name__)

EXCEL_FILE = "Job Card List 202605260830.xlsx"

@app.route("/")
def dashboard():
    df = pd.read_excel(EXCEL_FILE)

    total_tickets = len(df)
    active_tickets = len(df[df["Job Status"].notna()])

    stuck = df[df["Days In Status"] >= 3]

    avg_days = round(df["Days In Status"].mean(), 2)

    status_summary = (
        df.groupby("Job Status")
        .size()
        .reset_index(name="Count")
        .sort_values("Count", ascending=False)
    )

    tickets = df.to_dict(orient="records")

    return render_template(
        "index.html",
        total_tickets=total_tickets,
        active_tickets=active_tickets,
        stuck_count=len(stuck),
        avg_days=avg_days,
        status_summary=status_summary.to_dict(orient="records"),
        tickets=tickets
    )

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
