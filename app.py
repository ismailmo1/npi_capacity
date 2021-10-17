import pandas as pd
from flask import Flask, request
from flask.templating import render_template
from flask_wtf import FlaskForm
from flask_wtf.file import FileField, FileRequired
from wtforms.fields.core import IntegerField

app = Flask(__name__)
app.debug = True
app.secret_key = "secret"


class CapacityAnalysisForm(FlaskForm):
    sql_xl = FileField(label="ERP Query Results")
    work_to = FileField(label="Work to List")
    orderbook = FileField(label="Orderbook")
    mould_setup_mins = IntegerField(default=90)
    pre_mould_days = IntegerField(default=2)
    post_mould_days = IntegerField(default=9)
    doc_days = IntegerField(default=4)
    total_shift_mins = IntegerField(default=360)


@app.route("/", methods=["GET", "POST"])
def home():
    cap_analysis_form = CapacityAnalysisForm()
    table = None
    if cap_analysis_form.validate_on_submit():
        sql_data = request.files["sql_xl"].read()
        df = pd.read_excel(sql_data, header=8)
        table = df.to_html(
            classes=["table", "table-hover", "table-responsive"], border=0
        )
    return render_template("index.html", form=cap_analysis_form, table=table)


if __name__ == "__main__":
    app.run()
