import pandas as pd
from flask import Flask, request
from flask.templating import render_template
from flask_wtf import FlaskForm
from flask_wtf.file import FileField, FileRequired

app = Flask(__name__)
app.debug = True
app.secret_key = "secret"


class UploadFileForm(FlaskForm):
    sql_xl = FileField()
    work_to = FileField()
    orderbook = FileField()


@app.route("/", methods=["GET", "POST"])
def home():
    form = UploadFileForm()
    table = None
    if form.validate_on_submit():
        sql_data = request.files["sql_xl"].read()
        df = pd.read_excel(sql_data, header=8)
        table = df.to_html(
            classes=["table", "table-hover", "table-responsive"], border=0
        )
    return render_template("index.html", form=form, table=table)


if __name__ == "__main__":
    app.run()
