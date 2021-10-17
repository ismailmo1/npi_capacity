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
    if form.validate_on_submit():
        sql_data = request.files["sql_xl"].read()
    return render_template("index.html", form=form)


app.run()
