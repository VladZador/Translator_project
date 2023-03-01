from flask import Flask, render_template, redirect, url_for, request
from flask_wtf import FlaskForm
from wtforms import FileField, SubmitField
from wtforms.validators import InputRequired
from werkzeug.utils import secure_filename
import os
from dotenv import load_dotenv

from utils import translate

load_dotenv()

app = Flask(__name__)
app.config["SECRET_KEY"] = "translate_secret_key"
app.config["UPLOAD_FOLDER"] = os.environ.get("UPLOAD_FOLDER")


class UploadFileForm(FlaskForm):
    file = FileField("File", validators=[InputRequired()])
    submit = SubmitField("Upload File")


@app.route("/", methods=["GET", "POST"])
def file_upload():
    form = UploadFileForm()
    if form.validate_on_submit():
        file = form.file.data  # Grab the file
        file.save(os.path.join(
            os.path.abspath(os.path.dirname(__file__)),
            app.config["UPLOAD_FOLDER"],
            secure_filename(file.filename)
        ))  # Save the file
        simple = True if not request.form.get("complex") else False
        breakpoint()
        html_file = translate(secure_filename(file.filename), simple=simple)
        return redirect(
            url_for("translated_file", filename=html_file.name.split("/")[1])
        )
    return render_template("index.html", form=form)


@app.route("/<path:filename>")
def translated_file(filename):
    return render_template(filename)


if __name__ == "__main__":
    app.run(debug=True)


# todo: separate functions into different directories

# todo: add removal of old media files when starting a new translation
