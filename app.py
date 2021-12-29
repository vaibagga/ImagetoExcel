from flask import Flask, request, render_template, send_from_directory
import os
from utils import *


app = Flask(__name__)
APP_ROOT = os.path.dirname(os.path.abspath(__file__))


@app.route("/")
def main():
    return render_template('index.html')

@app.route('/static/images/<filename>')
def send_image(filename):
    return send_from_directory("static/images", filename)

@app.route("/upload", methods=["POST"])
def upload():
    target = os.path.join(APP_ROOT, 'static/images/')
    if not os.path.isdir(target):
        os.mkdir(target)

    # retrieve file from html file-picker
    upload = request.files.getlist("file")[0]
    print("File name: {}".format(upload.filename))
    filename = upload.filename
    ext = os.path.splitext(filename)[1]
    if (ext == ".jpg") or (ext == ".png") or (ext == ".bmp"):
        print("File accepted")
    else:
        return render_template("error.html", message="The selected file is not supported"), 400

    # save file
    destination = "/".join([target, filename])
    print("File saved to to:", destination)
    upload.save(destination)

    # forward to processing page
    return render_template("processing.html", image_name=filename)

@app.route("/process", methods=["POST"])
def process():
    filename = request.form['image']

    # open and process image
    target = os.path.join(APP_ROOT, 'static/images/')
    destination = "/".join([target, filename])
    image_processor = ImageProcessing(destination)
    #print(destination)
    output_width, output_height = request.form["width"], request.form["height"]
    output_sheet = request.form["output"]
    output_image = image_processor.toGrid((output_width, output_height))
    excel_processor = ExcelProcessing()
    target = os.path.join(APP_ROOT, 'output')
    destination = "/".join([target, output_sheet])
    if os.path.isfile(destination):
        os.remove(destination)
    excel_processor.createExcel(output_image, destination)
    return send_from_directory(target, output_sheet)


if __name__ == "__main__":
    app.run()