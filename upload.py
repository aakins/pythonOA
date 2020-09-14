import os
import general
import analysis
import analysis_special
import excel
from flask import Flask, flash, request, redirect, url_for
from werkzeug.utils import secure_filename
from flask import send_from_directory

UPLOAD_FOLDER = '/Projects/Uploads'

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024 * 1024
# app.run(host='0.0.0.0', port=5000,debug=True)

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    import datetime
    starttime = datetime.datetime.now()
    if request.method == 'POST':
        print("Start:" + str(starttime))
        # check if the post request has the file part
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        # if user does not select file, browser also
        # submit an empty part without filename
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        if file and general.allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            general.analyze_file(filename)
            excel.excel(general.xlanalysisfile)
            endtime = datetime.datetime.now()
            print("End:" + str(endtime))
            return redirect(url_for('uploaded_file',filename=general.xlanalysisfilename))
    return '''
    <!doctype html>
    <head>
        <title>Upload gsXXXXX_snapshot.bak.gz File</title>
        <link href="static/css/style.css" rel="stylesheet" type="text/css" media="all">
        <link href="//fonts.googleapis.com/css?family=Oxygen:400,700" rel="stylesheet" type="text/css">
    </head>
    <body>
    <h1>OA Analysis</h1>
    <div class="login-form w3-agile">
        <h2>Upload a Snapshot File</h2>
        <form method=post enctype=multipart/form-data>
          <p>Select your gsXXXXX_snapshot.bak.gz</p>
          <input type=file name=file />
          <input type=submit value=Upload />
        </form>
    </div>
    </body>
    '''

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'],
                               filename)