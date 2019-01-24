import webbrowser
import pandas as pd
from flask import Flask, render_template, request
import time
from AnalyticstoExcelClass import Analytics


app = Flask(__name__)

@app.route('/')
def index():
    return render_template("index.html", title='Google Analytics')

@app.route('/upload',methods=['POST'])
def upload():
    email = request.form['email']
    startDate = request.form['starttime']
    endDate = request.form['endtime']
    Dim = request.form['Dimentions']
    Met = request.form['Metrics']

    try:
        dimentions = [x.strip() for x in Dim.split(',')]
    except:
        dimentions = [Dim.strip()]
    try:
        metrics = [x.strip() for x in Met.split(',')]
    except:
        metrics = [Met.strip()]

    analytics = Analytics()
    analytics.main(email,startDate,endDate,dimentions,metrics)

    return render_template("index.html", title='Google Analytics')


def start():
    # time.sleep(2)
    webbrowser.open('http://127.0.0.1:5001/', new=2)




if __name__ == "__main__":
    start()
    app.run(port=5001)

