from flask import Flask, render_template, request
from routes import cota

app = Flask(__name__)
app.register_blueprint(cota)

app.config['UPLOAD_FOLDER'] = './archivos de texto'
        
if __name__ == '__main__':
    app.run(port = 2000, debug = True)