from flask import Flask

app = Flask(__name__)

@app.route("/")
def home():
    return "<h1>Bem-vindo à Loja de Livros 📚</h1><p>Nosso sistema está no ar!</p>"

if __name__ == "__main__":
    app.run(debug=True)
