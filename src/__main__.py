from flask import Flask

# TODO: avoid using local dev server when deploying to production: https://flask.palletsprojects.com/en/stable/deploying/

app = Flask(__name__)

@app.route("/")
def hello_world():
    return "<p>Hello, World!</p>"

# Run the app only if this file is executed directly
if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1", port=5000)