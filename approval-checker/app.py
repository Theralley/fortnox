from flask import Flask, jsonify

app = Flask(__name__)
link_status = {"clicked": False}

@app.route('/unique-link')
def unique_link():
    link_status["clicked"] = True
    return "Thank you for clicking the link."

@app.route('/check-status')
def check_status():
    return jsonify(link_status)

if __name__ == '__main__':
    app.run(debug=True)
