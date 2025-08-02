from flask import Flask, request, jsonify

app = Flask(__name__)

@app.route("/greet", methods=["POST"])
def greet():
    data = request.json
    name = data.get("name", "Anonymous")
    return jsonify({"greeting": f"Hello, {name}!"})
