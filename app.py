import csv
import os
import subprocess
import sys

from flask import Flask, jsonify, send_from_directory

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "capability1_news_output.csv")


@app.route("/")
def home():
    return send_from_directory(BASE_DIR, "dashboard.html")


@app.route("/sumridgelogo1.png")
def logo():
    return send_from_directory(BASE_DIR, "sumridgelogo1.png")


@app.route("/run-tool1", methods=["POST"])
def run_tool1():
    try:
        result = subprocess.run(
            [sys.executable, "tool1.py"],
            cwd=BASE_DIR,
            capture_output=True,
            text=True,
            timeout=180
        )

        if result.returncode != 0:
            return jsonify({
                "success": False,
                "error": result.stderr or result.stdout or "tool1.py failed."
            })

        if not os.path.exists(OUTPUT_FILE):
            return jsonify({
                "success": False,
                "error": "tool1.py ran, but capability1_news_output.csv was not created."
            })

        rows = []
        with open(OUTPUT_FILE, "r", encoding="utf-8") as file:
            reader = csv.DictReader(file)
            for row in reader:
                rows.append(row)

        return jsonify({
            "success": True,
            "rows": rows,
            "terminal_output": result.stdout
        })

    except subprocess.TimeoutExpired:
        return jsonify({
            "success": False,
            "error": "tool1.py timed out. Try reducing the number of bonds or headlines."
        })

    except Exception as error:
        return jsonify({
            "success": False,
            "error": str(error)
        })


if __name__ == "__main__":
    app.run(debug=True)