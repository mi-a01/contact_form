from flask import Flask, request, jsonify

app = Flask(__name__)

@app.route('/run-script', methods=['GET'])
def run_script():
    print("ああああてすとあああああ")

    result = {"message": "Pythonスクリプトが実行されました"}
    return jsonify(result)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8000)
