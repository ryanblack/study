#!/bin/bash
sudo pip install flask
cat <<EOF >> app.py
from flask import Flask, render_template, request
from flask import jsonify
app = Flask(__name__)

@app.route('/')
def hello_world():
  return 'Hello from Flask!'

@app.route('/builds', methods=['POST'])
def builds():
  data = request.get_json()
  for a in data:
    for b in data[a]:
      for c in data[a][b]:
        listdata = data[a][b][c]
  build_date = 0

  for i in listdata:
    if all (k in i for k in ('build_date','output')):
      if int(i['build_date']) > build_date:
        build_date = int(i['build_date'])
        output = i['output']
  return jsonify({'latest':{"build_date":build_date,"ami_id":output.split()[2],"commit_hash":output.split()[3]}})

if __name__ == '__main__':
  app.run(host="0.0.0.0", port=80)
EOF
sudo python app.py