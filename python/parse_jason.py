#!/usr/bin/python3
import sys
import json
from pprint import pprint
file_path = sys.argv[1]

with open(file_path) as json_file:
    data = json.loads(json_file.read())
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
latest = {"latest":{"build_date":build_date,"ami_id":output.split()[2],"commit_hash":output.split()[3]}}
pprint(latest)