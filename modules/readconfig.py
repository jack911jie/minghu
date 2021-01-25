import json
def exp_json(t):
    with open (t,'r',encoding='utf-8') as f:
        lines=f.readlines()
    _line=''
    for line in lines:
        newLine=line.strip('\n')
        _line=_line+newLine
    config=json.loads(_line) 

    return config