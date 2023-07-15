import json
def exp_json(t,wecomid_replace='no',wecomid_pair=['$wecomid$','1688850049985213']):
    with open (t,'r',encoding='utf-8') as f:
        lines=f.readlines()
    _line=''
    for line in lines:
        newLine=line.strip('\n')
        _line=_line+newLine
    if wecomid_replace=='yes':
        _line=_line.replace(wecomid_pair[0],wecomid_pair[1])
    config=json.loads(_line) 

    return config

def exp_json2(t):
    with open (t,'r',encoding='utf-8') as f:
        lines=f.readlines()
    _line=''
    for line in lines:
        newLine=line.strip('\n')
        _line=_line+newLine
    config=json.loads(_line)
    return config

def readColorConfig(fn):
    with open(fn,'r',encoding='utf-8') as f:
        lines=f.readlines()
        _line=''
        for line in lines:
            newLine=line.strip('\n')
            _line=_line+newLine
        config=_line
    return config