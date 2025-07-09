import requests


macro = 'var x = 1 + 2'
print('Checking macro: ' + macro)

r = requests.post('http://127.0.0.1:8000/check', json={ 'macro': macro })
print(r.json())
print()


macro = 'Api.GetActiveSheet().GetRange("A1").SetValue("Test");'
print('Checking macro: ' + macro)

r = requests.post('http://127.0.0.1:8000/check', json={ 'macro': macro })
print(r.json())
print()


macro = 'error ~!@'
print('Checking macro: ' + macro)

r = requests.post('http://127.0.0.1:8000/check', json={ 'macro': macro })
print(r.json())
print()
