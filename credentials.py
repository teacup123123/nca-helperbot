import json
import os.path

_root, _ = os.path.split(__file__)
with open(os.path.join(_root, r'secrets/credentials.json'), 'r') as json_data:
    _creds = json.load(json_data)

line_chan_acctoken = _creds['line_chan_acctoken']
line_chan_secret = _creds['line_chan_secret']
ngrok_token = _creds['ngrok_token']
public_url = _creds['public_url']
private_port = _creds['private_port']
gai_apikey = _creds['gai_apikey']

if __name__ == '__main__':
    print(_creds)
