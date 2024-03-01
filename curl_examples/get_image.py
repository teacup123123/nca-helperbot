import requests
from credentials import line_chan_acctoken
import requests

print(requests.get('https://api-data.line.me/v2/bot/message/497319186730320392/content',
                   headers={"Authorization": f"Bearer {line_chan_acctoken}"}))
