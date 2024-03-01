import pathlib
import textwrap
from credentials import *
import PIL.Image

img = PIL.Image.open('../imgsamples/155906.jpg')

import google.generativeai as genai

genai.configure(api_key=gai_apikey)

for m in genai.list_models():
    if 'generateContent' in m.supported_generation_methods:
        print(m.name)

model = genai.GenerativeModel('gemini-pro-vision')

# prompt = 'here is an application form for temporary leave, please generate from the photo a json formatted content'
# prompt = '這裡是請假申請表，申請人會在假別打勾，請問哪個被打勾?'
prompt = '這裡是請假申請表，申請人字比較潦草，但知道是張迪凱、顏孝丞、楊培原、周子瑜之一。這四個名字中，哪個比較符合瞭草的申請人欄位'
response = model.generate_content([prompt, img])
response.resolve()
print(response.text)
