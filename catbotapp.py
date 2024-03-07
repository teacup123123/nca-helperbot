import re
import time

from credentials import *
from ngrok_autolaunch import *
import json
from flask import Flask, request, abort

from linebot.v3 import (
    WebhookHandler
)
from linebot.v3.exceptions import (
    InvalidSignatureError
)
from linebot.v3.messaging import (
    Configuration,
    ApiClient,
    MessagingApi,
    ReplyMessageRequest,
    PushMessageRequest,
    TextMessage,
    ImageMessage,
    Emoji
)
from linebot.v3.webhooks import (
    MessageEvent,
    TextMessageContent,
    ImageMessageContent
)

app = Flask(__name__)
app.config["public_url"] = cat_tunnel.public_url
my_line_config = Configuration(access_token=line_chan_acctoken)
my_line_handler = WebhookHandler(line_chan_secret)


@app.route("/", methods=['POST'])
def callback():
    # get X-Line-Signature header value
    signature = request.headers['X-Line-Signature']

    # get request body as text
    body = request.get_data(as_text=True)
    bodyjson = json.loads(body)
    # print(request.headers)
    # print(bodyjson)
    app.logger.info(f"Request body: {json.dumps(bodyjson, indent=4)}")

    # handle webhook body
    try:
        my_line_handler.handle(body, signature)
    except InvalidSignatureError:
        app.logger.info("Invalid signature. Please check your channel access token/channel secret.")
        abort(400)

    return 'OK'


import catbot

palette_Emojis = {
    '(moon sunglasses)': Emoji(product_id='5ac1bfd5040ab15980c9b435', emoji_id='021'),
    '(cony cry)': Emoji(product_id='5ac1bfd5040ab15980c9b435', emoji_id='046'),
    '(moon halo)': Emoji(product_id='5ac1bfd5040ab15980c9b435', emoji_id='025'),
    '(heart)': Emoji(product_id='5ac1bfd5040ab15980c9b435', emoji_id='215'),
}


def handle_msg(event, **kwargs):
    if isinstance(event, str):
        assert re.match('\d{4}', event[:4])
        get = [uid for uid, cb in catbot.userbase.items() if cb.id == event[:4]]
        assert len(get) == 1
        uid = get[0]
        event = event[4:]
    else:
        uid = event.source.user_id
    if uid not in catbot.userbase:
        cbot = catbot.userbase[uid] = catbot.catbot()
    else:
        cbot = catbot.userbase[uid]

    if isinstance(event, str):
        if 'botmsg' in kwargs:
            _, reply = cbot.handle(kwargs['botmsg'], **kwargs)
        else:
            reply = event
    else:
        _, reply = cbot.handle(event.message.text if event.message.type == 'text' else 'IMG',
                               msgid=event.message.id, **kwargs)
    catbot.userbase[uid] = cbot  # force save

    reply: str
    used_emojis = []
    for a in re.compile('\([a-z\ ]+\)').finditer(reply):
        used_emojis.append(palette_Emojis[a.group(0)].copy())
    reply = re.sub('\([a-z\ ]+\)', '$', reply)

    for a, e in zip(re.compile('\$').finditer(reply), used_emojis):
        e: Emoji
        e.index = a.start()
        # e.length = a.end() - a.start()

    with ApiClient(my_line_config) as api_client:
        line_bot_api = MessagingApi(api_client)
        line_bot_api.push_message_with_http_info(
            PushMessageRequest(to=uid,
                               messages=[
                                   TextMessage(
                                       text=reply,
                                       quoteToken=None if isinstance(event, str) else event.message.quote_token,
                                       emojis=used_emojis if used_emojis else None
                                   )
                               ])
        )


@my_line_handler.add(MessageEvent, message=ImageMessageContent)
def handle_img_message(event):
    return handle_msg(event)


@my_line_handler.add(MessageEvent, message=TextMessageContent)
def handle_txt_message(event):
    return handle_msg(event)


from gcredentials import client
import time


# from datetime import datetime

def subscribe():
    # try:
    #     with open('secrets/resourceID.txt', 'r') as f:
    #         resourceId = f.read(-1).strip()
    #     response_clr = client.http_client.request(
    #         method='POST',
    #         endpoint='https://www.googleapis.com/drive/v3/channels/stop',
    #         json={
    #             "id": "holiday_form_watch",
    #             "resourceId": resourceId
    #         }
    #     )
    # except:
    #     pass

    try:
        response = client.http_client.request(
            method='POST',
            endpoint='https://www.googleapis.com/drive/v3/files/1JwFcXcyNIch2vfmATGI55nTSFDYJ7EUC5l764svBb_4/watch',
            json={
                "id": "holiday_form_watch4",
                "type": "web_hook",
                "address": cat_tunnel.public_url + '/filewatch',
                # 'expiration': int(time.time() * 1000) + 60 * 60 * 12
            }
        )
        print(response)
        print(response.content)
        resourceId = json.loads(response.content)['resourceId']
        with open('secrets/resourceID.txt', 'w') as f:
            f.write(resourceId)
    except:
        pass


holiday_old_timestamps = set()
spreadsheet = client.open('holiday3')
worksheet = spreadsheet.get_worksheet(0)


@app.route("/filewatch", methods=['GET', 'POST'])
def holiday_form_watcher():
    global holiday_old_timestamps
    if 'X-Goog-Resource-State' not in request.headers:
        g_resource_state = 'manual'
    else:
        g_resource_state = request.headers['X-Goog-Resource-State']
    if g_resource_state == 'sync':
        print("Syncing")
        new_timestamps = set((x[0], i + 1) for i, x in enumerate(worksheet.get('A:A')))
        assert ('Timestamp', 1) in new_timestamps
        diff, holiday_old_timestamps = new_timestamps.difference(holiday_old_timestamps), new_timestamps
        return 'OK'

    new_timestamps = set((x[0], i + 1) for i, x in enumerate(worksheet.get('A:A')))
    assert ('Timestamp', 1) in new_timestamps
    diff, holiday_old_timestamps = new_timestamps.difference(holiday_old_timestamps), new_timestamps
    if ('Timestamp', 1) in diff: return 'OK'
    print("UPDATE!", diff)
    if len(diff) == 0: return 'OK'
    ids = [id[0][0] for id in worksheet.batch_get([f'C{x}' for _, x in diff])]
    for id, (_ts, _) in zip(ids, diff):
        msg = f'{_ts}填寫/修改的表單看見了!'
        fabricated = f'{id}{msg}'
        print(fabricated)
        handle_msg(fabricated)

    return 'OK'


if __name__ == '__main__':
    subscription = subscribe()
    from waitress import serve

    print(f'serving at port {cat_tunnel.localport}')
    serve(app, host="0.0.0.0", port=cat_tunnel.localport)
    # app.run(host='0.0.0.0', port=private_port)
