import re

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


def handle_msg(event):
    uid = event.source.user_id
    if uid not in catbot.userbase:
        cbot = catbot.userbase[uid] = catbot.catbot()
    else:
        cbot = catbot.userbase[uid]

    _, reply = cbot.handle(event.message.text if event.message.type == 'text' else 'IMG', msgid=event.message.id)
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
        # line_bot_api.reply_message_with_http_info(
        #     ReplyMessageRequest(
        #         reply_token=event.reply_token,
        #         messages=[TextMessage(text=reply, quoteToken=event.message.quote_token)]
        #     )
        # )
        line_bot_api.push_message_with_http_info(
            PushMessageRequest(to=event.source.user_id,
                               messages=[
                                   TextMessage(
                                       text=reply,
                                       quoteToken=event.message.quote_token,
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


if __name__ == '__main__':
    from waitress import serve
    print(f'serving at port {cat_tunnel.localport}')
    serve(app, host="0.0.0.0", port=cat_tunnel.localport)
    # app.run(host='0.0.0.0', port=private_port)

# curl -v -X GET https://api-data.line.me/v2/bot/message/{messageId}/content \
# -H 'Authorization: Bearer {channel access token}'
