from credentials import *
import json

from flask import Flask

app = Flask(__name__)

# Open a ngrok tunnel to the HTTP server
# Update any base URLs to use the public ngrok URL
app.config["public_url"] = public_url

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
    TextMessage
)
from linebot.v3.webhooks import (
    MessageEvent,
    TextMessageContent
)

app = Flask(__name__)

my_line_config = Configuration(access_token=line_chan_acctoken)
my_line_handler = WebhookHandler(line_chan_secret)


@app.route("/", methods=['POST'])
def callback():
    # get X-Line-Signature header value
    signature = request.headers['X-Line-Signature']

    # get request body as text
    body = request.get_data(as_text=True)
    bodyjson = json.loads(body)

    print(request.headers)
    print(bodyjson)
    app.logger.info(f"Request body: {json.dumps(bodyjson, indent=4)}")

    # handle webhook body
    try:
        my_line_handler.handle(body, signature)
    except InvalidSignatureError:
        app.logger.info("Invalid signature. Please check your channel access token/channel secret.")
        abort(400)

    return 'OK'


@my_line_handler.add(MessageEvent, message=TextMessageContent)
def handle_message(event):
    print('event triggered')
    with ApiClient(my_line_config) as api_client:
        line_bot_api = MessagingApi(api_client)
        line_bot_api.reply_message_with_http_info(
            ReplyMessageRequest(
                reply_token=event.reply_token,
                messages=[TextMessage(text=event.message.text)]
            )
        )


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=private_port)

# curl -v -X GET https://api-data.line.me/v2/bot/message/{messageId}/content \
# -H 'Authorization: Bearer {channel access token}'
