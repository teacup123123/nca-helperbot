import asyncio
import websockets

async def hello(uri):
    async with websockets.connect(uri) as websocket:
        await websocket.send("Jimmy")
        print(f"(client) send to server: Jimmy")
        name = await websocket.recv()
        print(f"(client) recv from server {name}")

asyncio.get_event_loop().run_until_complete(
    hello('wss://39fa-114-42-135-18.ngrok-free.app/')
)