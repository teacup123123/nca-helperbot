# this one runs on a private pc, will tunnel to the random domain, e.g. ws://39fa-114-42-135-18.ngrok-free.app/
import asyncio
import websockets

async def echo(websocket, path):
    print('echo')
    async for message in websocket:
        print(message, 'received from client')
        greeting = f"Hello {message}!"
        await websocket.send(greeting)
        print(f"> {greeting}")


asyncio.get_event_loop().run_until_complete(
    websockets.serve(echo, 'localhost', 8080))
asyncio.get_event_loop().run_forever()
