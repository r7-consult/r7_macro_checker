#!/usr/bin/env python3
"""
Simple WebSocket test to debug connection issues
"""
import asyncio
import websockets
import json

async def simple_websocket_test():
    uri = "ws://localhost:8000/ws"
    
    print(f"🔗 Attempting to connect to {uri}")
    
    try:
        async with websockets.connect(uri) as websocket:
            print("✅ WebSocket connected successfully!")
            
            # Send a test message
            await websocket.send("Hello from test client")
            print("📤 Sent test message")
            
            # Try to receive a message
            try:
                message = await asyncio.wait_for(websocket.recv(), timeout=5.0)
                print(f"📨 Received: {message}")
            except asyncio.TimeoutError:
                print("⏱️ No message received within timeout")
            
    except Exception as e:
        print(f"❌ Error: {e}")
        print(f"   Exception type: {type(e)}")

if __name__ == "__main__":
    asyncio.run(simple_websocket_test())