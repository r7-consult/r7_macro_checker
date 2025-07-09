#!/usr/bin/env python3
"""
Working test for OnlyOffice Macro Tester
"""
import asyncio
import websockets
import json
import requests
import threading
import time

async def websocket_client():
    """WebSocket client that processes macros"""
    uri = "ws://localhost:8000/ws"
    
    print("🔗 Connecting to WebSocket server...")
    
    try:
        async with websockets.connect(uri) as websocket:
            print("✅ WebSocket connected successfully!")
            
            while True:
                try:
                    # Wait for macro from server
                    macro = await websocket.recv()
                    print(f"📨 Received macro: {macro[:100]}...")
                    
                    # Simulate macro execution
                    await asyncio.sleep(0.5)  # Simulate processing time
                    
                    # Create a realistic result
                    result = f"SUCCESS: Macro executed. Modified cells A1 and B1. Code: {macro[:50]}..."
                    
                    # Send result back
                    await websocket.send(result)
                    print(f"📤 Sent result: {result[:80]}...")
                    
                except websockets.exceptions.ConnectionClosed:
                    print("🔌 WebSocket connection closed")
                    break
                except Exception as e:
                    print(f"❌ WebSocket error: {e}")
                    break
                    
    except Exception as e:
        print(f"❌ Failed to connect: {e}")

def test_macro_execution():
    """Test macro execution via HTTP"""
    print("\n🧪 Testing macro execution...")
    
    # Wait for WebSocket to connect
    time.sleep(3)
    
    # Wait for WebSocket to be ready
    for i in range(10):
        try:
            response = requests.get("http://localhost:8000/status", timeout=2)
            status = response.json()
            if status.get('websocket_connected') and status.get('plugin_ready'):
                break
        except:
            pass
        time.sleep(0.5)
    else:
        print("⚠️ WebSocket connection timeout")
        return
    
    # Test status first
    response = requests.get("http://localhost:8000/status")
    status = response.json()
    print(f"📊 Server status: {status}")
    
    if not status.get('websocket_connected') or not status.get('plugin_ready'):
        print("⚠️ Server not ready for macro execution")
        return
    
    # Test macro execution
    macro_code = """
    (function() {
        let api = Api;
        let worksheet = api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("Hello OnlyOffice!");
        worksheet.GetRange("B1").SetValue("Macro test successful");
        return "Macro completed successfully";
    })();
    """
    
    print("🚀 Sending macro for execution...")
    
    try:
        response = requests.post(
            "http://localhost:8000/check",
            json={"macro": macro_code},
            timeout=15
        )
        
        if response.status_code == 200:
            result = response.json()
            print(f"✅ Macro execution result: {result}")
        else:
            print(f"❌ HTTP error {response.status_code}: {response.text}")
            
    except requests.exceptions.Timeout:
        print("⏱️ Macro execution timed out")
    except Exception as e:
        print(f"❌ Request failed: {e}")

async def run_test():
    """Run the complete test"""
    print("🚀 OnlyOffice Macro Tester - Working Test")
    print("=" * 50)
    
    # Start WebSocket client
    websocket_task = asyncio.create_task(websocket_client())
    
    # Start HTTP test in a separate thread
    http_thread = threading.Thread(target=test_macro_execution)
    http_thread.start()
    
    # Wait for HTTP test to complete
    http_thread.join()
    
    print("\n🛑 Stopping WebSocket client...")
    websocket_task.cancel()
    
    try:
        await websocket_task
    except asyncio.CancelledError:
        print("✅ WebSocket client stopped")
    
    print("\n✅ Test completed!")

if __name__ == "__main__":
    asyncio.run(run_test())