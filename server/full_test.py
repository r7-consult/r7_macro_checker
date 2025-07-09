#!/usr/bin/env python3
"""
Full test of the OnlyOffice Macro Tester workflow
"""
import asyncio
import websockets
import json
import threading
import time
import requests

class MacroTesterClient:
    def __init__(self):
        self.websocket = None
        self.running = False
        
    async def websocket_handler(self):
        """Handle WebSocket connection and macro execution"""
        uri = "ws://localhost:8000/ws"
        
        try:
            async with websockets.connect(uri) as websocket:
                self.websocket = websocket
                self.running = True
                print("✅ WebSocket client connected")
                
                while self.running:
                    try:
                        # Wait for macro from server
                        macro = await asyncio.wait_for(websocket.recv(), timeout=1.0)
                        print(f"📨 Received macro: {macro}")
                        
                        # Simulate macro execution
                        await asyncio.sleep(0.1)  # Simulate processing time
                        
                        # Send result back
                        result = f"Macro executed successfully: {macro[:50]}..."
                        await websocket.send(result)
                        print(f"📤 Sent result: {result}")
                        
                    except asyncio.TimeoutError:
                        # Continue listening for macros
                        continue
                    except websockets.exceptions.ConnectionClosed:
                        print("🔌 WebSocket connection closed")
                        break
                        
        except Exception as e:
            print(f"❌ WebSocket error: {e}")
        finally:
            self.running = False
            print("🔚 WebSocket client stopped")

def test_http_endpoints():
    """Test HTTP endpoints"""
    base_url = "http://localhost:8000"
    
    print("\n🔍 Testing HTTP endpoints:")
    
    # Test root endpoint
    response = requests.get(f"{base_url}/")
    print(f"📡 GET /: {response.json()}")
    
    # Test status endpoint
    response = requests.get(f"{base_url}/status")
    print(f"📡 GET /status: {response.json()}")
    
    # Test check endpoint with WebSocket connected
    macro_code = """
    let worksheet = Api.GetActiveSheet();
    worksheet.GetRange("A1").SetValue("Hello from OnlyOffice Macro!");
    worksheet.GetRange("B1").SetValue("Test successful");
    """
    
    try:
        response = requests.post(
            f"{base_url}/check",
            json={"macro": macro_code},
            timeout=10
        )
        print(f"📡 POST /check: {response.json()}")
    except requests.exceptions.Timeout:
        print("⏱️ POST /check: Request timed out")
    except Exception as e:
        print(f"❌ POST /check error: {e}")

async def run_full_test():
    """Run the complete test"""
    print("🚀 OnlyOffice Macro Tester - Full Test")
    print("=" * 50)
    
    # Create client
    client = MacroTesterClient()
    
    # Start WebSocket client in background
    websocket_task = asyncio.create_task(client.websocket_handler())
    
    # Wait for WebSocket to connect
    await asyncio.sleep(2)
    
    # Test HTTP endpoints in a separate thread
    def http_test():
        time.sleep(1)  # Wait for WebSocket to be fully established
        test_http_endpoints()
    
    http_thread = threading.Thread(target=http_test)
    http_thread.start()
    
    # Wait for HTTP tests to complete
    http_thread.join()
    
    # Stop WebSocket client
    client.running = False
    
    # Wait for WebSocket task to complete
    try:
        await asyncio.wait_for(websocket_task, timeout=3.0)
    except asyncio.TimeoutError:
        websocket_task.cancel()
    
    print("\n✅ Full test completed!")

if __name__ == "__main__":
    asyncio.run(run_full_test())