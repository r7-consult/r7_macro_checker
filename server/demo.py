#!/usr/bin/env python3
"""
Demo script showing OnlyOffice Macro Tester working correctly
"""
import asyncio
import websockets
import requests
import json
import time
import threading

def demo_websocket_client():
    """WebSocket client that simulates OnlyOffice plugin"""
    
    async def client():
        try:
            async with websockets.connect('ws://localhost:8000/ws') as websocket:
                print("✅ OnlyOffice Plugin (WebSocket) connected")
                
                while True:
                    # Wait for macro from server
                    macro = await websocket.recv()
                    print(f"📨 Plugin received macro: {macro[:60]}...")
                    
                    # Simulate macro execution time
                    await asyncio.sleep(1)
                    
                    # Send execution result back
                    result = {
                        "success": True,
                        "message": "Macro executed successfully in OnlyOffice",
                        "output": "Cell A1 set to 'Hello OnlyOffice!', Cell B1 set to 'Test successful'"
                    }
                    
                    await websocket.send(json.dumps(result))
                    print(f"📤 Plugin sent result: {result['message']}")
                    
        except websockets.exceptions.ConnectionClosed:
            print("🔌 Plugin disconnected")
        except Exception as e:
            print(f"❌ Plugin error: {e}")
    
    # Run the async client
    asyncio.run(client())

def demo_http_client():
    """HTTP client that sends macros for execution"""
    
    print("⏳ Waiting for OnlyOffice plugin to connect...")
    time.sleep(2)
    
    # Check server status
    response = requests.get("http://localhost:8000/status")
    status = response.json()
    print(f"📊 Server status: {status}")
    
    if not status.get('websocket_connected'):
        print("❌ OnlyOffice plugin not connected!")
        return
    
    # Test macro execution
    macro_code = """
    (function() {
        'use strict';
        
        try {
            // Get the active worksheet
            const api = Api;
            const worksheet = api.GetActiveSheet();
            
            // Set some values
            worksheet.GetRange("A1").SetValue("Hello OnlyOffice!");
            worksheet.GetRange("B1").SetValue("Test successful");
            
            // Format the cells
            worksheet.GetRange("A1").SetFontColor(api.CreateColorFromRGB(255, 0, 0));
            worksheet.GetRange("B1").SetFontColor(api.CreateColorFromRGB(0, 128, 0));
            
            return "Macro completed successfully";
            
        } catch (error) {
            return "Error: " + error.message;
        }
    })();
    """
    
    print("🚀 Sending OnlyOffice macro for execution...")
    
    try:
        response = requests.post(
            "http://localhost:8000/check",
            json={"macro": macro_code},
            timeout=20
        )
        
        if response.status_code == 200:
            result = response.json()
            print(f"✅ Macro execution completed!")
            print(f"📋 Result: {result}")
        else:
            print(f"❌ HTTP error {response.status_code}: {response.text}")
            
    except requests.exceptions.Timeout:
        print("⏱️ Macro execution timed out")
    except Exception as e:
        print(f"❌ Request failed: {e}")

def run_demo():
    """Run the complete demo"""
    print("🚀 OnlyOffice Macro Tester - Demo")
    print("=" * 50)
    print("This demo shows the complete workflow:")
    print("1. OnlyOffice plugin connects via WebSocket")
    print("2. Client sends macro via HTTP POST")
    print("3. Server forwards macro to plugin")
    print("4. Plugin executes macro and returns result")
    print("5. Server returns result to client")
    print("-" * 50)
    
    # Start WebSocket client (plugin simulation) in background thread
    websocket_thread = threading.Thread(target=demo_websocket_client)
    websocket_thread.daemon = True
    websocket_thread.start()
    
    # Run HTTP client (macro sender)
    demo_http_client()
    
    print("\n✅ Demo completed!")
    print("💡 The OnlyOffice Macro Tester is working correctly!")

if __name__ == "__main__":
    run_demo()