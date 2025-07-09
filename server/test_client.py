#!/usr/bin/env python3
"""
Test client for the OnlyOffice Macro Tester WebSocket API
"""
import asyncio
import websockets
import json
import time

async def test_websocket():
    uri = "ws://localhost:8000/ws"
    
    try:
        async with websockets.connect(uri) as websocket:
            print("✓ WebSocket connected successfully")
            
            # Wait for a macro to be sent by the server
            print("⏳ Waiting for macro from server...")
            
            # Listen for incoming messages with timeout
            try:
                message = await asyncio.wait_for(websocket.recv(), timeout=10.0)
                print(f"📨 Received macro: {message}")
                
                # Simulate macro execution and send result back
                result = f"Macro executed successfully at {time.strftime('%Y-%m-%d %H:%M:%S')}"
                await websocket.send(result)
                print(f"📤 Sent result: {result}")
                
                # Keep connection alive for a bit
                await asyncio.sleep(2)
                
            except asyncio.TimeoutError:
                print("⏱️ No macro received within timeout period")
                
    except websockets.exceptions.ConnectionRefused:
        print("❌ Could not connect to WebSocket server")
        print("   Make sure the server is running on localhost:8000")
    except websockets.exceptions.InvalidStatus as e:
        print(f"❌ WebSocket connection rejected: {e}")
        print("   Server might not have WebSocket support enabled")
    except ConnectionRefusedError:
        print("❌ Connection refused - server may not be running")
    except OSError as e:
        print(f"❌ Network error: {e}")
    except Exception as e:
        print(f"❌ WebSocket error: {e}")

async def test_http_endpoints():
    """Test HTTP endpoints using curl commands"""
    import subprocess
    
    endpoints = [
        ("GET", "/", "Root endpoint"),
        ("GET", "/status", "Status endpoint"),
        ("POST", "/check", "Check endpoint", '{"macro": "test macro"}')
    ]
    
    print("\n🔍 Testing HTTP endpoints:")
    
    for method, path, description, *data in endpoints:
        print(f"\n📡 {description} ({method} {path}):")
        
        if method == "GET":
            cmd = f"curl -s http://localhost:8000{path}"
        else:
            json_data = data[0] if data else "{}"
            cmd = f"curl -s -X {method} -H 'Content-Type: application/json' -d '{json_data}' http://localhost:8000{path}"
        
        try:
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
            if result.returncode == 0:
                print(f"✓ {result.stdout}")
            else:
                print(f"❌ Error: {result.stderr}")
        except Exception as e:
            print(f"❌ Exception: {e}")

async def test_full_workflow():
    """Test the complete workflow: WebSocket + HTTP"""
    print("\n🔄 Testing complete workflow:")
    
    # Start WebSocket connection in background
    websocket_task = asyncio.create_task(test_websocket())
    
    # Wait a bit for WebSocket to connect
    await asyncio.sleep(1)
    
    # Test HTTP endpoints
    await test_http_endpoints()
    
    # Send a macro via HTTP to test the full flow
    print("\n📨 Testing macro submission via HTTP:")
    import subprocess
    
    cmd = "curl -s -X POST -H 'Content-Type: application/json' -d '{\"macro\": \"console.log(\\\"Hello from macro!\\\");\"}' http://localhost:8000/check"
    result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
    
    if result.returncode == 0:
        print(f"✓ Macro result: {result.stdout}")
    else:
        print(f"❌ Error: {result.stderr}")
    
    # Wait for WebSocket task to complete
    try:
        await asyncio.wait_for(websocket_task, timeout=5.0)
    except asyncio.TimeoutError:
        websocket_task.cancel()

if __name__ == "__main__":
    print("🚀 OnlyOffice Macro Tester - Test Client")
    print("=" * 50)
    
    asyncio.run(test_full_workflow())