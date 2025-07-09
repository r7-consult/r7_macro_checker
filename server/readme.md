# OnlyOffice Macro Tester API

A FastAPI-based server for testing OnlyOffice macros with WebSocket communication.

## Features

- **WebSocket Communication**: Plugin connects via WebSocket to receive macros
- **HTTP API**: Submit macros for execution via REST API
- **Real-time Processing**: Asynchronous macro execution with timeouts
- **Status Monitoring**: Check connection status and plugin readiness
- **Error Handling**: Comprehensive error handling and logging

## API Endpoints

### HTTP Endpoints

- `GET /` - Root endpoint, returns API status
- `GET /status` - Check WebSocket connection and plugin readiness
- `POST /check` - Submit macro for execution
- `GET /docs` - FastAPI interactive documentation

### WebSocket Endpoint

- `WS /ws` - WebSocket connection for OnlyOffice plugin

## Quick Start

### 1. Setup and Install Dependencies

```bash
# Navigate to the project directory
cd  

# Create virtual environment (if not exists)
python3 -m venv venv

# Activate virtual environment
source venv/bin/activate

# Install dependencies
pip install fastapi uvicorn websockets requests
```

### 2. Start the Server

**Method 1: Using Python directly (Recommended)**
```bash
python server.py
```

**Method 2: Using uvicorn command**
```bash
uvicorn server:app --reload --host 0.0.0.0 --port 8000
```

**Method 3: One-liner**
```bash
  source venv/bin/activate && python server.py
```

**Method 4: Step by step (if Method 3 fails)**
```bash
cd  
source venv/bin/activate
python server.py
```

### 3. Verify Server is Running

```bash
# Test root endpoint
curl http://localhost:8000/
# Expected: {"message":"OnlyOffice Macro Tester API","status":"running"}

# Check status
curl http://localhost:8000/status
# Expected: {"websocket_connected":false,"plugin_ready":false}

# View interactive documentation
# Open browser to: http://localhost:8000/docs
```

## Usage

### 1. Connect Plugin (WebSocket)

The OnlyOffice plugin should connect to `ws://localhost:8000/ws` and:
- Listen for macro code
- Execute the macro
- Send results back

### 2. Submit Macros (HTTP)

```bash
curl -X POST "http://localhost:8000/check" \
  -H "Content-Type: application/json" \
  -d '{"macro": "Api.GetActiveSheet().GetRange(\"A1\").SetValue(\"Hello!\");"}'
```

## Demo

Run the included demo to see the complete workflow:

```bash
python demo.py
```

## Testing

Several test files are included:

### Run Complete Demo
```bash
python demo.py
```

### Test Files Available
- `demo.py` - Complete workflow demonstration
- `full_test.py` - Comprehensive testing
- `simple_ws_test.py` - Basic WebSocket connectivity test
- `test_client.py` - HTTP endpoint testing
- `working_test.py` - Production-ready workflow test

### Manual Testing
```bash
# Test WebSocket connection
python simple_ws_test.py

# Test full workflow
python working_test.py

# Test with demo client
python demo.py
```

## Fixed Issues

1. **Added proper error handling** - TimeoutError and exception handling
2. **Added logging** - Comprehensive logging for debugging
3. **Added connection tracking** - WebSocket connection state monitoring
4. **Added timeouts** - Prevents hanging (5s for queuing, 30s for execution)
5. **Added WebSocket disconnect handling** - Proper cleanup on disconnect
6. **Added status endpoints** - Monitor connection and readiness
7. **Improved WebSocket loop** - Better exception handling and connection management
8. **Added proper cleanup** - Reset global variables on disconnect

## Architecture

```
┌─────────────────┐    HTTP POST /check    ┌─────────────────┐
│   Client App    │ ───────────────────► │   FastAPI       │
│                 │                       │   Server        │
└─────────────────┘                       └─────────────────┘
                                                    │
                                                    │ WebSocket
                                                    │ /ws
                                                    ▼
                                          ┌─────────────────┐
                                          │  OnlyOffice     │
                                          │  Plugin         │
                                          └─────────────────┘
```

## Files Structure

```
macro_tester_v0.1/
├── server.py              # Main FastAPI server
├── README.md              # This documentation
├── demo.py                # Complete workflow demo
├── full_test.py           # Comprehensive testing
├── simple_ws_test.py      # Basic WebSocket test
├── test_client.py         # HTTP endpoint testing
├── working_test.py        # Production workflow test
└── venv/                  # Virtual environment
```

## Troubleshooting

### Server won't start
- Check if port 8000 is available: `netstat -tlnp | grep :8000`
- Kill existing server: `fuser -k 8000/tcp`
- Ensure virtual environment is activated: `source venv/bin/activate`
- Install dependencies: `pip install fastapi uvicorn websockets requests`

### Virtual environment activation fails
```bash
# Fix permissions issue
chmod +x venv/bin/activate

# Then activate
source venv/bin/activate
```

### Permission denied on venv/bin/activate
```bash
# Fix the permission issue
chmod +x venv/bin/activate

# Verify permissions
ls -la venv/bin/activate
# Should show: -rwxr-xr-x

# Then run server
source venv/bin/activate && python server.py
```

### WebSocket connection fails
- Verify server is running: `curl http://localhost:8000/status`
- Check firewall settings
- Test with simple client: `python simple_ws_test.py`

### Macro execution times out
- Check if WebSocket client is connected: `curl http://localhost:8000/status`
- Verify plugin is responding to messages
- Check server logs for error messages

 
