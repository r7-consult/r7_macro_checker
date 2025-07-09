from fastapi import FastAPI, WebSocket, HTTPException, WebSocketDisconnect
from pydantic import BaseModel
from asyncio import Queue
import asyncio
import json
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class Post(BaseModel):
    macro: str

class MacroResult(BaseModel):
    result: str
    success: bool = True
    error: str = None

# Global variables for queue management
macroQueue = None
resultQueue = None
websocket_connected = False

app = FastAPI(title="OnlyOffice Macro Tester", version="0.1")

@app.get("/")
async def root():
    return {"message": "OnlyOffice Macro Tester API", "status": "running"}

@app.get("/status")
async def status():
    return {
        "websocket_connected": websocket_connected,
        "plugin_ready": macroQueue is not None and resultQueue is not None
    }

@app.post("/check")
async def check(post: Post):
    global macroQueue, resultQueue, websocket_connected
    
    if not websocket_connected or macroQueue is None:
        raise HTTPException(status_code=503, detail='Plugin is not connected')
    
    try:
        # Add timeout to prevent hanging
        await asyncio.wait_for(macroQueue.put(post.macro), timeout=5.0)
        logger.info(f"Macro queued: {post.macro[:50]}...")
        
        # Wait for result with timeout
        result = await asyncio.wait_for(resultQueue.get(), timeout=30.0)
        logger.info(f"Result received: {result[:50]}...")
        
        return {"result": result, "success": True}
    
    except asyncio.TimeoutError:
        logger.error("Timeout waiting for macro execution")
        raise HTTPException(status_code=408, detail='Macro execution timeout')
    except Exception as e:
        logger.error(f"Error processing macro: {str(e)}")
        raise HTTPException(status_code=500, detail=f'Internal server error: {str(e)}')

@app.websocket("/ws")
async def websocket_endpoint(websocket: WebSocket):
    global macroQueue, resultQueue, websocket_connected
    
    await websocket.accept()
    websocket_connected = True
    
    macroQueue = Queue()
    resultQueue = Queue()
    
    logger.info("WebSocket connection established")
    
    try:
        while True:
            try:
                # Get macro from queue with timeout
                macro = await asyncio.wait_for(macroQueue.get(), timeout=1.0)
                logger.info(f"Sending macro to client: {macro[:50]}...")
                
                # Send macro to client
                await websocket.send_text(macro)
                
                # Wait for result from client
                result = await websocket.receive_text()
                logger.info(f"Received result from client: {result[:50]}...")
                
                # Put result in result queue
                await resultQueue.put(result)
                
            except asyncio.TimeoutError:
                # Keep connection alive, check for new macros periodically
                continue
            except WebSocketDisconnect:
                logger.info("WebSocket disconnected")
                break
            except Exception as e:
                logger.error(f"WebSocket error: {str(e)}")
                break
                
    except Exception as e:
        logger.error(f"WebSocket connection error: {str(e)}")
    finally:
        websocket_connected = False
        macroQueue = None
        resultQueue = None
        logger.info("WebSocket connection closed")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
