# server.py - To be run on the presenter computer
import uvicorn
from fastapi import FastAPI, HTTPException
import uno

# --- UNO Connection Setup ---
# --- UNO Connection Setup (NEW DEBUG VERSION) ---
def get_slideshow_controller():
    """Connects to LibreOffice and returns the slideshow controller, with debug prints."""
    print("\n--- Attempting to connect to LibreOffice ---")
    try:
        local_context = uno.getComponentContext()
        resolver = local_context.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_context)
        print("[SUCCESS] Got UNO component context and resolver.")

        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        smgr = ctx.ServiceManager
        print("[SUCCESS] Resolved socket and got service manager.")

        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        if not desktop:
            print("[FAILURE] Could not create Desktop object!")
            return None, None
        print(f"[SUCCESS] Got Desktop object: {desktop}")

        document = desktop.getCurrentComponent()
        if not document:
            print("[FAILURE] Desktop found, but could not get current document!")
            return None, None
        print(f"[SUCCESS] Got Document object: {document.Title}")

        if not document.supportsService("com.sun.star.presentation.PresentationDocument"):
            print(f"[FAILURE] The current document '{document.Title}' is not an Impress presentation.")
            return None, None
        print("[SUCCESS] Document is a valid Impress presentation.")
        
        # This is the critical part
        presentation_object = document.getPresentation()
        if not presentation_object:
            print("[FAILURE] Could not get the Presentation object from the document.")
            return None, None
        print("[SUCCESS] Got Presentation object.")

        show_controller = presentation_object.getController()
        if not show_controller:
            # THIS IS LIKELY WHERE IT IS FAILING
            print("[FAILURE] Got Presentation, but getController() returned None. Is slideshow *really* running in full-screen (F5)?")
            return None, None

        print("[SUCCESS] Successfully got the running slideshow controller!")
        return show_controller, document

    except Exception as e:
        print(f"[FATAL FAILURE] An exception occurred during connection: {e}")
        return None, None
# --- FastAPI Application ---
app = FastAPI()

@app.get("/control/next")
async def next_slide():
    """Advances to the next slide or animation."""
    controller, _ = get_slideshow_controller()
    if not controller:
        raise HTTPException(status_code=404, detail="Slideshow not running or not found.")
    controller.gotoNextEffect()
    return {"status": "advanced to next"}

@app.get("/control/previous")
async def previous_slide():
    """Goes to the previous slide or animation."""
    controller, _ = get_slideshow_controller()
    if not controller:
        raise HTTPException(status_code=404, detail="Slideshow not running or not found.")
    controller.gotoPreviousEffect()
    return {"status": "returned to previous"}

@app.get("/state")
async def get_slide_state():
    """Returns the current state of the slideshow, specifically the slide index."""
    controller, _ = get_slideshow_controller()
    if not controller:
        raise HTTPException(status_code=404, detail="Slideshow not running or not found.")

    # Get the current slide index (slide 1 is index 0)
    current_index = controller.getCurrentSlideIndex()

    # Return ONLY the slide number
    return {"slide_index": current_index}

# --- To run the server ---
if __name__ == "__main__":
    # Find your computer's IP address (e.g., 192.168.1.10)
    # and run the server on that host so other devices can see it.
    # Using "0.0.0.0" makes it accessible on your local network.
    print("Server starting...")
    print("Make sure LibreOffice is running with a listening port and a slideshow is active.")
    uvicorn.run(app, host="0.0.0.0", port=8000)
