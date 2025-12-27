import subprocess
import time
import sys
import os
import uno

def ignite(odp_path):
    # 1. Expand the path to be absolute (so LibreOffice can find it)
    abs_path = os.path.abspath(odp_path)
    
    # 2. Kill any old instances to free up the port
    print("--- Cleaning up environment ---")
    subprocess.run(["killall", "-9", "soffice.bin"], capture_output=True)

    # 3. Start LibreOffice with the socket and the file
    print(f"--- Starting LibreOffice with {os.path.basename(abs_path)} ---")
    subprocess.Popen([
        "libreoffice", 
        "--accept=socket,host=127.0.0.1,port=2002;urp;", 
        abs_path
    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    # 4. Wait for the socket to be alive
    print("--- Waiting for port 2002... ---")
    ctx = None
    for _ in range(20): # Try for 10 seconds
        try:
            local_context = uno.getComponentContext()
            resolver = local_context.ServiceManager.createInstanceWithContext(
                "com.sun.star.bridge.UnoUrlResolver", local_context)
            ctx = resolver.resolve("uno:socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext")
            if ctx: break
        except:
            time.sleep(0.5)

    if not ctx:
        print("[ERROR] Could not connect to LibreOffice!")
        return

    # 5. Trigger Fullscreen (F5)
    print("--- Triggering Fullscreen ---")
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    
    # Wait for the document to actually load in the GUI
    doc = None
    while doc is None:
        doc = desktop.getCurrentComponent()
        time.sleep(0.5)
        
    doc.getPresentation().start()

    # 6. Launch the FastAPI server
    print("--- Launching server.py ---")
    # Using os.execvp replaces this script with the server process
    server_path = os.path.join(os.path.dirname(__file__), "server.py")
    os.execvp("python", ["python", server_path])

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python ignition.py <path_to_slides.odp>")
    else:
        ignite(sys.argv[1])
