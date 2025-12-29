import sys
import os
import threading
import webbrowser
import time

# Prevent console window from appearing
if sys.platform == 'win32':
    import ctypes
    ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)

# Import Flask app
from server import app

# Try to import system tray (optional)
try:
    import pystray
    from PIL import Image, ImageDraw
    HAS_TRAY = True
except ImportError:
    HAS_TRAY = False
    print("pystray not installed - running without system tray icon")

def create_icon():
    """Create a simple icon for the system tray"""
    width = 64
    height = 64
    color1 = (52, 211, 153)  # Green
    color2 = (255, 255, 255)  # White
    
    image = Image.new('RGB', (width, height), color1)
    dc = ImageDraw.Draw(image)
    dc.rectangle([width // 4, height // 4, 3 * width // 4, 3 * height // 4], fill=color2)
    
    return image

def run_server():
    """Run the Flask server in a thread"""
    app.run(host='localhost', port=5000, debug=False, use_reloader=False)

def open_browser():
    """Open the web interface after a delay"""
    time.sleep(2)
    webbrowser.open('file://' + os.path.abspath('nameplate.html'))

def on_quit(icon, item):
    """Stop the server and exit"""
    icon.stop()
    os._exit(0)

def on_open(icon, item):
    """Open the web interface"""
    webbrowser.open('file://' + os.path.abspath('nameplate.html'))

if __name__ == '__main__':
    # Start Flask server in background thread
    server_thread = threading.Thread(target=run_server, daemon=True)
    server_thread.start()
    
    # Open browser
    browser_thread = threading.Thread(target=open_browser, daemon=True)
    browser_thread.start()
    
    if HAS_TRAY:
        # Create system tray icon
        icon = pystray.Icon(
            "nameplate_generator",
            create_icon(),
            "Nameplate Generator",
            menu=pystray.Menu(
                pystray.MenuItem("Open Interface", on_open),
                pystray.MenuItem("Quit", on_quit)
            )
        )
        
        # Run the icon (this blocks until quit)
        icon.run()
    else:
        # No tray icon - just keep server running
        print("Server running at http://localhost:5000")
        print("Press Ctrl+C to stop")
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            print("\nShutting down...")
            sys.exit(0)