"""
Launcher script that boots the server and automatically opens browser
"""
from app_server import start_server

if __name__ == "__main__":
    start_server(port=8505, open_browser=True)
