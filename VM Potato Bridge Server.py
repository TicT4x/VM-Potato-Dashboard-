import asyncio
import websockets
import json
import voicemeeterlib
import math
import http.server
import socketserver
import threading
import os
import base64
import datetime
import time

# ==========================================
# KONFIGURATION
# ==========================================
INPUTS_TO_MONITOR = [0, 1, 2, 3, 4, 5, 6] 
OUTPUTS_TO_MONITOR = [0] 

WS_PORT = 8765   
HTTP_PORT = 8080 
# ==========================================

try:
    from winrt.windows.media.control import GlobalSystemMediaTransportControlsSessionManager
    MEDIA_SUPPORT = True
except ImportError:
    MEDIA_SUPPORT = False
    print("\nWARNUNG: 'winrt' Pakete nicht vollständig installiert! Media-Tab wird leer bleiben.")

current_media = {
    "title": "",
    "artist": "",
    "playing": False,
    "art": "",
    "progress": 0,
    "duration": 0
}

# ==========================================
# ISOLIERTER THREAD: Media-Abfrage ohne Freeze!
# ==========================================
class MediaPollerThread(threading.Thread):
    def __init__(self):
        super().__init__(daemon=True)
        
    def run(self):
        if not MEDIA_SUPPORT:
            return
            
        from winrt.windows.media.control import GlobalSystemMediaTransportControlsSessionManager as MediaManager
        from winrt.windows.storage.streams import DataReader
        
        global current_media
        last_title = ""
        
        async def fetch_once(cached_title):
            manager = await MediaManager.request_async()
            session = manager.get_current_session()
            if not session: return None
            
            info = await session.try_get_media_properties_async()
            timeline = session.get_timeline_properties()
            playback = session.get_playback_info()
            
            title = info.title if info.title else ""
            artist = info.artist if info.artist else ""
            
            prog, dur = 0, 0
            if timeline:
                def ts_to_sec(ts):
                    try:
                        if isinstance(ts, datetime.timedelta): return ts.total_seconds()
                        if hasattr(ts, 'duration'): return ts.duration / 10000000
                        return float(ts) / 10000000
                    except: return 0
                prog = ts_to_sec(timeline.position)
                dur = ts_to_sec(timeline.end_time)
                
            playing = (int(playback.playback_status) == 4) if playback else False
            
            art_b64 = None 
            if title != cached_title:
                if info.thumbnail:
                    try:
                        stream = await info.thumbnail.open_read_async()
                        if stream:
                            reader = DataReader(stream)
                            await reader.load_async(stream.size)
                            try:
                                buf = bytearray(stream.size)
                                reader.read_bytes(buf)
                                art_b64 = base64.b64encode(buf).decode('utf-8')
                            except TypeError:
                                buffer = reader.read_buffer(stream.size)
                                art_b64 = base64.b64encode(bytes(buffer)).decode('utf-8')
                    except Exception:
                        art_b64 = None 
                else:
                    art_b64 = "" 
                    
            return {
                "title": title,
                "artist": artist,
                "playing": playing,
                "progress": prog,
                "duration": dur,
                "art": art_b64
            }

        while True:
            try:
                loop = asyncio.new_event_loop()
                asyncio.set_event_loop(loop)
                data = loop.run_until_complete(fetch_once(last_title))
                loop.close() 
                
                if data:
                    current_media["progress"] = data["progress"]
                    current_media["duration"] = data["duration"]
                    current_media["playing"] = data["playing"]
                    
                    if data["title"] != last_title:
                        current_media["title"] = data["title"]
                        current_media["artist"] = data["artist"]
                        
                        if data["art"] is not None:
                            current_media["art"] = data["art"]
                            last_title = data["title"] 
                        else:
                            pass
                else:
                    current_media["title"] = ""
                    current_media["artist"] = ""
                    current_media["playing"] = False
                    current_media["art"] = ""
                    last_title = ""
            except Exception:
                pass 
                
            time.sleep(0.5)

# ==========================================
# WS TASK: Pegel senden
# ==========================================
def level_to_percent(level_val):
    if level_val is None or level_val == 0: return 0
    if level_val < 0:
        if level_val < -60: return 0
        db = level_val
    else:
        if level_val <= 0.0001: return 0
        db = 20 * math.log10(level_val)
    percent = ((db + 60) / 72) * 100
    if percent < 0: return 0
    if percent > 100: return 100
    return percent

async def send_loop(websocket, vm):
    global current_media
    last_sent_title = None
    last_sent_art = None
    
    # Caches, damit wir Namen nur senden, wenn du sie in Voicemeeter umbenannt hast (spart Bandbreite)
    last_labels = []
    last_master_label = ""
    
    default_names = ["HW 1", "HW 2", "HW 3", "HW 4", "HW 5", "VAIO", "AUX", "VAIO3"]
    
    try:
        while True:
            try:
                _ = vm.pdirty
            except AttributeError:
                pass 
                
            levels_percent = []
            mutes_state = []
            current_labels = []
            
            for strip_index in INPUTS_TO_MONITOR:
                try:
                    # Pegel & Mute
                    raw_levels = getattr(vm.strip[strip_index].levels, 'postfader', None)
                    if raw_levels is None:
                        raw_levels = getattr(vm.strip[strip_index].levels, 'prefader', [0, 0])
                    max_level = max(raw_levels) if raw_levels else 0
                    levels_percent.append(level_to_percent(max_level))
                    mutes_state.append(bool(vm.strip[strip_index].mute))
                    
                    # Labels dynamisch abfragen
                    lbl = getattr(vm.strip[strip_index], 'label', '')
                    if not lbl:
                        lbl = default_names[strip_index] if strip_index < len(default_names) else f"In {strip_index+1}"
                    current_labels.append(lbl)
                    
                except Exception:
                    levels_percent.append(0)
                    mutes_state.append(False)
                    current_labels.append("...")

            current_master_label = "A1"
            for bus_index in OUTPUTS_TO_MONITOR:
                try:
                    raw_levels = getattr(vm.bus[bus_index].levels, 'all', [0, 0])
                    max_level = max(raw_levels) if raw_levels else 0
                    levels_percent.append(level_to_percent(max_level))
                    mutes_state.append(bool(vm.bus[bus_index].mute))
                    
                    lbl = getattr(vm.bus[bus_index], 'label', '')
                    current_master_label = lbl if lbl else f"A{bus_index+1}"
                except Exception:
                    levels_percent.append(0)
                    mutes_state.append(False)
            
            payload = {
                "levels": levels_percent,
                "mutes": mutes_state,
            }
            
            # Text-Labels nur ins Datenpaket packen, wenn sie sich geändert haben (oder ganz am Anfang)
            if current_labels != last_labels:
                payload["labels"] = current_labels
                last_labels = current_labels
                
            if current_master_label != last_master_label:
                payload["master_label"] = current_master_label
                last_master_label = current_master_label
            
            media_to_send = current_media.copy()
            if media_to_send["title"] != last_sent_title or media_to_send["art"] != last_sent_art:
                last_sent_title = media_to_send["title"]
                last_sent_art = media_to_send["art"]
            else:
                media_to_send["art"] = None 
                
            payload["media"] = media_to_send

            await websocket.send(json.dumps(payload))
            await asyncio.sleep(0.016)
            
    except websockets.exceptions.ConnectionClosed:
        pass 

# ==========================================
# WS TASK: Kommandos empfangen (Play/Pause)
# ==========================================
def execute_media_command(action):
    if not MEDIA_SUPPORT: return
    try:
        from winrt.windows.media.control import GlobalSystemMediaTransportControlsSessionManager as MediaManager
        
        async def _cmd():
            manager = await MediaManager.request_async()
            session = manager.get_current_session()
            if session:
                if action == "playpause": await session.try_toggle_play_pause_async()
                elif action == "next": await session.try_skip_next_async()
                elif action == "prev": await session.try_skip_previous_async()
                
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        loop.run_until_complete(_cmd())
        loop.close()
    except Exception:
        pass

async def receive_loop(websocket):
    try:
        if not MEDIA_SUPPORT:
            async for _ in websocket: pass
            return
            
        async for message in websocket:
            try:
                data = json.loads(message)
                action = data.get("action")
                if action in ["playpause", "next", "prev"]:
                    threading.Thread(target=execute_media_command, args=(action,), daemon=True).start()
            except Exception:
                pass
    except websockets.exceptions.ConnectionClosed:
        pass 

# ==========================================
# WS HANDLER: Vereint Senden und Empfangen
# ==========================================
async def handle_client(websocket):
    print(f"\n[{websocket.remote_address}] Web-Monitor verbunden!")
    
    try:
        with voicemeeterlib.api('potato') as vm:
            task_send = asyncio.create_task(send_loop(websocket, vm))
            task_recv = asyncio.create_task(receive_loop(websocket))
            
            done, pending = await asyncio.wait(
                [task_send, task_recv],
                return_when=asyncio.FIRST_COMPLETED,
            )
            for task in pending:
                task.cancel()
                
    except voicemeeterlib.error.VMError:
        print("FEHLER: Voicemeeter Potato läuft nicht!")
        try:
            await websocket.send(json.dumps({"levels": [0]*8, "mutes": [False]*8}))
        except: pass
    except websockets.exceptions.ConnectionClosed:
        pass
    except Exception as e:
        print(f"Server Fehler: {e}")
    finally:
        print(f"[{websocket.remote_address}] Verbindung getrennt.")


# ==========================================
# Webserver für die HTML-Datei
# ==========================================
def start_http_server():
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    class QuietHandler(http.server.SimpleHTTPRequestHandler):
        def log_message(self, format, *args): pass 
        def do_GET(self):
            if self.path == '/manifest.json':
                manifest = {
                    "name": "VM Potato Monitor",
                    "short_name": "VM Monitor",
                    "start_url": "/",
                    "display": "standalone",
                    "background_color": "#0f172a",
                    "theme_color": "#0f172a",
                    "orientation": "landscape",
                    "icons": [{"src": "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIxOTIiIGhlaWdodD0iMTkyIiB2aWV3Qm94PSIwIDAgMTkyIDE5MiI+PHJlY3Qgd2lkdGg9IjE5MiIgaGVpZ2h0PSIxOTIiIGZpbGw9IiMwZjE3MmEiLz48Y2lyY2xlIGN4PSI5NiIgY3k9Ijk2IiByPSI0OCIgZmlsbD0iIzIyYzU1ZSIvPjwvc3ZnPg==", "sizes": "192x192", "type": "image/svg+xml"}]
                }
                self.send_response(200)
                self.send_header("Content-type", "application/manifest+json")
                self.end_headers()
                self.wfile.write(json.dumps(manifest).encode("utf-8"))
                return
            if self.path == '/sw.js':
                sw_content = "self.addEventListener('install', (e) => { self.skipWaiting(); }); self.addEventListener('fetch', (e) => {});"
                self.send_response(200)
                self.send_header("Content-type", "application/javascript")
                self.end_headers()
                self.wfile.write(sw_content.encode("utf-8"))
                return
            if self.path == '/' or self.path == '/index.html':
                try:
                    with open('index.html', 'r', encoding='utf-8') as f:
                        content = f.read()
                    if '<link rel="manifest"' not in content:
                        injection = """<head>\n<link rel="manifest" href="/manifest.json">\n<script>if ('serviceWorker' in navigator) { window.addEventListener('load', () => { navigator.serviceWorker.register('/sw.js'); }); }</script>"""
                        content = content.replace('<head>', injection)
                    self.send_response(200)
                    self.send_header("Content-type", "text/html; charset=utf-8")
                    self.end_headers()
                    self.wfile.write(content.encode("utf-8"))
                    return
                except FileNotFoundError: pass 
            super().do_GET()
            
    with socketserver.TCPServer(("0.0.0.0", HTTP_PORT), QuietHandler) as httpd:
        httpd.serve_forever()

async def main():
    print("=====================================================")
    print(" VOICEMEETER WEB MONITOR & MEDIA SERVER")
    print("=====================================================")
    print(f"1. Öffne auf deinem Handy den Browser (Chrome/Safari)")
    print(f"2. Tippe diese Adresse ein: http://<DEINE-PC-IP>:{HTTP_PORT}")
    print("=====================================================\n")
    
    MediaPollerThread().start()
    async with websockets.serve(handle_client, "0.0.0.0", WS_PORT):
        await asyncio.Future()  

if __name__ == "__main__":
    http_thread = threading.Thread(target=start_http_server, daemon=True)
    http_thread.start()
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\nServer beendet.")