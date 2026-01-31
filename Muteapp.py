from flask import Flask, render_template_string, request, jsonify
import os, keyboard, mouse, comtypes, json, time, psutil
from pycaw.pycaw import AudioUtilities

app = Flask(__name__)

# --- CONFIG ---
SETTINGS_FILE = "settings.json"
config = {"target_app": "chrome.exe", "mute_hotkey": "f9", "reduce_hotkey": "f10", "reduce_level": 20}

def save_settings():
    with open(SETTINGS_FILE, "w") as f: json.dump(config, f)

if os.path.exists(SETTINGS_FILE):
    try:
        with open(SETTINGS_FILE, "r") as f: config.update(json.load(f))
    except: pass

MOUSE_BUTTONS = ['left', 'right', 'middle', 'x', 'x2']

def run_mute():
    comtypes.CoInitialize()
    try:
        sessions = AudioUtilities.GetAllSessions()
        for s in sessions:
            if s.Process and s.Process.name().lower() == config["target_app"].lower():
                vol = s.SimpleAudioVolume
                vol.SetMute(0 if vol.GetMute() else 1, None)
    finally: comtypes.CoUninitialize()

def run_reduce():
    comtypes.CoInitialize()
    try:
        sessions = AudioUtilities.GetAllSessions()
        for s in sessions:
            if s.Process and s.Process.name().lower() == config["target_app"].lower():
                vol = s.SimpleAudioVolume
                curr = round(vol.GetMasterVolume(), 2)
                low = float(config["reduce_level"]) / 100.0
                vol.SetMasterVolume(1.0 if curr <= (low + 0.05) else low, None)
                vol.SetMute(0, None)
    finally: comtypes.CoUninitialize()

def update_hotkeys():
    keyboard.unhook_all()
    try: mouse.unhook_all()
    except: pass
    
    def handle_event(e):
        if e.event_type == 'down':
            for action in ["mute", "reduce"]:
                hotkey = config.get(f"{action}_hotkey", "").lower()
                if not hotkey or hotkey == "none": continue
                
                # Support for combos like shift+f9
                parts = hotkey.split('+')
                main_key = parts[-1]
                modifiers = parts[:-1]

                # Check if the key pressed is our main key
                if e.name == main_key:
                    # Check if all required modifiers (shift, ctrl, etc) are held
                    if all(keyboard.is_pressed(mod) for mod in modifiers):
                        if action == "mute": run_mute()
                        else: run_reduce()

    keyboard.hook(handle_event)

    # Mouse handling (remains separate)
    for action in ["mute", "reduce"]:
        key = config.get(f"{action}_hotkey", "")
        func = run_mute if action == "mute" else run_reduce
        if key.lower() in MOUSE_BUTTONS:
            mouse.on_button(func, buttons=(key,), types=('down',))

# --- WEB UI ---
HTML_PAGE = """
<!DOCTYPE html>
<html>
<head>
    <title>Audio Controller Pro</title>
    <style>
        body { font-family: 'Segoe UI', sans-serif; text-align: center; padding: 20px; background: #1e1e2e; color: white; }
        .card { background: #313244; padding: 25px; border-radius: 15px; display: inline-block; width: 440px; box-shadow: 0 10px 30px rgba(0,0,0,0.5); }
        .section { background: #45475a; padding: 15px; border-radius: 10px; margin: 10px 0; text-align: left; }
        label { font-weight: bold; color: #fab387; display: block; margin-bottom: 5px; }
        .key-row { display: flex; gap: 8px; }
        .key-input { flex-grow: 1; padding: 10px; border-radius: 5px; border: 1px solid #6c7086; background: #1e1e2e; color: #a6e3a1; text-align: center; font-weight: bold; }
        .btn-rec { background: #fab387; border: none; padding: 10px; border-radius: 5px; cursor: pointer; color: #1e1e2e; font-weight: bold; flex: 1; }
        .btn-save { padding: 15px; background: #a6e3a1; color: #1e1e2e; border: none; border-radius: 8px; cursor: pointer; font-weight: bold; width: 100%; margin-top: 10px; font-size: 1.1em; }
        select { width: 100%; padding: 10px; border-radius: 5px; background: #1e1e2e; color: white; }
        #status-bar { color: #89b4fa; font-size: 0.9em; margin-bottom: 10px; height: 1.2em; font-weight: bold; }
    </style>
</head>
<body>
    <div class="card">
        <h2>Audio Controller</h2>
        <div id="status-bar">Ready</div>
        <form method="POST">
            <div class="section">
                <label>Target App:</label>
                <select name="app_name">
                    {% for app in apps %}<option value="{{ app }}" {% if app == current_app %}selected{% endif %}>{{ app }}</option>{% endfor %}
                </select>
            </div>
            <div class="section">
                <label>🔇 Mute Hotkey:</label>
                <div class="key-row">
                    <input type="text" name="mute_hotkey" id="mute_input" class="key-input" value="{{ mute_key }}" readonly>
                    <button type="button" class="btn-rec" onclick="startRecord('mute')">REC</button>
                </div>
            </div>
            <div class="section">
                <label>🔉 Reduce Hotkey:</label>
                <div class="key-row">
                    <input type="text" name="reduce_hotkey" id="reduce_input" class="key-input" value="{{ reduce_key }}" readonly>
                    <button type="button" class="btn-rec" onclick="startRecord('reduce')">REC</button>
                </div>
                <input type="range" name="reduce_level" min="0" max="80" value="{{ current_level }}" style="width:100%; margin-top:10px;">
            </div>
            <button type="submit" class="btn-save">Apply Settings</button>
        </form>
    </div>

    <script>
        async function startRecord(type) {
            const status = document.getElementById('status-bar');
            const input = document.getElementById(type + '_input');
            status.innerText = "Listening... Hold modifiers and press key.";
            
            const keyHandler = (e) => {
                e.preventDefault();
                let keys = [];
                if (e.ctrlKey) keys.push('ctrl');
                if (e.shiftKey) keys.push('shift');
                if (e.altKey) keys.push('alt');
                const main = e.key.toLowerCase();
                if (!['control', 'shift', 'alt'].includes(main)) keys.push(main);
                if (keys.length > 0) input.value = keys.join('+');
            };

            window.addEventListener('keydown', keyHandler);
            
            // Mouse backup
            const res = await fetch('/record_mouse');
            const data = await res.json();
            if (data.key) input.value = data.key;

            setTimeout(() => {
                window.removeEventListener('keydown', keyHandler);
                status.innerText = "Ready";
            }, 3000);
        }
    </script>
</body>
</html>
"""

@app.route("/record_mouse")
def record_mouse():
    captured = []
    def on_click(e):
        if not captured and hasattr(e, 'button') and e.event_type == 'down':
            captured.append(e.button)
    mouse.hook(on_click)
    time.sleep(2.5)
    mouse.unhook(on_click)
    return jsonify({"key": captured[0] if captured else None})

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        config.update({
            "target_app": request.form.get("app_name"),
            "mute_hotkey": request.form.get("mute_hotkey"),
            "reduce_hotkey": request.form.get("reduce_hotkey"),
            "reduce_level": int(request.form.get("reduce_level"))
        })
        save_settings()
        update_hotkeys()

    apps = sorted({p.info['name'] for p in psutil.process_iter(['name'])})
    return render_template_string(HTML_PAGE, apps=apps, current_app=config["target_app"], 
                                 mute_key=config["mute_hotkey"], reduce_key=config["reduce_hotkey"], current_level=config["reduce_level"])

if __name__ == "__main__":
    update_hotkeys()
    app.run(port=5000, debug=False, threaded=True)