# 🤖 Pecifics - AI Desktop Assistant

A powerful desktop AI assistant with Large Action Model (LAM) capabilities. Pecifics can see your screen, understand your requests, and take actions on your computer.

![Pecifics Screenshot](assets/screenshot.png)

## ✨ Features

- **Vision Understanding**: Captures and analyzes your screen in real-time
- **Natural Language Control**: Tell Pecifics what to do in plain English
- **Desktop Automation**: 
  - Create/delete files and folders
  - Open applications
  - Control mouse and keyboard
  - Type text automatically
  - Browse the web
  - And much more!
- **Powered by Open-Source AI**: Uses Qwen2-VL for vision and Qwen2.5 for reasoning
- **Free GPU via Google Colab**: No expensive hardware needed!

## 🏗️ Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                 Desktop App (Electron)                       │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────────┐   │
│  │   Chat UI    │  │  Screenshot  │  │ Action Executor  │   │
│  │              │◄─┤   Capture    │  │ (Mouse/Keyboard) │   │
│  └──────┬───────┘  └──────┬───────┘  └────────▲─────────┘   │
│         │                 │                    │              │
└─────────┼─────────────────┼────────────────────┼─────────────┘
          │                 │                    │
          │    ┌────────────▼────────────┐       │
          │    │        Internet         │       │
          │    └────────────┬────────────┘       │
          │                 │                    │
┌─────────▼─────────────────▼────────────────────┼─────────────┐
│              Google Colab (Free GPU)                         │
│  ┌──────────────────────────────────────────────────────┐   │
│  │                 FastAPI Server                        │   │
│  │  ┌──────────────┐      ┌──────────────────────────┐  │   │
│  │  │  Qwen2-VL    │      │      Qwen2.5-7B          │  │   │
│  │  │  (Vision)    │ ───► │  (Reasoning + Actions)   │──┼───┘
│  │  └──────────────┘      └──────────────────────────┘  │
│  └──────────────────────────────────────────────────────┘   │
│                    Exposed via ngrok                         │
└─────────────────────────────────────────────────────────────┘
```

## 📋 Requirements

### Desktop App
- Windows 10/11 (64-bit)
- Node.js 18+ (for development)
- Internet connection

### Colab Backend
- Google account (for Colab)
- ngrok account (free tier works)

## 🚀 Quick Start

### Step 1: Setup the Colab Backend

1. Open [Google Colab](https://colab.research.google.com/)
2. Upload `colab-backend/JARVIS_LAM_Backend.ipynb` or `JARVIS_LAM_Smart.ipynb` (recommended)
3. Select **GPU runtime**: `Runtime` → `Change runtime type` → `T4 GPU`
4. Get your free ngrok auth token from [ngrok.com](https://ngrok.com/)
5. Paste the token in the notebook
6. Run all cells in order
7. Copy the ngrok URL (looks like `https://xxxx.ngrok.io`)

### Step 2: Install the Desktop App

#### Option A: From Installer (Recommended)
1. Download the latest release from [Releases](releases/)
2. Run `Pecifics-Setup.exe` (formerly JARVIS-Setup.exe)
3. Follow the installation wizard

#### Option B: From Source (Development)
```bash
# Clone the repository
cd jarvis-desktop

# Install dependencies
npm install

# Run in development mode
npm run dev

# Build installer
npm run build:win
```

### Step 3: Configure Pecifics

1. Launch Pecifics from desktop shortcut
2. Click the ⚙️ settings icon
3. Paste the ngrok URL from Colab
4. Click "Test Connection" to verify
5. Save settings

### Step 4: Start Using!

Try these commands:
- "Create a folder called Projects on my Desktop"
- "Open Notepad and type Hello World"
- "Open Chrome and search for weather"
- "Create a file called notes.txt with my shopping list"

## 📁 Project Structure

```
jarvis-desktop/
├── src/
│   ├── main.js              # Electron main process
│   ├── preload.js           # Context bridge
│   ├── renderer/
│   │   ├── index.html       # Main UI
│   │   ├── styles.css       # Styles
│   │   └── renderer.js      # Frontend logic
│   └── modules/
│       └── action-executor.js  # Desktop automation
├── assets/
│   ├── icon.ico             # Windows icon
│   ├── icon.png             # PNG icon
│   └── icon.icns            # macOS icon
├── package.json
└── README.md

colab-backend/
└── JARVIS_LAM_Backend.ipynb  # Colab notebook with AI models
```

## ⚙️ Configuration

### Settings

| Setting | Description | Default |
|---------|-------------|---------|
| Colab URL | Backend ngrok URL | - |
| Screenshot Interval | How often to capture screen (ms) | 1000 |
| Screenshot Quality | JPEG quality (30-100) | 80 |
| Hotkey | Toggle window shortcut | Ctrl+Shift+J |

### Available Actions

Pecifics can perform these actions:

| Category | Actions |
|----------|---------|
| **Files** | create_file, delete_file, read_file, copy_file, move_file, rename_file |
| **Folders** | create_folder, list_directory |
| **Apps** | open_application, close_application |
| **Mouse** | click_at, move_mouse, scroll, drag |
| **Keyboard** | type_text, press_key |
| **Web** | open_url, search_web |
| **System** | run_command, speak |

## 🔧 Troubleshooting

### "Connection failed" error
- Make sure the Colab notebook is running
- Check that the ngrok URL is correct
- Colab sessions timeout after ~12 hours - restart if needed

### Mouse/keyboard not working
- Run as Administrator
- Some apps block automated input

### Screenshots not capturing
- Check if another app is blocking screen capture
- Try running as Administrator

### High CPU/memory usage
- Increase screenshot interval (2000ms+)
- Reduce screenshot quality

## 🛠️ Development

### Building from Source

```bash
# Install dependencies
npm install

# Run development mode (with DevTools)
npm run dev

# Build for Windows
npm run build:win

# Build for macOS
npm run build:mac

# Build for Linux
npm run build:linux
```

### Colab Notebook Modifications

To use different models, edit the notebook:

```python
# For smaller model (less VRAM):
vision_model = "Qwen/Qwen2-VL-2B-Instruct"
text_model = "Qwen/Qwen2.5-3B-Instruct"

# For better quality (more VRAM):
vision_model = "Qwen/Qwen2-VL-7B-Instruct"  # Default
text_model = "Qwen/Qwen2.5-7B-Instruct"     # Default
```

## 📝 License

MIT License - Feel free to use and modify!

## 🤝 Contributing

Contributions welcome! Please:
1. Fork the repo
2. Create a feature branch
3. Submit a pull request

## ⚠️ Disclaimer

Pecifics can control your computer. Use responsibly:
- Always review actions before confirming
- Be careful with file deletion commands
- Don't share your screen while using banking/sensitive apps
- The AI may occasionally misinterpret commands

## 🙏 Credits

- [Qwen2-VL](https://github.com/QwenLM/Qwen2-VL) - Vision Language Model
- [Qwen2.5](https://github.com/QwenLM/Qwen2.5) - Language Model
- [Electron](https://electronjs.org/) - Desktop framework
- [RobotJS](https://robotjs.io/) - Desktop automation
