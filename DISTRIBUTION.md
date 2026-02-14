# Shipping the app to customers

**Best option (fewest steps for customer):** You build once → ship one zip → customer unzips and double-clicks. No Python install, no terminal, no commands on their side.

---

## Recommended: One-click package (no Python on customer PC)

### What you do (once)

**On Windows:** Double-click **`build_for_customer.bat`**

**On Mac/Linux:** In terminal run:
```bash
chmod +x build_for_customer.sh
./build_for_customer.sh
```

This creates a folder **`BackDownCalculator_Ready`** with the app, a ready-made Python environment, and a launcher. Zip that folder (e.g. `BackDownCalculator_Ready.zip`) and send it to the customer.

### What the customer does

1. **Unzip** the folder (e.g. to Desktop or `C:\BackDownCalculator`).
2. **Double-click**  
   - **Windows:** `START_APP.bat`  
   - **Mac/Linux:** `START_APP.sh` (or in terminal: `./START_APP.sh`)
3. The browser opens with the app. They use it.
4. To stop: close the command/terminal window that opened.

No installation, no Python, no commands. You can include **README.txt** in the zip (it’s added automatically from `CUSTOMER_README.txt`) with these 3 steps.

**Note:** Build the package on the same type of OS as the customer (e.g. build on Windows for Windows users). The included `venv` is not cross-platform.

---

## Installable package (Windows setup.exe)

For a familiar “Install” experience (Next → choose folder → Start Menu shortcut, optional Desktop shortcut, Add/Remove Programs entry):

### What you do (on Windows)

1. **Build the app package:** Double-click **`build_for_customer.bat`** (creates `BackDownCalculator_Ready`).
2. **Install Inno Setup** (free): https://jrsoftware.org/isdl.php — use default install path.
3. **Build the installer:** Double-click **`build_installer.bat`**.  
   - If Inno Setup is installed, it will compile and create **`installer_output\BackDownCalculator_Setup.exe`**.  
   - If not, it will open `installer.iss` in Inno Setup; use **Build → Compile** there.
4. **Ship** `BackDownCalculator_Setup.exe` to the customer.

### What the customer does

1. Run **BackDownCalculator_Setup.exe**.
2. Click Next through the wizard (or change install folder if they want).
3. Finish. They get:
   - **Start Menu** entry: “Back Down Calculator” (launches the app).
   - **Optional Desktop shortcut** (if they ticked the option).
   - Entry in **Add or remove programs** (they can uninstall later).
4. To use the app: Start Menu → **Back Down Calculator**, or double-click the Desktop shortcut. Browser opens with the app.

No Python, no zip to unzip, no batch files to find — just install and run from the Start Menu.

---

## Other options

### Zip + Python (customer has Python installed)

If the customer already has Python 3.10+ and is okay with a few steps:

1. Send a zip of the project (all `.py` files, `requirements.txt`, `run.bat` / `run.sh`).
2. They unzip, then double-click **`run.bat`** (Windows) or run **`./run.sh`** (Mac/Linux).  
   The script creates a venv, installs dependencies, and starts the app.

### Standalone .exe (advanced)

You can use PyInstaller to build a single Windows `.exe`. The app would need a small launcher script that starts Streamlit and opens the browser. This is more work and can trigger antivirus; the “build once, ship zip” method above is simpler and very reliable.

### Docker

If the customer uses Docker: provide a Dockerfile and they run `docker build` / `docker run`. See the Docker section in the previous version of this doc if you need it.

---

## Summary

| Method | Customer steps | You do |
|--------|----------------|--------|
| **Installer (best for Windows)** | Run setup.exe → Next → Finish → Start Menu / Desktop | Run `build_for_customer.bat`, then `build_installer.bat` (needs Inno Setup); ship .exe |
| **One-click zip** | Unzip → double-click `START_APP` | Run `build_for_customer.bat` or `.sh`, zip the folder, send it |
| Zip + run script | Unzip → double-click `run.bat` / `run.sh` | Zip project, send (customer needs Python) |

For Windows customers, the **installable setup.exe** is the most familiar; for Mac/Linux, use the **zip + START_APP.sh**.

---

## Controlling the app menu (what customers see)

The top-right menu (Rerun, Settings, Print, Developer options, Clear cache, etc.) is controlled by **`.streamlit/config.toml`** in the project. It is included automatically when you run `build_for_customer.bat` or `build_for_customer.sh`.

**Current setting:** `toolbarMode = "viewer"` — customers see Rerun, Settings, Print, Record a screencast, About; they do **not** see Developer options or Clear cache.

To change what appears, edit **`.streamlit/config.toml`** before building the customer package:

| `toolbarMode`  | What customers see |
|----------------|---------------------|
| **viewer**     | Rerun, Settings, Print, Record screencast, About. No Developer options or Clear cache. (Recommended for customers.) |
| **minimal**    | Only options you set via the app or Community Cloud. |
| **developer**  | Full menu including Developer options and Clear cache. |
| **auto**       | Developer options only when run on localhost or as admin (Streamlit default). |

To hide the entire top bar (menu + Deploy), uncomment in `config.toml`:
```toml
[ui]
hideTopBar = true
```
