# Windows Audio Switch

Windows Audio Switch is a lightweight portable Windows tray app for switching between selected audio output devices.

It is designed for Windows 10/11, does not require admin rights, and can be distributed as a single `.exe`.

## What It Does

- Runs quietly from the system tray.
- Cycles through your chosen audio output devices with a global hotkey.
- Lets you switch directly to any detected output device from the tray menu.
- Lets you choose which output devices are included in the toggle loop.
- Shows optional Windows tray notifications after switching.
- Starts with Windows if you enable that option.
- Stores settings per user in `%LOCALAPPDATA%\Windows Audio Switch`.
- Includes uninstall cleanup from the Configure window.

The tray icon is a small blue speaker with teal sound waves. Windows may place it in the hidden tray overflow area.

## Download And Install

1. Download `audio-switch.exe`.
2. Put it anywhere you like, for example:

   ```text
   C:\Program Files\Personal Apps\Windows Audio Switch\audio-switch.exe
   ```

3. Double-click `audio-switch.exe`.
4. On first launch, the Configure window opens automatically.
5. Select at least two output devices, confirm the hotkey, choose startup/notification options, and click **Save**.

After saving, the app keeps running from the system tray.

Admin rights are not required to run the app. You may need admin rights only if you copy the `.exe` into a protected folder such as `C:\Program Files`.

## How To Use

Open the tray icon menu to access:

- **Toggle audio output**: switches to the next selected device.
- **Switch directly**: switches to a specific detected output device.
- **Configure**: opens settings.
- **Start with Windows**: toggles the per-user startup entry.
- **Open config folder**: opens saved settings/logs.
- **Quit**: exits the tray app.

The default hotkey is:

```text
ctrl+alt+a
```

You can change it in **Configure**. Supported hotkey examples include `ctrl+alt+a`, `ctrl+shift+f9`, and `win+alt+space`.

## Configure

The Configure window includes:

- Detected output devices with checkboxes.
- Hotkey field.
- Start with Windows checkbox.
- Enable notifications checkbox.
- Save / Cancel buttons.
- Uninstall cleanup button.

At least two devices must be selected before the toggle loop can be saved.

## Notifications

Notifications are optional. If **Enable notifications** is unchecked, Windows Audio Switch does not show tray notifications for switches, missing devices, or hotkey errors.

## Missing Devices

If a selected device is unplugged or unavailable, the app skips it. If notifications are enabled, you will see a short message. The device remains in your saved toggle loop and will work again when Windows reports it as active.

## Startup

When **Start with Windows** is enabled, the app writes this per-user registry value:

```text
HKCU\Software\Microsoft\Windows\CurrentVersion\Run\Windows Audio Switch
```

This does not require admin rights.

## Saved Files

Settings and logs are stored here:

```text
%LOCALAPPDATA%\Windows Audio Switch
```

Current config format:

```json
{
  "version": 2,
  "audio_toggle": {
    "hotkey": "ctrl+alt+a",
    "device_ids": [],
    "notifications_enabled": true
  }
}
```

The app stores stable Windows audio endpoint IDs, but displays friendly device names in the UI.

## Uninstall

To clean up startup and local settings:

1. Open the tray menu.
2. Choose **Configure**.
3. Click **Uninstall...**.

This removes the startup registry entry and deletes `%LOCALAPPDATA%\Windows Audio Switch`. It does not delete the portable `.exe`; delete that file manually from wherever you placed it.

You can also run:

```bat
uninstall.bat
```

## Command Line

The executable also supports CLI commands:

```bat
audio-switch.exe list
audio-switch.exe set <device-id-or-index>
audio-switch.exe toggle
audio-switch.exe configure
audio-switch.exe daemon
audio-switch.exe --help
```

`list` prints active output devices and marks the current default with `*`.

`set` accepts a device ID, exact friendly name, or the numeric index shown by `list`.

`toggle` cycles through the configured devices in order.

Running the executable with no arguments starts the tray daemon.

## Troubleshooting

- If the tray icon is missing, check the hidden tray overflow near the clock.
- If the hotkey does not work, another app may already be using it. Open Configure and choose a different hotkey.
- If you see a message that Windows Audio Switch is already running, use the existing tray icon or close `audio-switch.exe` from Task Manager.
- If setting the default output fails, check Windows sound settings and confirm the device is active.
- If the app cannot start, check `%LOCALAPPDATA%\Windows Audio Switch\WindowsAudioSwitch.log`.

## Build From Source

Install Python 3.10 or newer, then run:

```bat
build.bat
```

The single-file executable is written to:

```text
dist\audio-switch.exe
```

Runtime dependencies:

- `comtypes`: Windows Core Audio and PolicyConfig COM access.
- `pystray`: tray icon and menu.
- `Pillow`: tray icon image generation.

Build dependency:

- `pyinstaller`: single-file `.exe` packaging.
