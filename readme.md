# AudioNormalizer

A desktop GUI application for analyzing and normalizing audio loudness using FFmpeg. Built with PyQt5, it keeps the UI responsive with background threads, preserves metadata and album art, and lets you specify target LUFS, sample rate, bitrate mode (VBR/CBR), and bitrate.

## System Requirements (Runtime)

- Windows, macOS, or Linux

## System Requirements (Development)

- Windows, macOS, or Linux
- Python 3.10 or later
- PyQt5

## Development

### Install Dependencies

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

### Run Directly
```bash
python audio_normalizer.py
```

## Building Executables

### Windows
Create an executable with the PowerShell script:
```powershell
.\make.ps1
```

### Unix-like (Linux/macOS)
Create an executable with the shell script:
```bash
bash ./make.sh
```

The executable will be generated in the `dist/` folder.
