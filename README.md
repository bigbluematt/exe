# exe

Drop any `.py` Python script into this repo and GitHub Actions will automatically build it into a `.exe` file with the same name.

## How it works

1. Upload or push a `.py` file to this repo
2. GitHub Actions triggers automatically on any `.py` file change
3. PyInstaller builds a standalone `.exe` on a Windows runner
4. The `.exe` is available as a **downloadable artifact** in the Actions tab

## Download your EXE

1. Go to the [Actions tab](../../actions)
2. Click the latest workflow run
3. Download the `executables` artifact — it contains your `.exe` file

## Notes

- Uses Python 3.11 and PyInstaller (`--onefile` mode)
- Each `.exe` is named the same as the original `.py` file
- Only files changed in the latest commit are built
