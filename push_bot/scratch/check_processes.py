import subprocess
import sys

def main():
    try:
        cmd = ["powershell", "-Command", "Get-Process -Id 22732 | Select-Object Name, Id, StartTime | ConvertTo-Json"]
        result = subprocess.run(cmd, capture_output=True, text=True)
        print(result.stdout)
    except Exception as e:
        print("Exception:", e, file=sys.stderr)

if __name__ == "__main__":
    main()
