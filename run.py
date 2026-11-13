"""
run.py - 1-Click Root Launcher for VS Code and Terminal.
Click the '▷ Run' button in VS Code or run 'python run.py' to launch the web app.
"""

import os
import sys
import subprocess

if __name__ == "__main__":
    project_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(project_dir)
    app_path = os.path.join(project_dir, "app", "app.py")

    print("\n" + "=" * 65)
    print("  🧠 Autism Spectrum Disorder (ASD) Prediction System")
    print("  Launching Streamlit Clinical Web Dashboard on local browser...")
    print("=" * 65 + "\n")

    cmd = [sys.executable, "-m", "streamlit", "run", app_path, "--server.headless=false"]
    try:
        subprocess.run(cmd)
    except KeyboardInterrupt:
        print("\n[Application stopped by user]")
