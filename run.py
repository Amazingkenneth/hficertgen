import streamlit.web.cli as stcli
import os, sys

def resolve_path(path):
    if getattr(sys, "frozen", False):
        basedir = sys._MEIPASS
    else:
        basedir = os.path.dirname(__file__)
    return os.path.join(basedir, path)

if __name__ == "__main__":
    # Point to the main app file
    script_path = resolve_path("gen.py")
    
    # Fake the command line arguments
    sys.argv = [
        "streamlit",
        "run",
        script_path,
        "--global.developmentMode=false",
    ]
    
    sys.exit(stcli.main())