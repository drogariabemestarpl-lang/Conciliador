import sys, subprocess
from pathlib import Path
HERE=Path(__file__).resolve().parent
subprocess.call([sys.executable, str(HERE/'concilia_core.py'), '--provider','ALELO'])
