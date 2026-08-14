import os, subprocess, sys

ROOT = os.path.dirname(os.path.abspath(__file__))
os.chdir(ROOT)

def run(cmd):
    print('> ' + ' '.join(cmd))
    return subprocess.call(cmd)

print('Financial P&L Analysis - setup and launch')
if run([sys.executable, '-m', 'pip', 'install', '-r', 'requirements.txt']) != 0:
    print('\nPackage installation failed.')
    input('Press Enter to close...')
    raise SystemExit(1)

code = run([sys.executable, '-m', 'streamlit', 'run', 'app.py'])
if code != 0:
    print('\nStreamlit stopped with an error. Error code:', code)
    input('Press Enter to close...')
raise SystemExit(code)
