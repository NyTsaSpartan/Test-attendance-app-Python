from cx_Freeze import setup, Executable

base = None

executables = [Executable("votre_script.py", base=base)]

options = {
    'build_exe': {
        'excludes': ['tkinter'],
    },
}

setup(
    name="Nom de votre application",
    version="1.0",
    description="Description de votre application",
    options=options,
    executables=executables
)

# "python setup.py build" a executer dans le terminal

# "pyinstaller --onefile --noconsole votre_script.py" ou taper cete commande dans le terminal