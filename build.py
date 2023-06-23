import os
from pathlib import Path
from PyInstaller import __main__ as pyi

if __name__=='__main__':
    main_script = "frontend.py"

    # PyInstaller options
    pyinstaller_options = [
        "--onefile",  # Create a single executable file
        "--name=OneInNine",  # Name of the final executable
        "--distpath=dist",  # Directory where the executable will be created
        "--hidden-import=pulp"
    ]

    # Set the working directory to the script's directory
    os.chdir(Path(__file__).parent)

    # Build the executable
    pyi.run([
        *pyinstaller_options,
        main_script
    ])

    print("Build completed.")