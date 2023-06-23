To build:

to create build after important change use: - this creates a .spec file
pyinstaller -F -w frontend.py --additional-hooks-dir=./hooks_dir --hidden-import=pulp, spec --name OneInNine --onefile


after there is already a good .spec run:
pyinstaller -y OneInNine.spec



link to tables api: [link](https://ttkbootstrap.readthedocs.io/en/latest/api/tableview/tableview/#ttkbootstrap.tableview.Tableview)
link to 