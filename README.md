# mt_detect
A script to check if a given translation is copied from google translate

Wiki needs to be updated.

Compilation notes (these steps necessary to ensure that package is small; otherwise, 200+ MB size)
1. Created virtual environment. Deactivate base anaconda environment
2. In virtual env, installed all libraries (zhon, xlrd, python-pptx, pypiwin32, googletrans)
3. In virtual env, installed pyinstaller
4. Using pyinstaller -w -F (meaning not windowed and onefile), compiled script
