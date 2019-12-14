# mt_detect
A script to check if a given translation is copied from google translate

Wiki needs to be updated.

Compilation notes (these steps necessary to ensure that package is small; otherwise, 200+ MB size)
1. Created virtual environment (virtualenv mt_detect). Deactivate base anaconda environment
2. In virtual env, installed all libraries (xlrd, python-pptx, pypiwin32, googletrans. google-auth, google-cloud-translate)
    Notes:
    a. google-auth and google-cloud-translate only needed for cmt_detect.py
    b. pip install --upgrade for google-auth and google-cloud-translate)
3. In virtual env, installed pyinstaller
4. Using pyinstaller -w -F (meaning not windowed and onefile), compiled script
