# mt_detect
A script to check if a given translation is copied from google translate

Wiki needs to be updated.

Compilation notes (these steps necessary to ensure that package is small; otherwise, 200+ MB size)
1. Created virtual environment (virtualenv mt_detect). Deactivate base anaconda environment
2. Deactivated default environment (conda deactivate). 
3. Activated virtual environment by navigating to root of mt_detect folder, then Scripts\activate.
4. In virtual env, installed all libraries (xlrd, python-pptx, pypiwin32)
    Notes:  
    a. Additional google-auth and google-cloud-translate needed for cmt_detect.py  
    b. pip install --upgrade google-auth  
    c. pip install google-cloud-translate==2.0.0  
    d. Goto Pyinstaller hooks folder (~\Lib\site-packages\PyInstaller\hooks)  
    e. Find the file hook-google.cloud.py. Add the following code to this file  
            datas += copy_metadata('google-cloud-translate')  
            datas += copy_metadata('google-api-core')  
        This is necessary to solve Distributionnotfound errors in compilation.
5. In virtual env, installed pyinstaller
6. Using pyinstaller -w -F (meaning not windowed and onefile), compiled script
