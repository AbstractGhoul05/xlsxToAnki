import os
import shutil

if os.path.exists('output.apkg'):
    os.remove('output.apkg')
if os.path.exists('images'):
    shutil.rmtree('images')
