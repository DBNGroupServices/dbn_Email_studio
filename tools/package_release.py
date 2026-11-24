"""Create the user package, a source archive and dependency license notices."""
import importlib.metadata
from pathlib import Path
import zipfile
import hashlib

ROOT=Path(__file__).resolve().parents[1]
EXE=ROOT/'dist'/'DBN Email Studio.exe'
OUTPUT=ROOT/'DBN Email Studio.zip'
NOTICE=ROOT/'THIRD-PARTY-NOTICES.txt'
packages=list(importlib.metadata.distributions(path=[str(ROOT/'.build-deps')]))
for name in ('Pillow','lxml','charset-normalizer'):
    try:packages.append(importlib.metadata.distribution(name))
    except importlib.metadata.PackageNotFoundError:pass
notice=['DBN Email Studio - third-party components','', 'The executable bundles Python, extract-msg and their runtime dependencies. Original component licences apply. This list includes build tools used to produce the executable. Source and project links are preserved in package metadata.', '']
licences=[]
for dist in packages:
    name=dist.metadata.get('Name','package')
    notice.append(f'{name} {dist.version}')
    licence=dist.metadata.get('License-Expression') or dist.metadata.get('License') or ''
    if licence and len(licence)<300:notice.append('Licence: '+licence)
    for url in dist.metadata.get_all('Project-URL',[]):notice.append(url)
    if dist.metadata.get('Home-page'):notice.append(dist.metadata['Home-page'])
    for item in dist.files or []:
        if any(part.lower().startswith(('license','licence','copying','notice')) for part in item.parts) and ('.dist-info' in str(item) or 'pyinstaller' in name.lower()):
            path=Path(dist.locate_file(item))
            if path.is_file() and path.stat().st_size<300000:licences.append((path,'Licences/'+name+'/'+str(item).replace('\\','/')))
    notice.append('')
NOTICE.write_text('\n'.join(notice),encoding='utf-8')
with zipfile.ZipFile(OUTPUT,'w',zipfile.ZIP_DEFLATED) as package:
    package.write(EXE,'DBN Email Studio/DBN Email Studio.exe')
    for filename in ('START HERE.txt','README.md','THIRD-PARTY-NOTICES.txt','TEST-RESULTS.md'):
        package.write(ROOT/filename,'DBN Email Studio/'+filename)
    package.write(ROOT/'assets'/'email-studio-blue.ico','DBN Email Studio/DBN Email Studio Blue.ico')
    for file in (ROOT/'office-addin').rglob('*'):
        if file.is_file():package.write(file,'DBN Email Studio/Outlook add-in/'+str(file.relative_to(ROOT/'office-addin')).replace('\\','/'))
    for file in (ROOT/'examples').iterdir():package.write(file,'DBN Email Studio/examples/'+file.name)
    for file,name in licences:package.write(file,'DBN Email Studio/'+name)
    python_licence=Path(__import__('sys').base_prefix)/'LICENSE.txt'
    if python_licence.exists():package.write(python_licence,'DBN Email Studio/Licences/Python-LICENSE.txt')
with zipfile.ZipFile(OUTPUT) as checked:
    assert checked.testzip() is None
with zipfile.ZipFile(ROOT/'DBN Email Studio - Source.zip','w',zipfile.ZIP_DEFLATED) as source:
    for filename in ('app.py','chatgpt_auth.py','outlook_bridge.py','office_tls.py','windows_shell.py','README.md','START HERE.txt','requirements.txt','requirements-build.txt','build.ps1','.gitignore','THIRD-PARTY-NOTICES.txt','TEST-RESULTS.md'):
        source.write(ROOT/filename,filename)
    for folder in ('web','assets','tests','tools','examples','hooks','office-addin'):
        for file in (ROOT/folder).rglob('*'):
            if file.is_file() and '__pycache__' not in file.parts:source.write(file,file.relative_to(ROOT))
with zipfile.ZipFile(ROOT/'DBN Email Studio - Source.zip') as checked:
    assert checked.testzip() is None
digest=hashlib.sha256(EXE.read_bytes()).hexdigest()
(ROOT/'dist'/'SHA256.txt').write_text(digest+'  DBN Email Studio.exe\n',encoding='ascii')
print(f'Package: {OUTPUT.name} ({OUTPUT.stat().st_size/1024/1024:.1f} MB); ZIP integrity OK')
print('EXE SHA256:',digest)
