"""Build the two-file Apps Script distribution from the modular source files."""
from pathlib import Path
import re, shutil
root = Path(__file__).resolve().parent.parent
out = root / 'copy-paste'
out.mkdir(exist_ok=True)
files = ['Config.gs', 'Util.gs', 'Auth.gs', 'Drive.gs', 'Paths.gs', 'Quizzes.gs', 'Reports.gs', 'Admin.gs', 'Code.gs']
code = '\n\n'.join('// ===== ' + name + ' =====\n' + (root / name).read_text(encoding='utf8') for name in files)
(out / 'Code.gs').write_text(code, encoding='utf8')
html = (root / 'Index.html').read_text(encoding='utf8')
for name in ['Styles', 'Brand', 'Demo', 'AdminQuiz', 'Admin', 'App']:
    marker = "<?!= include_('" + name + "') ?>"
    if html.count(marker) != 1:
        raise RuntimeError('Missing or duplicate include: ' + name)
    html = html.replace(marker, (root / (name + '.html')).read_text(encoding='utf8'))
if re.search(r'<\?', html):
    raise RuntimeError('Unexpected template directive in bundled HTML')
(out / 'Index.html').write_text(html, encoding='utf8')
shutil.copyfile(root / 'appsscript.json', out / 'appsscript.json')
print('Built copy-paste/Code.gs, Index.html, appsscript.json')
