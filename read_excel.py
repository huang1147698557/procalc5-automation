import zipfile, xml.etree.ElementTree as ET, sys
try:
    z = zipfile.ZipFile('/Users/Apple/Downloads/lili_work/work1/procalc5/procalc5.proflute.xlsx')
    ss = []
    tree = ET.parse(z.open('xl/sharedStrings.xml'))
    ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    for si in tree.iter(ns+'si'):
        texts = list(si.iter(ns+'t'))
        ss.append(''.join((t.text or '') for t in texts))
    tree2 = ET.parse(z.open('xl/worksheets/sheet1.xml'))
    for i, row in enumerate(tree2.iter(ns+'row')):
        if i >= 3: break
        vals = []
        for cell in row.iter(ns+'c'):
            v = cell.find(ns+'v')
            t = cell.get('t','')
            if t == 's' and v is not None:
                vals.append(ss[int(v.text)])
            elif v is not None:
                vals.append(v.text)
            else:
                vals.append('')
        sys.stdout.write('Row %d: %s\n' % (i, vals))
except Exception as e:
    sys.stdout.write('ERROR: %s\n' % str(e))
sys.stdout.flush()
