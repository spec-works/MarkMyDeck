import zipfile, xml.etree.ElementTree as ET

ns = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'
}

issues = []

with zipfile.ZipFile('samples/demo.pptx') as z:
    # 1. Check [Content_Types].xml
    ct = ET.parse(z.open('[Content_Types].xml')).getroot()
    print('=== [Content_Types].xml ===')
    for child in ct:
        print(f'  {child.tag.split("}")[-1]}: {child.attrib}')
    
    # Check if slideLayout relationship exists
    has_layout_ct = False
    for child in ct:
        attrs = child.attrib
        if 'slideLayout' in str(attrs):
            has_layout_ct = True
    if not has_layout_ct:
        issues.append('Missing Content-Type for slideLayout')
    
    # 2. Check _rels/.rels
    print('\n=== _rels/.rels ===')
    rels = ET.parse(z.open('_rels/.rels')).getroot()
    for child in rels:
        print(f'  {child.attrib}')
    
    # 3. Check presentation.xml
    print('\n=== presentation.xml ===')
    pres = ET.parse(z.open('ppt/presentation.xml')).getroot()
    print(ET.tostring(pres, encoding='unicode')[:2000])
    
    # 4. Check presentation rels
    print('\n=== ppt/_rels/presentation.xml.rels ===')
    prels = ET.parse(z.open('ppt/_rels/presentation.xml.rels')).getroot()
    for child in prels:
        print(f'  {child.attrib}')
    
    # 5. Check slide master
    print('\n=== slideMaster1.xml ===')
    sm = ET.parse(z.open('ppt/slideMasters/slideMaster1.xml')).getroot()
    print(ET.tostring(sm, encoding='unicode')[:1500])
    
    # 6. Check slide master rels
    print('\n=== slideMaster1 rels ===')
    smrels = ET.parse(z.open('ppt/slideMasters/_rels/slideMaster1.xml.rels')).getroot()
    for child in smrels:
        print(f'  {child.attrib}')
    
    # 7. Check slideLayout1
    print('\n=== slideLayout1.xml ===')
    sl = ET.parse(z.open('ppt/slideLayouts/slideLayout1.xml')).getroot()
    print(ET.tostring(sl, encoding='unicode')[:1500])
    
    # 8. Check if slideLayout has rels file
    layout_rels = 'ppt/slideLayouts/_rels/slideLayout1.xml.rels'
    if layout_rels in z.namelist():
        print(f'\n=== {layout_rels} ===')
        lrels = ET.parse(z.open(layout_rels)).getroot()
        for child in lrels:
            print(f'  {child.attrib}')
    else:
        issues.append(f'MISSING: {layout_rels}')
        print(f'\n*** MISSING: {layout_rels} ***')
    
    # 9. Check slide rels for proper layout references
    print('\n=== Slide relationships ===')
    for name in sorted(z.namelist()):
        if name.startswith('ppt/slides/_rels/'):
            srels = ET.parse(z.open(name)).getroot()
            for child in srels:
                t = child.attrib.get('Type', '')
                target = child.attrib.get('Target', '')
                rtype = t.split('/')[-1]
                print(f'  {name}: {rtype} -> {target}')
    
    # 10. Check theme reference
    print('\n=== Theme ===')
    theme_file = 'ppt/slideMasters/theme/theme1.xml'
    if theme_file in z.namelist():
        print(f'  Found: {theme_file}')
    else:
        # Check common alternate locations
        for name in z.namelist():
            if 'theme' in name.lower():
                print(f'  Found theme at: {name}')
    
    print('\n=== ISSUES ===')
    for issue in issues:
        print(f'  *** {issue}')
    if not issues:
        print('  None found in basic checks')
