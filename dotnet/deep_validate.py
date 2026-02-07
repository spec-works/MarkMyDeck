"""Compare our PPTX structure against OOXML spec requirements."""
import zipfile, xml.etree.ElementTree as ET

ns = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

issues = []

with zipfile.ZipFile('samples/demo.pptx') as z:
    # Check required files
    required = [
        '[Content_Types].xml',
        '_rels/.rels',
        'ppt/presentation.xml',
        'ppt/_rels/presentation.xml.rels',
        'ppt/slideMasters/slideMaster1.xml',
        'ppt/slideMasters/_rels/slideMaster1.xml.rels',
        'ppt/slideLayouts/slideLayout1.xml',
        'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
    ]
    for f in required:
        if f not in z.namelist():
            issues.append(f'MISSING required file: {f}')
    
    # Validate each slide XML
    for name in sorted(z.namelist()):
        if name.startswith('ppt/slides/slide') and name.endswith('.xml') and '_rels' not in name:
            try:
                tree = ET.parse(z.open(name))
                root = tree.getroot()
                
                # Check for CommonSlideData
                cSld = root.find('p:cSld', ns)
                if cSld is None:
                    issues.append(f'{name}: missing cSld')
                
                # Check shape tree
                spTree = root.find('.//p:spTree', ns) if cSld is not None else None
                if spTree is None:
                    issues.append(f'{name}: missing spTree')
                
                # Check each shape has valid transform
                shapes = root.findall('.//p:sp', ns)
                for sp in shapes:
                    nvPr = sp.find('.//p:cNvPr', ns)
                    shape_name = nvPr.get('name', '?') if nvPr is not None else '?'
                    
                    # TextBody must have at least one paragraph
                    txBody = sp.find('p:txBody', ns)
                    if txBody is not None:
                        paras = txBody.findall('a:p', ns)
                        if len(paras) == 0:
                            issues.append(f'{name}/{shape_name}: TextBody has no paragraphs')
                        
                        # Each paragraph must have at least a run or endParaRPr
                        for i, p in enumerate(paras):
                            runs = p.findall('a:r', ns)
                            endPr = p.find('a:endParaRPr', ns)
                            if len(runs) == 0 and endPr is None:
                                # Empty paragraph - needs at least endParaRPr
                                issues.append(f'{name}/{shape_name}: paragraph {i} has no runs and no endParaRPr')
                
                # Check tables
                tables = root.findall('.//a:tbl', ns)
                for t in tables:
                    rows = t.findall('a:tr', ns)
                    grid = t.find('a:tblGrid', ns)
                    if grid is None:
                        issues.append(f'{name}: table missing tblGrid')
                    else:
                        cols = grid.findall('a:gridCol', ns)
                        for ri, row in enumerate(rows):
                            cells = row.findall('a:tc', ns)
                            if len(cells) != len(cols):
                                issues.append(f'{name}: table row {ri} has {len(cells)} cells but grid has {len(cols)} cols')
                            
                            for ci, cell in enumerate(cells):
                                tcPr = cell.find('a:tcPr', ns)
                                if tcPr is None:
                                    issues.append(f'{name}: table cell [{ri},{ci}] missing tcPr')
                                
                                txBody = cell.find('a:txBody', ns)
                                if txBody is None:
                                    issues.append(f'{name}: table cell [{ri},{ci}] missing txBody')
                                else:
                                    paras = txBody.findall('a:p', ns)
                                    if len(paras) == 0:
                                        issues.append(f'{name}: table cell [{ri},{ci}] has no paragraphs')
                
            except Exception as e:
                issues.append(f'{name}: parse error: {e}')
    
    # Check presentation.xml structure
    pres = ET.parse(z.open('ppt/presentation.xml')).getroot()
    sldMasterIdLst = pres.find('p:sldMasterIdLst', ns)
    if sldMasterIdLst is None:
        issues.append('presentation.xml: missing sldMasterIdLst')
    sldIdLst = pres.find('p:sldIdLst', ns)
    if sldIdLst is None:
        issues.append('presentation.xml: missing sldIdLst')
    sldSz = pres.find('p:sldSz', ns)
    if sldSz is None:
        issues.append('presentation.xml: missing sldSz')
    notesSz = pres.find('p:notesSz', ns)
    if notesSz is None:
        issues.append('presentation.xml: missing notesSz')

print('=== Validation Results ===')
if issues:
    for issue in issues:
        print(f'  *** {issue}')
else:
    print('  No issues found!')
print(f'\nTotal issues: {len(issues)}')
