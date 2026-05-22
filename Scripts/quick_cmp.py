"""Quick compare xlsx vs yaml context - just key values"""
import sys; sys.path.insert(0, '.')
from src.reader import ExcelDataReader
from src.mapper import DataMapper
from src.yaml_reader import YamlDataReader

for company, yml, xls in [
    ("A公司", "data/FDDv1_A公司.yaml", "data/A公司.xlsx"),
    ("B公司", "data/FDDv1_B公司.yaml", "data/B公司.xlsx"),
]:
    ry = YamlDataReader(yml)
    cy = ry.read_context()
    
    rx = ExcelDataReader(xls)
    mx = DataMapper(rx.read_all())
    cx = mx.build_context()

    diffs = []
    for domain in ['date', 'form']:
        for key in cy[domain]:
            yv = cy[domain][key]
            xv = cx[domain].get(key)
            if xv is None:
                diffs.append("%s.%s: YAML=%r, XLSX=MISSING" % (domain, key, type(yv).__name__))
            elif isinstance(yv, dict):
                for fk in yv:
                    yfv = yv[fk]
                    xfv = xv.get(fk) if isinstance(xv, dict) else '??'
                    if yfv != xfv:
                        diffs.append("  %s.%s.%s: YAML=%r (%s), XLSX=%r (%s)" % (
                            domain, key, fk, yfv, type(yfv).__name__, xfv, type(xfv).__name__))
            elif isinstance(yv, list):
                if len(yv) != len(xv):
                    diffs.append("  %s.%s rows: YAML=%d, XLSX=%d" % (domain, key, len(yv), len(xv)))
                elif yv and xv:
                    for ri, (yr, xr) in enumerate(zip(yv, xv)):
                        for fk in yr:
                            yfv = yr[fk]
                            xfv = xr.get(fk)
                            if yfv != xfv and type(yfv) != type(xfv):
                                diffs.append("  %s.%s[%d].%s: type YAML=%s XLSX=%s" % (
                                    domain, key, ri, fk, type(yfv).__name__, type(xfv).__name__))
    
    print("%s: %d diffs" % (company, len(diffs)))
    for d in diffs[:15]:
        print(d)
    print()
