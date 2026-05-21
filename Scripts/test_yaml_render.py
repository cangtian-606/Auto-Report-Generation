import sys
sys.path.insert(0, '.')
from src.yaml_reader import YamlDataReader
from src.generator import DocumentGenerator

reader = YamlDataReader('data/FDD_测试数据.yaml')
ctx = reader.read_context()
print('date keys:', list(ctx['date'].keys()))
print('form keys:', list(ctx['form'].keys()))

companies = ctx['form']['项目公司']
print(f'Companies: {len(companies)}')
for c in companies:
    shareholders = c.get('股东出资', [])
    print(f'  {c["公司简称"]}: {len(shareholders)} shareholders')
    for s in shareholders:
        print(f'    {s["股东"]}: {s["认缴出资额"]} ({s["认缴比例"]})')

gen = DocumentGenerator('templates/FDD模板.docx')
gen.render(ctx, 'output/FDD_YAML测试输出.docx')
print('Render OK')
