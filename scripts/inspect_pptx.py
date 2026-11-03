#!/usr/bin/env python3
"""只读清点PPTX及个人格式线索；不修改文件，不代替视觉或事实核验。"""
import argparse
import copy
import hashlib
import json
import posixpath
import sys
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile, BadZipFile

NS = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'pkg': 'http://schemas.openxmlformats.org/package/2006/relationships',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
}


def active_xml(root):
    """使用首个Choice，避免兼容分支重复统计；仍需实际PowerPoint复核。"""
    root = copy.deepcopy(root)
    def visit(parent):
        for child in list(parent):
            if child.tag == '{' + NS['mc'] + '}AlternateContent':
                branch = child.find('mc:Choice', NS)
                if branch is None:
                    branch = child.find('mc:Fallback', NS)
                index = list(parent).index(child)
                parent.remove(child)
                if branch is not None:
                    for offset, element in enumerate(list(branch)):
                        parent.insert(index + offset, element)
                        visit(element)
            else:
                visit(child)
    visit(root)
    return root


def rels(z, owner):
    path = posixpath.join(posixpath.dirname(owner), '_rels', posixpath.basename(owner) + '.rels')
    if path not in z.namelist():
        return {}
    result = {}
    for node in ET.fromstring(z.read(path)):
        if node.get('TargetMode') == 'External':
            continue
        target = node.get('Target', '')
        target = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(owner), target))
        result[node.get('Id')] = {'target': target, 'type': node.get('Type', '')}
    return result


def text(node):
    return ''.join(n.text or '' for n in node.findall('.//a:t', NS))


def paragraph_style(shape):
    paragraphs = shape.findall('p:txBody/a:p', NS)
    aligns, spacings, bolds = [], [], []
    for p in paragraphs:
        if not text(p).strip():
            continue
        prop = p.find('a:pPr', NS)
        aligns.append(prop.get('algn', 'inherited') if prop is not None else 'inherited')
        pct = prop.find('a:lnSpc/a:spcPct', NS) if prop is not None else None
        pts = prop.find('a:lnSpc/a:spcPts', NS) if prop is not None else None
        spacings.append({'proportional': int(pct.get('val')) / 100000} if pct is not None else
                        {'points': int(pts.get('val')) / 100} if pts is not None else {'inherited': True})
        default = prop.find('a:defRPr', NS) if prop is not None else None
        for run in p.findall('a:r', NS) + p.findall('a:fld', NS):
            if not text(run).strip():
                continue
            rp = run.find('a:rPr', NS)
            value = rp.get('b') if rp is not None else None
            if value is None and default is not None:
                value = default.get('b')
            bolds.append(True if value in ('1', 'true') else False if value in ('0', 'false') else 'inherited')
    body = shape.find('p:txBody/a:bodyPr', NS)
    return {'horizontal': aligns, 'vertical': body.get('anchor', 'inherited') if body is not None else 'inherited',
            'line_spacing': spacings, 'bold_runs': bolds}


def inspect(path):
    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    report = {'path': str(path.resolve()), 'sha256': digest, 'slides': [], 'warnings': [],
              'scope': '结构清点和显式格式线索；首个Choice视为激活分支；不验证事实、继承样式或真实渲染'}
    with ZipFile(path) as z:
        failed = z.testzip()
        if failed:
            raise ValueError('ZIP完整性错误：' + failed)
        presentation = ET.fromstring(z.read('ppt/presentation.xml'))
        size = presentation.find('p:sldSz', NS)
        width, height = int(size.get('cx')), int(size.get('cy'))
        report['size_emu'] = [width, height]
        mapping = rels(z, 'ppt/presentation.xml')
        for index, sid in enumerate(presentation.findall('p:sldIdLst/p:sldId', NS), 1):
            rid = sid.get('{' + NS['r'] + '}id')
            part = mapping[rid]['target']
            root = active_xml(ET.fromstring(z.read(part)))
            row = {'page': index, 'part': part, 'slide_id': sid.get('id'), 'hidden': root.get('show') in ('0', 'false'),
                   'native_formulas': len(root.findall('.//m:oMath', NS)), 'texts': [], 'bottom_candidates': []}
            for shape in root.findall('.//p:sp', NS):
                content = text(shape)
                if not content.strip():
                    continue
                row['texts'].append(content)
                for word in ('门控', '小时桶', '第0槽'):
                    if word in content:
                        report['warnings'].append({'page': index, 'kind': 'terminology', 'word': word})
                if '。' in content:
                    report['warnings'].append({'page': index, 'kind': 'sentence_full_stop', 'text': content})
                off = shape.find('p:spPr/a:xfrm/a:off', NS)
                ext = shape.find('p:spPr/a:xfrm/a:ext', NS)
                # 只识别未分组且宽而矮的底部文本框，不据此断言所有底部框均已检查。
                if off is not None and ext is not None and int(off.get('y')) >= height * .8 and int(ext.get('cx')) >= width * .6 and int(ext.get('cy')) <= height * .14:
                    style = paragraph_style(shape)
                    item = {'text': content, '_top': int(off.get('y')), **style}
                    row['bottom_candidates'].append(item)
            # 多个宽文本框位于底部时，优先最下方同一排，避免将来源注释误当总结框。
            if row['bottom_candidates']:
                lowest = max(item['_top'] for item in row['bottom_candidates'])
                row['bottom_candidates'] = [item for item in row['bottom_candidates'] if lowest - item['_top'] <= height * .015]
            for item in row['bottom_candidates']:
                item.pop('_top')
                issues = []
                if item['vertical'] != 'ctr' or any(v != 'ctr' for v in item['horizontal']):
                    issues.append('居中未显式满足或需检查继承')
                if not item['bold_runs'] or any(v is not True for v in item['bold_runs']):
                    issues.append('加粗未显式满足或需检查继承')
                if not item['line_spacing'] or any(v != {'proportional': 1.0} for v in item['line_spacing']):
                    issues.append('单倍行距未显式满足或需检查继承')
                if issues:
                    report['warnings'].append({'page': index, 'kind': 'bottom_style_review', 'issues': issues})
            note_targets = [v['target'] for v in rels(z, part).values() if v['type'].endswith('/notesSlide')]
            row['notes'] = []
            for target in note_targets:
                note = active_xml(ET.fromstring(z.read(target)))
                for s in note.findall('.//p:sp', NS):
                    ph = s.find('p:nvSpPr/p:nvPr/p:ph', NS)
                    if ph is not None and ph.get('type') == 'body' and text(s).strip():
                        row['notes'].append(text(s))
            report['slides'].append(row)
    report['counts'] = {'total': len(report['slides']), 'visible': sum(not s['hidden'] for s in report['slides']),
                        'hidden': sum(s['hidden'] for s in report['slides']),
                        'native_formulas': sum(s['native_formulas'] for s in report['slides']),
                        'bottom_candidates': sum(len(s['bottom_candidates']) for s in report['slides'])}
    if hashlib.sha256(path.read_bytes()).hexdigest() != digest:
        raise ValueError('检查期间输入文件发生变化，请基于最新保存版重试')
    return report


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('pptx', type=Path)
    parser.add_argument('--output', type=Path, help='可选：写出JSON清点报告，不覆盖输入文件')
    args = parser.parse_args()
    if args.output and args.output.resolve() == args.pptx.resolve():
        parser.error('报告路径不能与输入PPT相同')
    try:
        result = inspect(args.pptx)
    except (OSError, BadZipFile, ET.ParseError, KeyError, ValueError) as exc:
        print('检查失败：' + str(exc), file=sys.stderr)
        return 2
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({'counts': result['counts'], 'warning_count': len(result['warnings']),
                      'output': str(args.output.resolve()) if args.output else None,
                      'warning_is_not_failure': True}, ensure_ascii=False, indent=2))
    if not args.output:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == '__main__':
    sys.exit(main())
