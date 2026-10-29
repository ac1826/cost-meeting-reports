"""Cost report generator. Only standard-library OOXML edits; VBA is never executed.

Unchanged ZIP members are copied byte-for-byte. This deliberately avoids an XLSX
round-trip because it cannot retain this template's macros and WPS/Excel objects.
"""
from __future__ import annotations

import copy
import base64
import csv
import datetime as dt
import io
import json
import math
import os
from pathlib import Path
import posixpath
import re
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
N = {'m': NS}
TARGETS = {'PFS生产缩小差异': ('生产缩小差异', 6),
           'PFS 研发降本': ('研发降本', 7), 'MKA产品': ('MKA', 6),
           '观察产品': ('观察', 6), '肉松': (None, 4)}
REQUIRED = ['列表', '本月', '上月', '上上月', '前三月平均', '负责人填写', '降本目标']


class ReportError(Exception):
    """An actionable validation error, safe to display to a workbook operator."""


def tag(name):
    return '{' + NS + '}' + name


def colnum(s):
    v = 0
    for c in s:
        v = v * 26 + ord(c) - 64
    return v


def colname(v):
    s = ''
    while v:
        v, r = divmod(v - 1, 26)
        s = chr(65 + r) + s
    return s


def splitref(ref):
    m = re.fullmatch(r'\$?([A-Z]+)\$?(\d+)', ref)
    if not m:
        raise ReportError('无法读取单元格地址：' + ref)
    return colnum(m[1]), int(m[2])


def rangerefs(ref):
    ends = ref.split(':')
    return (*splitref(ends[0]), *splitref(ends[-1]))


def shifted(ref, delta):
    return re.sub(r'(\$?[A-Z]+\$?)(\d+)', lambda m: m[1] + str(int(m[2]) + delta), ref)


def translate_rows(formula, delta):
    """Translate relative A1 row references, leaving strings and sheet names intact."""
    # Excel strings escape quotes by doubling them. Quoted sheet names similarly
    # escape apostrophes, and must not be mistaken for cell references.
    tokens = re.split(r'("(?:[^"]|"")*"|\'(?:[^\']|\'\')*\')', formula)
    pattern = re.compile(r'(?<![A-Za-z0-9_.])(?P<col>\$?[A-Z]{1,3})(?P<absolute>\$?)(?P<row>\d+)(?![A-Za-z0-9_]|\s*[!(])')
    def replace(m):
        if m['absolute'] or colnum(m['col'].lstrip('$')) > 16384:
            return m.group()
        row = int(m['row']) + delta
        if not 1 <= row <= 1048576:
            raise ReportError('人工填写区的公式移动后超出工作表范围，请调整该公式后重试。')
        return m['col'] + str(row)
    return ''.join(token if i % 2 else pattern.sub(replace, token) for i, token in enumerate(tokens))


def xmlbytes(root, original):
    # ET drops unused namespace declarations, including those named in mc:Ignorable.
    # Re-register and retain all original declarations so Excel never repairs them.
    namespaces = list(ET.iterparse(io.BytesIO(original), events=['start-ns']))
    for _, (prefix, uri) in namespaces:
        if not re.fullmatch(r'ns\d+', prefix):
            ET.register_namespace(prefix, uri)
    result = ET.tostring(root, encoding='utf-8', xml_declaration=True)
    head = re.search(rb'<[^?!][^>]*>', result).group()
    extra = b''
    for _, (prefix, uri) in namespaces:
        attr = ('xmlns' + (':' + prefix if prefix else '')).encode()
        if not re.search(rb'\b' + re.escape(attr) + rb'=', head):
            extra += b' ' + attr + b'="' + uri.encode() + b'"'
    if extra:
        result = result.replace(head, head[:-1] + extra + b'>', 1)
    return result


def numeric(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value)


def norm(value):
    return str(value or '').strip()


def clean_path(value):
    # Windows Explorer's "Copy as path" surrounds the path with double quotes.
    return os.fspath(value).strip().strip('"')


@dataclass
class Result:
    output: Path
    log: Path
    archive: Path | None
    counts: dict
    warnings: list


class Sheet:
    def __init__(self, name, path, raw, strings):
        self.name, self.path, self.raw = name, path, raw
        self.root = ET.fromstring(raw)
        self.rows = {int(r.get('r')): r for r in self.root.findall('m:sheetData/m:row', N)}
        self.cells = {c.get('r'): c for r in self.rows.values() for c in r if c.tag == tag('c')}
        self.strings = strings

    def value(self, address):
        c = self.cells.get(address)
        if c is None:
            return None
        if c.get('t') == 'inlineStr':
            return ''.join(c.find('m:is', N).itertext())
        v = c.find('m:v', N)
        if v is None or v.text is None:
            return None
        if c.get('t') == 's':
            return self.strings[int(v.text)]
        if c.get('t') in ('str', 'e'):
            return v.text
        if c.get('t') == 'b':
            return v.text == '1'
        try:
            f = float(v.text)
            return int(f) if f.is_integer() else f
        except ValueError:
            return v.text

    def keyrows(self, factory='B', material='C', start=2):
        found = {}
        for r in sorted(self.rows):
            if r < start or not norm(self.value(material + str(r))):
                continue
            key = (norm(self.value(factory + str(r))), norm(self.value(material + str(r))))
            if not key[0]:
                raise ReportError(f'“{self.name}”第{r}行缺少工厂。')
            if key in found:
                raise ReportError(f'“{self.name}”存在重复工厂＋物料号：{key[0]} / {key[1]}，第{found[key]}、{r}行。')
            found[key] = r
        return found


_BUILTIN_TEMPLATE_B64 = "UEsDBBQAAAAIAOhcXV3f2CbmxBwAAGSMAgAKAAAAc3R5bGVzLnhtbO1dW4/ryHF+D5D/QOg4xtrYGfEu8pwzc7xzEbLwemN4zyIO9iwONCPODB1dxhJnPWPDL84FMRI/BMlLkvcECBIbeXO8WPu/BLte778IKWlGpMhiX9jdLM5p7WUkUaz+qrqqurq6uvn8xe10YnwSLZbxfHbQs/bNnhHNzufjeHZ50Pvw5XAv6BnLZDQbjybzWXTQu4uWvReHf/xHz5fJ3ST64CqKEiMlMVse9K6S5Pppv788v4qmo+X+/DqapVcu5ovpKEk/Li77y+tFNBovs5umk75tmn5/OopnvTWFp9NzGiLT0eIvb673zufT61ESn8WTOLlb0eoZ0/On717O5ovR2SSFemu5o3Pj1vIXtnG7uG9k9W2pnWl8vpgv5xfJfkq3P7+4iM+jMtywH/ZH51tKKWU+SpbXN+0C77cLTkpufxF9Emfd1zt8PruZDqfJ0jif38yStDu9h++M9Z93xwc91+kZa4Eez8epnF4b3zSevP3kiblvmq+NZ9nHV3s7X3z9hzfz5Nne+s+LF6tvv/Xa6PUr6FsDv9jAR1/bs8zAdD9eEX326q3V31ffAO4e7MB766Ov/d/P/mvP/jgH8xvP8l9vKKYXXhUv7MKGmgwqALtm+PE9WeC2sHTbMDDNj8fp69Xbr4xp+nplrN/fpa9qKoEJdEexL3g6IrCKpFNWUsl99L1o/PGrt9IPUBcEdum+4o3wnU7pzj8Bfunu/DLXAEjdgzW3rLp8IqvV3ZTqs7yyATR2NDintRsOiRRqFRK4J6wRToV0KMTT3/iTw+cX89nWraT2vf7m8Pnyx8Yno0naupX9/nw+mS+MJHVfafurb2ajabT+xZf//fMvf/O71a+uRotlOmysb3Tc7LuVz9v8chqnLnzV/LqNYkvhtqHF5dlBbzg0V69ia+8s4tEk++piNI0nd+sv7RLRs36OBbdI4vNPf/fFL/9+l4YDcQDSDeSQtXbkmxL9/LN/4qMrQwQPNG0JNE0JvFtUirVV46ZNcdOktxU5PDVqf41AVfsmFf9f/PaXf/jsV1/929989ek/VvoM6p7NOadqLyiuqUBdU4NWuFLSYYFUTZCsYjUmJbM9OT2SH9sUmFG+OQX6XQoIWmlOWQcOpDMkVfUU6LVU261vU373l1WuDT7t04F35Mtsk2XWIyi0KTevtDH1sSS7K+ugGODm6CJYUc0RhyZlshXSPNySrZzRqh6VHU2X5sniKLvS9bKqKdV9JVD9qlQPyZxc7hTLMve9UpPO6qWK5RatXX2EABk/T0qUbvT5/f/8/PNP//Pz//33LA/3m3/49p/LHeaM9+LLq4RPYqMfkMYgKQMu4zhIh+E79+uDxl+M/jSKjQ/f5TNUjvCGr8cVBjXMzkdECnl37tiAZL5PJJFtKzlLiPjcnpHE2SrTnrlvu2E4GFjZaxCEzYYLNpUfmMcq5dBakk1B6lNZdKbEKb2Mp9HSeD/6kfG9+XQ022XLakpfmHVLHFEVuxRZ+Q+KAUN1yEYMVCUMU9XcgCZfx83qT1YtEE8mD9UCttNbf3P4/HqUJNFiNkw/GJv3L++uUy5n81m0prP6HeHXl4vRnWV79Dcs55N4nKG4PC7qZqadGZmzzYV4No5uo/FBz1+LL0eRt63dYS3VrzB9OUHo22FgmW7g4UKwFY459ERhyEk8DKW3tuHYa0/mD2wdm9I1bNNWaJ+YniJtHmwlO0hDtDRcsm3HDE3HdH01CPwCgtA2fc8KBp7j266jBoGtpltPB6f+qWyhbto68U+OTxW5I/Mhzg/3UxUKfGdguo7j+4Hpne7JFm5Zj9v3ypklOQPPNQMr8EzfVYQgbN1LmuaROVTkuYLWe3zHcwWWHwRB6DqW66pyKe5xOBy2MVI089OrP2l4eTZfjKPFQ4DpDnr33x0+n0QXSXr/IpvQpH+T+XXWyDxJ5tP0zTgeXc5no0nWwv0d+TuN1a6Gg15yFWfl9IWZzImT/bMCl/100wblHavfruBQ3pD+8h435R3rH3PyOLpJ5vezVxJ3+d8S+Mr/lMRR/rcd4oVXq5qjbKzYAiHI7dFaC+zjMKv+G+EQaoSNydbZYQq2IUFDDx11PO60lGlszVfkIDQBTXkDpXHm7kDEY33HCIYue8zhh1vrwBHFPP1uuWLJ4pxG4/hmejJaXkXjXSXYrj1g4YwSbYf6rGSNgjsEmMtJbq29uQIHApYoowl5PDJHGPmItvv25nTNfVW3fDLfaHMyT9J/8XJXY7VyocuJvwUNnM2jw35Ldt32tEYK7poQtikY/ii56SiKejrJjbxjk0R1Ci4qPSLVCsSgYZlWqsx48LoRqUOEyBjVuL/04XV2E7qwiJo14g33P6TU5/X4y+bAOMxmpxmmiAt7hpNFLrymKXLw2LxZpndHk8kHGb3vXxSLDW8vcucirU7Xmj28jSeTzds1nfWH/s5N6yPS1nd58G3G6Pp6cjecr9a37z+lP91+Olr9evv5nUl8OZtG+Ru+u5gn0XmyOn/PTNGP7n+SHcuXxOdZoeV5+jFaV1reXuxgDbsDdXMy3BqsRd8dHWJRQ5UB1aSCuqs2b/ZdjaTshnLoMnC5OQhyc5+FQqZljIGkvtgcLrnpDZh7dspdGtu6itVy0GHt58OldfCUi5t8ny9wMm4vmtjT5vZBhWHlIwXgfuv+fqcgVWtHjlZZjhYgszRW7Rk/WoyuX0a3a0rZabvx7PLP7kk1cHAPzBLgXM0X8Y9TchmgLJQuIqr2FuuTXlkhvH8zPYsWw9WhoRvaSvkqCp4kWnuLwM4jcHYRFNVgbVzMSpFHuzGlsnGReyaH3+k4/vbkv6st1OgLduHS4t81C9S94nVRqzYnTDNy8Cj6xcZrLTTuvgRfCEQa9z+gRQRFAaUYo6CE1AxXayHcTDULInSCmqH6xgCzomwUUBoCf9Qc5YJqInVXjp+oyUZRQxLWeLU2OSq1qb4xKdpE4K+hNlVTb1Wb6iHJ0CYej08r4Gp6rQq4HpKwxqnmsT6WeezmWSSUs27bZOhAdUwU8jDUPLQncq/7Ioe03GXQcs7AV5rOO807oD2WqnWqwywBRk3gCAd2qjFggGcMYMm8dtkhEUTenvVWY++w9RZ8rLj0NNTLXieGndzsClwueghMLRlmRpX0IpUOkLHbyhN1xCGEqILVkqWdS1kurYrX0K9d1WNK5tWtRlbDqeONmRhzXpCJN9Gibrr6Ri+feltuVb3Ea5Q0s6ZzpIzCV4TWR45WrdnXjKnK0bHLBy6Y6qpNNhx0ICdFvV4jF17j5SQh/UdbNEQdlTRTIcELVELwgQ6+GiFLpKw8BgXi5mpOkMTNgE2j0Q8mmTZO8IvXA78jmPOVQFBeiGdCYdsiiTEjkzo7sRFMlgrJF7J8WPwnFczF/dnyhIkTfddxT5yEd4cCy1EzFRElfEVoqSfAraEddEi2av1ng2mecHTs8oGm7KKmDRLKdiHIAV7IuX4O1RZPimdAnZtngxi0bkvSrVmMnVh+9wyFhBmlpSg2dRF10qAOKx0PuoFSjDWqzc/RBckdWKYlY0eSbpIo8RZ4oZkGlHoAifbQTJxDXA4kn06nT3VIC3AgqVmECAeJ30WT4gV9QreXAMicoPHJgB+rR919+SNJvTPKX9LGCoRhKc9KgIyYl2NlIOdnLRZ87Thay3o0nhZgBburrdcRJCEjpf5YQScUiCqUs3CFcjRTPitEpvSFsJ2zCkaV3yaWGLTgtgF8gKdrGSColdTrETyjPj11ZeoD+cR2a2hovDYSt0GXKurIYEM38Je1BMnIT+XCbas9Hy5ijFEwFwAGakTJ7i2ox2BPpVAIiT3RqIeHUDu8R6QcXgd0Q+W2DIbjvVz5dX00Mx2zPmJ5ZLVURECi96iILe7Koetm0Q9UUaN2yVplDrQDiNi6ENoVRF2LrqraWmWVM8WYgMoXcmAKGTAVHY5KB9Oeh5QcCiDIZ/OncRSkAXm1Hc7HP56Vc4AVlDkq6qI2afrfwJAHFONkJwy52CdCz15pdY1K5d4ENi/EtKbJIu/2JIzY67e3XMWLOO/aiXuMkRTlUM1qlO6x6QhM/kjLYQHdbtQLxV0AC6jHGCjwqj6kGGXgRdZ81D1AWQAY7LpIlMxQrcQ5FopyD8eGFo4enk5hYYhMFOMUsXXKAUuq8EKG9p0jhux1DzISm8MMmnM64ICVBp3QBqtwfq9jEzC3OSVoF7Zo/4xI8I9BecQMOkQWkAi+O6ChBVzE3lGtQ5es0VvQ6PIQIGYp+ixVNfAihpyGWp+BC6bUqRYuxF7AJWSMS0x54UNVV9szWppKv6msqY/qxS5q2WcOK/N4jofXMhFrCy7XLX4FpONKLczhCT/GiagXOAJSN3yUI6T0SFUSPidEAlC6W5PdwdhcGJQ1RzYuU2UC2gwzCw6L3sy1w+LfR6FUEZjRkZ/+ZrkVRfUuI4VqNaopmmegyqJoJHEQWzWZjVJ+kyz9LYFBpvbdHAC/0lTNSltAk/7PM2D5tSKs46AtxIMuyjzPgFfPgKBTGbjhFTTCxYeP0P+C4OFDhEvfRAebIGAHnTwLBhJgg0cYkh6l8OT6aynmLRVy9QzCZMXYbuvKpVETy5PTIdXxh0xmC/uCyXMNhnkwvdNXDIwjAIVhAk9vthlO76nBU7uLW15v7TabZ7LBNFfQNFXAA7NRiJwBRH4WYxFzPNUbkahclEvO2DagbpGLUUkesIY6uUqqAXEQ+jbzKgE6P3FrEFDY7LZXAwmhO2l0tSBYQnoMIl6i3az/hRgG2amGQkLB6pGWSlXMkqp0DQqLlgoYY2SLx4XW37MRYkc+ymsbikihyow3HSnTZnXINKonlluYyvcE5oEG9DgRpVwhFjKwFQNJp3hA0w3ArI8vbKmOt1Vru+w4Swq0FmJAKwg5hldhSQcZYrTK3gCNuVRgQ24vTjsGk9+aD05OSsk6aUUh/Mc85IVMZqS1kIZOe8vhfmvaSzV5Lc3Z0NQHgZgJkB8LRDEGBQaEuINaliHXVzW40vS8jcuerEHYdojKfdQigkFN1JSGIXmIb78HhN014eSkJP2wyagolaPF3AIhPajSjYHrM+X1sXtYNYlh4snoLMTAxR16agVvQuH86pcCBRdWFcARDz0nYBO1BE/SF7IWt9UnfO24jdvJayy036Gx/hcfIQWTK0Zukgu1mZVHdHU450Ieuw63b0XypQKW+6gTi0V+oFRI5QWboWAfEhUsPpNFw+BaKAzRlm7qtgKxifA4QvbluE6pQSklicI9C1M2pHnERjNml6kx7FByoX1dXIZEzv+IpWYKIOa6EDWWUAc6ul4IdQGlNgLV2iUaFUu3EG2EgRhxvsVAC3yOC2F3cSvV9cIfhNNgLyb4PDC3HguHqMRhswmCkgaOsTi3qvm6el+KfAZZMeoaIA+wzehzyoTJtwluj6cSmmLgbCZHEfSFR4HUfpll9cNmF5RZSd6qfnyXWenoOdOSNA9YzEGyCeKjg8TdfH3rTdprYSMLBapqtQP0olrv2tCSoEW9rRZZ9Q5zlm6Ui1qmp+Q3djXdyOTKfYsan5zuY3xKSl6c1Tljs14JeTwqI8aCTLlA8oi6gRLkk292QQu2h8Egf6ouzIKc6Z1KDghTGSGmyIXYxSTzRlrfyhy7ECNRJMB5gySOBMvDJn4cRT4UiGWUJYk4A4OYwpJR48cLynLR9juUl9CaKl/GIVrEYEUCXrUA90UP0EKGNNlx0EIGt3Dj9XCgU8YrZZmDG8bBV0LRCSlTmg8RmdOuAnO+ih/VI/WsW2SQqU6/rXhSD8aicJqxuWuhuicIL49qKMljchy9UJ8AbpjAqq8qbnFhgeEcL5pCHwx7I/jzr1ystAoe9SYVURtuSt5KGiuyN8R2jRNwj0MHGGkvbBITbKh9WODjCfU6ojAKtid2wjahEBVZho5O9k5pIoBa9gOL3l7RjEn027hxyx7KPiCTfR4ycUGPMVuTB/iDm2USX9zxbIMiD5TVuNoTJISYJ4cGPuMUc3AjeyWpQ3mejgoUXeqJe8mnNRFJVzPkqzdy1sO2R1x0xkiFFSngK5uQ7YhRxEjQGf2WzJQy++YdiirO1rDlxSYemYg6O1CSOOHSjHydc5TSRp4mc378007KxFHJ5+PjhGb89zun1kpLu7jLIuQkJpQlyxGCZkiXIzE6xcoh9/nQmFUDOqwBsZzVQhZU0ZMbDIsPdCw/9K8mkkPJwbb8trss+B1hgUGNmkOW+iDTCpGjQ+xXxt8dwVs+bQc33vITwtDhlfMoXmWAsWoEXCAd4B3R4QgV8XIYLGmlmLlhet2AiTd2VipYub5h0EEzs9QqMNuBYOD6voQyQMmQPbSQ4Yo/TMWWNIBlyFgIYLUpuo6UgaqTsts5yIilDI5+3YOMWDHUSll2jCFFzrJHP8ySVlv4KESjIciIjZDsnatzs/gQW+VECxLEkPm5eOd/4KoaYjcH1tzj9RjwNgG0HsOGIG+fBIxu+gdjRjzLBjGLOsZcAmaLVNCpAnI3UNJsBXcpQJc0GOVOcCpWSordXVbKx/lj5KWtPKSQcQgCLMM9SgUs7NEPYqpJC/oNVf0g8DoN0CNwNA3QY/AtqKxTsgeUYp+CMIM7fEIFPkXg0286hdYKUMLN+QuHQbhoqux8MBdlKsjsiFMP/HAVPE5M7uRSM9BC/ArlBJFNm6mGdMJjItRPKgfVjtyuFTRN3IdH/dvVH47nouQRd70jioy5j4exkKW/bIQM4Fa5BnkEG7FL4wun6ObnpRIthNNzajXDZSkuYksRxhiY9ykFLQgVS2VtqOioVrElCErRQksQCNyQsNUUBJovaPEWb400vSKhXLyVkvHkh+l1cYBFpw1iIKtVDcQZCrn1P8hyKjSYEa9GQZAR2x9Yf9c9yNjW/PLHk5FLNTDOcag4cDvPgR1iZoFqLRbn6qbspWMp2ESJsgm43ATXYfDKrSXZaLATBkEU2ImHm+IYWaggy3AJ8rMGXUgaULJS6gCMvFDNa3BpEhVkdNoPVP1UC9pqTWOaoA/Qwrc6XR5Lgx5veSwNeruU7UECn2oLCjJfQzfnwTzloeOg5G4wsUBTI4l6GoSowKe1GbAAwL7sKYUUbAjmwA71FAFDqkO6jQtfEmizrL97uS8exLAGo8zWtarCAiv5FaHlBqhir2MjgOVgHBtCJY8RbLSZpyv5k3Yhi6zv5pq+4ym2hdjiqgjDwxa2GmLx2yLqlbBrDMhfpJFedd+WI5DOWFuuQPoOls73GMRY53sMUkWubHPbjFkUjHVwXCouzhAZY1mb6UIW2K/PAqMETSiZQol5W12HAzTNyozSIkZulEoLWrshSzHdj6xSWEQ9aItizo7IqPYMhKeatZSNAdOZHYNrYoQLJosRHWJEV7VSMjhplQeL+PKqKqfSAHwpkOgSeA8neJoENMEmW18Nl5Efl4lXxjHgMvGWn8CHGy/O9RKV2iAPrQxdaI42FxMXzq0iLD3jyVBDDGyfoas5kFq+YFGv3SDtgOzIR/SLm3jwcok4N4217S76GZiDzjiaGhYI27vbTlzbPN4Gf+I63zfd21FmodgG1A2UjHvF6EEj3ypmdaLctZCGII+u0tMQrDtn8u7xsS5Ykv1+dxYsC0vMj2lAKzJGHaYiMaS2/JZkV4t3QQ6EjPgcHHBxC9FcsbCASATc0N6aJJ9UCpN3bcjmwavQozFL3YO4KDkKdAhLC/XoEJY8FzqEMnsZZcei7EtuUCVX0wRUy0okIvssWaWkQhSjYFIhyvIWCLtaMkhZ3qQpSMK5YVgSpCDkcvEeesjl6Q2eWTv0IJdOnHoDPoy9y+ARnPXOBL7eobSdxOJnBd1eFn5W0O034nal5UxBy67U5QmqqsWvRFto8FJqPhq8lOqNRIVlxFyi5wTSIy6pgGXEW2Ie8EKvFICHVj+80KCnHSpxoqcdHZGMhyrn61in6Fhn5agGBfEolc69Ubl/NpS551jmQPqVO0VsiuBbdOK5eFA19eYAjMVHdKyUdg50lxWrtK8APy+Ij1YWxQqClJOwZ1MiOHRZuOGzVIu1y4r1+FlRt8NTOivqtgk3mPpg1qluq9E9+v55NJl8/2J5+Dx780FyN4mWxvn8JiNrO73c18ZsNI0Oeu9nICf3DGRatHh6E6dvf2JuXnvpf97e/bv1/xxrfe2nWcO7JL/8l99+8Xd/+/tf/XMutDy7iSdJPFtLq+4e4yGz6tJisWmwGPaWskdL2aGi/GB6Pi1dt4buF7/+9R/+469zsVVOdCZ8g2HlR346GB4RRkp1K7aAlq5PRzffJSEt7QEV7a03NGkJB1SEt47KoiUckglvayJtWqrvUFDNC9iiNu0jMuWteKmN9JiCag4ttYme0NDd4qU20VMy3a0uUFvckEx1O5jQ2tvGm9RS3T6LgtbS3DoX/8Uv/urzz36R+r+v/vXBy1sFV+WQbtt2tk1ro26dpy/Qfuhwm9ZM3TpfX6D90EE2rbG6eX/f3w7Mh8/Htxfb4dnNxudxNopnEUeKYj6ZL4zF5dlBbzgMMwPyVwTWV7NI5PD59ShJh/9ZlqYxNu9f3l2nkJfzSTxO6Z1dHufJDIfHg+PTFZncrRnR9Z9V8xqDxqAxaAwag8agMWgMGoPGoDFoDBqDxqAxaAwag8agMWgMGoPGoDFoDBqDxqAxaAwag8agMWgMGoPGoDFoDBqDxqAxaAwag8agMWgMGoPGoDFoDBqDxqAxaAwag8agMWgMGoPGoDFoDBqDxqAxaAwag8agMWDDwEY/uYqyk1PcnpHE2ekm5r4XhqHvOIPBIAh8Owg8ea16rbSqedW8al41r5J5LY8Fw9WxYrmxoOnQIXao0vCkwePSQL8VvW/c6or5s/SXsfFJdvLp+jDMfBO2gHDoQdi6p9rzUAJ/vj7LN3ekr+NlJ/pmJ/Mez8cpt6+NbxpP3n7yxNw3zdfGs+zjq72dL77+w5t58mxv/efFi9W333ptrFrUKqFbba7PaPyacnup5vwBd/ZS7tbR9Ia2GxnDhu5IVB35JonmTWoVQ9KyabDoAINfceSrGvbEBIk6fEWqylo0elxtXwsxOFSNQWPQGDQGjUFj0Bg0Bo1BY9AYNIY3C8ObyLPG0E7f6/a63R4GndUYNAaNQWOQjyH7//LweTI6m0TFB5KbPWMcXYxuJsnLh4sHve3770Tj+GZqP/zqu/En82Tzq+379+LLq8TyH+o00rbi2Ti6jcbHm48p2hxsc/PKbti9sq03KV+B7llfq76SXYPagRBA96zvgtp5TPwEID/raxC2oPJKAN4TgPes76q6crz6B2qn+p4wfVVzGoaO4/uQRI+PKxEcQ3Lz/ey/amoQtuwOqJ2sJTZZw70Na0i9HkB9WqchEKewJkKcwrLOrlTLLbsjDKt7G2onuwPqBUh3svar28l0qvoex8l6FcIGWTB8JQyhK5kuVuuo7wPS8bN/qvsHshLHCcPqK9m1agSOA13JrBG+AiHIMEBXnNVj0Ps749F0cXP/tjAqn7rHm74ufH1yfFT19TA4cod2+esHl9/PtdO/Hxqj2+S9ZbL6a9ws4oPeT06PBuHJ6dDeC8yjYM91Im8v9I5O9jz3+OjkZBiatnn8055xO53Mlk9vLfegd5Uk10/7/eX5VTQdLfen8flivpxfJPvn82l/fnERn0f95fUiGo2XV1GUTCd9OxVhP+xPR/EsjSJSIk+Xk/RXi008sBnfP9h+d9DLfViP8CuWUth57KHtm+94lrk3dExrz/VHwV7gO97e0LPsE989OvWGXg67x4fdMvuWtQXvPU3iaTSJZ1ER/sv8t2kck36sYaJ/3xP9ZXb5g6y1w/8HUEsDBBQAAAAIAOhcXV1Dv6ajBAAAAAIAAAAQAAAAcHJpbnRfbmFtZXMuanNvbquuBQBQSwMEFAAACAgA6FxdXRE18WDsDQAAqDUAABkAAABQRlPnlJ/kuqfnvKnlsI/lt67lvIIueG1svVtbc9vGFX7vr8BwOvWTSewu7pWUkagLqQtBi5Is5Q2mIAljkmBAyJbT6YPTuHEzje22TtrcJpdOmzTTcfqQmSRupr/GlOR/0bMAllwah6R8mWgkEvz23Pbsd84uRXDmtZN2S7nhR70g7MxeIkX1kuJ3muF+0DmcvXQcH1y2Lr0294uZm2F0vXfk+7EC8p3ebOEojrtOqdRrHvltr1cMu34HRg7CqO3F8DI6LPW6ke/tJ0rtVomqqlFqe0GnkFpw2s2LGGl70fXj7uVm2O56cXAtaAXxrcSWMHMS5cy0g2YU9sKDuAhqpfDgIGj6uWiIVor8GwGfdkFpN53qYSeMvGstf7ZwQjSvqZxE8Evhj4GryDkO9mcLv1Gzn8vwRy4nV4SmLwc/vxWh5SNDJpiGtxg2j9t+J07TFPktmGzY6R0F3d5govsXsrcfeTdh7UTGpWkvpiMDe0R7scypdskeWckkXy9hy2sOF5O+4Grqg9WkQ2PsBY0ZA2OsMDeTDNQjBYrCr3ltIEiDIyYMxd61ctgKIyUG8zCgF5Q46MSzBbWo2/DDLNugtkVUzYKh0txMKTM2N7MfwHpzF0rkH8wW5okzv2wlMonITuDf7EnXCrhq+C2/GfvAQ1JQ3gzDdqPpcb4Sokqva5wMrQzlZXstDK9zE1VQVCHortfxlZNGF0pptgAub2WXkLg47K77B3HZb4GBKtEgaK8ZBzf8OujMFq6FcRy2N4PDo7ig9GIvBuwgCt/0O2ngSXx8St1EHKxlsshgaot7mzCcqacxpFG5xIBAe28kWUtfDPLKZylfiwQuJ/UBS7jvH3jHrRjW7GqwHx/NFuyCwDbDmxWfu4PMsSJYDY/jVtDx1/0bfgsUkqSDp2bY6iWPSjvoJGDbO0meb6YmadGyDJNYps5zdIsvkAFumsc9mFDmdmAqNUIzI3RgRNKFFToK9vf91NkkMywzwwZmzKKmEU016NCeqV7YnpbZ0wb2tKKua4Y8NW7uWRty5nJG9cyoPjBKSJExqjIiRUkYnRybkZkxhinLxUYYm2zEzIyYUsKQUNjEjE2crZV5sAYe7HwKtbxRMNDyur2s1mWLdmbRHlg0iqZpMo3JJvXJE+etIaWtKpnJ8dY0ppgZsJ9MyuBUM4L/hErRENVmzxOLYD9hEzNjTjEjSE80aUrPxjLNiCA50ccbIQxZ9xErguNkSHILKehpiREsJ6aUmHytTKOMoDKxpHaQj+bi7YoIKhNbap7PpnqKESqYTFXJSH7Zp5kRTKZESlJudgmrLzg9OujqQ1ZrSGiEIQs4salQQXU6pDpRkUkj285kw4L8dEh+YhYN45k1fm67oh7osB5ovl8/v11RIdSYyMnntitKhg5LhvemzJ6F7Hkj6qJU6LBUiFGkI711qhFRHXRYHXAwyfXoaXaYKBCmyjtujiz8nDHRjqgQRmQ7z5TrVCuDVm8wa0gye8QAP8+lx6z0fYsXe3MzUXhTiRJ7PTghwhtQ4vAZ8Qt1UEDpQS+R4gc5Sov2oK0NTndwBG5yU/MkUddVqCLAgw4nRCOOYDwA3/Fc//t/9O+9NVOKIR6OlJqZ5kKqqdnjNJ88/qr/l9unf73f//2d9PrJD9+cvv8fxFb5wrYQ5cVpymd/+NfpBx/273+PKC8JZXOy8oP3zr7CIl8W+tr4yJ++c38LUV2Zpnr6yTdnX9/rv/f+k/9ioVemzbt/53eltRVEszpV895nMGlEczXTtMal6/x/d3DNtWma/T/eOf30AaK5Pk3zyePHwFJEc0NoWuN83v9b/92vz7/78ezh14h+bZr+6Sf/Pv3kbmrl9O4DeNl/cPf80ReILVfYssdm7sfz7744u/cOolwXC6aOCyRxfnbv29NvHyL6V4Q+Ga8P9dW/fb///aOnX36KmNgU8afv4ROsITBj3Mr8dLt/98/n3312/t3fYZUQs1vCxNji/fxh//6fJpnYzkyYY0s4MfH0w/d4ij5+dPo5luIdEci4FENm+j+9BZXR//jzX3nt7q9PH/7Y/+Dt/n8fnz96hBi8KgySYcJ2BUaH2F6KcV4I6PUUsiXN+axREzaOiyl/nvzwGHIFXfbso7dhouePsKY1vyCMSU7nyxiYNVc4TKRbEuxAg22IPs82RGiRs+QoaF5fCEd2IDrcgYTbBTpsUAIrI9gigi3RYUcX2DIdtlqBrSBYBbFXRbBVOmxJAltDsHUE26DDpiKwGoK5dNgwBFanwz4gsCt0WNsC26RIsdJhsQpsiw6rT2DbdFhOAtuhGTXGltjDz2Cjw6og1ZxS3oji7hTFtLUgintTFT+Cgj17/M35P28j6q9TpPYybFxrmHTAmF+gWK1J4IWPXSPFx54pvuSoZxY15KiHVh1LV1mXolpgWSVKcy8LTKqSxUzXkLAlgclVl2GmxNYVBKukGKFSLNUUGym6PLSWh9bz0EYequUhN4uCSLOqZxiT6y3FDF2uNyEnYQ0xU6nXbyHYtvAr1eqOkJP6wVWBSfHtItgegs1ny21J85hfQLDyCDZCOO0lCachhMswQyacwKQULWoI4TSEcBpCOASrpJgtuahqeb7lobU8tJ6HNvJQLQ+5GSTTLYMkJlzRMrZJsW5mYlIjb2gI2RBsW0PIJuSklbkqMEluF8H2EGw+W+pRsiFYWRtLNv0lyaYjZNMRsukI2XSEbDpCNh0hG4JV9OxkJGFVgUmxrCLYGoKtI9gGgtUQzBWYNN+6niedjpBOqEoUbmTTZVKb2tIR1ukI63SEdTrCOgTbQ7D5bM1HWYdgZX0s64yXZJ2BsM5AWGcgrDMQ1hkI6wyEdQhWMbKUS02immHyXr6KYGsIto5gGwhWQzDXQFhn5FlnIKwTqvJBNpsuk9KyZSCsMxDWGQjrDIR1CLaHYPPZmo+yDsHKxljWmS/JOhNhnZmlTVr9ssCktC2aCOtMhHUmwjoEq2Q+iKRbFX6l+FYRbA3B1hFsA8FqCOaKWKT46maedSbCOqEqH+fMPMO2EGzbRFhnIqwzEdYh2B6CzWdrPso6BCubY1lnvSTrrIwlkrsFK5u6tB2UBSZNfVHoSlNaEpiku5xhprQMKwhWyXxQeYcVfqWlWUWwNQRbR7ANBKshmCtikeKrC0ySu2JltJO2000hJ2+xYr4StoVg20JXquQdC6GdwKTOsItgewg2ny36KO0QrGzlaVeSPsbI7v8J4+xul/TWLX5rS4/fnaMEnZ4fxZVbXT+C98XXU9A7jsPloBX7UfISDA+R7PYhmt4+hN6oZl8WV+kDv2Mtu1ENTLX96DC5t6anNMNjfvMS47c2DeDUwRJzlqxR+cwzc+bRgTJxyhQdYE7ZwAYWiLOAaiwwZwHV4HdNoRqLxFnEB5iziIa7TJxlVGOZOcuoxhJxllCNCnEq6ECVOFV0YJU4q+jACnFW8AHmrKBRrRFnDdVYJ846OrBBnA10oEacGjrgEsdFB+rEqaMDV4hzBR9gzhV0HpvE2UQ1GsRp4APMaejogOE0UB9bxNlCTW0xZwvV2CbONqqxw5wdVGOHOLsEG9hlzi6qcZU5V9GBPebsZR+PDot2bqZ7FHb8OGjWI+gmnZjf5mfw/zdDD5ktdMJy2Mlu703bVNjZD3jr8Vrpv6njoHMo7qcD5taJUqGwvErdVCrMqelK1XZqhChVosOzqVQppwWDZxOebaXKQIqBGAM5DeQ0kNNATgc5HeR0kNNBzgA5A+QMkDNBzgQ5E+QskLNAzgI5C+RskLO5W+5X5Y5V7lkFUTiW8FgYvzL5FYjDngFXVOdXXINxDcY1GMyA6DZMid9/aRqAa1xf4/oa19e4vs71dV2pGA6IW4rLHBee+XZ/sHnc8rNM+ifdyO+ledw/OeBpJvwf7d0oCKMgvsU/7udnBN7Rj1ve3C8rZLbw5Id3+adkyYdLZ/feKcyUxPBMKbV+ES/GqBc6ycvTLz9FvZTQhZ/CB54KTXFNSIiquIQnxoBnG/hBFZfq8MzzBe2AgRgDOQ3kNK4FchrI6SCngxzPqwFyBsgZIGeCnAlysCquCXIWyFkgZ4GcDXI2yNmJX+5Y5Z5V7loFWTh1wBXQER5AHA4ccEU1fsU1WBIq12Bcg3ENjWtoXEPjGjrX0LmGbqKL3YTSqvakJdDlJdBhQULYob04hB25Beu1deR1pFVRX9kK1Am7EBE1OT6ijTKEvSoishEv+kQvr5aIOl9x+DOBdgz+bKAbYNBvXOg1LvQZF3qMC/3Fhd7iQl9xoae40E9cm+uqXBlaiMu7h8sbh8t7hsvbhcs7hcvbgss7AvBBn8oHKmdCs39GPthKHQKvQ9x1CLsOUdch6Do0sLoBmAGYCZgJmAWYDZgNGNQQfwAUKgceKOEPfIDxAcYHND6g8wHdvBDriJwFSx3lg/2qWKeOeCETvbxS1m2SaTzgn1AOY2P8g8OfiQhVvlvynY5vY/AMWx9VYVOGjY/CTgc9ucYoPIMYAznoyTUN5KAn13SQg55c00EOenLNADnoyTUD5KAn10yQg55cs0AOenLNAjnoyTUb5KAn12zuV+WOVe5ZTc4HKt+T+a5L+D7Lzwjwpo9f8X2WJpFyDcY1+EmBaFyDnxWIxjX4aYHoXIOfFy7QkvkHxlIJyon33zj2Wi+a9a536G940SG8E1Na/kHyZQ5Y1ij9l0ByHYfd5Ao2hfR7CuLVke/tJ2/SitApD8IwHrwopYYbfnzcVboehNoI3vSTO99gBn4nTr7wM1vohlEceUH6pYjS4FtXc/8HUEsDBBQAAAgIAOhcXV0e1qzQJQwAAHsxAAAUAAAAUEZTIOeglOWPkemZjeacrC54bWy9m1tz28YZhu/7KzC8qK8scrE4V1KGEnU+i9TBuoNJSMKYJBgQtGV3OlO7deym9SGtkzZ1MknaaZJmOk4vPJPYybR/xpTkf9EFsEsujZei5E6riS3wXXzv7n7fswtI3oy/c9SoK9e9sO0HzYlLZKxwSfGa1aDmNw8mLnWi/cvWpXcmfzJ+IwivtQ89L1LY/c32RO4wilpOPt+uHnoNtz0WtLwma9kPwoYbsY/hQb7dCj23lgQ16nm1UDDyDddv5lIHp1E9j0nDDa91WperQaPlRv5Vv+5HNxMvYXMUZmwafjUM2sF+NMbC8sH+vl/1MqMhWj70rvvxtHNKo+osHDSD0L1a9yZyR0Rzq8pRyP5T2R/Kugqdjl+byP28wL8usz/kcnJFaPqx9/ULMbTsyMAE0+GVgmqn4TWjNE2hV2eTDZrtQ7/V7k20di6/WujeYLUTGZemXUpben5Ee7vMFey8PVDJJF//hZdb7RdTfctq6r1qqn0z+pZmRs+M5ibHk4b1UGGLwlt1GwyQcqxQnbVF7tXpoB6ESsT8WYueUyK/GU3kCmO6zb6oZRuqbZGCZrGm/OR4nrtNjtd8VvC4DyX09idyReIUZ+3knuSWbd+70ZaulVtB0ChX3RhQQgq5/ufVuPp1rsbr9GoQXItDFhixBTbIltv0lKNyi62diZyVU27yS5apKGgte/vRtFdnBgtEZYN0q5F/3VtnMRO5q0EUBY1N/+AwyintyI2Yth8Gt7xmOlCv7lVjUJVWcjtz4/eCxtQr7u2MZh6ejiEd1QZRWVra7yZZij84G4TKyYynKl+LrM0mq4IVrubtu516xAq149eiQ5aDMdM0qUZNZsMbN4Mb817cOcsjHWOZCTpR3W96y951r84imZx0WQ3q7eRvpeE3E7HhHiXfb6TedEzTiFYw4lS2o5tJuTTWXO202fz4AHpeqYvKXdSeizxC7mKwyh36tZqXdnuWH+V+tOenjlGqFiiRRmUWzu2ncT+t56ePqcQyB83edJAzmLHUuaXeszTHLMswB11HjMvgJkbPRMvmjWjq2S4mdzGl2WWHQs/M1plztXgHllTeTAdEo1lbZlF3W22vlvG0uafd8zT6Y9XPnnC8T6TUqijaGBEt6CIaijZHRIvCE/1tRi4qTowzizXKRpScmBJ+hvHmstVG2IjCEusMG/P865aIqhJbWrkZQ7DWBnYTUd74oueS2ZVG2oi9TaWgUoSOWFWq2DTUPiYa2BxN+4KrSRUEqX2CiJFd9xc3FnCpfbiIDjbOCxsL3NQ+buwhk8GWErB4znYWBKqW7PzG3nzh8QoM1T6GhGQTYY0giAoQaeGsB1tsMxAmwKNEnpSua4ZlXqB3tbdXUEsb9WSN3ybSZ3v6ruxG7uR4GNxQwsS4zV5S2A89xIld45cElREn+u+9OLB3rWocUIwjGO42WzlM95txrstRyNp91kM02f3ub92Hd8bzEes1VvJVHjk1KvLVy6+6f7h9/MdH3ffuptevvv/m+MN/Aq/pc3uB4NKo4JPf/P34o4+7j74DwTPnDH784OQrNPLZNF431OEjf33vUQWEzolQMiT0+JNvTr5+2H3w4asf0NDnR5bu7q/yS3MgcoH3rGvDIh9+xiYNIhdFpDEk8vRfd3Hk0qjI7u/uHn/6GEQuj4p89fIloxRErohIa1ifj/7Uff/r0+cvTp58DeJXR8Uff/KP40/upy7H9x+zj93H90+ffQG81kZn7sXp8y9OHt4DwesieCgqSecnD789/vYJiN8Q8cNQYfFsfXVvP+p+9+z1Xz4FFpucNv7DYaKVhW1hWGV+vN29//vT55+dPv8rqxKwrXALzR62/j5/0n30wVkWW9yCDKtSavH64wdxip4+O/4cpXh7lAsL6957jyXq1fe/ZV6v732AE7UzyogP5NNfdt9/enLnxfGXd07/jbaWXZHxYdllter+eIet1e7Tz3/qNlo/O37yovvRr7s/vDx99gwYXhGGar+Ee0Ar8ocCjV9deyLf74kmAVDkG7cMRZFviVYanGdPpt7jSQWPJ2KNGfLvpApk+MNK7e94orspoE0DrQS0GaDNqv0NXWhzan+nFto8iF3g9+lSOhaFZva1JaAtA21FaHZfWwXamhhfoa+ti/ukeWyo/W1AaJsqWNdqf10LraL2F6rQtrhGJG0baDtA21U5TsP2w5Mnn7FnJ8I4jbSGPrGT5Q4C90YEprsVCCwWR0b+mS25k5ffnH55G8VPqWjtpKI9bHmf9dpS5DgTKqW0OCOJ536ZG1igFC9QOnxJ0rS0qoTKVKpZ0mqZzkolESkBOiM0TVqQQpM2ozmhWdKCpDzHkrYgNClLi0BbAtoy0FaAtgq0Na7JC2hdaFIONlLN0KX5bor7pLyUU02TN6UK0LZ4XqikbXONSNoO0Ha5nykxegVoe0ArCgwGnhi86FTeNou8xPwnqQH4tAvDpwH4tCx8WamkAfg0AJ8G4BOaVPB5DcCnAfiAtgS0ZaCtAG0VaGsagE9o0s6/oXH4pEJuagA+DcAHtC0NwMc1IvW7A7RdDcAHtD2gFQUGg/BpCD5tKHz6heHTAXw6T4tMn55NVUkH+OkAPx3gpwP8dF42KXYh1eT3k8WstJSVlrPSSlZazUprXJIGsa4D7nTAHQ+V5lkGeasAbUsH2OkAO6Dt6gA7oO0BrSgAGMROR9jpQ7EzLoydAbAzAHYGwM4A2BkAOwNgZwDsDICdkcUuKy1lpeWstJKVVrPSmpHFzgDYGQA7cZ/88gsSVwHalgG4MwB3QNs1AHdA2wNaURAwyJ2BuDOGcmdemDsTcGcC7kzAnQm4MwF3JuDOBNyZvGzSfQtcK0hjWQTaEtCWgbYCtFWgrQlNmtt6qtlS1TZMQJ8Ilcgog/RVgLZlAvpMQB/Qdk1AH9D2gFYUHAzSZyL6zKH0WRemzwL0WYA+C9BnAfosQJ8F6LMAfRanT/pJfYFrBclvEWhLQFsG2oroQ3qfXAXamgXos7L0WYA+ESrvfSB9FaBtWYA+C9AHtF0L0Ae0PaAVBQeD9FmIPmsofTagTwX/RKJEh3712lQwAKLNgZD6mrJ5YSRIprkm/0qmJGLlX0EJTZrlrNAkwOa4RiW/edGHNJYFoUmVWQTaEtCWgbYCtFWgrQlNmtu60CSwN2xOolTATXGfTKKYrzS3CtC2hCb1u801IvW7A7RdockkAm0PaEUBAjFlEm1Eop0lMS/90x0/ZhVE/JBRekQuPlHUjg9FKX6z7YXR/M2WF9b95rVUdDtRMOvXIy9MPjLjvsJPaanpKS14INC+LK7Sv+KTgfxAILNqeOFBcqSprVSDTnxGTFNzkpx2sEucPTJ4f/98mAobqFO0UcM0caZhxLTuTFuoYYY4MzBihjozsI8p4kzBiCndmYJ9lIhTghEl6pRgH7PEmYURs9SZhRHzxJmHEQvEWYANi8RZhA1LxFmCDcvEWYYNK8RZgQ2rxFmFDWvEWYMNc8SZww3UmYMzXyfOOozYIM4GbqDOBrTaJM4mjCgTp4wbqFPWYYPhlGEfFeJUoFWFOhUYsUWcLRixRZ0tGLFLnV3YcIU6V2DDHnX2YEMxXmt4sW0TZxuOa4c4O7Bhmzrb0GqHOjtpQ76/Z0yOtw6Dphf51fWQbWbNKD7cacS/5mdb2ESuGUwHTX6KO90lg2bNj3c+t56eg4z85kH/FCU1nM3kUKppKGxRrPHr+IG8v9mpe9y2yrpeaOeU2tF+3B+x2L7cCv0g9KOb8W/uWbUDtoO6UcB2zLrXblcO3SYziXfbTt2dLIznxeV4PnVOzlmAsY0Ycgz2eQdpDA6SyIP03u249f/JCDfJyIHp8sCIpv3/0leKK14SFY/3YKKzbybbi1WlpGrsu62UKNMpa6aWU9KoUtJ0p6QXlJKusu8m+247JUNTSgbzMolSMtmWbVpKySo4JYv5WczPZn4287OZH/uZLu407rDALAlhnoQwU6IyV6Kq8ZUZXzFjQrU3E1jrtOp+1Y28bbfe8eRManImqclfZ4elpuUeeCtueMAe+krd20+OZzOSw/SdNLmOglZyxYzSk8ji06Hn1pL3gfgHqf0giHof8qlx2Ys6LaXlskKW/Vte8ibMBuY1o+QM/0SuFYRR6Prpsed873+kmPwPUEsDBBQAAAgIAOhcXV3enl/gOw0AAIUyAAANAAAATUtB5Lqn5ZOBLnhtbL1bW3PbxhV+76/AcDr1k0XuLrC4VFJGd1I3UHc5bzAFSRiTBAtCtpxOH5zGjZtpbLd10uY2uXTapJmO0wfPJHEy/TWmJP+LngWwxNI4JBXXU40lgt+e79vds99ZgNJ68rXTVlO76UfdIGxPXSETlSua326EB0H7aOrKSXx41bry2vTPJm+F0Y3use/HGsS3u1Ol4zjuOOVyt3Hst7zuRNjx29ByGEYtL4a30VG524l87yAhtZplWqnwcssL2qVUwWk1LiPS8qIbJ52rjbDV8eLgetAM4tuJlpQ5jQoyraARhd3wMJ4AWjk8PAwafmE0RC9H/s1ATLuktRpO7agdRt71pj9VOiW619BOI/hH4ZtBV5FzEhxMlX5dyb6uwje5mlwRPX3b//qNHFpxZMgE0+HNh42Tlt+O0zRFfhMmG7a7x0Gn25/owaX0DiLvFqydzLgy7fm0pa9H9JfLXMUu2wMrmeTrf9DyGvli0pdcTaO/mjQXYy8pxvtirDQ9mTTUIw2Kwl/3WmCQLYHoBNpi7/pc2AwjLQZ9aDFKWhy046lSZcKw4YtZNqe2RSq6BU3l6clypjY9eRDAgos+tMg/nCrNEGdm0UpikpDdwL/VVa61N8KwtdXwhEEJqZTy9+ti9ZsZKur0ehjeEJQaOLYCg+x4bV873epA7UyVoIvb2SVkKg47q/5hPOc3QaBGOAzSa8TBTb8OnKnS9TCOw9ZmcHQcl7Ru7MWAHUbhG347Hajf9BvCqFonCQe1LBZpTLVEbyOaM3o6hnRUGzZ0/askRxu2s0F0JY9iluq1TNhiUhCwZgf+oXfSjGGN9oKD+HiqBGIZthneqvqiO8gcm4BchCdxM2j7q/5NvwkEgJOeGmGzm/zUWkE7AVveafJ6K5WkE5bFTWKZhsjRbbFAHLppnHRhQlm3falUhGYitC+icGGFjoODAz/tbJQMy2RYX8ac0HXYkDjN9czKpfX0TE/v6+kTdGBiQuxFBTVvBUkjkzT6knyC88EREkZHj4tnIrwvwoo5J4yNVjEzFVPJFmO0wgh9QWVEukZO1sp6sPo9kEqxC1MvyoJE0+t0/YOCpp1p2sqoDUPnA6tijJ652BhS01JlGQoZNPkYGek3wkbKmGNkpM2IrszJNE2ms58iI61FVG/RFzyB5HpARFqL5N6yCgYdmxhpLZJ7y0BsPm6ZpH+IpRRgMb+X3yCIdA+xRyqSSytSaSVxMVpxpIzcRilRMlasaTomZbS/k+bO1nUk9cgKjixlKr1Oc68TZKd/CWVpf5rbn5EJUrEH3I/cQUbLynKgeTkQGymrnywsS4RyZcWLt5mfrCuLhppKhvt6FnKnGaDLYqGWukAvrvxYGVkhNK8QgtxdxukwWRcsrwtCiukXd/eROrIwGFF1XnDHWBXa39aYldvMHhAQT1Hpw0368cCLvenJKLylRYleF57L4HMeccSMxEWl7/T08SqJSh6fKDxkNgRthiShRgWKBhqCtlj+rTiC9gD6iad73/69d//NyXIMfQuk3MiYsylTt4cxnz39svfnO2d/edD73d30+tl3X5+9929Ea+7SWgh5fhz5/Pf/PHv/g96DbxHygiSbo8kP3z3/Ehv5ouTrw0f+/O0H2wh1aRz17OOvz7+633v3vWc/YEOvjpt37+5vyytLCLM2lnn/U5g0wlzOmNawdF385y7OXBnH7P3h7tknDxHm6jjms6dPwaUIc00yrWF9Pvhr752vLp58f/7oK4S/Po5/9vG/zj6+l6qc3XsIb3sP7108/hzRcqWWPTRz3188+fz8/tsIuS4XrDJsIEnn5/e/OfvmEcLfkHwynA/11bvzoPft4+dffIJIbMrxp5+KE2xLYnzYyvx4p3fvTxdPPr148jdYJUR2W0oMLd7PHvUe/HGUxE4mYQ4t4UTi+QfvihR99PjsMyzFu3Igw1IMmen9+CZURu+jz37htTq/PHv0fe/9t3o/PL14/BgR3FMEZcL2JUZz7FqKCV9I6PUUskkOzWQbNWHDvJj659l3TyFXsMuef/gWTPTiMbZpzcxKMaXTmTkMzDZXeHRIbz9wt+nfcuhL3nJofsuR/czSfEeS2ByCzSPYAs23cIkt0nxvldgSglURvRqCLdN8D5LYCoKtItgazXcRia0jmEvzHUJidZoXvsQ2aF7MEtukSHXSvDoltk3zcpPYDs3rR2K7NPPC0Jp69Cnc2TDbj2Um1Ygw91Pm0J0g3UwQ4jXar6AhxA+hRM+ffn3xjzsI/XWKVFuGDdsMRj1SzMxSrLoU8NIPWgPlxgbLTQqzdPUMpbNZllWYMqU5iSnun8+4XMEWJKZWU4aZiguXEKyaYoQqY6ml2EAxFaGVIrRahNaK0HoRcrNREGVW9Qxjah2lGDeUmW7KOCVzW3Kmyqa9jWA7WYbVjOzKOGUsewi2L7lKv9ckpqvGTEH1pjIzN4ANeEbHPaMjnskwrnpGYsos53XEMzriGR3xDIJVU8xWuqjpRcsUoZUitFqE1orQehFyM0h1TAYpG+qGnhlG2bQ3szBlj93SEb8g2E6KGWomd2Wc0u0egu3riF+QuJlsqQf9og/1i4H7xUD8YiB+MRC/GIhfDMQvBuIXBKsa2YOGgtUkpoxlGcFWEGwVwdYQbB3BXIkp860bRd8YiG8kVXHhVjZdpsRtG4hxDMQ4BmIcBNs3EOMgcTPZmg8axxhqHI4bhyPG4YhxOGIcjhiHI8bhiHEQrMqz24JSqrUMU2+Kywi2gmCrCLaGYOsI5nLEOLxoHI4YR1LVJ71sukxJyzZHjMMR43DEOAi2zxHjIHEz2ZoPGocPNY6JG8dEjGNmM1cWcE5iysznTcQ4JmIcEzEOglWzPojCrcl+lfEtI9gKgq0i2BqCrSOYK8eijK9uFo1jIsaRVPXRxiyaZBvBdkzEOCZiHATbNxHjIHEz2ZoPGsccahxr0Dha9mdh8fH4OGjcmA2Tz55SyEoHod41Zi3ETxbipxTTbSWZCxJTuIsSUxK3hGBVC/GTxJS4ZQRbQbBVBFtDsHUEczOMKotTt4p+sjI/KVbctBA/WcgdDMF2suXgSh+7WZyppHQPwfYtxE9I3Ey25oN+sop+Kiu/TM/OeoRxdtIhPacjjjV0xckMLWh3/Siu3u74EXx+u5GC3kkcLgbN2I+StyCcI9lREZoeFUFPJdlX5VX6QxxPyk4lgVTLj46ScxVdrRGeiIMqkEgFTjvYJc4+GYzPD6lQtIE5MxbWMEucWZQxy5xZjjXMEWcOZcwxZw5lzBNnHmXMM2ceHdUCcRZQxgJzFlDGInEWUcYicxZRxhJxllDGEnOWUEaVOFWUUSNODW1YJs4y2rBCnBW0YZU4q2jDGnHW0IZ14qyjDS5xXLShTpw62rBBnA28gTkbaEo2ibOJMraIs4U3MGfLQBu4s4X2sU2cbVRqmznbKGOHODsoY4c5Oyhjlzm7aMMec/bQhn3m7KMN15hzLfvrXV7N05Od47Dtx0GjHsE2047F2S8ufpEOm8tUqR3Ohe3skGe6f4Xtg0DsSV4z/c1qHLSP5DGrKoWF1eqmVicMvm2tTg2tzgh8A6YDpgNmAMYB44CZgJmAWYDZgNmCWxHkCqCEAAw3BfEjERUNTDToosEgWpWYplNPTsmZXEAijIswLsJMQ6syB4K45jLHhVdxJz7cPGn62QT9007kd9PpHZweitkT8au/ThSEURDfFk+G05Ni/z1petM/r9Kp0rPv3hF/bUn+SHF+/+3SZFk2T5ZT7cv0QdU+rFF9PP/iE7SPMroW45YISoxcKglEHSCjAyMkryoLlYFO2KhOXmkaNgs5aEA11Lr50MSxBGVo8MQZwt3Wi0O4uzZhHtvHXlsZbuWVDa3GnHVDq9nOOiFajRjwamo1KjZTBq8mvNpajUEUgzAGcTrE6RCnQ5wBcQbEGRBnQByHOA5xHOJMiDMhzoQ4C+IsiLMgzoI4G+Js0a3otyI6roieKxAKj1ZiLExcmeIKwgmFeALVDVeCwQSDCQYTDF0wdMHQBUMXDEMwDMEwBIMLBhcMLhimYJjsMs4Ux5jylTEGPMNekTHFHyHyPvioPl6pL8UupWuuCXtVRXOJ2LNg6yI27KtUc6kBrxZsZXADZRDGIE6HOF2wIE6HOAPiDIgzII5DHIc4DnEmxJkQJ7ZCE+IsiLMgzoI4G+JsiLOTfkXHFdFzRXRdgVh4VocrsCL8gHDYkeGK6uJKMFgyVMFggsEEQxcMXTB0wTAEwxAMQzC4YPBkdoLBBcMUDNNAd+lCcXJ1ef5/tekaIg9inLAYDL5tWATAoAJdqD4XKs+FqnOh4lyoNhcqzYUqc6HCXFtwK4IMReWKenJFKbmiilxRQK6oHVeUjSsqxhXF4oo6cUWJuEh1FNJiDOxZrzItHe/IX/OiI/jkoTX9w+SgOtRIlJ6DTq7jsJNcwSDSM9ny3bHvHSQfSiZgSIdhGPfflFPhLT8+6WgdD0a7FbzhJ+eNYBJ+O07+N8NUqRNGceQF6QHwcv+/lEz/F1BLAwQUAAAICADoXF1dgQS6J6sMAAAeNgAAEAAAAOinguWvn+S6p+WTgS54bWzdW1tz28YVfu+vwHA69ZNFLhZYgKikjO6iLqRESrasN5iCJIxJggEhW06nD07rxs00tts66S2ZXDpt0kzH6YNnkriZ/hpTUv5FzwK7xMI4JCXH1UM1pgF+e657vj1YicvJN47bLe22F/b8oDN1hUyUrmhepxns+Z2DqStH0f5V+8ob0z+avBOEt3qHnhdpIN/pTRUOo6jrFIu95qHXdnsTQdfrwMh+ELbdCN6GB8VeN/TcvVip3SrqpRIrtl2/U0gsOO3meYy03fDWUfdqM2h33ci/6bf86G5sS5o5DnNm2n4zDHrBfjQBasVgf99verloiFEMvds+T7ugtZtO5aAThO7NljdVOCaG29SOQ/inw4uCq9A58vemCj8riZ+r8CJX4ztiJm8HPz+XoeUjQxJMwpsPmkdtrxMl0xR6LUg26PQO/W5vkCiP6tVyLZWL5aLbTKdMf8U5MwdzpqfG6CsaYwNjtDA9GQ9shBpQz6u6bShDgyOUwVjk3pwLWkGoRWAfRsyCFvmdaKpQmjDL8EPtMtPLNikZNgwVpyeLwtr05J4P08p9aKG3P1WYIc7Moh3LxCLXfO9OT7nX3gqCdqPptrwqr1FrqkBKpYLG2X8zCG5xkQrwoARBdd2Opx03usDIqQKYvCtuYWaioLvm7UdzXgsMVCAktxn5t70N0Jgq3AyiKGjX/YPDqKD1IjcCbD8M3vI6SVhey2vy4mvdWBxsCVlkMLHFfY0YFupJDElM2wSC6r0ZT0l8P5gznqF6LydnMaYs1GfP23ePWhHU47q/Fx1OFcoFidWDO8sedwazRidgHoKjqOV3vDXvttcCBYBjT82g1Yv/19p+Jwbb7nF8vZOY1Cdsm1nEtniU0V2+Jhm4aR71IB3hdmAqMaILI/rAiKIL1Tn09/a8xNkoM1SYoQMz1oRhEKPE9NSeVTq3PUPYMwb2jAk9kxg39rIFdd5yJk1h0lRCtCyLGlS1OiYuJoywgRH2clz6aAuWsGClEz5BqV6iRJ0pOnKmRuZpCwf2wAEpIR6MvFkw0XK7PW8vZ7MsbJYHNm1k7szRmZOSJKw+0gwbY0ZyjRgjzVhjzEg+EPaDkpL1JNZI6o8zI6tGbIUXpFTOxHL+9UhkxUhaMnpuluuyUrraWy68WPRBc0kLrtupOtKaRjJbl5XXqcptxl6a6wvbld1GTxlFMC5c2LDkmJ42HYI06Ysbln1IV8mr/0Cjksq6pUQ7sGcjDTejLims26M67Fgzkrl6ylx4KubmbJwdKjnMb6Qdko+HP+RG2pGPWUpUOy8tzbFW9EG3oXbKsnLGAN9MJM/4eOsw70bu9GQY3NHC2F4PNifwCwRxeCTpZiFxO9hAwA6ryRVmuAY0kRKsFcD9Di98Iwph3AcP0XT/67/1H749WYzAK0eKTaE5m2ga5WGaL55/3v/9vZM/POr/6n5y/+KbL0/e/xdia+7cthDl+XHKp7/+x8kHf+o/+hpRXjin8uP3Tj/HIl8U88f04ZF//86jLUR1SaqSIaonH355+sXD/nvvv/g3FvryuND7939RXF1CNCtjNR9+DEkjmitC07aGaJ795z6uuTpOs/+b+ycfPUY018Zpvnj+HFiKaK5LTXuYz0d/7L/7xdmzb0+ffIHoV8fpn3z4z5MPHyRWTh48hrf9xw/Onn6K2KqNn7lvz559evrwHUR5Q1DFHEqV2Pnpw69OvnqC6G9K/WEFB31YX/17j/pfP/3+s48QE3UZf/LrTIw1JMaGVea7e/0Hvzt79vHZs79ClRCzW9LE0PX3yZP+o9+OMrGdmLD1NLBrIl8e7JCu9rT/3dtA8/5fPvmJ2+7+9OTJt/0Pftn/9/Ozp08RH9elQZY62ZGYlWI3RCw0hXYFZKTQjOi66lzOiDYmPBShoQ+6up7t6lJDT1u3tDKrpytbYnMINo9gCwi2qKftTWJLetq3JLaM6FYQbEVPV4HEVhFsDcHW9XQ1SqyKYDVEd0PEbCp5bOrpopBYXUdYrqcsl9iWntJWYtsJprq9JqBh/eP0ycfwfMD4pg94NGJVIIo7QrE8ckUiijfGeHzx3Z9haZw+//Ls7/cQ9V0dYXmCWcMW4ajn8oxgLd+/nXtfklk0FNkK6aUJ3Ry+F6IJIfhvoIMFlWCGSvY5iSlkmpeYuqAkphBnUfiwFIItSbmysqASrKy4qAhIXU55aDUPreWh9TxUzUM1ASlF3RCQEv5mAjGmTFtdiCmZN0SWVFkfWwi2LVTVZSQnTZmg6wi2IzD1UXADkduVmK2ylQ4IN8DmEExUuqzqLmTkMjQ0Lk5DQ/R1xcOsgdDQQGgoMaVkCxJTZnRR+LCUAi0JOUvJdtnI09DI0zAPreahtTy0noeqeahm5Glo5GloCBoqU1Q38jSUWSopbSHYtlBVaHNNTJpKh+sItiMxJYcbiNyugdHQQGiIYKLSWRoaQ2loXpyGJkJDE6GhidDQRGhoIjQ0ERqaCA0TLP6wYMBDiSmxrCDYKoKtIdg6glURrCYxJd8NM09HE6GjVFWI0RDpUmWat0yEj2aejybCRwTbMRE+InK7JsZHE+EjgomSZ/loDuUjuzgfGcJHhvCRIXxkCB8ZwkeG8JEhfGQIHyWm+FhBsFUEW0OwdQSrIliNIXxkeT4yhI9SVd3tMuQxzRA+sjwfGcJHBNthCB8RuV2G8ZEhfEQwUfIsH9lQPloX56OF8NESE6rwZ05iyoTOWwgfLYSPFsJHC+Gj9KHEUpGYUqEVBFtFsDUEW0ewKoLVBKbuojesPB8thI9SVcEaVp57Wwi2beX5aCF8RLAdC+EjIrdrYXy0ED4imCh5lo/WUD7a5+ejFh36zVuzQYactiCnMh2zNkJOGyFnghlqqAsSU3+VET4sRXdJyimVWbYRcgqMKHIrCLaKYGsIto5gVQSr2Qg57Tw5bUFORbVuI+S0kYc3gm1LXYVh12yEnQi2YyPsROR2bYydNsJOBBM1z7LTzrOzqHzkIA6BBJE4FJEcu+EnIHr8CIfmd3peGC3f7Xoh/PJ+KwHdoyhY9FuRF8ZvwXCKiDMkenKGBD0UVL4q75L/+OkgcSgITLW98CA+gtHTmsERP8FC+dmXAZweUtGz8mKAOjM2NjBHnDlUY446cwwbmCfOPKoxT5151McscWZRjVnqzKI+FomziGosUmcR9bFAnAVUY4E6C6jGMnGWUY0l4izhA9RZQk1ViFNBNVaIs4IOrBJnFR1YI84aOrBOnHV0oEqcKjpQI04NHdggzgY6sEmcTXyAOpto5teIs0OwgTpx6qipBnEa+AB1GiY6wJwG6nyLOFuoqS3qbKEau9TZFR8zpgtqerJ7GHS8yG9uhLDSOxE/p8X/sBXB+p4qdIK5oCOOOSYtJOjs+bwtuK3kbFPkdw7kmagKr4ZWsUynGh8qs5i2CW/q4g1/iO3Xj1qeMN6EACq9grZ3vM+9En4opRv6QehHd/mGqaAF0GPcKICe0vJ6va1DtwM2eD86arnTpcmivJ0sJobjT1CRAMfFDQGbWoVA+ITA1YKrpVV0CulQuJbhWtYqFMQoyBkgZ4CcAXIGyJkgZ4KcCXImyDGQYyBngZxFXs7aO+6GkEw8oYPMqZq5bitZ/njZnCq8+OZd/nlU/DHO6cN3Cljm453oGSflUU6+/+wj1MkrTW+NOjVDq5WdGilpNWLClWk1na9QHa4WXG2tRkGKghgFOQPkDJAzQM4EORPkTJAzQY6BHAM5BnJWaSynSDZp69JIVbMgVwqvMuRoQl4EXoAZgBmAmYAxwBhgFh2bRymbR/kS83BqNk8FymbwdKBcJZ4SXBlPC8qm89TgavP0oGwGTxHKVeJpwpXxVKFsOk8XrjZPGcpnjEub/+lcTdtW0/bePHJb/4c529mcL68PwkLdOEd8lhof5X8ruXh8Y7sVfwapXrI9kb6mnsjMrJfyKC+vtSlunOvBwDKPRMqMTHzkdc0CzXoxR3kZNgujOaNnPVwepzfOx+nMY4JQ43L6zCbVNsvaJjG1TZ3Ay9I2KUAUMAMwEzATMAYYK4/NIfOIoJf4qKvnuPxybGamjxvm5T2+6tSpj6+/mem5Br28rSfsHQ2tUuY7Zth58p0zgx0n30jDThr2RlXYG1Vgb1SFvVEF9kZV2BtVYG9Uhb1RBfZGVdgbVWBvVIW9UQX2RlXYG1Vgb1TN742wpW9mmrn+P2myZraVv84W23UPvHU3PPA7Pa3l7cffbYGMwuRvaPF9FHTjO+j0yRc75LtDz92L/1wxAeXeD4Jo8KaYGG540VFX67rAhIb/lhef14QkvE4Uf81oqtANwih0/eRbJMXBd72m/wtQSwMEFAAACAgA6FxdXRNlwL1dCQAAYiQAAAoAAADogonmnb4ueG1srVpbc9vGFX7vr8DwIX4SyQUWSxIVlRFJUXeJInVxHiFyKWFMAgywtGR38uCkbtxOY7mNk97csTOdNq6n4+mDZ3yJ2/4ZU1L+RReXBVbGIelY5ugCfAfn27Nnv3MIcjH78XG/p1ynrmc5dvkKyuavKNRuOx3LPihfGbLuTPHKx3M/mz1y3GveIaVM4dfbXjlzyNjAyOW89iHtm17WGVCbW7qO2zcZP3UPct7ApWYncOr3cmo+T3J907IzIYPRb78LSd90rw0HM22nPzCZtW/1LHYj4BI0x26Kpm+1XcdzuizL3XJOt2u1aSoahHMuvW75084o/baxfGA7rrnfo+XMMcJmWzl2+Y/KfzU+lGsMrU4584si0RfmK4TMFPK4NoNLhYWZeVJUZ4o1nNeqWK+XqvOfidDSkQETDMOrOe1hn9osTJNLe3yyju0dWgMvnqgf1fvNNV/KlXJmO0mZ+p450+OcqQmZ9p5kJCbTMnOzgaHhKlx6dMPs82Vo+YhvYuZ+1ek5rsI4PTfoGYVZNitn8lm9xF9asUTUUhHlcZGbcnOzuYhsbrZj8az6Qygu7ZYz88iYr5PgmuCSXYseedKxctNx+q226asAoXwmOd/wl6wXoX4x7DvONd9lmcsiz4McmDZVjlsDLtByppRRbkSHPFHMGazRLqvSHidY0XiMZptZ12mDu5Qz+w5jTr9pHRyyjOIxk3Gs6zo3qR3GSXu07YtBGQSXc7LoWsAYcvmDTTBH7mEMYVC1Ah/60yBFtYJR8+cYp9GfpHws8lUPRMxXrEO75rDH+BLtWR12WM4Us4VCQcNagc80MjadoyXqj8szqGWLGcUZsp5l0zV6nfa4J4eDIdtOzwv+Kn3LDsC+eRz8Pwq5cRZjhPNE1f1k3fAXCvMFPbQ6HRo6tIcen2QUSswa8qkRnyrxFYukgIoFia8wmUSLSLSYRJ6wRPKOQeGID0/kK74znx7x6TGfBvO9TSKvSYqVRKwkZkV5IHelybEVIpZCzKKnF1TzV3QSSzFiKUosqVD0/MSMTZxsKRqgJE02IQbYuGfPHHi0k6LyHUMN52MyAkSrTp4yikshkV0hq+uYXGDhxn3qsbrfeabIBAndITyJEf8ERqE8lEivlCXk4ur6jBNZhNIQkRKWEsnUhAmpoYJU7G8HA5XBBRYhNVSUqikt+2k0QlBIUhRQ5br+/ppVhdDURGgIiPVSYwgVqgisjMtQx705ac4qyaJ86YMlSDRuNSkhpH7gBImiUrG8COqHG0DUmJrUGEKpNJHJVTu+Xami+lQy6e1IJ2+5iWpTC/K8U91kSmNX43IjWnHy+2EhvEUJbxiCG5Kaycy5Wdc5UtyA2OM3PfyDCjJ81ujOQ+V3am3fPu9fwAMqchs3Wbaf8hZzud3ihGzu9MWjH299PZtjfBAfybUjz8o0z9Hzv4/ufg54Vqd5vnn1/ejrW6d/OBn96nZ4/ObFk9Nv/g1w1d6ZC3BemOZ89ut/nn77p9HJc8C5/o7O9746+x6KfDH0J3l1fOQ/fnmyDbguCVdt3KI9eHL2+O7oq2/e/ACFvjx16W5/kVtdBDxXpnrefcgnDXiuRp4lPMbz/L+3Yc+1aZ6j394+/es9wHN9muebV6+4SgHPDZFhNG7Mkz+OfvP4/NnLs/uPAf/Naf6nD/51+uBOyHJ65x4/Hd27c/70O4CrMT1zL8+ffXd290vAeUssWGlcCl7fGt35/fmzh+fP/sbTAVA0xfj5cUJ/dH908rtJFC1BMTadz5+OXn/OxTP6y6OPzP7g56f3X46+/eXoh1fnT58ChNuCUA0bp4/tANiuwLQE2wsxDUvXXRWYdN0nkW8hgebjphd2Xd5k406rTuu0alI6gq8CYFUAqwHYAoDVAWxRTfqMwJbUpIEIbBnwXQGwVTWRo8DWAGwdwDbUpCwEtglgDcB3S02ULLCmmkhTYC01Wd4JcoVEJTzHqfTs/kPelAHPnWljhlUGeO4Kz7Et6vWfeQmc/+8/5y+eAP5700f2/c9ePTn/xy3A/6oKCD/K6tgETnhXnBcyJz/hruBCHWlAHan5rFa68BL3T/G3KXGZaZEkpCoTkC5VmcCIVGUCkyp+QWBylQlMGmNRi1QsKXFJYJKKlwUmFVkIySuwmobW0tB6GtpIQ5tpqBFBcnWJuKR8NAUm5a0VYUiq1m1xnZSjHQDbFZiUtz3BJyXkqhZpUhr3Ey2WVdyM5wGsIvjSDRpfUlg4LSwMCAsDwsKAsDAgLAwICwPCwoCwcFpYOC2sNLSWhtbT0EYa2kxDDZwWlohLmnsTA8LCacFsA9gOgO1iQFgYEBYGhIUBYQFYBY8Vln5JYelpYemAsHRAWDogLB0Qlg4ISweEpQPCijDp9mElgmRhpaG1NLSehjbS0GYaakSQLCwRq9yxdEBYOiAsANsBsF0dEJYOCEsHhKUDwgKwij5WWOSSwiJpYRFAWAQQFgGERQBhEUBYBBAWAYQVYVIoKxEkRbKahtbS0Hoa2oggaQabaaiRdtwSscodiwDCIoCwAGwHwHYJICwCCIsAwiKAsACsQtLCyknfEEWbgw6L9sbC3Vh/I8zzt/IUy/aoy5ZuDKjLb+KuhaA5ZE7d6jHqBqecOEGivUU13FuM94rz0WuG/6KZ8KgensavzwKqPnUPgp04T2k7Q39nU8MZCQ4HaGlGi1y8PjRsa8Y2aNjRjB3QsKsZu6BhTzP2QMOWbmyBhhYydhFkqGhGBfSoakYVNNQ0owYaFjRjATTUNaMOGhY1YxE0LGnGEjxBzdjCkKGpGU3Qo4GMhgpSIWMLNDSR0QQNi8hYBA1LyFgCDcvIWAYNK8hYAQ2ryFgFDWvIWAMN/m45aKggowIaqsiogoYaMmqgYQEZC6Chjow6aFhHxjpo2EDGBmjYRMZmaMgllTY3Ozh0bMqsdsPlLcBm/ka+3w4ZL/xyxnaqjh09FhP2FsfuWH6/MHvhzjez7AOxac6T3ggePygQ/02o2xz2aETU5oMtexmlc9z1R0D+7uHAtRzXYjeCD30O7zMmc3hf6VHP2z40bU7h96Rhz5zLz+bEIf/EGPAGX3oDsfiPIBzQddM94A1M6dFu8IAE7+Vu+PYYHDNnEBzxtho+DCDODqnZCXpblofUdRwWn+RC4hZlw4EyMHm0LesmDfZG+SSozYJHVcqZgeMy17TCJw9y8fNCc/8HUEsBAhQAFAAAAAgA6FxdXd/YJubEHAAAZIwCAAoAAAAAAAAAAAAAAIABAAAAAHN0eWxlcy54bWxQSwECFAAUAAAACADoXF1dQ7+mowQAAAACAAAAEAAAAAAAAAAAAAAAgAHsHAAAcHJpbnRfbmFtZXMuanNvblBLAQIUABQAAAgIAOhcXV0RNfFg7A0AAKg1AAAZAAAAAAAAAAAAAACAAR4dAABQRlPnlJ/kuqfnvKnlsI/lt67lvIIueG1sUEsBAhQAFAAACAgA6FxdXR7WrNAlDAAAezEAABQAAAAAAAAAAAAAAIABQSsAAFBGUyDnoJTlj5HpmY3mnKwueG1sUEsBAhQAFAAACAgA6FxdXd6eX+A7DQAAhTIAAA0AAAAAAAAAAAAAAIABmDcAAE1LQeS6p+WTgS54bWxQSwECFAAUAAAICADoXF1dgQS6J6sMAAAeNgAAEAAAAAAAAAAAAAAAgAH+RAAA6KeC5a+f5Lqn5ZOBLnhtbFBLAQIUABQAAAgIAOhcXV0TZcC9XQkAAGIkAAAKAAAAAAAAAAAAAACAAddRAADogonmnb4ueG1sUEsFBgAAAAAHAAcAsAEAAFxbAAAAAA=="


def add_builtin_templates(parts, infos):
    """Add missing output sheets, remapping all template style IDs safely."""
    raw_workbook=parts['xl/workbook.xml']
    workbook=ET.fromstring(raw_workbook)
    sheets=workbook.find('m:sheets',N)
    present={s.get('name') for s in sheets}
    missing_sources=[name for name in REQUIRED if name not in present]
    if missing_sources:
        raise ReportError('缺少必要来源表：'+'、'.join(missing_sources)+'。只需提供7张来源表，结果表由程序自动创建。')
    missing=[name for name in TARGETS if name not in present]
    if not missing:return
    if not _BUILTIN_TEMPLATE_B64:raise ReportError('程序内置模板缺失，请使用完整的新版app.py。')
    with zipfile.ZipFile(io.BytesIO(base64.b64decode(_BUILTIN_TEMPLATE_B64))) as assets:
        template_styles=ET.fromstring(assets.read('styles.xml'))
        styles_raw=parts['xl/styles.xml'];styles=ET.fromstring(styles_raw)
        order=['numFmts','fonts','fills','borders','cellStyleXfs','cellXfs','cellStyles','dxfs','tableStyles','colors','extLst']
        def section(name):
            node=styles.find('m:'+name,N)
            if node is None:
                node=ET.Element(tag(name),{'count':'0'})
                rank=order.index(name)
                before=next((x for x in styles if x.tag.split('}')[-1] in order and order.index(x.tag.split('}')[-1])>rank),None)
                styles.insert(list(styles).index(before) if before is not None else len(styles),node)
            return node
        num_map={}
        formats=section('numFmts')
        codes={x.get('formatCode'):int(x.get('numFmtId')) for x in formats}
        next_id=max([163]+[int(x.get('numFmtId')) for x in formats])+1
        for x in template_styles.findall('m:numFmts/m:numFmt',N):
            original=int(x.get('numFmtId'));code=x.get('formatCode')
            if code not in codes:
                child=copy.deepcopy(x);child.set('numFmtId',str(next_id));formats.append(child)
                codes[code]=next_id;next_id+=1
            num_map[original]=codes[code]
        formats.set('count',str(len(formats)))
        offsets={}
        for name in ['fonts','fills','borders','cellStyleXfs','cellXfs','dxfs']:
            destination=section(name);offsets[name]=len(destination)
            source=template_styles.find('m:'+name,N)
            if source is None:continue
            for original in source:
                child=copy.deepcopy(original)
                if name in ['cellStyleXfs','cellXfs']:
                    for attr,kind in [('fontId','fonts'),('fillId','fills'),('borderId','borders')]:
                        if attr in child.attrib:child.set(attr,str(int(child.get(attr))+offsets[kind]))
                    if 'numFmtId' in child.attrib:
                        number=int(child.get('numFmtId'));child.set('numFmtId',str(num_map.get(number,number)))
                    if name=='cellXfs' and 'xfId' in child.attrib:
                        child.set('xfId',str(int(child.get('xfId'))+offsets['cellStyleXfs']))
                for fmt in child.findall('.//m:numFmt',N):
                    number=int(fmt.get('numFmtId'));fmt.set('numFmtId',str(num_map.get(number,number)))
                destination.append(child)
            destination.set('count',str(len(destination)))
        parts['xl/styles.xml']=xmlbytes(styles,styles_raw)
        relpath='xl/_rels/workbook.xml.rels';relraw=parts[relpath];rels=ET.fromstring(relraw)
        relnamespace=rels.tag.split('}')[0]+'}'
        content_raw=parts['[Content_Types].xml'];content=ET.fromstring(content_raw)
        content_namespace=content.tag.split('}')[0]+'}'
        used_rids={r.get('Id') for r in rels}
        max_id=max(int(s.get('sheetId')) for s in sheets)
        used_codes=set()
        print_names=json.loads(assets.read('print_names.json')) if 'print_names.json' in assets.namelist() else {}
        for name,raw in parts.items():
            if re.fullmatch(r'xl/worksheets/[^/]+\.xml',name):
                pr=ET.fromstring(raw).find('m:sheetPr',N)
                if pr is not None and pr.get('codeName'):used_codes.add(pr.get('codeName'))
        for index,name in enumerate(missing):
            raw=assets.read(name+'.xml');root=ET.fromstring(raw)
            for c in root.findall('m:sheetData/m:row/m:c',N):
                if c.get('s') is not None:c.set('s',str(int(c.get('s'))+offsets['cellXfs']))
            for row in root.findall('m:sheetData/m:row',N):
                if row.get('s') is not None:row.set('s',str(int(row.get('s'))+offsets['cellXfs']))
            for col in root.findall('m:cols/m:col',N):
                if col.get('style') is not None:col.set('style',str(int(col.get('style'))+offsets['cellXfs']))
            for rule in root.findall('m:conditionalFormatting/m:cfRule',N):
                if rule.get('dxfId') is not None:rule.set('dxfId',str(int(rule.get('dxfId'))+offsets['dxfs']))
            pr=root.find('m:sheetPr',N)
            if pr is not None and pr.get('codeName'):
                code=pr.get('codeName')
                if code in used_codes:pr.attrib.pop('codeName',None)
                else:used_codes.add(code)
            number=1
            while 'xl/worksheets/sheet'+str(number)+'.xml' in parts:number+=1
            path='xl/worksheets/sheet'+str(number)+'.xml'
            ridnumber=1
            while 'rId'+str(ridnumber) in used_rids:ridnumber+=1
            rid='rId'+str(ridnumber);used_rids.add(rid)
            max_id+=1
            ET.SubElement(sheets,tag('sheet'),{'name':name,'sheetId':str(max_id),'{'+REL+'}id':rid})
            if print_names.get(name):
                definitions=workbook.find('m:definedNames',N)
                if definitions is None:
                    definitions=ET.Element(tag('definedNames'))
                    before=next((x for x in workbook if x.tag in [tag('calcPr'),tag('extLst')]),None)
                    workbook.insert(list(workbook).index(before) if before is not None else len(workbook),definitions)
                for definition in print_names[name]:
                    ET.SubElement(definitions,tag('definedName'),{'name':definition['name'],'localSheetId':str(len(sheets)-1)}).text=definition['text']
            ET.SubElement(rels,relnamespace+'Relationship',{'Id':rid,'Type':REL+'/worksheet','Target':path[3:]})
            ET.SubElement(content,content_namespace+'Override',{'PartName':'/'+path,'ContentType':'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'})
            parts[path]=xmlbytes(root,raw)
            info=zipfile.ZipInfo(path);info.compress_type=zipfile.ZIP_DEFLATED;infos[path]=info
        parts['xl/workbook.xml']=xmlbytes(workbook,raw_workbook)
        parts[relpath]=xmlbytes(rels,relraw)
        parts['[Content_Types].xml']=xmlbytes(content,content_raw)
        if 'docProps/app.xml' in parts:
            raw=parts['docProps/app.xml'];props=ET.fromstring(raw)
            vt={'v':'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'}
            for vector in props.findall('.//v:vector',vt):
                if vector.get('baseType')=='variant':
                    children=list(vector)
                    for i in range(0,len(children)-1,2):
                        heading=children[i].find('v:lpstr',vt);number=children[i+1].find('v:i4',vt)
                        if heading is not None and heading.text in ['Worksheets','工作表'] and number is not None:number.text=str(len(sheets))
            title=next((x for x in props if x.tag.split('}')[-1]=='TitlesOfParts'),None)
            vector=title.find('v:vector',vt) if title is not None else None
            if vector is not None:
                for i,name in enumerate(missing):
                    node=ET.Element('{'+vt['v']+'}lpstr');node.text=name;vector.insert(len(present)+i,node)
                vector.set('size',str(len(vector)))
            parts['docProps/app.xml']=xmlbytes(props,raw)


class Generator:
    def __init__(self, source):
        self.source = Path(clean_path(source)).resolve()
        if self.source.suffix.lower() != '.xlsm':
            raise ReportError('请选择包含7张来源表的 .xlsm 工作簿。')
        try:
            with zipfile.ZipFile(self.source) as z:
                self.parts = {x.filename: z.read(x) for x in z.infolist()}
                self.infos = {x.filename: x for x in z.infolist()}
        except (zipfile.BadZipFile, OSError) as e:
            raise ReportError(f'无法读取文件，请确认文件存在、未加密且可访问。\n{e}') from e
        add_builtin_templates(self.parts,self.infos)
        ss = self.parts.get('xl/sharedStrings.xml')
        strings = [''.join(x.itertext()) for x in ET.fromstring(ss)] if ss else []
        self.workbook = ET.fromstring(self.parts['xl/workbook.xml'])
        rels = {x.get('Id'): x.get('Target') for x in ET.fromstring(self.parts['xl/_rels/workbook.xml.rels'])}
        self.sheets = {}
        self.sheet_order = []
        for sh in self.workbook.find('m:sheets', N):
            dest = rels[sh.get('{' + REL + '}id')]
            path = dest.lstrip('/') if dest.startswith('/') else posixpath.normpath('xl/' + dest)
            self.sheets[sh.get('name')] = Sheet(sh.get('name'), path, self.parts[path], strings)
            self.sheet_order.append(sh.get('name'))
        missing = [s for s in REQUIRED if s not in self.sheets]
        if missing:
            raise ReportError('缺少必要来源表：' + '、'.join(missing))
        self.changed = {}
        self.deleted = set()
        self.warnings = []
        self.archived = []
        self.counts = {}
        self.indices = {s: self.sheets[s].keyrows('A', 'B') if s == '前三月平均' else self.sheets[s].keyrows()
                        for s in ['列表', '本月', '上月', '上上月', '前三月平均']}
        self.validate_headers()
        self.month = self.period_month()
        self.old_manual = {}

    def warn(self, text):
        if text not in self.warnings:
            self.warnings.append(text)

    def validate_headers(self):
        specs = {'列表': {'A': '渠道', 'B': '工厂', 'C': '物料号', 'D': '成品名', 'E': '产量'},
                 '前三月平均': {'A': '工厂', 'B': '物料号'}}
        for s in ['本月', '上月', '上上月']:
            specs[s] = {'B': '工厂', 'C': '物料号', 'D': '成品名', 'E': '实际入库数量'}
        for name, cols in specs.items():
            for col, expected in cols.items():
                if norm(self.sheets[name].value(col + '1')) != expected:
                    raise ReportError(f'“{name}”的{col}1应为“{expected}”，当前表结构与模板不同。')
        for name in ['本月', '上月', '上上月']:
            for col in 'FGHIJKLMNOPQ':
                if not self.sheets[name].value(col + '1'):
                    raise ReportError(f'“{name}”缺少成本列：{col}1。')

    def period_month(self):
        def parse(value):
            text = norm(value)
            if re.fullmatch(r'\d{4}|\d{6}', text) and 1 <= int(text[-2:]) <= 12:
                return (2000 + int(text[:2]) if len(text) == 4 else int(text[:4]), int(text[-2:]))
            dated = re.search(r'(\d{4})\s*[年/-]\s*(\d{1,2})(?:月|\b)', text)
            if dated and 1 <= int(dated[2]) <= 12:
                return int(dated[1]), int(dated[2])
            m = re.search(r'(\d{1,2})月', text)
            if m and 1 <= int(m[1]) <= 12:
                return None, int(m[1])
            return None
        raw = self.sheets['本月'].value('AE1')
        current = parse(raw)
        if current is None and norm(raw):
            raise ReportError('“本月”AE1的期间无效，请填写YYMM或YYYYMM，例如2610或202610。')
        if current is None:
            current = parse(self.source.stem)
            if current:
                self.warn('本月AE1未标注期间，已根据文件名确定月份。下次建议在本月AE1填写如2609的期间。')
        if current is None:
            raise ReportError('无法确定报告月份。请在“本月”AE1填写期间，例如2609表示2026年9月。')
        self.year, month = current
        for name, offset in [('上月', 1), ('上上月', 2)]:
            previous_raw = self.sheets[name].value('AE1')
            previous = parse(previous_raw)
            if previous is None and norm(previous_raw):
                raise ReportError(f'“{name}”AE1的期间无效，请检查。')
            expected_month = (month - 1 - offset) % 12 + 1
            expected_year = ((self.year * 12 + month - 1 - offset) // 12) if self.year else None
            if previous and (previous[1] != expected_month or
                             (self.year and previous[0] and previous[0] != expected_year)):
                label = f'{expected_year}年{expected_month}月' if expected_year else f'{expected_month}月'
                raise ReportError(f'期间不一致：“{name}”AE1应为{label}，请检查年份和月份。')
        return month

    def has_variance_amount(self, name):
        return name != '肉松' and re.sub(r'\s+', '', norm(self.sheets[name].value('Q1'))) == '成品偏差额'

    def read(self, sheet, address, numeric_only=False, context=''):
        value = self.sheets[sheet].value(address)
        c = self.sheets[sheet].cells.get(address)
        if value is None or value == '' or (c is not None and c.get('t') == 'e') or (numeric_only and not numeric(value)):
            self.warn(f'{context}：来源“{sheet}”{address}缺少有效' + ('数值。' if numeric_only else '数据。'))
            return None
        return value

    def source_value(self, sheet, key, col, numeric_only=False):
        r = self.indices[sheet].get(key)
        if r is None:
            self.warn(f'{key[0]} / {key[1]}：未匹配到“{sheet}”，对应数据留空。')
            return None
        return self.read(sheet, col + str(r), numeric_only, ' / '.join(key))

    def manual_cols(self, name):
        if name == '肉松':
            return ['S', 'T', 'U', 'V', 'W']
        if name == 'PFS 研发降本':
            return ['Q', 'R', 'V', 'W', 'X', 'Y', 'Z', 'AA']
        return ['Q', 'R', 'V', 'W', 'X', 'Y', 'Z']

    def capture_manual(self, name):
        sh = self.sheets[name]
        size = TARGETS[name][1]
        keycol, faccol = ('Y', 'AA') if name == '肉松' else ('AC', 'AE') if name == 'PFS 研发降本' else ('AA', 'AC')
        result = {}
        for start in range(3, max(sh.rows, default=2) + 1, size):
            key = (norm(sh.value(faccol + str(start))), norm(sh.value(keycol + str(start))))
            if not all(key):
                # Visible identifiers are a fallback for manually trimmed helper columns.
                key = (norm(sh.value(('B' if name == '肉松' else 'A') + str(start))),
                       norm(sh.value(('E' if name == '肉松' else 'D') + str(start))))
            if not all(key) or key[1].startswith('#'):
                continue
            if key in result:
                raise ReportError(f'“{name}”原结果中存在重复产品块：{key}，无法安全保留人工内容。')
            saved = []
            for offset in range(size):
                for col in self.manual_cols(name):
                    if col == 'Q' and self.has_variance_amount(name):
                        continue
                    address = col + str(start + offset)
                    c = sh.cells.get(address)
                    if c is not None and sh.value(address) not in (None, '') and c.find('m:f', N) is None:
                        saved.append((col, offset, copy.deepcopy(c), sh.value(address)))
                    elif c is not None and c.find('m:f', N) is not None:
                        # Existing broken lookup placeholders are not user input.
                        f = c.find('m:f', N).text or ''
                        if '#REF!' not in f and '[' not in f:
                            saved.append((col, offset, copy.deepcopy(c), '=' + f))
                        elif '#REF!' not in f or not f.startswith('IFERROR('):
                            self.archived.append([name, *key, col, offset + 1, '=' + f])
                            self.warn(f'“{name}”{address}人工区公式存在失效或外部引用，已另存记录，请检查。')
            result[key] = (start, saved)
        return result

    def set_cell(self, rows, address, value=None, formula=None, cache=None):
        col, row = splitref(address)
        node = rows[row]
        c = next((c for c in node if c.get('r') == address), None)
        if c is None:
            c = ET.SubElement(node, tag('c'), {'r': address})
        for child in list(c):
            c.remove(child)
        c.attrib.pop('t', None)
        c.attrib.pop('cm', None)
        c.attrib.pop('vm', None)
        if formula is not None:
            ET.SubElement(c, tag('f')).text = formula.lstrip('=')
            value = cache
            if value is None or value == '':
                c.set('t', 'str')
                ET.SubElement(c, tag('v'))
            elif numeric(value):
                ET.SubElement(c, tag('v')).text = repr(value)
            else:
                c.set('t', 'str')
                ET.SubElement(c, tag('v')).text = str(value)
        elif value is not None and value != '':
            if numeric(value):
                ET.SubElement(c, tag('v')).text = repr(value)
            else:
                c.set('t', 'inlineStr')
                t = ET.SubElement(ET.SubElement(c, tag('is')), tag('t'))
                t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                t.text = str(value)
        return c

    def link(self, rows, address, source, key, col, number=False, divide=None):
        value = self.source_value(source, key, col, number)
        r = self.indices[source].get(key)
        if r is None:
            self.set_cell(rows, address)
            return None
        ref = f"'{source}'!{col}{r}"
        if number:
            formula = f'IF(ISNUMBER({ref}),{ref}' + (f'/{divide}' if divide else '') + ',"")'
        else:
            formula = f'IF({ref}="","",{ref})'
        if divide and numeric(value):
            value /= divide
        self.set_cell(rows, address, formula=formula, cache=value)
        return value

    def diff(self, rows, dest, actual, standard, values, ratio=False):
        a, b = values.get(actual), values.get(standard)
        formula = f'IF(AND(ISNUMBER({actual}),ISNUMBER({standard})' + (f',{standard}<>0' if ratio else '') + '),'
        formula += f'({actual}-{standard})/{standard},"")' if ratio else f'{actual}-{standard},"")'
        val = ((a - b) / b if ratio else a - b) if numeric(a) and numeric(b) and (not ratio or b != 0) else None
        values[dest] = val
        self.set_cell(rows, dest, formula=formula, cache=val)

    def rate(self, rows, dest, cost, price, values):
        a, b = values.get(cost), values.get(price)
        val = (b - a) / b if numeric(a) and numeric(b) and b != 0 else None
        self.set_cell(rows, dest, formula=f'IF(AND(ISNUMBER({cost}),ISNUMBER({price}),{price}<>0),({price}-{cost})/{price},"")', cache=val)
        values[dest] = val

    def lookup(self, rows, address, material_ref, material, source, keycol, valuecol, start=2):
        sh = self.sheets[source]
        end = max(20000, max(sh.rows, default=1) + 1000)
        keys = f"'{source}'!${keycol}${start}:${keycol}${end}"
        vals = f"'{source}'!${valuecol}${start}:${valuecol}${end}"
        match = f'MATCH({material_ref},{keys},0)'
        index = f'INDEX({vals},{match})'
        formula = f'IF(COUNTIF({keys},{material_ref})=0,"",IF({index}="","",{index}))'
        matches = [r for r in sh.rows if r >= start and norm(sh.value(keycol + str(r))) == material]
        if len(matches) > 1:
            raise ReportError(f'“{source}”物料号{material}有多条匹配记录，无法确定负责人／目标。')
        value = sh.value(valuecol + str(matches[0])) if matches else None
        if not matches or value in (None, ''):
            self.warn(f'{material}：“{source}”{valuecol}列尚未填写，关联结果留空，后续填写可更新。')
        elif isinstance(value, str) and value.startswith('#'):
            self.warn(f'{material}：“{source}”{valuecol}列存在错误，关联结果需检查。')
        self.set_cell(rows, address, formula=formula, cache=value)

    def build(self, name):
        sh = self.sheets[name]
        root = copy.deepcopy(sh.root)
        data = root.find('m:sheetData', N)
        category, size = TARGETS[name]
        if name == '肉松':
            keys = [key for key in self.indices['本月'] if key[0] == '蚌埠二厂肉松'
                    and '推广包' not in norm(self.source_value('本月', key, 'D'))]
        else:
            keys = [key for key, r in self.indices['列表'].items() if norm(self.sheets['列表'].value('A' + str(r))) == category]
        self.counts[name] = len(keys)
        manual = self.capture_manual(name)
        self.old_manual[name] = manual
        for key, (_, cells) in manual.items():
            if key not in keys:
                for col, off, _, val in cells:
                    self.archived.append([name, *key, col, off + 1, str(val)])
        rows = {}
        for child in list(data):
            data.remove(child)
        for r in (1, 2):
            row = copy.deepcopy(sh.rows[r])
            data.append(row)
            rows[r] = row
        # First product block owns the reference style and merge structure.
        merges = [x.get('ref') for x in sh.root.findall('m:mergeCells/m:mergeCell', N)]
        header_merges = [m for m in merges if rangerefs(m)[3] <= 2]
        block_merges = [m for m in merges if rangerefs(m)[1] >= 3 and rangerefs(m)[3] < 3 + size]
        # RD identity and family fields in the old template start midway through a
        # block. Normalize their merges to the full product block, preserving style.
        if name == 'PFS 研发降本':
            for col in ['B', 'C']:
                block_merges = [m for m in block_merges if rangerefs(m)[0] != colnum(col)]
                block_merges.append(f'{col}3:{col}{2 + size}')
        all_merges = header_merges[:]
        values = {}
        for i, key in enumerate(keys):
            start = 3 + i * size
            old_start = manual.get(key, (3, []))[0]
            for offset in range(size):
                source_row = sh.rows.get(old_start + offset, sh.rows.get(3 + offset))
                if source_row is None:
                    raise ReportError(f'“{name}”缺少完整产品版式块，请使用保留五张结果表模板的完整工作簿。')
                row = copy.deepcopy(source_row)
                r = start + offset
                row.set('r', str(r))
                row.attrib.pop('hidden', None)
                for c in row:
                    if c.tag != tag('c'):
                        continue
                    c.set('r', re.sub(r'\d+$', str(r), c.get('r')))
                    for child in list(c):
                        c.remove(child)
                    for attr in ['t', 'cm', 'vm']:
                        c.attrib.pop(attr, None)
                data.append(row)
                rows[r] = row
            all_merges.extend(shifted(m, start - 3) for m in block_merges)
            self.populate(name, rows, start, key, values)
            for col, offset, original, _ in manual.get(key, (0, []))[1]:
                address = col + str(start + offset)
                restored = self.set_cell(rows, address)
                for child in original:
                    restored.append(copy.deepcopy(child))
                if original.get('t'):
                    restored.set('t', original.get('t'))
                f = restored.find('m:f', N)
                if f is not None and f.text:
                    f.text = translate_rows(f.text, start - manual[key][0])
                    f.attrib.pop('t', None)
                    f.attrib.pop('si', None)
                    f.attrib.pop('ref', None)
        # Keep one empty styled block when a category has no products, allowing
        # next month's workbook to add products without losing the template.
        if not keys:
            for offset in range(size):
                r = 3 + offset
                source_row = sh.rows.get(r)
                if source_row is None:
                    raise ReportError(f'“{name}”缺少产品版式块，请使用完整模板。')
                row = copy.deepcopy(source_row)
                row.attrib.pop('hidden', None)
                for c in row:
                    for child in list(c):
                        c.remove(child)
                    for attr in ['t', 'cm', 'vm']:
                        c.attrib.pop(attr, None)
                data.append(row)
                rows[r] = row
            all_merges.extend(block_merges)
        last = 2 + size * max(1, len(keys))
        for row in data:
            row[:] = sorted(row, key=lambda c: splitref(c.get('r'))[0] if c.tag == tag('c') else 99999)
        # Every merged child must be empty; otherwise Excel drops content on open.
        for merge in all_merges:
            c1, r1, c2, r2 = rangerefs(merge)
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    if (c, r) != (c1, r1):
                        self.set_cell(rows, colname(c) + str(r))
        # Merge normalization can materialize previously absent blank cells.
        # OOXML requires cells in ascending column order, including those blanks.
        for row in data:
            row[:] = sorted(row, key=lambda c: splitref(c.get('r'))[0] if c.tag == tag('c') else 99999)
        oldmc = root.find('m:mergeCells', N)
        if oldmc is not None:
            oldmc.clear()
        else:
            oldmc = ET.Element(tag('mergeCells'))
            insert_after = root.find('m:autoFilter', N) or root.find('m:sheetProtection', N) or data
            root.insert(list(root).index(insert_after) + 1, oldmc)
        oldmc.set('count', str(len(all_merges)))
        for ref in all_merges:
            ET.SubElement(oldmc, tag('mergeCell'), {'ref': ref})
        maxcol = 32 if name == 'PFS 研发降本' else 28 if name == '肉松' else 31
        root.find('m:dimension', N).set('ref', f'A1:{colname(maxcol)}{max(last, 2)}')
        af = root.find('m:autoFilter', N)
        if af is not None:
            af.set('ref', f'A2:{colname(maxcol)}{max(last, 2)}')
            for child in list(af):
                af.remove(child)
        # Existing full-column colour rules are retained; block-specific rules
        # are expanded from the first product so newly added rows are included.
        for cf in root.findall('m:conditionalFormatting', N):
            if name == 'MKA产品' and any('INDEX($N:$N,ROW()-MOD(ROW()-3,6)+3)' in (f.text or '')
                                        for f in cf.findall('m:cfRule/m:formula', N)):
                root.remove(cf)
                continue
            refs = cf.get('sqref', '').split()
            newrefs = []
            for ref in refs:
                c1, r1, c2, r2 = rangerefs(ref)
                if r2 >= 1048576:
                    if r1 <= 2:
                        newrefs.append(f'{colname(c1)}{r1}:{colname(c2)}{max(last, r1)}')
                elif r1 <= 2 and r2 <= 2:
                    newrefs.append(ref)
                elif r1 >= 3 and r2 < 3 + size:
                    newrefs.extend(shifted(ref, size * i) for i in range(len(keys)))
                elif r1 <= 3 and r2 >= size + 2:
                    newrefs.append(f'{colname(c1)}{r1}:{colname(c2)}{max(last, r1)}')
            if newrefs:
                cf.set('sqref', ' '.join(dict.fromkeys(newrefs)))
            else:
                root.remove(cf)
        if name == 'MKA产品' and keys:
            self.mka_conditional(root, last)
        # Keep frozen headers/identity columns, but show the first generated
        # product instead of the template's previously scrolled position.
        for view in root.findall('m:sheetViews/m:sheetView', N):
            view.set('topLeftCell', 'A1')
            pane = view.find('m:pane', N)
            x, y = 0, 0
            if pane is not None and pane.get('state') in ('frozen', 'frozenSplit'):
                x, y = int(float(pane.get('xSplit', '0'))), int(float(pane.get('ySplit', '0')))
                pane.set('topLeftCell', colname(x + 1) + str(y + 1))
            for sel in view.findall('m:selection', N):
                p = sel.get('pane', '')
                address = colname(x + 1 if p in ('topRight', 'bottomRight') else 1) + str(y + 1 if p in ('bottomLeft', 'bottomRight') else 1)
                sel.set('activeCell', address)
                sel.set('sqref', address)
        for elname in ['rowBreaks']:
            for el in root.findall('m:' + elname, N):
                root.remove(el)
        self.changed[sh.path] = xmlbytes(root, sh.raw)
        return last

    def populate(self, name, rows, s, key, values):
        meat = name == '肉松'
        rd = name == 'PFS 研发降本'
        source = '本月' if meat else '列表'
        factory, material = key
        fc, mc, nc, qc, pc, familyc = ('B', 'E', 'F', 'G', 'H', 'D') if meat else ('A', 'D', 'E', 'F', 'G', 'C')
        def put(col, value):
            self.set_cell(rows, col + str(s), value)
            values[col + str(s)] = value
        put(fc, factory)
        put(mc, material)
        for dest, srccol, number, div in [(nc, 'D', False, None), (qc, 'E', True, 1000)]:
            values[dest + str(s)] = self.link(rows, dest + str(s), source, key, srccol, number, div)
        # List H/I are lookup formulas; resolve with the full factory/material key
        # to avoid first-match errors when the same SKU is made at several plants.
        price_source, price_col = '本月', 'T'
        family_source, family_col = '本月', 'AA'
        if not meat:
            listrow = self.indices['列表'][key]
            for col, field in [('H', 'price'), ('I', 'family')]:
                c = self.sheets['列表'].cells.get(col + str(listrow))
                if c is not None and c.find('m:f', N) is None and self.sheets['列表'].value(col + str(listrow)) not in (None, ''):
                    if field == 'price':
                        price_source, price_col = '列表', col
                    else:
                        family_source, family_col = '列表', col
        values[pc + str(s)] = self.link(rows, pc + str(s), price_source, key, price_col, True)
        self.link(rows, familyc + str(s), family_source, key, family_col)
        if values[pc + str(s)] == 0:
            self.warn(f'{factory} / {material}：未税单价为0，边贡率留空。')
        if meat:
            if key in self.indices['列表']:
                self.link(rows, 'A' + str(s), '列表', key, 'A')
            hc, costcols, totalcol, ratecol = 'I', 'JKLMNO', 'O', 'P'
            specs = [(0, '本月', 'LMNOPQ', f'{self.month}月标准成本'),
                     (1, '本月', 'FGHIJK', f'{self.month}月实际成本')]
            diff_off, ratio_off, actual_off, standard_off = 2, 3, 1, 0
        elif rd:
            hc, costcols, totalcol, ratecol = 'H', 'IJKLMN', 'N', 'O'
            specs = [(0, '上上月', 'LMNOPQ', f'{(self.month-3)%12+1}月标准成本'),
                     (1, '上月', 'LMNOPQ', f'{(self.month-2)%12+1}月标准成本'),
                     (2, '本月', 'LMNOPQ', f'{self.month}月标准成本'),
                     (3, '本月', 'FGHIJK', f'{self.month}月实际成本'),
                     (6, '前三月平均', 'DEFGHI', '前三月平均实际成本')]
            diff_off, ratio_off, actual_off, standard_off = 4, 5, 3, 2
        else:
            hc, costcols, totalcol, ratecol = 'H', 'IJKLMN', 'N', 'O'
            specs = [(0, '本月', 'LMNOPQ', f'{self.month}月标准成本'),
                     (1, '本月', 'FGHIJK', f'{self.month}月实际成本'),
                     (4, '上月', 'FGHIJK', f'{(self.month-2)%12+1}月实际成本'),
                     (5, '上上月', 'FGHIJK', f'{(self.month-3)%12+1}月实际成本')]
            diff_off, ratio_off, actual_off, standard_off = 2, 3, 1, 0
        for off, sheet, cols, label in specs:
            row = s + off
            self.set_cell(rows, hc + str(row), label)
            for dest, src in zip(costcols, cols):
                address = dest + str(row)
                # Monthly cost rows are editable snapshots. Calculated rows
                # below reference these cells and retain their Excel formulas.
                values[address] = self.source_value(sheet, key, src, True)
                self.set_cell(rows, address, values[address])
            self.rate(rows, ratecol + str(row), totalcol + str(row), pc + str(s), values)
        self.set_cell(rows, hc + str(s + diff_off), f'{self.month}月实际-标准' if rd else '实际-标准')
        self.set_cell(rows, hc + str(s + ratio_off), f'{self.month}月偏差率' if rd or meat else '偏差率')
        for col in costcols:
            actual, standard = col + str(s + actual_off), col + str(s + standard_off)
            self.diff(rows, col + str(s + diff_off), actual, standard, values)
            self.diff(rows, col + str(s + ratio_off), actual, standard, values, True)
            if values.get(standard) == 0:
                self.warn(f'{factory} / {material}：{col}列本月标准成本为0，该项偏差率留空。')
        self.diff(rows, ratecol + str(s + diff_off), ratecol + str(s + actual_off), ratecol + str(s + standard_off), values)
        if self.has_variance_amount(name):
            quantity, difference = 'F' + str(s), 'N' + str(s + diff_off)
            a, b = values.get(quantity), values.get(difference)
            self.set_cell(rows, 'Q' + str(s),
                          formula=f'IF(AND(ISNUMBER({quantity}),ISNUMBER({difference})),{quantity}*{difference},"")',
                          cache=a * b if numeric(a) and numeric(b) else None)
        if not meat:
            if rd:
                self.diff(rows, 'P' + str(s + 1), 'N' + str(s + 1), 'N' + str(s), values)
                self.diff(rows, 'P' + str(s + 2), 'N' + str(s + 2), 'N' + str(s + 1), values)
            else:
                self.diff(rows, 'P' + str(s + 1), 'N' + str(s + 1), 'N' + str(s + 4), values)
                self.diff(rows, 'P' + str(s + 4), 'N' + str(s + 4), 'N' + str(s + 5), values)
        mr = f'${mc}{s}'
        for dest, src in [('Q' if meat else 'S', 'C'), ('R' if meat else 'T', 'D')]:
            self.lookup(rows, dest + str(s), mr, material, '负责人填写', 'A', src)
        self.lookup(rows, ('Q' if meat else 'S') + str(s + (2 if meat else 3)), mr, material, '负责人填写', 'A', 'E')
        if name in ['PFS生产缩小差异', 'PFS 研发降本', 'MKA产品']:
            self.lookup(rows, 'U' + str(s), mr, material, '降本目标', 'B', 'D', 1)
        keycol, faccol = ('Y', 'AA') if meat else ('AC', 'AE') if rd else ('AA', 'AC')
        for off in range(TARGETS[name][1]):
            self.set_cell(rows, keycol + str(s + off), material)
            self.set_cell(rows, faccol + str(s + off), factory)

    def mka_conditional(self, root, last):
        raw = self.changed.get('xl/styles.xml', self.parts['xl/styles.xml'])
        styles = ET.fromstring(raw)
        dxfs = styles.find('m:dxfs', N)
        if dxfs is None:
            dxfs = ET.Element(tag('dxfs'), {'count': '0'})
            # Per the stylesheet schema dxfs precedes tableStyles/colors/extLst.
            before = next((x for x in styles if x.tag in [tag('tableStyles'), tag('colors'), tag('extLst')]), None)
            styles.insert(list(styles).index(before) if before is not None else len(styles), dxfs)
        dxf = ET.Element(tag('dxf'))
        font = ET.SubElement(dxf, tag('font'))
        ET.SubElement(font, tag('b'))
        ET.SubElement(font, tag('color'), {'rgb': 'FF9C0006'})
        fill = ET.SubElement(dxf, tag('fill'))
        pattern = ET.SubElement(fill, tag('patternFill'), {'patternType': 'solid'})
        ET.SubElement(pattern, tag('fgColor'), {'rgb': 'FFFFC7CE'})
        ET.SubElement(pattern, tag('bgColor'), {'rgb': 'FFFFC7CE'})
        # Reuse an existing identical style on repeated generation.
        target = ET.tostring(dxf)
        index = next((i for i, x in enumerate(dxfs) if ET.tostring(x) == target), None)
        if index is None:
            index = len(dxfs)
            dxfs.append(dxf)
        dxfs.set('count', str(len(dxfs)))
        self.changed['xl/styles.xml'] = xmlbytes(styles, self.parts['xl/styles.xml'])
        cf = ET.Element(tag('conditionalFormatting'), {'sqref': f'A3:P{last}'})
        # Highlight the complete six-row block; formula reacts to edited costs.
        rule = ET.SubElement(cf, tag('cfRule'), {'type': 'expression', 'dxfId': str(index), 'priority': '1', 'stopIfTrue': '1'})
        rate = 'INDEX($N:$N,ROW()-MOD(ROW()-3,6)+3)'
        ET.SubElement(rule, tag('formula')).text = f'AND(ISNUMBER({rate}),ABS({rate})>0.05)'
        for existing in root.findall('m:conditionalFormatting/m:cfRule', N):
            existing.set('priority', str(int(existing.get('priority', '1')) + 1))
        before = next((x for x in root if x.tag in [tag('conditionalFormatting'), tag('dataValidations'), tag('pageMargins'), tag('pageSetup')]), None)
        root.insert(list(root).index(before) if before is not None else len(root), cf)

    def metadata(self, lasts):
        names = self.workbook.find('m:definedNames', N)
        if names is not None:
            for defined in names:
                sid = defined.get('localSheetId')
                if sid is not None and int(sid) < len(self.sheet_order):
                    name = self.sheet_order[int(sid)]
                    if name in lasts and defined.get('name') in ('_xlnm.Print_Area', '_xlnm._FilterDatabase') and defined.text:
                        defined.text = re.sub(r'(\$[A-Z]+\$)\d+(?=,|$)', lambda m: m[1] + str(lasts[name]), defined.text)
        calc = self.workbook.find('m:calcPr', N)
        if calc is None:
            calc = ET.SubElement(self.workbook, tag('calcPr'))
        calc.set('calcMode', 'auto')
        # New formulas already contain independently calculated caches. Forcing
        # every unrelated old whole-column array formula to recalculate on every
        # open makes this large legacy workbook unnecessarily slow. Auto mode
        # still updates report formulas when owners, targets or costs are edited.
        calc.set('fullCalcOnLoad', '0')
        calc.set('forceFullCalc', '0')
        calc.set('calcCompleted', '1')
        self.changed['xl/workbook.xml'] = xmlbytes(self.workbook, self.parts['xl/workbook.xml'])
        # A stale calc chain refers to removed/rearranged target rows. Excel will
        # rebuild it from formula cells, without touching any external links.
        if 'xl/calcChain.xml' in self.parts:
            self.deleted.add('xl/calcChain.xml')
            for part in ['xl/_rels/workbook.xml.rels', '[Content_Types].xml']:
                raw = self.parts[part]
                root = ET.fromstring(raw)
                for child in list(root):
                    if 'calcChain' in child.get('Type', '') or 'calcChain' in child.get('PartName', ''):
                        root.remove(child)
                self.changed[part] = xmlbytes(root, raw)

    def run(self, output_dir=None, progress=None):
        dest = Path(clean_path(output_dir)).resolve() if output_dir else self.source.parent
        if not dest.is_dir():
            raise ReportError('输出文件夹不存在：' + str(dest))
        timestamp = dt.datetime.now().strftime('%Y%m%d_%H%M%S_%f')
        period_label = f'{self.year}年{self.month}月' if self.year else f'{self.month}月'
        stem = period_label + '成本会议报表'
        output = dest / f'{stem}_自动生成_{timestamp}.xlsm'
        log = output.with_suffix('.生成记录.txt')
        archive = output.with_suffix('.退出产品填写记录.csv') if self.archived else None
        lasts = {}
        for name in TARGETS:
            if progress:
                progress('正在生成：' + name)
            lasts[name] = self.build(name)
        self.metadata(lasts)
        archive = output.with_suffix('.退出产品填写记录.csv') if self.archived else None
        # Write to temporary files, then publish a complete workbook. Never use a
        # user-specified existing filename, never alter the input in place.
        staged = []
        published = []
        try:
            with tempfile.NamedTemporaryFile(dir=dest, suffix='.xlsm', delete=False) as f:
                temp = Path(f.name)
            staged.append(temp)
            with zipfile.ZipFile(temp, 'w', compression=zipfile.ZIP_DEFLATED) as z:
                for name, raw in self.parts.items():
                    if name not in self.deleted:
                        z.writestr(self.infos[name], self.changed.get(name, raw))
            with zipfile.ZipFile(temp) as z:
                if z.testzip():
                    raise ReportError('生成文件完整性校验失败，未输出结果。')
                if z.read('xl/vbaProject.bin') != self.parts.get('xl/vbaProject.bin'):
                    raise ReportError('宏保留校验失败，未输出结果。')
                for name in TARGETS:
                    r = ET.fromstring(z.read(self.sheets[name].path))
                    for f in r.findall('m:sheetData/m:row/m:c/m:f', N):
                        if f.text and ('#REF!' in f.text or re.search(r'\[\d+\]', f.text)):
                            raise ReportError(f'“{name}”仍存在失效或外部公式，未输出结果。')
            report = ['五张成本分析表生成记录', '来源：' + str(self.source), '输出：' + str(output),
                      '生成时间：' + dt.datetime.now().isoformat(timespec='seconds'), '报告期间：' + period_label, '',
                      *[f'{name}：{count}个产品' for name, count in self.counts.items()], '',
                      '提示共' + str(len(self.warnings)) + '条：', *self.warnings,
                      '', '处理说明：原文件未修改；宏原样保留且未运行；负责人和目标保留关联公式。',
                      '列表按渠道全部生成；MKA总成本偏差率绝对值超过5%时突出显示。',
                      '肉松取本月“蚌埠二厂肉松”，排除成品名包含“推广包”的产品。',
                      '缺失数值及分母为0的比例显示为空；产品族内产品个数留空。',
                      '人工区内容按目标表＋工厂＋物料号＋字段恢复。',
                      '前三月平均直接读取现有来源，程序不改变其统计期间和计算口径。']
            if archive:
                report.append('退出产品的人工内容已保存：' + archive.name)
            log_temp = temp.with_suffix('.txt')
            log_temp.write_text('\n'.join(report), encoding='utf-8-sig')
            staged.append(log_temp)
            if archive:
                archive_temp = temp.with_suffix('.csv')
                staged.append(archive_temp)
                with archive_temp.open('w', newline='', encoding='utf-8-sig') as f:
                    writer = csv.writer(f)
                    writer.writerow(['目标表', '工厂', '物料号', '字段列', '产品块内行号', '原填写内容'])
                    # Avoid CSV formula execution when a note starts with =,+,-,@.
                    writer.writerows([["'" + v if isinstance(v, str) and v[:1] in '=+-@' else v for v in row] for row in self.archived])
            for src, target in [(log_temp, log)] + ([(archive_temp, archive)] if archive else []) + [(temp, output)]:
                if target.exists():
                    raise ReportError('输出文件名冲突，请重新生成。')
                os.rename(src, target)
                published.append(target)
            return Result(output, log, archive, self.counts, self.warnings)
        except PermissionError as e:
            for p in published:
                p.unlink(missing_ok=True)
            raise ReportError('无法写入输出目录或文件被占用。请关闭同名文件，或选择有写入权限的输出目录。') from e
        except Exception:
            for p in published:
                p.unlink(missing_ok=True)
            raise
        finally:
            for p in staged:
                p.unlink(missing_ok=True)


def generate(source, output_dir=None, progress=None):
    return Generator(source).run(output_dir, progress)


# Browser application. The complete delivered app.py includes the engine above;
# no separate engine.py, desktop executable or Excel installation is required.
import hashlib
import socket
import subprocess
import sys
import threading
import time
import urllib.request
import webbrowser


@dataclass
class BrowserResult:
    filename: str
    workbook: bytes
    records_filename: str
    records_zip: bytes
    counts: dict
    warnings: list
    archived_count: int


def process_upload(filename, content, progress=None):
    """Pure browser boundary: uploaded bytes in, ready-to-download bytes out."""
    safe_name = re.split(r'[\\/]', filename)[-1].strip()
    if not safe_name or Path(safe_name).suffix.lower() != '.xlsm':
        raise ReportError('请上传完整的 .xlsm 工作簿。')
    if len(content) > 200 * 1024 * 1024:
        raise ReportError('文件超过200MB，请减少文件中的无关内容后再上传。')
    # No persistent upload folder; each operation owns a separate temporary
    # directory and deletes it before returning the download buffers.
    with tempfile.TemporaryDirectory(prefix='cost_report_') as directory:
        source = Path(directory) / safe_name
        source.write_bytes(content)
        generator = Generator(source)
        result = generator.run(directory, progress)
        workbook = result.output.read_bytes()
        log = result.log.read_text(encoding='utf-8-sig')
        log = log.replace(str(generator.source), safe_name).replace(str(result.output), result.output.name)
        records = io.BytesIO()
        with zipfile.ZipFile(records, 'w', compression=zipfile.ZIP_DEFLATED) as z:
            z.writestr('生成记录.txt', b'\xef\xbb\xbf' + log.encode('utf-8'))
            if result.archive:
                z.writestr('人工内容保留记录.csv', result.archive.read_bytes())
        return BrowserResult(result.output.name, workbook,
                             result.output.stem + '_生成记录.zip', records.getvalue(),
                             dict(result.counts), list(result.warnings), len(generator.archived))


def sync_upload_state(state, filename, content):
    """Discard stale downloads on removal, replacement, rename or content change."""
    fingerprint = hashlib.sha256(filename.encode('utf-8') + b'\0' + content).hexdigest() if filename else None
    if state.get('cost_upload_id') != fingerprint:
        state.pop('cost_result', None)
        state.pop('cost_error', None)
        state['cost_upload_id'] = fingerprint
    return fingerprint


def render_page():
    import streamlit as st
    st.set_page_config(page_title='成本会议报表', page_icon='📊', layout='wide')
    st.title('成本会议报表')
    st.caption('内置五张结果表模板：每月上传7张来源表，自动生成生产缩小差异、研发降本、MKA、观察和肉松五张表。')
    st.caption('公式保留版：负责人、降本目标及差额／偏差率等计算保留公式；各月成本和前三月平均成本写为可编辑数值。')
    uploaded = st.file_uploader('上传成本工作簿（.xlsm）', type=['xlsm'],
                               help='只需包含：列表、本月、上月、上上月、前三月平均、负责人填写、降本目标。无需附带五张结果表。',
                               key='cost_upload')
    content = uploaded.getvalue() if uploaded is not None else b''
    sync_upload_state(st.session_state, uploaded.name if uploaded else '', content)
    if uploaded:
        st.caption(f'已选择：{uploaded.name}　·　{len(content) / 1024 / 1024:.2f} MB')
    if st.button('生成五张表', type='primary', disabled=uploaded is None, key='generate_report'):
        st.session_state.pop('cost_result', None)
        st.session_state.pop('cost_error', None)
        with st.status('正在检查工作簿…', expanded=False) as status:
            try:
                result = process_upload(uploaded.name, content,
                                        lambda label: status.update(label=label, state='running'))
                st.session_state['cost_result'] = result
                status.update(label='五张表生成完成', state='complete', expanded=False)
            except ReportError as e:
                st.session_state['cost_error'] = str(e)
                status.update(label='未生成，请检查文件', state='error', expanded=False)
            except Exception as e:
                st.session_state['cost_error'] = '无法生成，请确认文件未加密、未损坏，并包含7张来源表。具体原因：' + str(e)
                status.update(label='未生成，请检查文件', state='error', expanded=False)
    if st.session_state.get('cost_error'):
        st.error(st.session_state['cost_error'])
    result = st.session_state.get('cost_result')
    if result:
        st.success(f'已生成五张表，共 {sum(result.counts.values())} 个产品块。原文件未修改。')
        labels = {'PFS生产缩小差异': '生产缩小差异', 'PFS 研发降本': '研发降本',
                  'MKA产品': 'MKA产品', '观察产品': '观察产品', '肉松': '肉松'}
        columns = st.columns(5)
        for column, (name, count) in zip(columns, result.counts.items()):
            column.metric(labels[name], f'{count} 个')
        left, right = st.columns(2)
        left.download_button('下载结果工作簿（.xlsm）', result.workbook,
                             file_name=result.filename,
                             mime='application/vnd.ms-excel.sheet.macroEnabled.12',
                             key='download_workbook', type='primary')
        right.download_button('下载生成记录（.zip）', result.records_zip,
                              file_name=result.records_filename, mime='application/zip',
                              key='download_records')
        if result.archived_count:
            st.info(f'{result.archived_count}条人工内容已另存到生成记录包，请下载保存。')


def launch_local_browser_app():
    import streamlit  # Validate installation before starting a background browser timer.
    port = None
    for candidate in range(8501, 8521):
        with socket.socket() as sock:
            try:
                sock.bind(('127.0.0.1', candidate))
                port = candidate
                break
            except OSError:
                continue
    if port is None:
        raise SystemExit('8501—8520端口均被占用，请关闭不再使用的本地应用后重试。')
    url = f'http://127.0.0.1:{port}'
    def open_when_ready():
        for _ in range(150):
            try:
                with urllib.request.urlopen(url + '/_stcore/health', timeout=1) as response:
                    if response.status == 200:
                        webbrowser.open(url)
                        return
            except Exception:
                time.sleep(0.2)
    threading.Thread(target=open_when_ready, daemon=True).start()
    print('成本会议报表正在启动：' + url + '\n使用结束后按 Ctrl+C 关闭服务。', flush=True)
    command = [sys.executable, '-m', 'streamlit', 'run', str(Path(__file__).resolve()),
               '--server.address=127.0.0.1', f'--server.port={port}',
               '--server.headless=true', '--browser.gatherUsageStats=false',
               '--server.fileWatcherType=none']
    try:
        return subprocess.call(command)
    except KeyboardInterrupt:
        return 0


if __name__ == '__main__':
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx
    except ModuleNotFoundError:
        print('请先安装依赖：python -m pip install -r requirements.txt', file=sys.stderr)
        raise SystemExit(1)
    if get_script_run_ctx(suppress_warning=True) is None:
        raise SystemExit(launch_local_browser_app())
    render_page()
