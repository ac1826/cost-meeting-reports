"""Portable UI regression tests; no business workbook or private data required."""
from pathlib import Path
from types import SimpleNamespace
import unittest
from unittest.mock import patch

from streamlit.testing.v1 import AppTest

APP=Path(__file__).resolve().parents[1]/'app.py'


class PageTests(unittest.TestCase):
    def test_initial_page_has_no_bottom_expanders(self):
        page=AppTest.from_file(str(APP)).run(timeout=15)
        self.assertFalse(page.exception)
        self.assertEqual(page.title[0].value,'成本会议报表')
        self.assertTrue(page.button[0].disabled)
        self.assertEqual(len(page.expander),0)

    def test_results_keep_downloads_but_hide_warning_details(self):
        upload=SimpleNamespace(name='sample.xlsm',getvalue=lambda:b'ui-test')
        with patch('streamlit.file_uploader',return_value=upload):
            page=AppTest.from_file(str(APP)).run(timeout=15)
            self.assertFalse(page.exception)
            page.session_state['cost_result']=SimpleNamespace(
                counts={'PFS生产缩小差异':1,'PFS 研发降本':1,'MKA产品':1,'观察产品':1,'肉松':1},
                workbook=b'workbook-download',filename='report.xlsm',
                records_zip=b'record-download',records_filename='records.zip',
                warnings=['Warning details stay in the downloadable record.'],archived_count=0)
            page.run(timeout=15)
            self.assertFalse(page.exception)
            self.assertEqual(len(page.metric),5)
            self.assertEqual(len(page.expander),0)
            labels=[element.label for element in page.get('download_button')]
            self.assertEqual(labels,['下载结果工作簿（.xlsm）','下载生成记录（.zip）'])


if __name__=='__main__':unittest.main()
