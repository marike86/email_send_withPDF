"""Qt 离屏交互测试，可在无显示器的 CI 环境运行。"""

import os
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import time
import unittest
from unittest.mock import patch

try:
    from PySide6.QtWidgets import QApplication, QMessageBox
    from PySide6.QtTest import QTest
    import send_email_gui as gui
except ImportError:
    gui = None

import send_email as core
import test_send_email


@unittest.skipIf(gui is None, "未安装可选桌面依赖")
class GuiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])
        cls.app.setStyle("Fusion")
        cls.app.setStyleSheet(gui.STYLE)

    def setUp(self):
        self.fixture = test_send_email.BatchTests()
        self.fixture.setUp()
        self.addCleanup(self.fixture.doCleanups)
        self.config_path = self.fixture.root / "config.json"
        core.save_config(dict(self.fixture.config, attachment_folder=str(self.fixture.folder)), self.config_path)
        self.window = gui.MainWindow(self.config_path, self.fixture.state)
        self.window.fields["sender_password"].setText("test-secret")
        self.window.show()
        self.addCleanup(self.window.close)

    def wait_worker(self):
        deadline = time.monotonic() + 10
        while self.window.busy() and time.monotonic() < deadline:
            QTest.qWait(10)
        self.assertFalse(self.window.busy(), "工作线程未及时退出")

    def test_preflight_preview_filter_and_invalidation(self):
        self.window.preflight()
        self.wait_worker()
        self.assertEqual(self.window.table.rowCount(), 1)
        self.assertIn("张三的资料", self.window.preview.toPlainText())
        self.assertTrue(self.window.start.isEnabled())
        self.window.search.setText("不存在的学生")
        self.assertTrue(self.window.table.isRowHidden(0))
        self.window.search.clear()
        self.assertFalse(self.window.table.isRowHidden(0))
        self.window.fields["smtp_port"].setValue(465)
        self.assertIsNone(self.window.snapshot)
        self.assertFalse(self.window.start.isEnabled())

    def test_mock_send_and_resume(self):
        self.window.preflight()
        self.wait_worker()
        with patch.object(QMessageBox, "question", return_value=QMessageBox.StandardButton.Yes), \
             patch.object(core, "send_email", return_value=True) as send:
            self.window.send()
            self.wait_worker()
        send.assert_called_once()
        self.assertEqual(self.window.jobs[0]["status"], "success")
        self.assertFalse(self.window.start.isEnabled())
        self.window.preflight()
        self.wait_worker()
        self.assertEqual(self.window.jobs[0]["status"], "success")

    def test_preflight_error_is_recoverable(self):
        self.window.fields["email_template"].setText(str(self.fixture.root / "absent.txt"))
        with patch.object(QMessageBox, "warning") as warning:
            self.window.preflight()
            self.wait_worker()
        warning.assert_called_once()
        self.assertTrue(self.window.check.isEnabled())
        self.assertFalse(self.window.start.isEnabled())

    def test_uncertain_record_can_be_resolved(self):
        jobs = core.prepare_batch(self.fixture.config, self.fixture.folder, self.fixture.state)
        core.save_status(self.fixture.state, {"version": 2, "records": {
            jobs[0]["key"]: {"status": "sending", "error": ""}}})
        self.window.preflight()
        self.wait_worker()
        self.assertTrue(self.window.resolve.isEnabled())
        with patch.object(gui, "QMessageBox") as dialog:
            sent, retry, cancel = object(), object(), object()
            box = dialog.return_value
            box.addButton.side_effect = [sent, retry, cancel]
            box.clickedButton.return_value = sent
            self.window.resolve_uncertain()
        self.assertEqual(self.window.jobs[0]["status"], "success")
        self.assertFalse(self.window.start.isEnabled())


if __name__ == "__main__":
    unittest.main()
