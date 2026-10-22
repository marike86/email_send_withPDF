"""离线回归测试；所有 SMTP 调用均为模拟，不会发送邮件。"""

import csv
import io
import json
import os
from pathlib import Path
import smtplib
import tempfile
import threading
import unittest
from unittest.mock import patch

import openpyxl
import send_email as app


class BatchTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        self.folder = self.root / "pdf"
        self.folder.mkdir()
        self.pdf = self.folder / "001_证书.PDF"
        self.pdf.write_bytes(b"%PDF-1.4\nfixture")
        self.excel = self.root / "students.xlsx"
        self.write_students([("001", "张三", "student@example.com")])
        self.template = self.root / "template.txt"
        self.template.write_text("Subject: {姓名}的资料\n\n{姓名}同学，学号 {学号}", encoding="utf-8-sig")
        self.state = self.root / "status.json"
        self.config = dict(app.DEFAULT_CONFIG, sender_email="sender@example.com", sender_password="secret",
                           student_info_excel=str(self.excel), email_template=str(self.template),
                           send_interval=0, retry_interval=0)

    def write_students(self, rows):
        book = openpyxl.Workbook()
        book.active.append(["学号", "姓名", "邮箱"])
        for row in rows:
            book.active.append(row)
        book.save(self.excel)
        book.close()

    def prepare(self):
        return app.prepare_batch(self.config, self.folder, self.state)

    def test_preflight_personalizes_and_is_offline(self):
        with patch.object(app.smtplib, "SMTP") as smtp:
            jobs = self.prepare()
        smtp.assert_not_called()
        self.assertEqual(jobs[0]["status"], "ready")
        self.assertEqual(jobs[0]["subject"], "张三的资料")
        self.assertIn("001", jobs[0]["body"])
        self.assertFalse(self.state.exists())

    def test_pdf_filter_and_unmatched_are_visible(self):
        (self.folder / "001_password.txt").write_text("secret")
        (self.folder / "unmatched.pdf").write_bytes(b"%PDF-x")
        (self.folder / "001_invalid.pdf").write_bytes(b"not pdf")
        jobs = self.prepare()
        self.assertEqual(len(jobs), 3)
        self.assertEqual(sum(j["status"] == "blocked" for j in jobs), 2)

    def test_duplicate_students_rejected(self):
        self.write_students([("001", "A", "a@example.com"), ("001", "B", "b@example.com")])
        with self.assertRaisesRegex(ValueError, "重复"):
            self.prepare()

    def test_invalid_email_is_blocked(self):
        self.write_students([("001", "A", "a@example.com,b@example.com")])
        self.assertEqual(self.prepare()[0]["status"], "blocked")

    def test_numeric_id_and_multi_letter_columns(self):
        self.assertEqual(app.column_number("AA"), 27)
        self.write_students([(123.0, "A", "a@example.com")])
        self.assertIn("123", app.load_students(self.excel))
        self.assertNotIn("12", app.load_students(self.excel))

    def test_missing_column_rejected(self):
        with self.assertRaises(ValueError):
            app.load_students(self.excel, 1, 2, 27)
        with self.assertRaises(ValueError):
            app.column_number(0)

    def test_resume_and_content_identity(self):
        jobs = self.prepare()
        with patch.object(app, "send_email", return_value=True) as send:
            app.run_batch(self.config, jobs, self.state)
            app.run_batch(self.config, self.prepare(), self.state)
        self.assertEqual(send.call_count, 1)
        self.assertEqual(self.prepare()[0]["status"], "success")
        self.pdf.write_bytes(b"%PDF-1.4 changed")
        self.assertEqual(self.prepare()[0]["status"], "ready")

    def test_stale_preflight_rechecks_disk_status(self):
        stale = self.prepare()
        with patch.object(app, "send_email", return_value=True) as send:
            app.run_batch(self.config, self.prepare(), self.state)
            app.run_batch(self.config, stale, self.state)
        self.assertEqual(send.call_count, 1)

    def test_changed_attachment_never_sends(self):
        jobs = self.prepare()
        self.pdf.write_bytes(b"%PDF-1.4 changed")
        with patch.object(app, "send_email") as send:
            app.run_batch(self.config, jobs, self.state)
        send.assert_not_called()
        self.assertEqual(jobs[0]["status"], "failed")

    def test_failed_then_success_has_single_record(self):
        jobs = self.prepare()
        with patch.object(app, "send_email", side_effect=smtplib.SMTPAuthenticationError(535, b"denied")) as send:
            app.run_batch(self.config, jobs, self.state)
        self.assertEqual(send.call_count, 1)
        self.assertEqual(jobs[0]["status"], "failed")
        with patch.object(app, "send_email", return_value=True):
            app.run_batch(self.config, self.prepare(), self.state, retry_only=True)
        records = app.load_status(self.state)["records"]
        self.assertEqual(len(records), 1)
        self.assertEqual(next(iter(records.values()))["status"], "success")

    def test_uncertain_delivery_not_retried(self):
        with patch.object(app, "send_email", side_effect=app.DeliveryUncertain("核实")) as send:
            app.run_batch(self.config, self.prepare(), self.state)
            app.run_batch(self.config, self.prepare(), self.state)
        self.assertEqual(send.call_count, 1)
        self.assertEqual(self.prepare()[0]["status"], "uncertain")

    def test_crash_checkpoint_blocks_resend(self):
        with patch.object(app, "send_email", side_effect=KeyboardInterrupt):
            with self.assertRaises(KeyboardInterrupt):
                app.run_batch(self.config, self.prepare(), self.state)
        self.assertEqual(self.prepare()[0]["status"], "uncertain")

    def test_stop_before_send(self):
        stop = threading.Event()
        stop.set()
        with patch.object(app, "send_email") as send:
            app.run_batch(self.config, self.prepare(), self.state, stop)
        send.assert_not_called()

    def test_stop_between_messages(self):
        (self.folder / "001_second.pdf").write_bytes(b"%PDF-1.4 second")
        stop = threading.Event()
        with patch.object(app, "send_email", return_value=True) as send:
            jobs = app.run_batch(self.config, self.prepare(), self.state, stop,
                                 on_update=lambda *_: stop.set())
        self.assertEqual(send.call_count, 1)
        self.assertEqual(sum(j["status"] == "ready" for j in jobs), 1)

    def test_state_write_failure_prevents_send(self):
        with patch.object(app, "save_status", side_effect=OSError("disk full")), patch.object(app, "send_email") as send:
            with self.assertRaises(OSError):
                app.run_batch(self.config, self.prepare(), self.state)
        send.assert_not_called()

    def test_corrupt_and_legacy_state_do_not_reset(self):
        for text in ('{"processed_files": []}', '{bad json'):
            self.state.write_text(text)
            with self.assertRaises(ValueError):
                self.prepare()

    def test_config_password_not_saved_and_relative_paths(self):
        config_path = self.root / "config.json"
        app.save_config(dict(self.config, email_template="template.txt"), config_path)
        self.assertNotIn("secret", config_path.read_text())
        self.assertNotIn("sender_password", json.loads(config_path.read_text()))
        self.assertTrue(Path(app.load_config(config_path)["email_template"]).samefile(self.template))

    def test_environment_uses_correct_names(self):
        with patch.dict(os.environ, {"SMTP_SERVER": "smtp.example.com", "SMTP_PORT": "465", "USE_SSL": "true"}):
            config = app.load_config(self.root / "missing.json")
        self.assertEqual(config["smtp_server"], "smtp.example.com")
        self.assertEqual(config["smtp_port"], 465)
        self.assertTrue(config["use_ssl"])

    def test_ssl_transport_and_attachment(self):
        config = dict(self.config, smtp_port=465)
        with patch.object(app.smtplib, "SMTP_SSL") as ssl, patch.object(app.smtplib, "SMTP") as plain:
            server = ssl.return_value
            server.send_message.return_value = {}
            app.send_email(config, "student@example.com", "主题", "正文", self.pdf)
        plain.assert_not_called()
        server.starttls.assert_not_called()
        message = server.send_message.call_args.args[0]
        attachment = list(message.iter_attachments())[0]
        self.assertEqual(attachment.get_filename(), self.pdf.name)
        self.assertEqual(attachment.get_content_type(), "application/pdf")
        server.quit.assert_called_once()

    def test_starttls_uses_validated_context(self):
        with patch.object(app.smtplib, "SMTP") as smtp:
            smtp.return_value.send_message.return_value = {}
            app.send_email(self.config, "student@example.com", "主题", "正文", self.pdf)
        context = smtp.return_value.starttls.call_args.kwargs["context"]
        self.assertTrue(context.check_hostname)

    def test_missing_attachment_fails_before_network(self):
        with patch.object(app.smtplib, "SMTP") as smtp:
            with self.assertRaises(FileNotFoundError):
                app.send_email(self.config, "a@example.com", "s", "b", self.root / "absent.pdf")
        smtp.assert_not_called()

    def test_disconnect_during_submission_is_uncertain(self):
        with patch.object(app.smtplib, "SMTP") as smtp:
            smtp.return_value.send_message.side_effect = smtplib.SMTPServerDisconnected("lost")
            with self.assertRaises(app.DeliveryUncertain):
                app.send_email(self.config, "a@example.com", "s", "b", self.pdf)

    def test_transient_connection_error_retries(self):
        with patch.object(app, "send_email", side_effect=[OSError("offline"), True]) as send:
            success, _ = app.send_email_with_retry(self.config, "a@example.com", "s", "b", self.pdf, 3, 0)
        self.assertTrue(success)
        self.assertEqual(send.call_count, 2)

    def test_csv_formula_and_bom(self):
        jobs = self.prepare()
        jobs[0]["name"] = "=1+1"
        output = self.root / "results.csv"
        app.export_results(jobs, output)
        self.assertTrue(output.read_bytes().startswith(b"\xef\xbb\xbf"))
        with output.open(encoding="utf-8-sig", newline="") as stream:
            self.assertEqual(list(csv.reader(stream))[1][2], "'=1+1")

    def test_dry_run_cli_no_network(self):
        config_path = self.root / "config.json"
        app.save_config(self.config, config_path)
        with patch.object(app, "send_email") as send:
            code = app.main(["--config", str(config_path), "-p", str(self.folder),
                             "--state", str(self.state), "--dry-run"])
        self.assertEqual(code, 0)
        send.assert_not_called()

    def test_invalid_parameters(self):
        for key, value in (("send_interval", -1), ("retry_interval", float("nan")), ("max_retries", 0)):
            with self.assertRaises(ValueError):
                app.validate_config(dict(self.config, **{key: value}))

    def test_attachment_bytes_rechecked_before_transport(self):
        digest = app.file_digest(self.pdf)
        self.pdf.write_bytes(b"%PDF-1.4 changed")
        with patch.object(app.smtplib, "SMTP") as smtp:
            with self.assertRaisesRegex(ValueError, "更改"):
                app.send_email(self.config, "a@example.com", "s", "b", self.pdf, digest)
        smtp.assert_not_called()

    def test_connection_cleanup_cannot_cause_duplicate(self):
        with patch.object(app.smtplib, "SMTP") as smtp:
            server = smtp.return_value
            server.send_message.return_value = {}
            server.quit.side_effect = OSError("quit failed")
            server.close.side_effect = OSError("close failed")
            self.assertTrue(app.send_email(self.config, "a@example.com", "s", "b", self.pdf))

    def test_lock_rejects_simultaneous_sender(self):
        with app.status_lock(self.state):
            with self.assertRaises(RuntimeError):
                with app.status_lock(self.state):
                    self.fail("第二个发送任务不应取得锁")

    def test_xlsx_export_preserves_text_ids(self):
        jobs = self.prepare()
        jobs[0]["name"] = "=1+1"
        output = self.root / "results.xlsx"
        app.export_results(jobs, output)
        book = openpyxl.load_workbook(output)
        try:
            self.assertEqual(book.active["B2"].value, "001")
            self.assertEqual(book.active["C2"].value, "'=1+1")
            self.assertNotEqual(book.active["C2"].data_type, "f")
        finally:
            book.close()

    def test_invalid_workbook_has_friendly_error(self):
        self.excel.write_text("not a workbook")
        with self.assertRaisesRegex(ValueError, "无法读取 Excel"):
            self.prepare()

    def test_cli_handles_non_chinese_windows_output_encoding(self):
        config_path = self.root / "config.json"
        app.save_config(self.config, config_path)
        buffer = io.BytesIO()
        stream = io.TextIOWrapper(buffer, encoding="cp1252")
        with patch.object(app.sys, "stdout", stream):
            code = app.main(["--config", str(config_path), "-p", str(self.folder),
                             "--state", str(self.state), "--dry-run"])
        stream.flush()
        self.assertEqual(code, 0)
        self.assertIn("待发送", buffer.getvalue().decode("utf-8"))


if __name__ == "__main__":
    unittest.main()
