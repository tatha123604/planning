import unittest
from unittest.mock import patch

from starlette.requests import Request

from src import app as report


class PfResultPageTests(unittest.TestCase):
    def render(self, task, **kwargs):
        request = Request({"type": "http", "method": "GET", "path": "/ssts-report", "headers": []})
        with (
            patch.object(report, "refresh_ssts_snapshot", side_effect=AssertionError("Remote sync on result page")),
            patch.object(report, "build_ssts_report_context", side_effect=AssertionError("Unrelated history on result page")),
            patch.object(report, "_build_ssts_pf_speed_analysis_result", side_effect=AssertionError("Synchronous rebuild")),
            patch.object(report, "_get_ssts_pf_analysis_task", side_effect=lambda task_id: task if task_id == "test-task" else None),
            patch.object(report, "_ssts_cleanup_junk_summary", return_value={"items": [], "item_count": 0}),
        ):
            return report._build_ssts_report_response(
                request, None, report_tab="pf_entering", pf_task_id="test-task",
                pf_day="2026-09-22", **kwargs,
            )

    def test_completed_result_renders_without_remote_sync_or_history(self):
        task = {
            "status": "completed", "report_day": "2026-09-22", "speed_threshold": 40,
            "result": {
                "pf_analysis_summary_rows": [{"train_no": "12345"}],
                "pf_analysis_detail_rows_by_train": {"12345": [{"station": "TEST-STATION", "pf_enter_speed": 45}]},
            },
        }
        response = self.render(task)
        self.assertEqual(response.status_code, 200)
        self.assertIn("TEST-STATION", response.body.decode())
        self.assertEqual(response.context["pf_analysis_status"], "completed")

    def test_lost_task_shows_recovery_message_without_rebuilding(self):
        response = self.render(None)
        self.assertEqual(response.status_code, 200)
        self.assertIn("Press Analysis to run it again", response.body.decode())
        self.assertEqual(response.context["pf_analysis_status"], "error")

    def test_pending_task_can_resume_polling_without_sync(self):
        response = self.render({"status": "running", "report_day": "2026-09-22"}, force=True)
        self.assertIn('data-status="running"', response.body.decode())
        self.assertIn('data-task-id="test-task"', response.body.decode())

    def test_online_tab_still_refreshes_and_builds_history(self):
        request = Request({"type": "http", "method": "GET", "path": "/ssts-report", "headers": []})
        with (
            patch.object(report, "refresh_ssts_snapshot", return_value={}) as sync,
            patch.object(report, "build_ssts_report_context", return_value={}) as history,
            patch.object(report, "_get_ssts_pf_analysis_task", return_value=None),
            patch.object(report, "_ssts_cleanup_junk_summary", return_value={}),
            patch.object(report.templates, "TemplateResponse", return_value=None),
        ):
            report._build_ssts_report_response(request, None, report_tab="online_offline", force=True)
            sync.assert_called_once_with(None, force=True)
            history.assert_called_once()


if __name__ == "__main__":
    unittest.main()
