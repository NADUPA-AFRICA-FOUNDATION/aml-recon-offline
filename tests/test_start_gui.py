import unittest

from start_gui import build_command


class LauncherSecurityTests(unittest.TestCase):
    def test_streamlit_is_loopback_only_and_telemetry_is_disabled(self):
        command = build_command()
        self.assertIn("--server.address=127.0.0.1", command)
        self.assertIn("--server.enableXsrfProtection=true", command)
        self.assertIn("--server.enableCORS=true", command)
        self.assertIn("--browser.gatherUsageStats=false", command)
        self.assertNotIn("--server.address=0.0.0.0", command)


if __name__ == "__main__":
    unittest.main()
