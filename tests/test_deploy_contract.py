import ast
import re
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def parse_python(path: str) -> ast.Module:
    return ast.parse((ROOT / path).read_text(encoding="utf-8"))


class DeployContractTest(unittest.TestCase):
    def test_fastapi_entrypoint_exists(self):
        tree = parse_python("app.py")

        assigned_names = {
            target.id
            for node in tree.body
            if isinstance(node, ast.Assign)
            for target in node.targets
            if isinstance(target, ast.Name)
        }

        self.assertIn("app", assigned_names)
        self.assertTrue(
            any(
                isinstance(node, ast.ImportFrom)
                and node.module == "main"
                and any(alias.name == "scrape_subject" for alias in node.names)
                for node in tree.body
            )
        )

    def test_scraper_contract_exists(self):
        tree = parse_python("main.py")

        functions = {
            node.name
            for node in tree.body
            if isinstance(node, ast.FunctionDef)
        }

        self.assertIn("scrape_subject", functions)
        self.assertIn("create_driver", functions)
        self.assertIn("run_source", functions)
        self.assertIn("format_source_error", functions)

    def test_scraper_uses_source_level_error_handling(self):
        main_py = (ROOT / "main.py").read_text(encoding="utf-8")

        self.assertIn("SCRAPER_WAIT_SECONDS", main_py)
        self.assertIn("source_key", main_py)
        self.assertIn("selenium_error", main_py)
        self.assertIn('type(error).__name__ == "TimeoutException"', main_py)

    def test_dockerfile_uses_railway_port_and_chromium(self):
        dockerfile = (ROOT / "Dockerfile").read_text(encoding="utf-8")

        self.assertIn("${PORT:-8000}", dockerfile)
        self.assertIn("chromium", dockerfile)
        self.assertIn("chromium-driver", dockerfile)
        self.assertIn("uvicorn app:app", dockerfile)

    def test_fly_config_matches_docker_port(self):
        dockerfile = (ROOT / "Dockerfile").read_text(encoding="utf-8")
        fly_config = (ROOT / "fly.toml").read_text(encoding="utf-8")

        exposed_port = re.search(r"^EXPOSE\s+(\d+)$", dockerfile, re.MULTILINE)

        self.assertIsNotNone(exposed_port)
        self.assertIn(f"internal_port = {exposed_port.group(1)}", fly_config)
        self.assertIn('dockerfile = "Dockerfile"', fly_config)
        self.assertIn("force_https = true", fly_config)

    def test_deploy_files_are_not_markdown_snippets(self):
        filenames = [
            "app.py",
            "main.py",
            "requirements.txt",
            "Dockerfile",
            "docker-compose.yaml",
            "fly.toml",
        ]

        for filename in filenames:
            with self.subTest(filename=filename):
                content = (ROOT / filename).read_text(encoding="utf-8")

                self.assertNotIn("```", content)
                self.assertNotIn("Copy this", content)


if __name__ == "__main__":
    unittest.main()
