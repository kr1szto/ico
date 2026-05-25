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
        self.assertIn("selected_sources", main_py)
        self.assertIn("normalize_selected_sources", main_py)
        self.assertIn('SCRAPER_WAIT_SECONDS = int(os.getenv("SCRAPER_WAIT_SECONDS", "10"))', main_py)
        self.assertIn('PAGE_LOAD_TIMEOUT_SECONDS = int(os.getenv("PAGE_LOAD_TIMEOUT_SECONDS", "20"))', main_py)
        self.assertIn('driver.execute_script("window.stop();")', main_py)

    def test_ui_has_service_picker_and_no_rendered_raw_json(self):
        app_py = (ROOT / "app.py").read_text(encoding="utf-8")

        self.assertIn('name="services" value="finstat" checked', app_py)
        self.assertIn('name="services" value="orsr"', app_py)
        self.assertIn('name="services" value="rpvs"', app_py)
        self.assertIn('name="services" value="ruz"', app_py)
        self.assertIn("markSubmitting", app_py)
        self.assertIn("selected_sources=selected_services", app_py)
        self.assertIn("render_source_sections", app_py)
        self.assertNotIn("<h2>Raw JSON</h2>", app_py)
        self.assertIn("novalidate", app_py)
        self.assertIn("return markSubmitting(this)", app_py)
        self.assertNotIn("required\n                autofocus", app_py)
        self.assertIn('name="services" value="orsr" checked', app_py)
        self.assertIn('name="services" value="rpvs" checked', app_py)
        self.assertIn('name="services" value="ruz" checked', app_py)
        self.assertIn('Form(["finstat", "orsr", "rpvs", "ruz"])', app_py)
        self.assertIn("normalizeServices(item.services)", app_py)
        self.assertNotIn('<input type="hidden" name="services" value="finstat">', app_py)

    def test_vendor_intelligence_sections_exist(self):
        app_py = (ROOT / "app.py").read_text(encoding="utf-8")

        self.assertIn("build_vendor_intelligence", app_py)
        self.assertIn("build_confidence_model", app_py)
        self.assertIn("build_key_observations", app_py)
        self.assertIn("build_verification_prompts", app_py)
        self.assertIn("build_registry_coverage", app_py)
        self.assertIn("Prevádzkový prehľad", app_py)
        self.assertIn("Kvalita dostupných údajov", app_py)
        self.assertIn("Kľúčové pozorovania", app_py)
        self.assertNotIn("Otázky pre vendor manažéra", app_py)
        self.assertIn("Pokrytie registrov", app_py)
        self.assertIn("Stiahnuť podklady", app_py)
        self.assertIn("PDF je v backloge", app_py)
        self.assertIn("SOURCE_SHORT_LABELS", app_py)

    def test_slovak_validation_and_no_visible_raw_errors(self):
        app_py = (ROOT / "app.py").read_text(encoding="utf-8")

        self.assertIn("Zadajte IČO spoločnosti.", app_py)
        self.assertIn("Neplatný formát IČO.", app_py)
        self.assertIn("best_company_name", app_py)
        self.assertIn("escapeHtml", app_py)
        self.assertIn("Zdroj {SOURCE_LABELS.get(source_key, source_key)} je momentálne nedostupný.", app_py)
        self.assertNotIn("Selenium setup failed", app_py)
        self.assertNotIn("No basic risk flags", app_py)

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
