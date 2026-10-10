from datetime import datetime
from html import escape
from pathlib import Path

import pytest
from selenium import webdriver
from selenium.webdriver.chrome.options import Options


_test_reports = {}


def pytest_addoption(parser):
    parser.addoption(
        "--headed",
        action="store_true",
        default=False,
        help="Run Chrome with a visible browser window.",
    )


@pytest.fixture
def browser(request):
    options = Options()
    if not request.config.getoption("--headed"):
        options.add_argument("--headless=new")
    options.add_argument("--window-size=1440,1000")
    options.add_argument("--disable-notifications")

    driver = webdriver.Chrome(options=options)
    driver.set_page_load_timeout(30)
    yield driver
    driver.quit()


def pytest_runtest_logreport(report):
    _test_reports.setdefault(report.nodeid, []).append(report)


def pytest_sessionfinish(session, exitstatus):
    results = []
    for nodeid, reports in _test_reports.items():
        failed = next((report for report in reports if report.failed), None)
        skipped = next((report for report in reports if report.skipped), None)
        call_report = next(
            (report for report in reports if report.when == "call"), None
        )
        if failed is not None:
            status = "Failed"
            details = getattr(failed, "longreprtext", str(failed.longrepr))
        elif skipped is not None:
            status = "Skipped"
            details = getattr(skipped, "longreprtext", str(skipped.longrepr))
        else:
            status = "Passed"
            details = ""
        duration = sum(report.duration for report in reports)
        results.append((nodeid, status, duration, details))

    counts = {
        status: sum(result[1] == status for result in results)
        for status in ("Passed", "Failed", "Skipped")
    }
    rows = []
    for nodeid, status, duration, details in results:
        detail_cell = (
            f"<details><summary>Chi tiết</summary><pre>{escape(details)}</pre></details>"
            if details
            else ""
        )
        rows.append(
            "<tr>"
            f'<td class="status {status.lower()}">{status}</td>'
            f"<td><code>{escape(nodeid)}</code>{detail_cell}</td>"
            f"<td>{duration:.2f}s</td>"
            "</tr>"
        )

    generated_at = datetime.now().astimezone().strftime("%Y-%m-%d %H:%M:%S %Z")
    report_html = f"""<!doctype html>
<html lang="vi">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Báo cáo kiểm thử Demoblaze</title>
  <style>
    :root {{ color-scheme: light; font-family: "Segoe UI", Arial, sans-serif; color: #243247; background: #f3f6fa; }}
    body {{ margin: 0; padding: 32px; }}
    main {{ max-width: 1180px; margin: 0 auto; }}
    h1 {{ margin: 0 0 8px; color: #182d4d; }}
    .meta {{ color: #64748b; margin: 0 0 24px; }}
    .summary {{ display: flex; flex-wrap: wrap; gap: 12px; margin-bottom: 20px; }}
    .metric {{ flex: 1 1 150px; min-width: 120px; padding: 14px 18px; border-radius: 8px; background: #fff; box-shadow: 0 1px 4px #182d4d18; }}
    .metric strong {{ display: block; font-size: 24px; }}
    .passed {{ color: #18794e; }} .failed {{ color: #c62828; }} .skipped {{ color: #9a6700; }}
    .table-wrap {{ overflow-x: auto; background: #fff; border-radius: 8px; box-shadow: 0 1px 4px #182d4d18; }}
    table {{ width: 100%; table-layout: fixed; border-collapse: collapse; }}
    th, td {{ padding: 12px 14px; border-bottom: 1px solid #e5eaf0; text-align: left; vertical-align: top; }}
    th {{ background: #182d4d; color: #fff; }}
    th:first-child, td:first-child {{ width: 110px; }}
    th:last-child, td:last-child {{ width: 100px; white-space: nowrap; }}
    tbody tr:nth-child(even) {{ background: #f7f9fc; }}
    td:first-child {{ font-weight: 700; white-space: nowrap; }}
    code {{ overflow-wrap: anywhere; }}
    details {{ margin-top: 8px; }}
    pre {{ max-width: 900px; overflow: auto; white-space: pre-wrap; color: #8b1e1e; }}
    @media (max-width: 640px) {{ body {{ padding: 16px; }} th, td {{ padding: 9px; }} }}
  </style>
</head>
<body>
  <main>
    <h1>Báo cáo kiểm thử tự động — Demoblaze</h1>
    <p class="meta">Selenium WebDriver + Python + pytest · Tạo lúc {escape(generated_at)}</p>
    <section class="summary" aria-label="Tổng kết">
      <div class="metric"><strong>{len(results)}</strong>Tổng số test</div>
      <div class="metric passed"><strong>{counts["Passed"]}</strong>Đạt</div>
      <div class="metric failed"><strong>{counts["Failed"]}</strong>Lỗi</div>
      <div class="metric skipped"><strong>{counts["Skipped"]}</strong>Bỏ qua</div>
    </section>
    <div class="table-wrap">
      <table>
        <thead><tr><th>Kết quả</th><th>Test case</th><th>Thời gian</th></tr></thead>
        <tbody>
          {"".join(rows) if rows else '<tr><td colspan="3">Chưa có test nào được chạy.</td></tr>'}
        </tbody>
      </table>
    </div>
  </main>
</body>
</html>
"""
    results_dir = Path(session.config.rootpath) / "results"
    results_dir.mkdir(parents=True, exist_ok=True)
    (results_dir / "report.html").write_text(report_html, encoding="utf-8")
