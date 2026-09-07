from __future__ import annotations

import json
import os
import shutil
import sys
import time
import urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from tempfile import SpooledTemporaryFile

from core import JOBS_DIR, build_docx, create_analysis, ensure_dirs, load_analysis


HOST = "0.0.0.0" if os.environ.get("RENDER") else "127.0.0.1"
PORT = int(os.environ.get("PORT", "8765"))


INDEX_HTML = r"""<!doctype html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>مولد تقارير تقييم المقررات</title>
  <style>
    :root {
      --ink: #18202a;
      --muted: #5f6b7a;
      --line: #d7dde5;
      --soft: #f5f7fa;
      --brand: #0b5b3c;
      --brand-2: #1f4e79;
      --warn: #9b5c00;
      --bad: #9a1e1e;
      --ok: #0b6b44;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: Arial, "Segoe UI", Tahoma, sans-serif;
      color: var(--ink);
      background: #ffffff;
      line-height: 1.55;
    }
    header {
      padding: 22px 28px 14px;
      border-bottom: 1px solid var(--line);
      display: grid;
      grid-template-columns: 1fr auto;
      gap: 16px;
      align-items: end;
    }
    h1 {
      margin: 0;
      font-size: 24px;
      color: #000;
      font-weight: 800;
    }
    .subtle { color: var(--muted); font-size: 14px; }
    main { padding: 22px 28px 32px; max-width: 1440px; margin: 0 auto; }
    .steps {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      border: 1px solid var(--line);
      margin-bottom: 18px;
    }
    .step {
      padding: 12px 14px;
      background: var(--soft);
      border-left: 1px solid var(--line);
      font-weight: 700;
      color: var(--muted);
    }
    .step:last-child { border-left: 0; }
    .step.active { color: #fff; background: var(--brand-2); }
    section {
      border: 1px solid var(--line);
      padding: 18px;
      margin-bottom: 18px;
      background: #fff;
    }
    h2 { margin: 0 0 14px; font-size: 19px; color: #000; }
    label { display: block; font-weight: 700; margin: 12px 0 6px; }
    input[type="file"], input[type="text"], select {
      width: 100%;
      border: 1px solid var(--line);
      padding: 10px;
      font: inherit;
      background: #fff;
    }
    .grid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 14px; }
    .terms { display: flex; gap: 14px; flex-wrap: wrap; margin-top: 10px; }
    .terms label { margin: 0; font-weight: 600; }
    button, .download {
      border: 0;
      background: var(--brand);
      color: #fff;
      padding: 10px 16px;
      font: inherit;
      font-weight: 700;
      cursor: pointer;
      text-decoration: none;
      display: inline-block;
    }
    button.secondary { background: var(--brand-2); }
    button:disabled { opacity: .55; cursor: not-allowed; }
    table {
      width: 100%;
      border-collapse: collapse;
      table-layout: fixed;
      margin-top: 12px;
      font-size: 14px;
    }
    th {
      background: var(--brand-2);
      color: #fff;
      text-align: right;
      padding: 8px;
      border: 1px solid var(--line);
      vertical-align: middle;
    }
    td {
      border: 1px solid var(--line);
      padding: 8px;
      vertical-align: top;
      word-break: break-word;
    }
    tr.disabled { opacity: .55; }
    .status { font-weight: 800; }
    .ok { color: var(--ok); }
    .warn { color: var(--warn); }
    .bad { color: var(--bad); }
    .toolbar { display: flex; gap: 10px; align-items: center; flex-wrap: wrap; margin-top: 14px; }
    .message { padding: 12px; background: #f0f4f8; border: 1px solid var(--line); margin: 12px 0; }
    .hidden { display: none; }
    .small { font-size: 12px; color: var(--muted); }
    @media (max-width: 900px) {
      header, .grid, .steps { grid-template-columns: 1fr; }
      .step { border-left: 0; border-bottom: 1px solid var(--line); }
      table { font-size: 12px; }
      th, td { padding: 6px; }
    }
  </style>
</head>
<body>
  <header>
    <div>
      <h1>مولد تقارير تقييم المقررات</h1>
      <div class="subtle">مطابقة الخطة الدراسية مع التقرير التفصيلي ثم توليد تقرير Word قابل للمراجعة</div>
    </div>
    <div class="subtle">يعمل محلياً على جهازك</div>
  </header>
  <main>
    <div class="steps">
      <div id="stepUpload" class="step active">1. رفع الملفات</div>
      <div id="stepConfirm" class="step">2. تأكيد المطابقة</div>
      <div id="stepGenerate" class="step">3. توليد التقرير</div>
    </div>

    <section id="uploadPanel">
      <h2>رفع الملفات واختيار الفصول</h2>
      <form id="uploadForm">
        <div class="grid">
          <div>
            <label for="plan">الخطة الدراسية PDF</label>
            <input id="plan" name="plan" type="file" accept=".pdf" required>
          </div>
          <div>
            <label for="report">التقرير التفصيلي PDF</label>
            <input id="report" name="report" type="file" accept=".pdf" required>
          </div>
          <div>
            <label for="template">نموذج Word الرسمي DOCX</label>
            <input id="template" name="template" type="file" accept=".docx" required>
          </div>
        </div>
        <div class="toolbar">
          <button id="scanTermsBtn" type="button" class="secondary">قراءة فصول الخطة</button>
          <span id="termsMsg" class="small">اقرأ فصول الخطة أولاً أو اترك الاختيار فارغاً لتحليل كل الفصول.</span>
        </div>
        <label>الفصول المطلوبة من الخطة</label>
        <div id="termsBox" class="terms"></div>
        <div class="toolbar">
          <button id="analyzeBtn" type="submit">تحليل الملفات</button>
          <span id="uploadMsg" class="small"></span>
        </div>
      </form>
    </section>

    <section id="confirmPanel" class="hidden">
      <h2>تأكيد المقررات وصفحات القياس</h2>
      <div class="message">
        راجع المقررات قبل التوليد. الصفحات التفصيلية فقط تدخل في الحساب، أما صفحات الملخص فتظهر هنا للمراجعة ولا تحسب مرة ثانية.
      </div>
      <div style="overflow:auto">
        <table id="matchesTable">
          <thead>
            <tr>
              <th style="width:70px">اعتماد</th>
              <th>رمز الخطة</th>
              <th>اسم المقرر في الخطة</th>
              <th>الفصل</th>
              <th>الحالة</th>
              <th>صفحات القياس التفصيلية</th>
              <th>صفحات الملخص المستبعدة</th>
              <th>الأنشطة</th>
            </tr>
          </thead>
          <tbody></tbody>
        </table>
      </div>
      <div class="toolbar">
        <button id="generateBtn" class="secondary">توليد تقرير Word</button>
        <span id="generateMsg" class="small"></span>
      </div>
    </section>

    <section id="resultPanel" class="hidden">
      <h2>التقرير جاهز</h2>
      <p id="summaryText"></p>
      <a id="downloadLink" class="download" href="#">تحميل ملف Word</a>
    </section>
  </main>
  <script>
    let currentJob = null;
    let currentMatches = [];

    const $ = (id) => document.getElementById(id);
    function setStep(n) {
      $('stepUpload').classList.toggle('active', n === 1);
      $('stepConfirm').classList.toggle('active', n === 2);
      $('stepGenerate').classList.toggle('active', n === 3);
    }
    function statusClass(status) {
      if (status.includes('غير')) return 'bad';
      if (status.includes('تكرار')) return 'warn';
      return 'ok';
    }
    function renderMatches(matches) {
      const tbody = $('matchesTable').querySelector('tbody');
      tbody.innerHTML = '';
      matches.forEach((match, index) => {
        const tr = document.createElement('tr');
        if (!match.selected) tr.className = 'disabled';
        tr.innerHTML = `
          <td><input type="checkbox" data-index="${index}" ${match.selected ? 'checked' : ''} ${match.detail_pages.length ? '' : 'disabled'}></td>
          <td>${match.plan_code}</td>
          <td>${match.name}</td>
          <td>${match.term}</td>
          <td class="status ${statusClass(match.status)}">${match.status}</td>
          <td>${match.detail_pages.join(', ') || '-'}</td>
          <td>${match.summary_pages.join(', ') || '-'}</td>
          <td>${match.activities.join(', ') || '-'}</td>
        `;
        tbody.appendChild(tr);
      });
      tbody.querySelectorAll('input[type="checkbox"]').forEach(input => {
        input.addEventListener('change', (event) => {
          const index = Number(event.target.dataset.index);
          currentMatches[index].selected = event.target.checked;
          event.target.closest('tr').classList.toggle('disabled', !event.target.checked);
        });
      });
    }

    $('uploadForm').addEventListener('submit', async (event) => {
      event.preventDefault();
      $('analyzeBtn').disabled = true;
      $('uploadMsg').textContent = 'جاري التحليل...';
      const formData = new FormData(event.target);
      const terms = Array.from(document.querySelectorAll('input[name="terms"]:checked')).map(item => item.value);
      formData.set('terms', terms.join(','));
      try {
        const response = await fetch('/api/analyze', { method: 'POST', body: formData });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || 'تعذر تحليل الملفات');
        currentJob = data.job_id;
        currentMatches = data.matches;
        renderMatches(currentMatches);
        $('confirmPanel').classList.remove('hidden');
        $('uploadMsg').textContent = `تم العثور على ${data.matches.length} مقررات في الفصول المحددة.`;
        setStep(2);
      } catch (error) {
        $('uploadMsg').textContent = error.message;
      } finally {
        $('analyzeBtn').disabled = false;
      }
    });

    $('scanTermsBtn').addEventListener('click', async () => {
      const planInput = $('plan');
      if (!planInput.files.length) {
        $('termsMsg').textContent = 'اختر ملف الخطة الدراسية أولاً.';
        return;
      }
      $('scanTermsBtn').disabled = true;
      $('termsMsg').textContent = 'جاري قراءة فصول الخطة...';
      const formData = new FormData();
      formData.set('plan', planInput.files[0]);
      try {
        const response = await fetch('/api/terms', { method: 'POST', body: formData });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || 'تعذر قراءة الفصول');
        const box = $('termsBox');
        box.innerHTML = '';
        data.terms.forEach((term, index) => {
          const label = document.createElement('label');
          const checked = index === 0 || index === 2;
          label.innerHTML = `<input type="checkbox" name="terms" value="${term}" ${checked ? 'checked' : ''}> ${term}`;
          box.appendChild(label);
        });
        $('termsMsg').textContent = `تم العثور على ${data.terms.length} فصول في الخطة.`;
      } catch (error) {
        $('termsMsg').textContent = error.message;
      } finally {
        $('scanTermsBtn').disabled = false;
      }
    });

    $('generateBtn').addEventListener('click', async () => {
      $('generateBtn').disabled = true;
      $('generateMsg').textContent = 'جاري توليد ملف Word...';
      try {
        const response = await fetch('/api/generate', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ job_id: currentJob, matches: currentMatches })
        });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || 'تعذر توليد التقرير');
        $('downloadLink').href = data.download_url;
        $('summaryText').textContent = `تم توليد التقرير من ${data.selected_count} مقررات مؤكدة.`;
        $('resultPanel').classList.remove('hidden');
        $('generateMsg').textContent = 'تم التوليد بنجاح.';
        setStep(3);
      } catch (error) {
        $('generateMsg').textContent = error.message;
      } finally {
        $('generateBtn').disabled = false;
      }
    });
  </script>
</body>
</html>
"""


class AppHandler(BaseHTTPRequestHandler):
    server_version = "CourseReportApp/0.1"

    def log_message(self, format: str, *args: object) -> None:
        sys.stdout.write("%s - %s\n" % (self.address_string(), format % args))

    def send_json(self, payload: dict, status: int = 200) -> None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/":
            data = INDEX_HTML.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        if parsed.path.startswith("/download/"):
            job_id = parsed.path.rsplit("/", 1)[-1]
            path = JOBS_DIR / job_id / "تقرير_تحليل_استبانة_المقررات.docx"
            if not path.exists():
                self.send_json({"error": "الملف غير موجود"}, 404)
                return
            data = path.read_bytes()
            filename = urllib.parse.quote(path.name)
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{filename}")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        self.send_json({"error": "المسار غير موجود"}, 404)

    def do_POST(self) -> None:
        try:
            if self.path == "/api/analyze":
                self.handle_analyze()
                return
            if self.path == "/api/terms":
                self.handle_terms()
                return
            if self.path == "/api/generate":
                self.handle_generate()
                return
            self.send_json({"error": "المسار غير موجود"}, 404)
        except Exception as exc:
            self.send_json({"error": str(exc)}, 500)

    def parse_multipart(self) -> dict[str, object]:
        content_type = self.headers.get("Content-Type", "")
        if "multipart/form-data" not in content_type:
            raise ValueError("يرجى إرسال النموذج بصيغة multipart/form-data")
        match = re_search_boundary(content_type)
        if not match:
            raise ValueError("تعذر قراءة حدود نموذج الرفع")
        boundary = b"--" + match
        content_length = int(self.headers.get("Content-Length", "0"))
        body = self.rfile.read(content_length)
        fields: dict[str, object] = {}
        for part in body.split(boundary):
            part = part.strip(b"\r\n")
            if not part or part == b"--":
                continue
            header_block, _, payload = part.partition(b"\r\n\r\n")
            if not payload:
                continue
            payload = payload.rstrip(b"\r\n")
            headers = header_block.decode("utf-8", errors="replace")
            disposition = next((line for line in headers.splitlines() if line.lower().startswith("content-disposition:")), "")
            name_match = re_search_param(disposition, "name")
            if not name_match:
                continue
            filename_match = re_search_param(disposition, "filename")
            name = name_match.decode("utf-8", errors="replace")
            if filename_match is None:
                fields[name] = payload.decode("utf-8", errors="replace")
                continue
            temp = SpooledTemporaryFile(max_size=8 * 1024 * 1024)
            temp.write(payload)
            temp.seek(0)
            fields[name] = temp
        return fields

    def handle_terms(self) -> None:
        ensure_dirs()
        form = self.parse_multipart()
        job_dir = JOBS_DIR / "_terms_preview"
        job_dir.mkdir(parents=True, exist_ok=True)
        field = form["plan"]
        plan_path = job_dir / "plan.pdf"
        with plan_path.open("wb") as output:
            shutil.copyfileobj(field, output)
        from core import extract_plan_courses, term_sort_key

        courses, meta = extract_plan_courses(plan_path)
        terms = []
        for course in courses:
            if course.term not in terms:
                terms.append(course.term)
        self.send_json({"terms": sorted(terms, key=term_sort_key), "meta": meta})

    def handle_analyze(self) -> None:
        ensure_dirs()
        form = self.parse_multipart()
        job_id = time.strftime("%Y%m%d-%H%M%S")
        job_dir = JOBS_DIR / job_id
        counter = 1
        while job_dir.exists():
            counter += 1
            job_dir = JOBS_DIR / f"{job_id}-{counter}"
        job_dir.mkdir(parents=True)
        for field_name, filename in [("plan", "plan.pdf"), ("report", "report.pdf"), ("template", "template.docx")]:
            field = form[field_name]
            with (job_dir / filename).open("wb") as output:
                shutil.copyfileobj(field, output)
        terms_value = str(form.get("terms", ""))
        selected_terms = [item.strip() for item in terms_value.split(",") if item.strip()]
        analysis = create_analysis(job_dir, selected_terms)
        self.send_json(
            {
                "job_id": job_dir.name,
                "matches": analysis["matches"],
                "all_terms": analysis["all_terms"],
                "meta": analysis["meta"],
            }
        )

    def handle_generate(self) -> None:
        content_length = int(self.headers.get("Content-Length", "0"))
        payload = json.loads(self.rfile.read(content_length).decode("utf-8"))
        job_id = payload["job_id"]
        job_dir = JOBS_DIR / job_id
        if not job_dir.exists():
            self.send_json({"error": "لم يتم العثور على عملية التحليل"}, 404)
            return
        load_analysis(job_dir)
        matches = payload.get("matches") or []
        selected_count = sum(1 for match in matches if match.get("selected"))
        if selected_count == 0:
            self.send_json({"error": "اختر مقرراً واحداً على الأقل قبل التوليد"}, 400)
            return
        output = build_docx(job_dir, matches)
        self.send_json({"download_url": f"/download/{job_id}", "path": str(output), "selected_count": selected_count})


def main() -> None:
    ensure_dirs()
    server = ThreadingHTTPServer((HOST, PORT), AppHandler)
    print(f"Course report app running at http://{HOST}:{PORT}")
    server.serve_forever()


def re_search_boundary(content_type: str) -> bytes | None:
    import re

    match = re.search(r'boundary="?([^";]+)"?', content_type)
    return match.group(1).encode("utf-8") if match else None


def re_search_param(header: str, name: str) -> bytes | None:
    import re

    match = re.search(name + r'="([^"]*)"', header)
    return match.group(1).encode("utf-8") if match else None


if __name__ == "__main__":
    main()
