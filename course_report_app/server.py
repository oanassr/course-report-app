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


HOST = "0.0.0.0" if os.environ.get("PORT") or os.environ.get("RENDER") or os.environ.get("HF_SPACE") else "127.0.0.1"
PORT = int(os.environ.get("PORT", "7860" if os.environ.get("HF_SPACE") else "8765"))


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
    h1 { margin: 0; font-size: 24px; color: #000; font-weight: 800; }
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
      border-radius: 3px;
    }
    input[type="text"]:focus { outline: 2px solid var(--brand-2); border-color: var(--brand-2); }
    .hint { font-size: 12px; color: var(--muted); margin-top: 4px; }
    .grid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 14px; }
    .grid-2 { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 14px; }
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
      border-radius: 3px;
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
    .message { padding: 12px; background: #f0f4f8; border: 1px solid var(--line); margin: 12px 0; border-radius: 3px; }
    .message.warn-box { background: #fff8e6; border-color: #e6c26e; }
    .hidden { display: none; }
    .small { font-size: 12px; color: var(--muted); }
    .spinner { display:inline-block; width:14px; height:14px; border:2px solid #ccc; border-top-color:var(--brand-2); border-radius:50%; animation:spin .7s linear infinite; vertical-align:middle; margin-left:6px; }
    @keyframes spin { to { transform: rotate(360deg); } }
    .progress-bar-wrap { background: var(--line); border-radius: 4px; height: 6px; margin-top: 8px; overflow: hidden; display: none; }
    .progress-bar { height: 100%; background: var(--brand-2); width: 0%; transition: width .4s; }
    @media (max-width: 900px) {
      header, .grid, .grid-2, .steps { grid-template-columns: 1fr; }
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
    <div class="subtle" id="hostBadge">يعمل محلياً على جهازك</div>
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
            <label for="semesterLabel">رمز الفصل الدراسي</label>
            <input id="semesterLabel" name="semester_label" type="text" placeholder="مثال: 461 أو الفصل الأول 1446">
            <div class="hint">يظهر في عنوان التقرير بدلاً من اسم الفصل</div>
          </div>
        </div>
        <div class="toolbar">
          <button id="scanTermsBtn" type="button" class="secondary">قراءة فصول الخطة</button>
          <span id="termsMsg" class="small">اقرأ فصول الخطة أولاً أو اترك الاختيار فارغاً لتحليل كل الفصول.</span>
        </div>
        <label>الفصول المطلوبة من الخطة</label>
        <div id="termsBox" class="terms"></div>
        <div class="progress-bar-wrap" id="analyzeProgress">
          <div class="progress-bar" id="analyzeBar"></div>
        </div>
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
      <div id="ocrWarning" class="message warn-box hidden">
        ⚠️ تم استخدام OCR لقراءة الخطة المصورة — راجع أسماء ورموز المقررات بعناية قبل التوليد.
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
      <button id="resetBtn" style="margin-right:14px;background:var(--muted)">تقرير جديد</button>
    </section>
  </main>
  <script>
    let currentJob = null;
    let currentMatches = [];

    // Show whether running on HF Spaces, Render, or locally.
    if (location.hostname !== '127.0.0.1' && location.hostname !== 'localhost') {
      document.getElementById('hostBadge').textContent = 'نسخة سحابية';
    }

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

    async function readJson(response) {
      const body = await response.text();
      if (!body.trim()) {
        throw new Error(
          `لم يصل رد من الخادم (HTTP ${response.status}). ` +
          'إذا كان الملف ممسوحاً ضوئياً فقد تستغرق المعالجة دقيقتين أو أكثر — أعد المحاولة.'
        );
      }
      try {
        return JSON.parse(body);
      } catch (_) {
        throw new Error(
          `استجابة الخادم غير صالحة (HTTP ${response.status}). ` +
          'تحقق من سجل الخادم أو أعد المحاولة.'
        );
      }
    }

    function renderMatches(matches) {
      const tbody = $('matchesTable').querySelector('tbody');
      tbody.innerHTML = '';
      matches.forEach((match, index) => {
        const tr = document.createElement('tr');
        if (!match.selected) tr.className = 'disabled';
        tr.innerHTML = `
          <td><input type="checkbox" data-index="${index}"
            ${match.selected ? 'checked' : ''}
            ${match.detail_pages.length ? '' : 'disabled'}></td>
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
        input.addEventListener('change', (e) => {
          const idx = Number(e.target.dataset.index);
          currentMatches[idx].selected = e.target.checked;
          e.target.closest('tr').classList.toggle('disabled', !e.target.checked);
        });
      });
    }

    // Animate a fake progress bar while waiting for the server.
    let _progressTimer = null;
    function startProgress(barId, wrapId) {
      const bar = $(barId), wrap = $(wrapId);
      bar.style.width = '0%';
      wrap.style.display = 'block';
      let pct = 0;
      _progressTimer = setInterval(() => {
        // Slow logarithmic advance — never reaches 100% until done.
        pct = Math.min(pct + (100 - pct) * 0.03, 92);
        bar.style.width = pct + '%';
      }, 400);
    }
    function stopProgress(barId, wrapId) {
      clearInterval(_progressTimer);
      const bar = $(barId), wrap = $(wrapId);
      bar.style.width = '100%';
      setTimeout(() => { wrap.style.display = 'none'; bar.style.width = '0%'; }, 500);
    }

    $('uploadForm').addEventListener('submit', async (e) => {
      e.preventDefault();
      $('analyzeBtn').disabled = true;
      $('uploadMsg').innerHTML = 'جاري التحليل… <span class="spinner"></span>';
      startProgress('analyzeBar', 'analyzeProgress');
      const formData = new FormData(e.target);
      const terms = Array.from(document.querySelectorAll('input[name="terms"]:checked')).map(i => i.value);
      formData.set('terms', terms.join(','));
      try {
        const response = await fetch('/api/analyze', { method: 'POST', body: formData });
        const data = await readJson(response);
        if (!response.ok) throw new Error(data.error || 'تعذر تحليل الملفات');
        currentJob = data.job_id;
        currentMatches = data.matches;
        renderMatches(currentMatches);
        $('confirmPanel').classList.remove('hidden');
        const useOcr = data.meta && data.meta.ocr_used === 'true';
        $('ocrWarning').classList.toggle('hidden', !useOcr);
        const matched = data.matches.filter(m => m.detail_pages.length > 0).length;
        $('uploadMsg').textContent =
          `تم العثور على ${data.matches.length} مقرر في الفصول المحددة (${matched} مطابق).`;
        setStep(2);
        $('confirmPanel').scrollIntoView({ behavior: 'smooth' });
      } catch (err) {
        $('uploadMsg').textContent = err.message;
      } finally {
        $('analyzeBtn').disabled = false;
        stopProgress('analyzeBar', 'analyzeProgress');
      }
    });

    $('scanTermsBtn').addEventListener('click', async () => {
      const planInput = $('plan');
      if (!planInput.files.length) {
        $('termsMsg').textContent = 'اختر ملف الخطة الدراسية أولاً.';
        return;
      }
      $('scanTermsBtn').disabled = true;
      $('termsMsg').innerHTML = 'جاري قراءة فصول الخطة… <span class="spinner"></span>';
      const formData = new FormData();
      formData.set('plan', planInput.files[0]);
      try {
        const response = await fetch('/api/terms', { method: 'POST', body: formData });
        const data = await readJson(response);
        if (!response.ok) throw new Error(data.error || 'تعذر قراءة الفصول');
        const box = $('termsBox');
        box.innerHTML = '';
        data.terms.forEach((term, i) => {
          const label = document.createElement('label');
          // Pre-select first and third semesters (typical for odd academic year).
          const checked = i === 0 || i === 2;
          label.innerHTML = `<input type="checkbox" name="terms" value="${term}" ${checked ? 'checked' : ''}> ${term}`;
          box.appendChild(label);
        });
        const ocrNote = data.meta && data.meta.ocr_used === 'true'
          ? ' ⚠️ تم استخدام OCR — راجع الفصول قبل المتابعة.'
          : '';
        $('termsMsg').textContent = `تم العثور على ${data.terms.length} فصل.${ocrNote}`;
      } catch (err) {
        $('termsMsg').textContent = err.message;
      } finally {
        $('scanTermsBtn').disabled = false;
      }
    });

    $('generateBtn').addEventListener('click', async () => {
      const selected = currentMatches.filter(m => m.selected);
      if (selected.length === 0) {
        $('generateMsg').textContent = 'اختر مقرراً واحداً على الأقل.';
        return;
      }
      $('generateBtn').disabled = true;
      $('generateMsg').innerHTML = 'جاري توليد ملف Word… <span class="spinner"></span>';
      try {
        const response = await fetch('/api/generate', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ job_id: currentJob, matches: currentMatches })
        });
        const data = await readJson(response);
        if (!response.ok) throw new Error(data.error || 'تعذر توليد التقرير');
        $('downloadLink').href = data.download_url;
        $('summaryText').textContent =
          `تم توليد التقرير من ${data.selected_count} مقرر مؤكد.`;
        $('resultPanel').classList.remove('hidden');
        $('generateMsg').textContent = 'تم التوليد بنجاح ✓';
        setStep(3);
        $('resultPanel').scrollIntoView({ behavior: 'smooth' });
      } catch (err) {
        $('generateMsg').textContent = err.message;
      } finally {
        $('generateBtn').disabled = false;
      }
    });

    $('resetBtn').addEventListener('click', () => {
      currentJob = null;
      currentMatches = [];
      $('uploadForm').reset();
      $('termsBox').innerHTML = '';
      $('termsMsg').textContent = 'اقرأ فصول الخطة أولاً أو اترك الاختيار فارغاً لتحليل كل الفصول.';
      $('uploadMsg').textContent = '';
      $('generateMsg').textContent = '';
      $('confirmPanel').classList.add('hidden');
      $('resultPanel').classList.add('hidden');
      $('ocrWarning').classList.add('hidden');
      $('matchesTable').querySelector('tbody').innerHTML = '';
      setStep(1);
      window.scrollTo({ top: 0, behavior: 'smooth' });
    });
  </script>
</body>
</html>
"""


class AppHandler(BaseHTTPRequestHandler):
    server_version = "CourseReportApp/1.1"

    def log_message(self, format: str, *args: object) -> None:  # noqa: A002
        sys.stdout.write("%s - %s\n" % (self.address_string(), format % args))
        sys.stdout.flush()

    def send_json(self, payload: dict, status: int = 200) -> None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/health":
            self.send_json({"status": "ok", "service": "course-report-app"})
            return
        if parsed.path == "/":
            data = INDEX_HTML.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Cache-Control", "no-store")
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
            file_bytes = path.read_bytes()
            filename = urllib.parse.quote(path.name)
            self.send_response(200)
            self.send_header(
                "Content-Type",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
            self.send_header(
                "Content-Disposition", f"attachment; filename*=UTF-8''{filename}"
            )
            self.send_header("Content-Length", str(len(file_bytes)))
            self.end_headers()
            self.wfile.write(file_bytes)
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
        boundary_match = _re_boundary(content_type)
        if not boundary_match:
            raise ValueError("تعذر قراءة حدود نموذج الرفع")
        boundary = b"--" + boundary_match
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
            disposition = next(
                (
                    line
                    for line in headers.splitlines()
                    if line.lower().startswith("content-disposition:")
                ),
                "",
            )
            name_match = _re_param(disposition, "name")
            if not name_match:
                continue
            filename_match = _re_param(disposition, "filename")
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
        plan_path = job_dir / "plan.pdf"
        with plan_path.open("wb") as out:
            shutil.copyfileobj(form["plan"], out)
        from core import extract_plan_courses, term_sort_key

        courses, meta = extract_plan_courses(plan_path)
        seen: list[str] = []
        for course in courses:
            if course.term not in seen:
                seen.append(course.term)
        terms = seen
        if meta.get("ocr_used") == "true":
            # Expose all possible terms when OCR is used so missed ones are visible.
            all_eight = [
                "الأول", "الثاني", "الثالث", "الرابع",
                "الخامس", "السادس", "السابع", "الثامن",
            ]
            terms = sorted(
                {*all_eight, *terms},
                key=term_sort_key,
            )
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
        for field_name, filename in [("plan", "plan.pdf"), ("report", "report.pdf")]:
            field = form.get(field_name)
            if field is None:
                self.send_json({"error": f"الحقل '{field_name}' مفقود"}, 400)
                return
            with (job_dir / filename).open("wb") as out:
                shutil.copyfileobj(field, out)
        terms_value = str(form.get("terms", ""))
        selected_terms = [t.strip() for t in terms_value.split(",") if t.strip()]
        semester_label = str(form.get("semester_label", "")).strip()
        analysis = create_analysis(job_dir, selected_terms, semester_label=semester_label)
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
        job_id = payload.get("job_id", "")
        job_dir = JOBS_DIR / job_id
        if not job_dir.exists():
            self.send_json({"error": "لم يتم العثور على عملية التحليل"}, 404)
            return
        load_analysis(job_dir)
        matches = payload.get("matches") or []
        selected_count = sum(1 for m in matches if m.get("selected"))
        if selected_count == 0:
            self.send_json({"error": "اختر مقرراً واحداً على الأقل قبل التوليد"}, 400)
            return
        output = build_docx(job_dir, matches)
        self.send_json(
            {
                "download_url": f"/download/{job_id}",
                "path": str(output),
                "selected_count": selected_count,
            }
        )


def main() -> None:
    ensure_dirs()
    server = ThreadingHTTPServer((HOST, PORT), AppHandler)
    print(f"Course report app running at http://{HOST}:{PORT}", flush=True)
    server.serve_forever()


def _re_boundary(content_type: str) -> bytes | None:
    import re
    match = re.search(r'boundary="?([^";]+)"?', content_type)
    return match.group(1).encode("utf-8") if match else None


def _re_param(header: str, name: str) -> bytes | None:
    import re
    match = re.search(name + r'="([^"]*)"', header)
    return match.group(1).encode("utf-8") if match else None


if __name__ == "__main__":
    main()
