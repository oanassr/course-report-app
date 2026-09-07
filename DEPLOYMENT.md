# نشر التطبيق على GitHub وRender

## الخيار الموصى به

استخدم GitHub لحفظ الكود، وRender لتشغيل التطبيق. GitHub Pages لا يناسب هذا التطبيق لأنه يحتاج Python لمعالجة ملفات PDF وWord.

## الملفات الجاهزة للنشر

- `requirements.txt`: مكتبات Python المطلوبة.
- `.python-version`: نسخة Python المقترحة.
- `render.yaml`: إعداد Render Blueprint.
- `.gitignore`: يمنع رفع ملفات التقارير والتجارب.

## خطوات الرفع إلى GitHub

إذا كان GitHub CLI مسجل الدخول:

```bat
git init
git add .
git commit -m "Initial course report app"
gh repo create course-report-app --private --source=. --remote=origin --push
```

إذا لم يكن مسجل الدخول:

```bat
gh auth login -h github.com
```

ثم أعد تنفيذ أوامر إنشاء المستودع.

## خطوات Render

1. افتح Render.
2. اختر New ثم Blueprint.
3. اختر مستودع GitHub.
4. Render سيقرأ `render.yaml`.
5. بعد النشر افتح رابط الخدمة.

## ملاحظات تشغيل

Render يمرر متغير `PORT` تلقائياً، والتطبيق يستمع على `0.0.0.0` عند التشغيل على Render.

في الخطة المجانية قد ينام التطبيق بعد فترة عدم استخدام، وقد يستغرق أول فتح بعد النوم وقتاً قصيراً.
