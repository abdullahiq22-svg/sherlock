# Engineering Drawing & Document Control System

نظام محلي لفهرسة ومتابعة مخططات ومستندات المشاريع الهندسية (MEP / إنشائي) يعمل على جهازك بدون إنترنت.

## التشغيل / Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

ثم الصق مسار مجلد المشروع في الشريط الجانبي واضغط **Scan & Index**.
قاعدة البيانات `project_drawings_db.sqlite` تُنشأ تلقائياً بجانب `app.py`.

## ما يفعله البرنامج / What it does

| Feature | Details |
|---|---|
| Multi-format indexing | PDF, DWG, DXF, XLSX/XLS/XLSM, CSV, DOCX/DOC |
| Incremental scan | Only new or changed files are re-read; deleted files are removed from the index |
| Multi-threaded | PDF / DXF parsing runs in parallel |
| Title-block parsing | Drawing No, Title, Revision, Date are read from the PDF title block when the file name has none |
| Revision ranking | `A < B < P01 < 00 < 01 < 02 < C01` (letters, preliminary, issued, contract) |
| Status | Current / Superseded per drawing number, computed across all folders |
| Issues tab | Duplicate copies (same content), revision conflicts (same No + Rev, different content), files without number |
| Search | Token search: `gf ltg 02` finds `GS-E-LTG-GF-02` |
| Export | Excel drawing register and CSV of the filtered list |
| Open buttons | Open file / open folder / copy path (works when the app runs on the same PC as the browser) |

## ملاحظات / Notes

* ملفات `.dwg` و`.doc` و`.xls` تُفهرس من اسم الملف فقط (صيغ ثنائية مغلقة).
* ملف PDF ممسوح ضوئياً (بدون نص) يُفهرس من اسم الملف فقط. يمكن إضافة OCR لاحقاً.
* استخدم **Deep scan** بعد تحديث البرنامج لإعادة قراءة كل الملفات بالمنطق الجديد.
