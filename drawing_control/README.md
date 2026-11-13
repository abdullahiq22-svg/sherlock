# Drawing & Document Control

نظام محلي لإدارة مخططات ومستندات المشاريع الهندسية (MEP / إنشائي) بأسلوب أنظمة EDMS الحديثة. يعمل على جهازك بدون إنترنت.

## التشغيل / Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

على ويندوز: اضغط مرتين على `run_app.bat`. ثم الصق مسار مجلد المشروع واضغط **Scan & index**.

## الصفحات / Pages

| Page | What it does |
|---|---|
| 📂 Files | Folder tree like Explorer. List, thumbnail grid or table. Click a file for preview and details. |
| 📋 Register | Full drawing register with filters (zone, discipline, status, CDE state, approval, flags) and Excel / CSV export. |
| 📨 Transmittals | Collect drawings in a basket, issue a numbered transmittal, export the Excel sheet and copy the files to a folder. |
| ⚠️ Issues | Revision conflicts, naming-standard violations, missing CAD or PDF, duplicate copies, files without a number. |
| 📊 Dashboard | KPIs and charts per zone, discipline, CDE state and approval code. |
| 🕘 Activity | Audit trail: scans, opened files, state changes, transmittals. |
| ⚙️ Settings | Scanner, naming standard, display, data management. |

## الميزات / Features

* Incremental multi-threaded indexing of PDF, DWG, DXF, Excel, Word, CSV. Deleted files are removed from the index.
* Title-block reading for PDF: drawing number, title, revision, date.
* Automatic zone (Basement, Ground Floor, Roof ...) and discipline (Lighting, Power, HVAC, Fire Alarm ...) classification.
* Revision ranking `A < B < P01 < 00 < 01 < C01`, Current / Superseded status, full version history per drawing.
* Built-in drawing viewer with zoom, and visual comparison of two revisions (red = removed, blue = added).
* Workflow per drawing revision: CDE state (WIP / Shared / Published / Archived), approval code A-D, notes, favourites.
* Global search: `b1 ltg 02` finds `GS-E-LTG-B1-02`.

## ملاحظات / Notes

* البيانات تُحفظ في `project_drawings_db.sqlite` بجانب `app.py`، ومعاينات المخططات في `.dc_cache`.
* البرنامج لا يعدّل ولا يحذف أي ملف من مجلد المشروع. التصدير ينسخ الملفات فقط.
* ملفات DWG و DOC و XLS تُفهرس من اسم الملف. معاينة DWG تؤخذ من ملف PDF لنفس المخطط إن وُجد.
* بعد تحديث البرنامج شغّل **Deep scan** من Settings مرة واحدة.
* نظام لمستخدم واحد على جهاز واحد. لا يوجد تعدد مستخدمين ولا مزامنة سحابية.
