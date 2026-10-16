"""Run: python tests/test_app.py  - exercises the web endpoints with the sample files."""
import io
import os
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import app  # noqa: E402

c = app.app.test_client()
D = os.path.expanduser("~/Downloads")

r = c.post("/api/generate", content_type="multipart/form-data", data={
    "master": (open(f"{D}/June_Infinity.xlsx", "rb"), "June_Infinity.xlsx"),
    "resources": [(open(f"{D}/Resource Sheet-Rahul-June.xlsx", "rb"), "r.xlsx")], "month": ""})
j = r.get_json()
print("errors:", j["errors"])
print("warnings:", j["warnings"])
print("employees:", j["employees"])
z = c.get(j["download"])
print("download:", z.status_code, len(z.data))
print(zipfile.ZipFile(io.BytesIO(z.data)).namelist())

bad = c.post("/api/generate", content_type="multipart/form-data", data={
    "master": (io.BytesIO(b"junk"), "x.xlsx"), "resources": [(io.BytesIO(b"junk"), "y.xlsx")]})
print("bad file:", bad.get_json()["errors"])

# unmatched employee must block the download
import openpyxl  # noqa: E402
wb = openpyxl.load_workbook(f"{D}/Resource Sheet-Rahul-June.xlsx")
wb.active["B5"].value = "Nobody Here"
wb.active["D5"].value = "TE999999"
b = io.BytesIO(); wb.save(b); b.seek(0)
r = c.post("/api/generate", content_type="multipart/form-data", data={
    "master": (open(f"{D}/June_Infinity.xlsx", "rb"), "June_Infinity.xlsx"), "resources": [(b, "r.xlsx")]})
j = r.get_json()
print("unmatched:", j["errors"], "download" in j)
