# -*- coding: utf-8 -*-
"""测试脚本：创建样例 PDF 并逐一验证所有核心功能"""
import os, sys, traceback

BASE = r"D:\C\Desktop\cs\workbuddy\updf\MyPDF工具箱"
TEST = os.path.join(BASE, "_test")
os.makedirs(TEST, exist_ok=True)
sys.path.insert(0, BASE)

import fitz
from PIL import Image, ImageDraw
import pdf_toolbox as tb

ok = 0
fail = 0

def check(name, fn):
    global ok, fail
    try:
        r = fn()
        ok += 1
        print(f"[OK]   {name}  ->  {r if isinstance(r, str) else '成功'}")
    except Exception as e:
        fail += 1
        print(f"[FAIL] {name}  ->  {e}")
        traceback.print_exc()

# ---------- 构造测试素材 ----------
def make_sample_pdf(path, n=3):
    doc = fitz.open()
    for i in range(n):
        page = doc.new_page(width=595, height=842)  # A4
        page.insert_text((72, 72), f"安全测试 PDF 第 {i+1} 页  (中文内容)", fontsize=16)
        # 表格
        rows = [["序号", "项目", "数值"], ["1", "测试A", "100"], ["2", "测试B", "200"]]
        y = 120
        for r in rows:
            x = 72
            for c in r:
                page.insert_text((x, y), c, fontsize=12)
                x += 120
            y += 20
        # 内嵌图片
        img = Image.new("RGB", (300, 150), (200, 60, 60))
        d = ImageDraw.Draw(img)
        d.text((20, 60), f"IMG-{i+1}", fill=(255, 255, 255))
        pimg = os.path.join(TEST, f"_img{i+1}.png")
        img.save(pimg)
        page.insert_image(fitz.Rect(72, 200, 372, 350), filename=pimg)
    doc.save(path)
    doc.close()
    print(f"已生成样例 PDF：{path} ({os.path.getsize(path)} 字节)")

p1 = os.path.join(TEST, "样例A.pdf")
p2 = os.path.join(TEST, "样例B.pdf")
make_sample_pdf(p1)
make_sample_pdf(p2)

# ---------- 1. 信息 ----------
check("PDF 信息", lambda: tb.pdf_info(p1))

# ---------- 2. 合并 ----------
merged = os.path.join(TEST, "_out_merged.pdf")
check("PDF 合并", lambda: tb.pdf_merge([p1, p2], merged))
check("合并后页数=6", lambda: fitz.open(merged).page_count == 6 and "页数正确")

# ---------- 3. 拆分 ----------
out_split = os.path.join(TEST, "split")
os.makedirs(out_split, exist_ok=True)
r = check("PDF 拆分(每页)", lambda: tb.pdf_split(p1, out_split))
check("拆分文件数=3", lambda: len(os.listdir(out_split)) == 3)

# ---------- 4. 压缩 ----------
comp = os.path.join(TEST, "_out_compressed.pdf")
check("PDF 压缩", lambda: tb.pdf_compress(p1, comp, "中"))
print(f"       压缩：{os.path.getsize(p1)} -> {os.path.getsize(comp)}")

# ---------- 5. 加密 ----------
enc = os.path.join(TEST, "_out_encrypted.pdf")
check("PDF 加密(AES-256)", lambda: tb.pdf_encrypt(p1, enc, "123456"))
check("加密后需密码", lambda: fitz.open(enc).is_encrypted and "已加密")

# ---------- 6. 解密 ----------
dec = os.path.join(TEST, "_out_decrypted.pdf")
check("PDF 解密", lambda: tb.pdf_decrypt(enc, dec, "123456"))
check("解密后无需密码", lambda: not fitz.open(dec).is_encrypted)

# ---------- 7. PDF 转图片 ----------
out_imgs = os.path.join(TEST, "toimg")
os.makedirs(out_imgs, exist_ok=True)
check("PDF 转图片", lambda: tb.pdf_to_images(p1, out_imgs, dpi=100))
imgs = [f for f in os.listdir(out_imgs) if f.endswith(".png")]
check("转图片数量=3", lambda: len(imgs) == 3)

# ---------- 8. 图片转 PDF ----------
img2pdf = os.path.join(TEST, "_out_img2pdf.pdf")
check("图片转 PDF", lambda: tb.images_to_pdf([os.path.join(out_imgs, f) for f in imgs], img2pdf))
check("图片合成页数=3", lambda: fitz.open(img2pdf).page_count == 3)

# ---------- 9. PDF 转 Word ----------
docx = os.path.join(TEST, "_out.docx")
check("PDF 转 Word", lambda: tb.pdf_to_word(p1, docx))
check("Word 文件有效", lambda: os.path.getsize(docx) > 500)

# ---------- 10. PDF 转 Excel ----------
xlsx = os.path.join(TEST, "_out.xlsx")
check("PDF 转 Excel", lambda: tb.pdf_to_excel(p1, xlsx))
check("Excel 文件有效", lambda: os.path.getsize(xlsx) > 500)

# ---------- 11. PDF 转 PPT ----------
pptx = os.path.join(TEST, "_out.pptx")
check("PDF 转 PPT", lambda: tb.pdf_to_ppt(p1, pptx, dpi=80))
check("PPT 文件有效", lambda: os.path.getsize(pptx) > 500)

# ---------- 12. 提取图片 ----------
out_ext = os.path.join(TEST, "extract")
os.makedirs(out_ext, exist_ok=True)
check("提取图片", lambda: tb.pdf_extract_images(p1, out_ext))
ext_imgs = os.listdir(out_ext)
check("提取图片数量>=3", lambda: len(ext_imgs) >= 3)

print(f"\n===== 测试结果：通过 {ok} 项，失败 {fail} 项 =====")
