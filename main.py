from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import StreamingResponse
from openpyxl import load_workbook
import csv
from datetime import datetime
import io
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import tempfile
import os
import uuid
import json
import base64
from typing import List, Optional

app = FastAPI(title="ERP Product Import/Export Service")

# -----------------------------
# IMPORT: Excel/CSV
# -----------------------------
@app.post("/import/excel")
async def import_excel(file: UploadFile = File(...)):
    if not file.filename.endswith(('.xlsx', '.csv')):
        raise HTTPException(400, "Only .xlsx or .csv allowed")

    contents = await file.read()
    products = []

    try:
        if file.filename.endswith('.csv'):
            text = contents.decode("utf-8").splitlines()
            reader = csv.DictReader(text)

            for row in reader:
                products.append({
                    'name': row.get('Name', ''),
                    'category_name': row.get('Category', ''),
                    'subcategory_name': row.get('Subcategory', ''),
                    'description': row.get('Description', ''),
                    'price': float(row.get('Price') or 0),
                    'stock_quantity': int(row.get('Stock') or 0),
                    'sku': row.get('SKU', ''),
                    'cft': float(row.get('CFT') or 0),
                    'material': row.get('Material', ''),
                    'finish': row.get('Finish', ''),
                    'specifications': row.get('Specifications', ''),
                    'main_image': None,
                    'image_data': None
                })

        else:
            wb = load_workbook(io.BytesIO(contents), data_only=True)
            ws = wb.active
            headers = [cell.value for cell in ws[1]]

            for row in ws.iter_rows(min_row=2, values_only=True):
                row_data = dict(zip(headers, row))
                products.append({
                    'name': row_data.get('Name', ''),
                    'category_name': row_data.get('Category', ''),
                    'subcategory_name': row_data.get('Subcategory', ''),
                    'description': row_data.get('Description', ''),
                    'price': float(row_data.get('Price') or 0),
                    'stock_quantity': int(row_data.get('Stock') or 0),
                    'sku': row_data.get('SKU', ''),
                    'cft': float(row_data.get('CFT') or 0),
                    'material': row_data.get('Material', ''),
                    'finish': row_data.get('Finish', ''),
                    'specifications': row_data.get('Specifications', ''),
                    'main_image': None,
                    'image_data': None
                })

    except Exception as e:
        raise HTTPException(400, f"Parse error: {str(e)}")

    return {"products": products}

# -----------------------------
# IMPORT: PowerPoint (images only)
# -----------------------------
@app.post("/import/pptx")
async def import_pptx(file: UploadFile = File(...), category_id: int = Form(0), subcategory_id: Optional[str] = Form(None)):
    if not file.filename.endswith('.pptx'):
        raise HTTPException(400, "Only .pptx allowed")
    
    contents = await file.read()
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp:
        tmp.write(contents)
        tmp_path = tmp.name

    prs = Presentation(tmp_path)
    image_dir = f"/tmp/ppt_images_{uuid.uuid4().hex}"
    os.makedirs(image_dir, exist_ok=True)
    products = []

    for i, slide in enumerate(prs.slides, 1):
        for shape in slide.shapes:
            if hasattr(shape, "image"):
                img = shape.image
                ext = img.ext
                img_name = f"product_{i}_{uuid.uuid4().hex}.{ext}"
                img_path = os.path.join(image_dir, img_name)
                with open(img_path, "wb") as f:
                    f.write(img.blob)
                
                products.append({
                    "name": f"Imported Image {i}",
                    "description": "",
                    "category_id": category_id,
                    "subcategory_id": int(subcategory_id) if subcategory_id and subcategory_id.isdigit() else None,
                    "price": 0.0,
                    "stock_quantity": 0,
                    "main_image": img_name
                })
                break  # one image per slide

    os.unlink(tmp_path)
    return {"products": products, "image_dir": image_dir}

# -----------------------------
# EXPORT: Excel
# -----------------------------
@app.post("/export/excel")
async def export_excel(products: str = Form(...)):
    try:
        data = json.loads(products)
    except:
        raise HTTPException(400, "Invalid JSON")

    wb = load_workbook(io.BytesIO(), write_only=True)
    ws = wb.create_sheet("Products")

    if not data:
        raise HTTPException(400, "No data to export")

    headers = list(data[0].keys())
    ws.append(headers)

    for item in data:
        ws.append([item.get(h, "") for h in headers])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=products_export.xlsx"}
    )

# -----------------------------
# EXPORT: PowerPoint
# -----------------------------
@app.post("/export/pptx")
async def export_pptx(products: str = Form(...)):
    try:
        data = json.loads(products)
    except:
        raise HTTPException(400, "Invalid JSON")

    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    prs.removeSlideByIndex(0)

    # Title slide
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(11), Inches(2))
    tf = title.text_frame
    tf.text = "Product Catalog"
    tf.paragraphs[0].font.size = Pt(44)
    tf.paragraphs[0].font.bold = True

    subtitle = slide.shapes.add_textbox(Inches(1), Inches(4), Inches(11), Inches(1))
    tf2 = subtitle.text_frame
    tf2.text = f"Generated on {datetime.now().strftime('%B %d, %Y')}"

    # Product slides (2 per slide)
    chunks = [data[i:i+2] for i in range(0, len(data), 2)]
    for chunk in chunks:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        for idx, prod in enumerate(chunk):
            x = Inches(0.5 + idx * 6.5)
            y = Inches(1.5)

            # Product name
            name_box = slide.shapes.add_textbox(x, y, Inches(6), Inches(0.5))
            name_tf = name_box.text_frame
            name_tf.text = prod.get('name', 'Unnamed')
            name_tf.paragraphs[0].font.bold = True
            name_tf.paragraphs[0].font.size = Pt(16)

            # Price & details
            y += Inches(0.6)
            details = f"Price: ${prod.get('price', 0):.2f}\n"
            if prod.get('sku'): details += f"SKU: {prod['sku']}\n"
            if prod.get('cft'): details += f"CFT: {prod['cft']}\n"
            if prod.get('material'): details += f"Material: {prod['material']}\n"
            if prod.get('finish'): details += f"Finish: {prod['finish']}\n"

            details_box = slide.shapes.add_textbox(x, y, Inches(6), Inches(2))
            details_tf = details_box.text_frame
            details_tf.text = details
            details_tf.paragraphs[0].font.size = Pt(12)

            # Image (placeholder only – full image export requires more work)
            if prod.get('main_image'):
                # In real use, you'd download the image from your PHP server
                pass

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return StreamingResponse(
        output,
        media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation',
        headers={"Content-Disposition": "attachment; filename=product_catalog.pptx"}
    )
