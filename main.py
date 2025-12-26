# main.py — ERP Product Import/Export Service (Render-compatible)
from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import StreamingResponse
from openpyxl import Workbook, load_workbook
import csv
from datetime import datetime
import io
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import tempfile
import os
import uuid
import json
import shutil
from typing import List, Optional

app = FastAPI(title="ERP Product Import/Export Service")

# -----------------------------
# IMPORT: Excel/CSV
# -----------------------------
@app.post("/import/excel")
async def import_excel(file: UploadFile = File(...)):
    if not file.filename.endswith(('.xlsx', '.csv')):
        raise HTTPException(400, "Only .xlsx or .csv files allowed")

    contents = await file.read()
    products = []

    try:
        if file.filename.endswith('.csv'):
            text = contents.decode("utf-8-sig").splitlines()  # Handle BOM
            reader = csv.DictReader(text)

            for row in reader:
                products.append({
                    'name': str(row.get('Name', '')).strip(),
                    'category_name': str(row.get('Category', '')).strip(),
                    'subcategory_name': str(row.get('Subcategory', '')).strip(),
                    'description': str(row.get('Description', '')).strip(),
                    'price': float(row.get('Price') or 0),
                    'stock_quantity': int(row.get('Stock') or 0),
                    'sku': str(row.get('SKU', '')).strip(),
                    'cft': float(row.get('CFT') or 0),
                    'material': str(row.get('Material', '')).strip(),
                    'finish': str(row.get('Finish', '')).strip(),
                    'specifications': str(row.get('Specifications', '')).strip(),
                    'main_image': None,
                    'image_data': None
                })

        else:
            wb = load_workbook(io.BytesIO(contents), data_only=True, read_only=True)
            ws = wb.active
            headers = [str(cell.value).strip() if cell.value else '' for cell in ws[1]]

            for row in ws.iter_rows(min_row=2, values_only=True):
                if not any(row):  # Skip empty rows
                    continue
                row_data = {}
                for i, value in enumerate(row):
                    key = headers[i] if i < len(headers) else f"col_{i}"
                    row_data[key] = value

                products.append({
                    'name': str(row_data.get('Name', '')).strip(),
                    'category_name': str(row_data.get('Category', '')).strip(),
                    'subcategory_name': str(row_data.get('Subcategory', '')).strip(),
                    'description': str(row_data.get('Description', '')).strip(),
                    'price': float(row_data.get('Price') or 0),
                    'stock_quantity': int(row_data.get('Stock') or 0),
                    'sku': str(row_data.get('SKU', '')).strip(),
                    'cft': float(row_data.get('CFT') or 0),
                    'material': str(row_data.get('Material', '')).strip(),
                    'finish': str(row_data.get('Finish', '')).strip(),
                    'specifications': str(row_data.get('Specifications', '')).strip(),
                    'main_image': None,
                    'image_data': None
                })
            wb.close()

    except Exception as e:
        raise HTTPException(400, f"Failed to parse file: {str(e)}")

    return {"products": products}

# -----------------------------
# IMPORT: PowerPoint (images only)
# -----------------------------
@app.post("/import/pptx")
async def import_pptx(file: UploadFile = File(...), category_id: int = Form(0), subcategory_id: Optional[str] = Form(None)):
    if not file.filename.endswith('.pptx'):
        raise HTTPException(400, "Only .pptx files allowed")
    
    contents = await file.read()
    temp_dir = tempfile.mkdtemp()
    tmp_path = os.path.join(temp_dir, "uploaded.pptx")
    
    try:
        with open(tmp_path, "wb") as f:
            f.write(contents)

        prs = Presentation(tmp_path)
        image_dir = os.path.join(temp_dir, "extracted_images")
        os.makedirs(image_dir, exist_ok=True)
        products = []

        for i, slide in enumerate(prs.slides, 1):
            image_found = False
            for shape in slide.shapes:
                if hasattr(shape, "image"):
                    img = shape.image
                    ext = img.ext or "png"
                    img_name = f"product_{i}_{uuid.uuid4().hex}.{ext}"
                    img_path = os.path.join(image_dir, img_name)
                    with open(img_path, "wb") as f_img:
                        f_img.write(img.blob)
                    
                    products.append({
                        "name": f"Product from Slide {i}",
                        "description": "",
                        "category_id": category_id,
                        "subcategory_id": int(subcategory_id) if subcategory_id and subcategory_id.isdigit() else None,
                        "price": 0.0,
                        "stock_quantity": 0,
                        "main_image": img_name,
                        "image_path": img_path  # For internal use only
                    })
                    image_found = True
                    break  # One image per slide
            if not image_found:
                # Add placeholder product if no image
                products.append({
                    "name": f"Product from Slide {i} (no image)",
                    "description": "",
                    "category_id": category_id,
                    "subcategory_id": int(subcategory_id) if subcategory_id and subcategory_id.isdigit() else None,
                    "price": 0.0,
                    "stock_quantity": 0,
                    "main_image": None
                })

        # Return image_dir so client can request image files if needed (optional)
        # But for security, we don't expose paths. So we just return product stubs.
        # Actual image export should be handled separately.
        return {"products": products}

    finally:
        # Clean up temp directory
        shutil.rmtree(temp_dir, ignore_errors=True)

# -----------------------------
# EXPORT: Excel
# -----------------------------
@app.post("/export/excel")
async def export_excel(products: str = Form(...)):
    try:
        data = json.loads(products)
    except Exception as e:
        raise HTTPException(400, f"Invalid JSON: {str(e)}")

    if not isinstance(data, list):
        raise HTTPException(400, "Products must be a JSON array")

    if not data:
        raise HTTPException(400, "No products to export")

    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Products")

    # Define headers from first product or standard set
    sample = data[0]
    headers = [
        'name', 'category_name', 'subcategory_name', 'description',
        'price', 'stock_quantity', 'sku', 'cft', 'material', 'finish', 'specifications'
    ]
    ws.append(headers)

    for item in data:
        row = [
            item.get('name', ''),
            item.get('category_name', ''),
            item.get('subcategory_name', ''),
            item.get('description', ''),
            float(item.get('price', 0)),
            int(item.get('stock_quantity', 0)),
            item.get('sku', ''),
            float(item.get('cft', 0)),
            item.get('material', ''),
            item.get('finish', ''),
            item.get('specifications', '')
        ]
        ws.append(row)

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
    except Exception as e:
        raise HTTPException(400, f"Invalid JSON: {str(e)}")

    if not isinstance(data, list):
        raise HTTPException(400, "Products must be a JSON array")

    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    # Clear default slide if any
    if len(prs.slides) > 0:
        prs.slides._sldIdLst = prs.slides._sldIdLst[:0]

    # Title slide
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
    left = top = Inches(1)
    width = Inches(11.33)
    height = Inches(2)
    title_box = slide.shapes.add_textbox(left, top, width, height)
    tf = title_box.text_frame
    tf.text = "Product Catalog"
    tf.paragraphs[0].font.size = Pt(44)
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER

    subtitle_box = slide.shapes.add_textbox(left, top + Inches(2.5), width, Inches(1))
    tf2 = subtitle_box.text_frame
    tf2.text = f"Generated on {datetime.now().strftime('%B %d, %Y')}"
    tf2.paragraphs[0].font.size = Pt(24)
    tf2.paragraphs[0].alignment = PP_ALIGN.CENTER

    # Product slides: 2 per slide
    chunks = [data[i:i+2] for i in range(0, len(data), 2)]
    for chunk in chunks:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        for idx, prod in enumerate(chunk):
            x = Inches(0.5 + idx * 6.5)
            y = Inches(1.5)

            # Product name
            name_box = slide.shapes.add_textbox(x, y, Inches(6), Inches(0.5))
            name_tf = name_box.text_frame
            name_tf.text = str(prod.get('name', 'Unnamed'))
            name_tf.paragraphs[0].font.bold = True
            name_tf.paragraphs[0].font.size = Pt(16)

            # Details
            y += Inches(0.6)
            details = f"Price: ${float(prod.get('price', 0)):.2f}\n"
            if prod.get('sku'):
                details += f"SKU: {prod['sku']}\n"
            if prod.get('cft'):
                details += f"CFT: {prod['cft']}\n"
            if prod.get('material'):
                details += f"Material: {prod['material']}\n"
            if prod.get('finish'):
                details += f"Finish: {prod['finish']}\n"

            details_box = slide.shapes.add_textbox(x, y, Inches(6), Inches(2))
            details_tf = details_box.text_frame
            details_tf.text = details
            details_tf.paragraphs[0].font.size = Pt(12)

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)

    return StreamingResponse(
        output,
        media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation',
        headers={"Content-Disposition": "attachment; filename=product_catalog.pptx"}
    )
