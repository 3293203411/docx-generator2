# -*- coding: utf-8 -*-
"""
制式文档生成API服务
使用FastAPI + python-docx实现Word模板替换功能
"""

import os
import re
import uuid
import json
import base64
from typing import List, Optional, Union
from urllib.parse import urlparse

import requests
from fastapi import FastAPI, HTTPException, Query, BackgroundTasks
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tc
from docx.oxml.shared import qn

# 配置
TEMP_DIR = os.environ.get("TEMP_DIR", "/tmp/docx_generated")
HOST = os.environ.get("HOST", "0.0.0.0")
PORT = int(os.environ.get("PORT", "8000"))
BASE_URL = os.environ.get("BASE_URL", f"http://localhost:{PORT}")

# 确保临时目录存在
os.makedirs(TEMP_DIR, exist_ok=True)

app = FastAPI(
    title="制式文档生成API",
    description="Word模板占位符替换服务，支持扣子工作流调用",
    version="1.1.0"
)

# 添加CORS中间件
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


class GenerateRequest(BaseModel):
    """生成文档请求模型"""
    template_file_url: str
    text_keys: Union[List[str], str]
    text_values: Union[List[str], str]
    filename: Optional[str] = None


def parse_json_param(param: Union[List, str, None]) -> List[str]:
    """解析JSON字符串或列表参数"""
    if param is None:
        return []
    if isinstance(param, list):
        return [str(item) for item in param]
    if isinstance(param, str):
        try:
            parsed = json.loads(param)
            if isinstance(parsed, list):
                return [str(item) for item in parsed]
            return [str(parsed)]
        except json.JSONDecodeError:
            return [param]
    return [str(param)]


def download_file(url: str) -> bytes:
    """从URL下载文件"""
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
        response = requests.get(url, headers=headers, timeout=60)
        response.raise_for_status()
        return response.content
    except requests.exceptions.RequestException as e:
        raise HTTPException(status_code=400, detail=f"下载模板文件失败: {str(e)}")


def merge_runs_in_paragraph(paragraph: Paragraph) -> dict:
    """合并段落中所有run的文本"""
    if not paragraph.runs:
        return {"text": "", "runs": [], "formatting": None}

    full_text = ""
    for run in paragraph.runs:
        full_text += run.text if run.text else ""

    first_run = paragraph.runs[0]
    formatting = {
        "bold": first_run.bold,
        "italic": first_run.italic,
        "underline": first_run.underline,
        "font.name": first_run.font.name,
        "font.size": first_run.font.size,
        "font.color.rgb": first_run.font.color.rgb if first_run.font.color and first_run.font.color.rgb else None,
    }

    return {
        "text": full_text,
        "runs": list(paragraph.runs),
        "formatting": formatting
    }


def replace_placeholders_in_text(text: str, keys: List[str], values: List[str]) -> str:
    """在文本中替换占位符 {{key}} -> value"""
    result = text
    for key, value in zip(keys, values):
        pattern = r'\{\{\s*' + re.escape(key) + r'\s*\}\}'
        result = re.sub(pattern, str(value), result)
    return result


def apply_formatting_to_run(run, formatting: dict):
    """将格式应用到run"""
    if formatting is None:
        return
    if formatting.get("bold") is not None:
        run.bold = formatting["bold"]
    if formatting.get("italic") is not None:
        run.italic = formatting["italic"]
    if formatting.get("underline") is not None:
        run.underline = formatting["underline"]
    if formatting.get("font.name"):
        run.font.name = formatting["font.name"]
    if formatting.get("font.size"):
        run.font.size = formatting["font.size"]
    if formatting.get("font.color.rgb"):
        from docx.shared import RGBColor
        run.font.color.rgb = formatting["font.color.rgb"]


def process_paragraph(paragraph: Paragraph, keys: List[str], values: List[str]):
    """处理段落中的占位符替换"""
    if not paragraph.runs:
        return

    para_info = merge_runs_in_paragraph(paragraph)
    original_text = para_info["text"]

    if not original_text:
        return

    has_placeholder = False
    for key in keys:
        if re.search(r'\{\{\s*' + re.escape(key) + r'\s*\}\}', original_text):
            has_placeholder = True
            break

    if not has_placeholder:
        return

    new_text = replace_placeholders_in_text(original_text, keys, values)

    if new_text == original_text:
        return

    if paragraph.runs:
        first_run = paragraph.runs[0]
        first_run.text = new_text
        for run in paragraph.runs[1:]:
            run.text = ""
        apply_formatting_to_run(first_run, para_info["formatting"])


def process_table(table: Table, keys: List[str], values: List[str]):
    """处理表格中的占位符替换"""
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                process_paragraph(paragraph, keys, values)


def process_document(doc: Document, keys: List[str], values: List[str]):
    """处理整个文档的占位符替换"""
    for element in doc.element.body:
        if isinstance(element, CT_P):
            paragraph = Paragraph(element, doc)
            process_paragraph(paragraph, keys, values)

    for table in doc.tables:
        process_table(table, keys, values)


def generate_filename(original_url: str, custom_filename: Optional[str] = None) -> str:
    """生成输出文件名"""
    if custom_filename:
        if not custom_filename.lower().endswith('.docx'):
            custom_filename += '.docx'
        return custom_filename

    parsed = urlparse(original_url)
    basename = os.path.basename(parsed.path)

    if basename and basename.lower().endswith('.docx'):
        unique_id = uuid.uuid4().hex[:8]
        name, ext = os.path.splitext(basename)
        return f"{name}_{unique_id}{ext}"

    unique_id = uuid.uuid4().hex[:8]
    return f"generated_{unique_id}.docx"


@app.get("/")
async def root():
    """健康检查接口"""
    return {
        "service": "制式文档生成API",
        "version": "1.1.0",
        "status": "running",
        "endpoints": {
            "generate": "/generate",
            "generate_base64": "/generate-base64",
            "download": "/download/{filename}",
            "health": "/health"
        }
    }


@app.get("/health")
async def health_check():
    """健康检查"""
    return {"status": "healthy"}


@app.post("/generate")
async def generate_document(request: GenerateRequest):
    """生成文档，返回下载链接"""
    try:
        keys = parse_json_param(request.text_keys)
        values = parse_json_param(request.text_values)

        if len(keys) != len(values):
            raise HTTPException(
                status_code=400,
                detail=f"text_keys数量({len(keys)})与text_values数量({len(values)})不匹配"
            )

        if not keys:
            raise HTTPException(status_code=400, detail="text_keys不能为空")

        file_content = download_file(request.template_file_url)

        temp_id = uuid.uuid4().hex
        temp_input_path = os.path.join(TEMP_DIR, f"temp_{temp_id}.docx")

        with open(temp_input_path, 'wb') as f:
            f.write(file_content)

        try:
            doc = Document(temp_input_path)
            process_document(doc, keys, values)

            output_filename = generate_filename(request.template_file_url, request.filename)
            output_path = os.path.join(TEMP_DIR, output_filename)

            doc.save(output_path)

            download_url = f"{BASE_URL}/download/{output_filename}"

            return JSONResponse({
                "success": True,
                "message": "文档生成成功",
                "full_download_url": download_url,
                "filename": output_filename,
                "replaced_count": len(keys)
            })

        finally:
            if os.path.exists(temp_input_path):
                os.remove(temp_input_path)

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"生成文档失败: {str(e)}")


@app.post("/generate-base64")
async def generate_document_base64(request: GenerateRequest):
    """生成文档，直接返回base64编码的文件内容"""
    try:
        keys = parse_json_param(request.text_keys)
        values = parse_json_param(request.text_values)

        if len(keys) != len(values):
            raise HTTPException(
                status_code=400,
                detail=f"text_keys数量({len(keys)})与text_values数量({len(values)})不匹配"
            )

        if not keys:
            raise HTTPException(status_code=400, detail="text_keys不能为空")

        file_content = download_file(request.template_file_url)

        temp_id = uuid.uuid4().hex
        temp_input_path = os.path.join(TEMP_DIR, f"temp_{temp_id}.docx")
        temp_output_path = os.path.join(TEMP_DIR, f"output_{temp_id}.docx")

        with open(temp_input_path, 'wb') as f:
            f.write(file_content)

        try:
            doc = Document(temp_input_path)
            process_document(doc, keys, values)

            output_filename = generate_filename(request.template_file_url, request.filename)

            doc.save(temp_output_path)

            # 读取生成的文件并转base64
            with open(temp_output_path, 'rb') as f:
                file_bytes = f.read()

            file_base64 = base64.b64encode(file_bytes).decode('utf-8')

            return JSONResponse({
                "success": True,
                "message": "文档生成成功",
                "filename": output_filename,
                "file_base64": file_base64,
                "replaced_count": len(keys)
            })

        finally:
            if os.path.exists(temp_input_path):
                os.remove(temp_input_path)
            if os.path.exists(temp_output_path):
                os.remove(temp_output_path)

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"生成文档失败: {str(e)}")


@app.get("/download/{filename}")
async def download_file_endpoint(filename: str):
    """下载生成的文件"""
    filename = os.path.basename(filename)
    file_path = os.path.join(TEMP_DIR, filename)

    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="文件不存在或已过期")

    return FileResponse(
        path=file_path,
        filename=filename,
        media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )


@app.delete("/cleanup")
async def cleanup_old_files(max_age_hours: int = Query(default=24, ge=1, le=168)):
    """清理过期文件"""
    import time

    current_time = time.time()
    max_age_seconds = max_age_hours * 3600
    cleaned_count = 0

    try:
        for filename in os.listdir(TEMP_DIR):
            file_path = os.path.join(TEMP_DIR, filename)

            if os.path.isfile(file_path):
                file_age = current_time - os.path.getmtime(file_path)

                if file_age > max_age_seconds:
                    os.remove(file_path)
                    cleaned_count += 1

        return {
            "success": True,
            "message": f"清理完成，删除 {cleaned_count} 个过期文件",
            "cleaned_count": cleaned_count
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"清理失败: {str(e)}")


def start_server():
    """启动服务器"""
    import uvicorn
    uvicorn.run(app, host=HOST, port=PORT)


if __name__ == "__main__":
    start_server()
