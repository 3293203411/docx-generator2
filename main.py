# -*- coding: utf-8 -*-
import os
import re
import uuid
import json
from typing import Optional
from urllib.parse import quote
from datetime import datetime, timedelta
import io

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

# htmldocx 用于把 HTML 直接转 Word
from htmldocx import HtmlToDocx

HOST = os.environ.get("HOST", "0.0.0.0")
PORT = int(os.environ.get("PORT", 8000))
BASE_URL = os.environ.get("BASE_URL", f"http://localhost:{PORT}")

# 内存文件存储（2小时过期）
file_store = {}

app = FastAPI(
    title="HTML转Word文档API",
    description="输入HTML代码，直接返回Word下载链接",
    version="2.0.0-html-to-word"
)

# 全局异常捕获
@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    print("❌ 全局错误：", exc)
    return JSONResponse(
        status_code=500,
        content={"detail": f"服务器错误: {str(exc)}"}
    )

# 跨域
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ====================== 请求模型 ======================
class HtmlToWordRequest(BaseModel):
    html_content: str                # 你的 HTML 代码
    filename: Optional[str] = None   # 自定义文件名


# ====================== 核心：HTML → Word ======================
def html_to_word_bytes(html_content: str) -> bytes:
    """把 HTML 直接转换成 docx 字节流"""
    parser = HtmlToDocx()
    
    # 生成空白文档 + 写入 HTML
    doc = parser.parse_html_string(html_content)
    
    # 保存到内存
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.read()


# ====================== 生成文件名 ======================
def generate_filename(custom_name=None):
    if custom_name:
        name = custom_name.replace(".docx", "")
        return f"{name}.docx"
    return f"html_word_{uuid.uuid4().hex[:8]}.docx"


# ====================== 接口 ======================
@app.get("/")
async def root():
    return {"service": "HTML转Word文档API", "version": "2.0.0", "status": "running"}

@app.get("/health")
async def health_check():
    return {"status": "healthy"}

@app.post("/generate")
async def generate_document(request: HtmlToWordRequest):
    try:
        html = request.html_content.strip()
        if not html:
            raise HTTPException(status_code=400, detail="HTML 内容不能为空")

        # 1. HTML → Word
        word_bytes = html_to_word_bytes(html)

        # 2. 生成文件名
        real_filename = generate_filename(request.filename)
        encoded_filename = quote(real_filename, encoding="utf-8")

        # 3. 存入内存（2小时过期）
        file_store[encoded_filename] = {
            "data": word_bytes,
            "expires": datetime.now() + timedelta(hours=2),
            "real_name": real_filename
        }

        # 4. 返回下载链接
        download_url = f"{BASE_URL}/download/{encoded_filename}"

        return {
            "success": True,
            "message": "HTML 转 Word 成功",
            "download_url": download_url,
            "filename": real_filename
        }

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"转换失败: {str(e)}")


@app.get("/download/{filename}")
async def download_file_endpoint(filename: str):
    filename = os.path.basename(filename)
    if filename not in file_store:
        raise HTTPException(status_code=404, detail="文件不存在")

    item = file_store[filename]
    if datetime.now() >= item["expires"]:
        del file_store[filename]
        raise HTTPException(status_code=404, detail="文件已过期")

    real_name = item["real_name"]
    return JSONResponse(
        content=item["data"],
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={
            "Content-Disposition": f"attachment; filename*=utf-8''{quote(real_name)}"
        }
    )


def start_server():
    import uvicorn
    uvicorn.run(app, host=HOST, port=PORT)

if __name__ == "__main__":
    start_server()
