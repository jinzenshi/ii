import os
from typing import Optional
from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import StreamingResponse, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from core import fill_form

app = FastAPI()

# 检查静态文件目录是否存在，如果存在则挂载
if os.path.exists("static"):
    app.mount("/static", StaticFiles(directory="static"), name="static")

@app.get("/")
def root():
    """主页 - 返回 HTML 界面"""
    if os.path.exists(os.path.join("static", "index.html")):
        path = os.path.join("static", "index.html")
        return FileResponse(path)
    return {
        "status": "ok",
        "message": "智能填表服务正在运行",
        "docs": "/docs"
    }

@app.get("/index")
def index():
    """主页 - 返回 HTML 界面（备用）"""
    if os.path.exists(os.path.join("static", "index.html")):
        path = os.path.join("static", "index.html")
        return FileResponse(path)
    return {"message": "请访问 /docs 查看API文档"}

@app.post("/process")
async def process(
    docx: UploadFile = File(...),
    user_info_text: str = Form(...),
    photo: Optional[UploadFile] = File(None)
):
    try:
        docx_bytes = await docx.read()
        photo_bytes = await photo.read() if photo else None
        output_bytes = fill_form(docx_bytes, user_info_text, photo_bytes)
        headers = {"Content-Disposition": "attachment; filename=filled.docx"}
        return StreamingResponse(
            iter([output_bytes]),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers=headers
        )
    except Exception as e:
        print(f"❌ Error in /process: {e}")
        import traceback
        traceback.print_exc()
        return JSONResponse(status_code=500, content={"error": str(e)})

# 导出 app 变量，供 uvicorn 使用
app_instance = app

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("server:app_instance", host="0.0.0.0", port=port, reload=True)
