from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
import io
import datainsert 

app = FastAPI()

@app.post('/generate-hwp')
async def generate_hwp(request: Request):
    try:
        json_data = await request.json()
    except Exception:
        raise HTTPException(status_code=400, detail='유효한 JSON 데이터를 제공해 주세요.')

    if not json_data:
        raise HTTPException(status_code=400, detail='JSON 데이터가 제공되지 않았습니다.')

    # HWP 파일 생성 및 바이너리 데이터 얻기
    hwp_content = datainsert.create_hwp_file(json_data)
    if hwp_content is None:
        raise HTTPException(status_code=500, detail='HWP 파일 생성에 실패했습니다.')

    # 바이너리 데이터를 클라이언트에게 반환
    return StreamingResponse(
        io.BytesIO(hwp_content),
        media_type='application/octet-stream',
        headers={'Content-Disposition': 'attachment; filename="generated.hwp"'}
    )