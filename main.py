from fastapi import FastAPI, UploadFile, File
from paddleocr import PaddleOCR
from openpyxl import Workbook
import cv2
import numpy as np
import io
from fastapi.responses import StreamingResponse

app = FastAPI()

ocr = PaddleOCR(
    use_angle_cls=True,
    lang="japan",
    show_log=False
)

@app.post("/extract")
async def extract(images: list[UploadFile] = File(...)):
    wb = Workbook()
    ws = wb.active
    ws.append(["Rank", "Alliance", "Name", "Warzone", "Power"])

    rank = 1

    for img in images:
        content = await img.read()
        np_img = np.frombuffer(content, np.uint8)
        image = cv2.imdecode(np_img, cv2.IMREAD_GRAYSCALE)

        result = ocr.ocr(image, cls=True)

        for line in result[0]:
            text = line[1][0]

            alliance = ""
            name = text
            power = ""
            warzone = ""

            if "[" in text and "]" in text:
                alliance = text[text.find("[")+1:text.find("]")]
                name = text[text.find("]")+1:]

            digits = "".join(c for c in text if c.isdigit() or c == ",")
            power = digits

            ws.append([rank, alliance, name.strip(), warzone, power])
            rank += 1

    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)

    return StreamingResponse(
        stream,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=output.xlsx"}
    )
