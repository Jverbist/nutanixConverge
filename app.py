from fastapi import FastAPI, UploadFile, File
import shutil

app = FastAPI()

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


@app.get("/")
async def root():
    return {"message": "Hello FastAPI"}

@app.post("/upload")
async def upload_file(file: UploadFile = File(...)):
    with open(f"uploaded_files/{file.filename}", "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
    return {"filename": file.filename, "status": "uploaded"}


