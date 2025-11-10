# PDF Processor API

A FastAPI service for merging and compressing PDF files.

## Setup
1. Clone repo: `git clone <repo-url> && cd pdf-processor-api`
2. Create venv: `python -m venv venv && source venv/bin/activate`
3. Install deps: `pip install -r requirements.txt`
4. Run: `uvicorn src.main:app --reload --host 0.0.0.0 --port 8000`

## API Endpoints
- POST `/merge`: Upload multiple PDFs to merge.
- POST `/compress`: Upload one PDF to compress.
- GET `/download/{filename}`: Download processed file.



python3 -m venv venv
source venv/bin/activate
(venv) ➜  pdf-tools git:(main)
pip install --upgrade pip
pip install -r src/requirements.txt
uvicorn src.main:app --reload


source venv/bin/activate && pip install -r src/requirements.txt && uvicorn src.main:app --reload
## Testing
Use http://localhost:8000/docs for interactive Swagger UI.

run command
uvicorn src.main:app --reload --host 0.0.0.0 --port 8000


python3 -m ensurepip --upgrade
python3 -m venv venv
source venv/bin/activate
/Users/ktnaveen/Documents/Python-pdf-merge/pdf-tools/venv/bin/python3 -m pip install --upgrade pip
pip install PyPDF2  # Or pip3 if needed
pip freeze > requirements.txt

git pull origin main
 git add .
 git commit -m "all the commands"
 git push origin main
pip freeze > requirements.txt