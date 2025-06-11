# Excel Protection Remover Web App

A web application that allows users to upload Excel files and remove their protection (both workbook and worksheet protection).

## Features

- Upload multiple Excel files (.xlsx, .xlsm)
- Remove workbook and worksheet protection
- Modern drag-and-drop interface with progress bars
- Download unprotected files individually
- Files are processed in memory and deleted after download
- Session isolation: only you can download your files

## Requirements

- Python 3.7 or higher
- Flask
- lxml
- Werkzeug

## Installation (Local)

1. Clone this repository:
```bash
git clone <repository-url>
cd xlsx_protection_remover
```

2. Create a virtual environment and activate it:
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install the required packages:
```bash
pip install -r requirements.txt
```

4. Run the application:
```bash
python app.py
```

5. Open your browser and go to `http://localhost:5000`

## Deployment

### Deploying to GitHub
- Make sure your `.gitignore` excludes secrets, venv, and temp files.
- Do **not** commit `.env` or any secrets.
- Push your code to a new GitHub repository.

### Deploying to Vercel
- Vercel is optimized for Node.js, but you can deploy Python using serverless functions.
- You may need to move your Flask app to `/api/index.py` and adjust imports/paths.
- Add a `vercel.json` file:
  ```json
  {
    "functions": {
      "api/index.py": {
        "runtime": "python3.9"
      }
    }
  }
  ```
- See [Vercel Python Flask Example](https://github.com/vercel/examples/tree/main/python/flask) for details.
- **Note:** Vercel has a 4MB request/response size limit and 10s execution time. For larger files, use [Render](https://render.com/), [Railway](https://railway.app/), or [Fly.io](https://fly.io/).

### Alternative: Render, Railway, Fly.io
- These platforms support Python web apps out of the box and are recommended for file uploads.

## Privacy Policy
- Uploaded files are stored only temporarily in memory and deleted after processing or download.
- No files are retained on the server after your session.
- Only you can download your processed files during your session.
- If you refresh or close your browser, your session and files may be lost.

## License

MIT License 