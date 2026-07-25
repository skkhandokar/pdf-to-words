FROM python:3.10-slim

# Working Directory সেট করা
WORKDIR /app

# Python Output & Bytecode অপটিমাইজেশন
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

# ১. সিস্টেম ডিপেন্ডেন্সি ইনস্টলেশন (Office to PDF এবং PDF to PPT-এর জন্য আবশ্যক)
RUN apt-get update && apt-get install -y \
    libreoffice \
    poppler-utils \
    ffmpeg \
    libsm6 \
    libxext6 \
    libgl1 \
    libglib2.0-0 \
    curl \
    gcc \
    && rm -rf /var/lib/apt/lists/*

# ২. Requirements ইনস্টল
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# ৩. প্রজেক্ট ফাইল কপি
COPY . .

# ৪. Render-এর অটোমেটিক PORT অনুযায়ী FastAPI Uvicorn চালুকরণ
CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port ${PORT:-10000} --workers 1"]