# CloudConvertX
<img width="1055" height="1491" alt="Project_Intro" src="https://github.com/user-attachments/assets/3e9c024f-fc78-4554-aec3-c9d65aeb548d" />


**CloudConvertX** is a privacy-focused, serverless AWS application that converts Microsoft Word `.docx` documents into print-ready PDF files.

The project uses **AWS Lambda, Amazon S3, Docker, AWS CodeBuild, Amazon ECR, and LibreOffice**. Files are uploaded directly from the browser to a private S3 bucket using a temporary pre-signed URL. A Lambda function then converts the document inside a containerized LibreOffice environment and stores the resulting PDF in S3.


---

## Features

- Convert `.docx` Word documents to PDF.
- Serverless architecture with AWS Lambda.
- Direct browser-to-S3 uploads using pre-signed URLs.
- Private S3 bucket for uploaded and converted documents
- LibreOffice running inside a Lambda container image
- Automatic PDF download through a temporary pre-signed URL
- S3 lifecycle rule automatically deletes files after 1 day
- No always-on application server
- Static frontend hosted from Amazon S3
- Maximum frontend upload size: **20 MB**
- Lambda converter configured with **3008 MB memory**, **2048 MB ephemeral storage**, and a **3-minute timeout**

---

## Architecture

```text
                         ┌─────────────────────────┐
                         │       User Browser       │
                         │       index.html         │
                         └────────────┬────────────┘
                                      │
                                      │ 1. Request upload URL
                                      ▼
                         ┌─────────────────────────┐
                         │ Lambda: docpdf-presign  │
                         │ Python 3.12              │
                         └────────────┬────────────┘
                                      │
                                      │ 2. Pre-signed PUT URL
                                      ▼
                         ┌─────────────────────────┐
                         │ Private S3 Files Bucket │
                         │ uploads/{id}.docx       │
                         └────────────┬────────────┘
                                      │
                                      │ 3. S3 Object Created
                                      ▼
                         ┌─────────────────────────┐
                         │ Lambda: docpdf-converter│
                         │ Container Image         │
                         │ Docker + LibreOffice    │
                         └────────────┬────────────┘
                                      │
                                      │ 4. Convert DOCX → PDF
                                      ▼
                         ┌─────────────────────────┐
                         │ Private S3 Files Bucket │
                         │ converted/{id}.pdf      │
                         └────────────┬────────────┘
                                      │
                                      │ 5. Poll status
                                      ▼
                         ┌─────────────────────────┐
                         │ Lambda: docpdf-presign  │
                         │ Returns signed GET URL  │
                         └────────────┬────────────┘
                                      │
                                      │ 6. Temporary download URL
                                      ▼
                         ┌─────────────────────────┐
                         │       User Browser       │
                         │      Download PDF        │
                         └─────────────────────────┘
```

### Supporting AWS services

```text
S3 Static Website
      │
      └── Hosts index.html

AWS Lambda
      ├── docpdf-presign
      └── docpdf-converter

Amazon ECR
      └── docpdf-converter:latest

AWS CodeBuild
      └── Builds and pushes the Lambda container image

S3 Lifecycle
      └── Deletes objects after 1 day
```

---

# AWS Deployment Guide

## Prerequisites

You need:

- An AWS account
- Access to the AWS Management Console
- A browser
- A local text editor
- A Word `.docx` file for testing
- AWS Region set to **us-east-1 (N. Virginia)**

This deployment uses the AWS Console and does not require Docker to be installed locally. The converter Docker image is built using AWS CodeBuild.

---

# Step 1 — Create the Private S3 Files Bucket

Open:

**AWS Console → S3 → Create bucket**

Use:

```text
Bucket name: docpdf-files-YOURNAME123
Region: us-east-1
```

Replace `YOURNAME123` with a unique value because S3 bucket names must be globally unique.

### Public access

Keep:

```text
Block all public access: Enabled
```

Then click **Create bucket**.

The files bucket must remain private.

---

## Configure S3 CORS

Open the bucket:

**Permissions → Cross-origin resource sharing (CORS) → Edit**

Use:

```json
[
  {
    "AllowedHeaders": ["*"],
    "AllowedMethods": ["GET", "PUT"],
    "AllowedOrigins": ["*"],
    "ExposeHeaders": [],
    "MaxAgeSeconds": 3000
  }
]
```

Save the CORS configuration.

This allows the browser to upload and download files directly through S3.

---

## Configure automatic deletion

Go to:

**Management → Create lifecycle rule**

Use:

```text
Rule name: expire-1d
Scope: Apply to all objects in the bucket
Action: Expire current versions of objects
Days after object creation: 1
```

Create the rule.

This provides automatic cleanup of uploaded and converted files after one day.

---

# Step 2 — Create the IAM Role

Go to:

**AWS Console → IAM → Roles → Create role**

### Trusted entity

```text
Trusted entity: AWS service
Use case: Lambda
```

Attach:

```text
AWSLambdaBasicExecutionRole
```

Create the role with:

```text
Role name: docpdf-lambda-role
```

---

## Add S3 permissions

Open:

**IAM → Roles → docpdf-lambda-role → Add permissions → Create inline policy → JSON**

Use the following policy and replace the bucket name:

```json
{
  "Version": "2012-10-17",
  "Statement": [
    {
      "Effect": "Allow",
      "Action": [
        "s3:GetObject",
        "s3:PutObject"
      ],
      "Resource": "arn:aws:s3:::docpdf-files-YOURNAME123/*"
    }
  ]
}
```

Policy name:

```text
s3access
```

Create the policy.

---

# Step 3 — Create Lambda 1: `docpdf-presign`

Go to:

**AWS Console → Lambda → Create function → Author from scratch**

Use:

```text
Function name: docpdf-presign
Runtime: Python 3.12
Architecture: x86_64
Execution role: Use an existing role
Role: docpdf-lambda-role
```

Create the function.

---

## Lambda code

Open:

**Code → lambda_function.py**

Delete the existing code and use:

```python
import json, os, re, uuid
import boto3
from botocore.config import Config
from botocore.exceptions import ClientError

REGION = os.environ["AWS_REGION"]
BUCKET = os.environ["FILES_BUCKET"]
DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

s3 = boto3.client(
    "s3",
    region_name=REGION,
    endpoint_url=f"https://s3.{REGION}.amazonaws.com",
    config=Config(signature_version="s3v4"),
)

def resp(code, body):
    # CORS headers come from the Function URL settings.
    return {
        "statusCode": code,
        "headers": {
            "Content-Type": "application/json"
        },
        "body": json.dumps(body)
    }

def lambda_handler(event, context):
    q = event.get("queryStringParameters") or {}
    action = q.get("action")

    if action == "upload":
        file_id = uuid.uuid4().hex

        url = s3.generate_presigned_url(
            "put_object",
            Params={
                "Bucket": BUCKET,
                "Key": f"uploads/{file_id}.docx",
                "ContentType": DOCX
            },
            ExpiresIn=300,
        )

        return resp(200, {
            "id": file_id,
            "url": url
        })

    if action == "status":
        file_id = q.get("id", "")

        if not re.fullmatch(r"[0-9a-f]{32}", file_id):
            return resp(400, {
                "error": "bad id"
            })

        key = f"converted/{file_id}.pdf"

        try:
            s3.head_object(
                Bucket=BUCKET,
                Key=key
            )
        except ClientError as e:
            if e.response["Error"]["Code"] in (
                "404",
                "NoSuchKey",
                "403"
            ):
                return resp(200, {
                    "ready": False
                })
            raise

        url = s3.generate_presigned_url(
            "get_object",
            Params={
                "Bucket": BUCKET,
                "Key": key,
                "ResponseContentDisposition":
                    'attachment; filename="converted.pdf"',
                "ResponseContentType": "application/pdf"
            },
            ExpiresIn=300,
        )

        return resp(200, {
            "ready": True,
            "url": url
        })

    return resp(400, {
        "error": "unknown action"
    })
```

Click **Deploy**.

---

## Configure the environment variable

Go to:

**Configuration → Environment variables → Edit → Add**

Add:

```text
Key: FILES_BUCKET
Value: docpdf-files-YOURNAME123
```

Save.

---

## Configure Lambda timeout

Go to:

**Configuration → General configuration → Edit**

Set:

```text
Timeout: 10 seconds
```

Save.

---

## Create the Lambda Function URL

Go to:

**Configuration → Function URL → Create function URL**

Use:

```text
Auth type: NONE
```

Enable CORS and configure:

```text
Allow origin: *
Allow methods: GET
Allow headers: content-type
Max age: 3000
```

Save.

The Function URL will look similar to:

```text
https://abcd1234.lambda-url.us-east-1.on.aws/
```

Keep this URL. It is required by the frontend.

> **Security note:** The Function URL is intentionally public in this project because it acts as the browser-facing API. The S3 files bucket itself remains private.

---

## Test the Function URL

Open:

```text
YOUR_FUNCTION_URL?action=upload
```

You should receive JSON containing an `id` and a pre-signed S3 upload `url`.

If you receive `403 Forbidden`, check:

**Configuration → Permissions → Resource-based policy statements**

Confirm that the function URL has permission for public invocation with the selected `NONE` authentication configuration.

---

# Step 4 — Build the Converter Container

LibreOffice is too large for a normal ZIP-based Lambda deployment, so the converter runs as a **Lambda container image**.

The image contains:

- Python 3.12
- LibreOffice Writer
- Compatible fonts
- AWS Lambda Runtime Interface Client
- boto3

The image is built using **AWS CodeBuild** and stored in **Amazon ECR**.

---

## 4A — Create the converter files

Create a local folder:

```text
converter/
├── Dockerfile
└── app.py
```

### Dockerfile

Create a file named exactly:

```text
Dockerfile
```

Use:

```dockerfile
FROM public.ecr.aws/docker/library/python:3.12-slim-bookworm

# LibreOffice + fonts metric-compatible with Word fonts
# Carlito = Calibri
# Caladea = Cambria
# Liberation = Arial / Times New Roman / Courier New
RUN apt-get update && apt-get install -y --no-install-recommends \
      libreoffice-writer \
      fonts-liberation fonts-crosextra-carlito fonts-crosextra-caladea \
      fonts-dejavu fonts-noto-core fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

RUN pip install --no-cache-dir awslambdaric boto3

WORKDIR /var/task

COPY app.py .

ENTRYPOINT ["python", "-m", "awslambdaric"]
CMD ["app.handler"]
```

---

## `app.py`

Create:

```text
app.py
```

Use:

```python
import os
import shutil
import subprocess
import urllib.parse

import boto3

s3 = boto3.client("s3")

def handler(event, context):
    for rec in event["Records"]:
        bucket = rec["s3"]["bucket"]["name"]
        key = urllib.parse.unquote_plus(
            rec["s3"]["object"]["key"]
        )

        # uploads/ID.docx
        base = os.path.splitext(
            os.path.basename(key)
        )[0]

        work = f"/tmp/{base}"
        os.makedirs(work, exist_ok=True)

        src = f"{work}/{base}.docx"

        s3.download_file(
            bucket,
            key,
            src
        )

        subprocess.run(
            [
                "soffice",
                "--headless",
                "--norestore",
                "--nologo",
                "-env:UserInstallation=file:///tmp/lo-profile",
                "--convert-to",
                "pdf:writer_pdf_Export",
                "--outdir",
                work,
                src
            ],
            check=True,
            timeout=150,
            env={
                **os.environ,
                "HOME": "/tmp"
            },
        )

        s3.upload_file(
            f"{work}/{base}.pdf",
            bucket,
            f"converted/{base}.pdf",
            ExtraArgs={
                "ContentType": "application/pdf"
            }
        )

        shutil.rmtree(
            work,
            ignore_errors=True
        )

    return {
        "ok": True
    }
```

Zip the two files:

```text
converter.zip
```

The ZIP must have `Dockerfile` at its top level:

```text
converter.zip
├── Dockerfile
└── app.py
```

---

# 4B — Upload `converter.zip`

Open:

**S3 → docpdf-files-YOURNAME123 → Upload**

Upload:

```text
converter.zip
```

Keep it at the bucket root.

The lifecycle rule may delete this file after one day, which is acceptable because it is only used as the CodeBuild source.

---

# 4C — Create the ECR Repository

Open:

**AWS Console → ECR → Private repositories → Create repository**

Use:

```text
Repository name: docpdf-converter
```

Create the repository.

Also note your **12-digit AWS Account ID** from the AWS console.

---

# 4D — Create the CodeBuild Project

Open:

**AWS Console → CodeBuild → Create project**

Use:

```text
Project name: docpdf-converter-build
```

### Source

```text
Source provider: Amazon S3
Bucket: docpdf-files-YOURNAME123
S3 object key: converter.zip
```

### Environment

Use a managed environment:

```text
Operating system: Amazon Linux
Runtime: Standard
Image: latest AWS CodeBuild Amazon Linux standard image
Environment type: Linux
```

Enable:

```text
Privileged
```

This is required because CodeBuild builds a Docker image.

Use a new service role and keep the suggested role name.

---

## Environment variable

Add:

```text
Name: ACCOUNT_ID
Value: YOUR_12_DIGIT_AWS_ACCOUNT_ID
Type: Plaintext
```

---

## Buildspec

Choose:

**Insert build commands → Switch to editor**

Use:

```yaml
version: 0.2

phases:
  pre_build:
    commands:
      - aws ecr get-login-password --region $AWS_DEFAULT_REGION | docker login --username AWS --password-stdin $ACCOUNT_ID.dkr.ecr.$AWS_DEFAULT_REGION.amazonaws.com

  build:
    commands:
      - docker build --provenance=false -t $ACCOUNT_ID.dkr.ecr.$AWS_DEFAULT_REGION.amazonaws.com/docpdf-converter:latest .

  post_build:
    commands:
      - docker push $ACCOUNT_ID.dkr.ecr.$AWS_DEFAULT_REGION.amazonaws.com/docpdf-converter:latest
```

Create the build project.

---

# 4E — Give CodeBuild Permission

Open:

**IAM → Roles**

Open the CodeBuild service role. It will normally start with something similar to:

```text
codebuild-docpdf-converter-build-service-role
```

Attach:

```text
AmazonEC2ContainerRegistryPowerUser
```

If the build fails during source download, also attach:

```text
AmazonS3ReadOnlyAccess
```

---

## Start the build

Go to:

**CodeBuild → docpdf-converter-build → Start build**

The build normally takes several minutes.

Wait until the build status is:

```text
Succeeded
```

Then open:

**ECR → docpdf-converter**

You should see:

```text
latest
```

as the image tag.

---

# Step 5 — Create Lambda 2: `docpdf-converter`

Open:

**AWS Console → Lambda → Create function → Container image**

Use:

```text
Function name: docpdf-converter
Container image: docpdf-converter:latest
Architecture: x86_64
Execution role: docpdf-lambda-role
```

Create the function.

---

## Configure Lambda resources

Go to:

**Configuration → General configuration → Edit**

Use:

```text
Memory: 3008 MB
Ephemeral storage: 2048 MB
Timeout: 3 minutes
```

Save.

These settings provide enough temporary storage and execution time for LibreOffice document conversion.

---

# Step 6 — Connect S3 to the Converter Lambda

Open:

**S3 → docpdf-files-YOURNAME123 → Properties**

Find:

**Event notifications → Create event notification**

Use:

```text
Name: on-docx-upload
Prefix: uploads/
Suffix: .docx
```

### Event type

Select:

```text
All object create events
```

### Destination

Select:

```text
Lambda function
Function: docpdf-converter
```

Save.

### Why the prefix matters

Uploaded Word files are stored under:

```text
uploads/
```

Converted PDFs are stored under:

```text
converted/
```

Therefore, the S3 notification only reacts to `.docx` files under `uploads/` and does not re-trigger when the PDF is created.

---

# Step 7 — Create the Static Website

The frontend is a simple HTML/CSS/JavaScript application hosted by Amazon S3.

Create:

```text
index.html
```

Replace:

```text
PASTE_FUNCTION_URL_HERE
```

with the Function URL from Step 3.

Use:

```html
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Word to PDF</title>

<style>
  body {
    font-family: system-ui, sans-serif;
    background: #f4f6f8;
    margin: 0;
    display: flex;
    justify-content: center;
    padding: 48px 16px;
  }

  .card {
    background: #fff;
    max-width: 460px;
    width: 100%;
    padding: 32px;
    border-radius: 12px;
    box-shadow: 0 2px 12px rgba(0,0,0,.08);
  }

  h1 {
    margin: 0 0 8px;
    font-size: 22px;
  }

  p {
    color: #555;
    margin: 0 0 20px;
  }

  input[type=file] {
    width: 100%;
    margin-bottom: 16px;
  }

  button,
  a.dl {
    display: inline-block;
    background: #2563eb;
    color: #fff;
    border: 0;
    padding: 10px 20px;
    border-radius: 8px;
    font-size: 15px;
    cursor: pointer;
    text-decoration: none;
  }

  button:disabled {
    background: #9db7ef;
    cursor: not-allowed;
  }

  #status {
    margin: 16px 0;
    color: #333;
  }

  .err {
    color: #b91c1c;
  }
</style>
</head>

<body>
<div class="card">

  <h1>Word to PDF</h1>

  <p>
    Upload a .docx file (max 20 MB) and download the PDF.
  </p>

  <input type="file" id="file" accept=".docx">

  <button id="go">Convert</button>

  <div id="status"></div>
  <div id="result"></div>

</div>

<script>
const API = "PASTE_FUNCTION_URL_HERE".replace(/\/$/, "");

const DOCX =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

const MAX = 20 * 1024 * 1024;

const $ = id => document.getElementById(id);

const say = (t, err) => {
  $("status").textContent = t;
  $("status").className = err ? "err" : "";
};

const sleep = ms =>
  new Promise(r => setTimeout(r, ms));

$("go").onclick = async () => {

  const f = $("file").files[0];

  $("result").innerHTML = "";

  if (!f) {
    return say("Please choose a .docx file.", true);
  }

  if (!f.name.toLowerCase().endsWith(".docx")) {
    return say("Only .docx files are supported.", true);
  }

  if (f.size > MAX) {
    return say("File is larger than 20 MB.", true);
  }

  $("go").disabled = true;

  try {

    say("Preparing upload...");

    const r1 = await fetch(`${API}/?action=upload`);

    if (!r1.ok) {
      throw new Error("Could not get upload URL");
    }

    const { id, url } = await r1.json();

    say("Uploading...");

    const up = await fetch(url, {
      method: "PUT",
      headers: {
        "Content-Type": DOCX
      },
      body: f
    });

    if (!up.ok) {
      throw new Error("Upload failed (" + up.status + ")");
    }

    say("Converting to PDF...");

    for (let i = 0; i < 60; i++) {

      await sleep(3000);

      const r2 =
        await fetch(`${API}/?action=status&id=${id}`);

      const s = await r2.json();

      if (s.ready) {

        say("Done! Your PDF is ready.");

        const name =
          f.name.replace(/\.docx$/i, "") + ".pdf";

        $("result").innerHTML =
          `<a class="dl" href="${s.url}" download="${name}">
             Download PDF
           </a>`;

        return;
      }
    }

    throw new Error(
      "Conversion timed out, please try again."
    );

  } catch (e) {

    say(e.message, true);

  } finally {

    $("go").disabled = false;
  }
};
</script>

</body>
</html>
```

The browser:

1. Requests a pre-signed upload URL.
2. Uploads the `.docx` directly to S3.
3. Polls the status endpoint.
4. Receives a temporary PDF download URL.
5. Downloads the converted PDF.

---

# Step 7A — Create the Website S3 Bucket

Create another S3 bucket:

**AWS Console → S3 → Create bucket**

Use:

```text
Bucket name: docpdf-site-YOURNAME123
Region: us-east-1
```

Because this bucket hosts the public static website, configure its public-access settings separately from the private files bucket.

For the static website bucket:

```text
Block all public access: Disabled
```

Acknowledge the warning and create the bucket.

---

## Upload `index.html`

Open the website bucket:

**Upload → Add files → index.html → Upload**

---

## Enable static website hosting

Go to:

**Properties → Static website hosting → Edit**

Configure:

```text
Hosting type: Host a static website
Index document: index.html
```

Save.

---

## Add the bucket policy

Go to:

**Permissions → Bucket policy → Edit**

Replace the bucket name and use:

```json
{
  "Version": "2012-10-17",
  "Statement": [
    {
      "Sid": "PublicRead",
      "Effect": "Allow",
      "Principal": "*",
      "Action": "s3:GetObject",
      "Resource": "arn:aws:s3:::docpdf-site-YOURNAME123/*"
    }
  ]
}
```

Save.

---

## Open the website

Go back to:

**Properties → Static website hosting**

Copy the **Bucket website endpoint**.

It will use the S3 website endpoint format and can be opened in a browser.

---

# Step 8 — Test the Application

Open the website endpoint.

Choose a `.docx` file and click:

```text
Convert
```

The expected flow is:

```text
Preparing upload...
        ↓
Uploading...
        ↓
Converting to PDF...
        ↓
Done! Your PDF is ready.
        ↓
Download PDF
```

The first conversion can take approximately **20–40 seconds** because the converter Lambda may experience a cold start. Later conversions can be faster.

---

# S3 Object Structure

The private files bucket uses separate prefixes:

```text
docpdf-files-YOURNAME123/
│
├── uploads/
│   └── {file-id}.docx
│
├── converted/
│   └── {file-id}.pdf
│
└── converter.zip
```

The important separation is:

```text
uploads/      → input Word documents
converted/    → generated PDF documents
```

The S3 trigger only watches:

```text
uploads/*.docx
```

so generated PDFs do not trigger another conversion.

---

# API Flow

The `docpdf-presign` Lambda exposes two actions.

## Generate upload URL

Request:

```text
GET /?action=upload
```

Example response:

```json
{
  "id": "32-character-file-id",
  "url": "https://..."
}
```

The browser then uploads the `.docx` to the returned URL using HTTP `PUT`.

---

## Check conversion status

Request:

```text
GET /?action=status&id=FILE_ID
```

While conversion is running:

```json
{
  "ready": false
}
```

When the PDF is available:

```json
{
  "ready": true,
  "url": "https://..."
}
```

The returned PDF URL is temporary and expires after 5 minutes.

---

# Configuration Summary

| Component | Configuration |
|---|---|
| AWS Region | `us-east-1` (N. Virginia) |
| Files bucket | `docpdf-files-YOURNAME123` |
| Website bucket | `docpdf-site-YOURNAME123` |
| IAM role | `docpdf-lambda-role` |
| Presign Lambda | `docpdf-presign` |
| Presign runtime | Python 3.12 |
| Presign architecture | x86_64 |
| Presign timeout | 10 seconds |
| Converter Lambda | `docpdf-converter` |
| Converter architecture | x86_64 |
| Converter memory | 3008 MB |
| Converter ephemeral storage | 2048 MB |
| Converter timeout | 3 minutes |
| ECR repository | `docpdf-converter` |
| CodeBuild project | `docpdf-converter-build` |
| Upload prefix | `uploads/` |
| PDF prefix | `converted/` |
| Upload limit | 20 MB |
| Presigned URL expiry | 5 minutes |
| S3 lifecycle cleanup | 1 day |

---

# Privacy and Security Model

CloudConvertX is designed around temporary access and automatic cleanup.

### Private document storage

The files bucket is configured with:

```text
Block all public access: Enabled
```

Users do not receive direct public access to the S3 bucket.

### Temporary upload access

The frontend receives a short-lived pre-signed S3 `PUT` URL.

```text
Expiry: 5 minutes
```

### Temporary download access

When conversion completes, the frontend receives a short-lived pre-signed S3 `GET` URL.

```text
Expiry: 5 minutes
```

### Automatic cleanup

The S3 lifecycle rule removes objects after:

```text
1 day
```

### Direct browser-to-S3 transfer

The application does not send the document through a traditional always-on web server.

The main data path is:

```text
Browser → S3
S3 → Lambda
Lambda → S3
S3 → Browser
```

---

# AWS Services Used

| AWS Service | Purpose |
|---|---|
| Amazon S3 | Private file storage and static website hosting |
| AWS Lambda | API/presigned URLs and document conversion |
| AWS IAM | Lambda and S3 permissions |
| AWS CodeBuild | Builds the Docker container image |
| Amazon ECR | Stores the converter container image |
| S3 Lifecycle | Automatic file cleanup |
| S3 Event Notifications | Starts conversion after `.docx` upload |

External conversion software:

```text
LibreOffice Writer
```

---

# Project Structure

A recommended GitHub repository structure is:

```text
CloudConvertX/
│
├── README.md
│
├── frontend/
│   └── index.html
│
├── presign/
│   └── lambda_function.py
│
└── converter/
    ├── Dockerfile
    └── app.py
```

The `converter.zip` deployment package and generated build artifacts do not need to be committed if you prefer to keep the repository source-only.

---

# Repository Setup

After creating the project files locally:

```bash
git init
git add .
git commit -m "Initial CloudConvertX serverless document converter"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/CloudConvertX.git
git push -u origin main
```

Replace:

```text
YOUR_USERNAME
```

with your GitHub username.

---

# Important Deployment Notes

## 1. Replace bucket placeholders

Anywhere you see:

```text
docpdf-files-YOURNAME123
docpdf-site-YOURNAME123
```

replace them with your actual globally unique S3 bucket names.

---

## 2. Replace the Function URL

In `index.html`, replace:

```javascript
const API = "PASTE_FUNCTION_URL_HERE".replace(/\/$/, "");
```

with your actual Lambda Function URL, for example:

```javascript
const API =
  "https://abcd1234.lambda-url.us-east-1.on.aws/"
  .replace(/\/$/, "");
```

Do not commit credentials or AWS access keys to GitHub.

---

## 3. Keep the files bucket private

Do **not** disable Block Public Access on:

```text
docpdf-files-YOURNAME123
```

Only the static website bucket is configured for public website hosting.

---

## 4. AWS Region

All resources in this deployment should be created in:

```text
us-east-1
```

which is **N. Virginia**.

If resources are accidentally created in another region, the deployment can fail because Lambda, S3, ECR, CodeBuild, and related resources need to be configured consistently.

---

# Troubleshooting

## Upload URL returns an error

Check:

- `FILES_BUCKET` environment variable
- S3 bucket name
- Lambda IAM role
- S3 `PutObject` permission
- S3 CORS configuration
- Lambda Function URL configuration

---

## S3 upload fails

Check:

- S3 CORS allows `PUT`
- Browser is sending the correct DOCX `Content-Type`
- Pre-signed URL has not expired
- Files bucket name and region are correct

---

## Converter Lambda does not run

Check:

- S3 event notification
- Prefix is exactly:

```text
uploads/
```

- Suffix is exactly:

```text
.docx
```

- Destination is `docpdf-converter`
- Lambda execution role has S3 access

---

## CodeBuild fails

Check:

- CodeBuild has `AmazonEC2ContainerRegistryPowerUser`
- Source S3 object exists
- `ACCOUNT_ID` environment variable is correct
- Privileged mode is enabled
- ECR repository exists
- CodeBuild and ECR are in `us-east-1`

If the failure occurs while downloading the S3 source, check whether the CodeBuild role needs:

```text
AmazonS3ReadOnlyAccess
```

---

## PDF conversion fails

Check:

- Lambda memory is `3008 MB`
- Ephemeral storage is `2048 MB`
- Timeout is `3 minutes`
- ECR image tag is `latest`
- LibreOffice is installed in the container
- Lambda architecture is `x86_64`

---

## Website does not open

Check:

- `index.html` was uploaded
- Static website hosting is enabled
- Website bucket is configured for public read
- Bucket policy uses the correct bucket name
- The website endpoint is the one shown by S3

---

# Limitations

The current implementation is intentionally simple.

- Only `.docx` files are accepted.
- Frontend upload size is limited to 20 MB.
- Conversion status is handled by polling.
- The static website uses an S3 website endpoint.
- The API Function URL is publicly reachable.
- The files bucket is automatically cleaned after one day.
- Conversion performance can vary because Lambda may start from a cold state.
- LibreOffice conversion fidelity can vary for complex Word documents.

---

# Future Improvements

Potential improvements include:

- Amazon CloudFront for HTTPS delivery
- AWS WAF for additional API protection
- Amazon Cognito authentication
- More restrictive CORS settings
- Per-user document isolation
- Better job status tracking with DynamoDB
- Amazon EventBridge or SQS for more robust job processing
- CloudWatch alarms and dashboards
- Support for additional document formats
- Custom domain and HTTPS frontend
- Improved conversion error reporting
- Virus/malware scanning before conversion

---

# Project Highlights

CloudConvertX demonstrates a practical serverless document-processing architecture using AWS.

### Core concepts demonstrated

```text
Serverless Architecture
        +
Direct-to-S3 Uploads
        +
Pre-signed URLs
        +
Lambda Functions
        +
Docker Container Images
        +
LibreOffice Automation
        +
AWS CodeBuild
        +
Amazon ECR
        +
S3 Event Notifications
        +
Automatic Lifecycle Cleanup
```

The result is a lightweight document conversion platform without a permanently running application server.

---

# License

Add the license you want to use for your GitHub repository.

For example:

```text
MIT License
```

if you choose to release the project under the MIT License.

---

## Author

**CloudConvertX**

A serverless AWS project for converting Word documents into PDF files with a privacy-focused temporary-storage workflow.
