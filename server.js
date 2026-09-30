```javascript
const express = require("express");
const multer = require("multer");
const fs = require("fs");
const path = require("path");
const { execFile } = require("child_process");

const app = express();

app.use((req, res, next) => {
  res.header("Access-Control-Allow-Origin", "*");
  res.header("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
  res.header("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") {
    return res.sendStatus(204);
  }

  next();
});

const upload = multer({
  dest: "/tmp/uploads"
});

const jobs = new Map();

app.get("/", (req, res) => {
  res.json({
    success: true,
    service: "iLovePDF4 Free Conversion Engine",
    status: "online"
  });
});

app.get("/status/:jobId", (req, res) => {
  const job = jobs.get(req.params.jobId);

  if (!job) {
    return res.status(404).json({
      status: "error",
      error: "Job not found."
    });
  }

  if (job.status === "finished") {
    return res.json({
      status: "finished",
      url: `${getBaseUrl(req)}/download/${req.params.jobId}`,
      filename: job.filename
    });
  }

  if (job.status === "error") {
    return res.json({
      status: "error",
      error: job.error || "Conversion failed."
    });
  }

  return res.json({
    status: "processing"
  });
});

app.get("/download/:jobId", (req, res) => {
  const job = jobs.get(req.params.jobId);

  if (!job || job.status !== "finished" || !job.outputBuffer) {
    return res.status(404).send("Output file no longer exists.");
  }

  res.setHeader(
    "Content-Type",
    job.contentType || "application/pdf"
  );

  res.setHeader(
    "Content-Disposition",
    `attachment; filename="${job.filename}"`
  );

  res.setHeader(
    "Content-Length",
    job.outputBuffer.length
  );

  res.setHeader(
    "Cache-Control",
    "no-store, no-cache, must-revalidate"
  );

  res.setHeader(
    "Pragma",
    "no-cache"
  );

  res.send(job.outputBuffer);
});

app.post("/convert", upload.single("file"), (req, res) => {
  if (!req.file) {
    return res.status(400).json({
      success: false,
      error: "No file uploaded."
    });
  }

  const tool = String(req.body.tool || "")
    .trim()
    .toLowerCase();

  const jobId = String(req.body.jobId || "");

  if (!jobId) {
    fs.unlink(req.file.path, () => {});

    return res.status(400).json({
      success: false,
      error: "Missing job id."
    });
  }

  const allowedTools = {
    "word-to-pdf": "pdf",
    "powerpoint-to-pdf": "pdf",
    "excel-to-pdf": "pdf"
  };

  const outputFormat = allowedTools[tool];

  if (!outputFormat) {
    fs.unlink(req.file.path, () => {});

    return res.status(400).json({
      success: false,
      error: "Unsupported conversion tool: " + tool
    });
  }

  const originalName = String(
    req.body.filename ||
    req.file.originalname ||
    "file"
  );

  const safeName = path.basename(originalName);

  const inputPath = path.join(
    "/tmp/uploads",
    `${jobId}-${safeName}`
  );

  const outputDir = path.join(
    "/tmp/output",
    jobId
  );

  fs.mkdirSync(outputDir, {
    recursive: true
  });

  fs.renameSync(
    req.file.path,
    inputPath
  );

  const outputFilename =
    `${path.parse(safeName).name}.${outputFormat}`;

  jobs.set(jobId, {
    status: "processing",
    filename: outputFilename
  });

  res.json({
    success: true,
    jobId,
    status: "processing"
  });

  execFile(
    "libreoffice",
    [
      "--headless",
      "--convert-to",
      outputFormat,
      "--outdir",
      outputDir,
      inputPath
    ],
    (error) => {
      try {
        fs.unlinkSync(inputPath);
      } catch {}

      if (error) {
        jobs.set(jobId, {
          status: "error",
          error: "LibreOffice conversion failed."
        });

        return;
      }

      let files = [];

      try {
        files = fs.readdirSync(outputDir);
      } catch {
        jobs.set(jobId, {
          status: "error",
          error: "Unable to read conversion output."
        });

        return;
      }

      if (!files.length) {
        jobs.set(jobId, {
          status: "error",
          error: "Conversion output was not created."
        });

        return;
      }

      const outputFile = path.join(
        outputDir,
        files[0]
      );

      try {
        const outputBuffer =
          fs.readFileSync(outputFile);

        jobs.set(jobId, {
          status: "finished",
          outputBuffer,
          filename: outputFilename,
          contentType: "application/pdf"
        });

        try {
          fs.unlinkSync(outputFile);
        } catch {}

        setTimeout(() => {
          jobs.delete(jobId);
        }, 10 * 60 * 1000);

      } catch {
        jobs.set(jobId, {
          status: "error",
          error: "Unable to read conversion output."
        });
      }
    }
  );
});

function getBaseUrl(req) {
  return `${req.protocol}://${req.get("host")}`;
}

const PORT =
  process.env.PORT || 10000;

app.listen(
  PORT,
  "0.0.0.0",
  () => {
    console.log(
      `Conversion engine running on port ${PORT}`
    );
  }
);
```
