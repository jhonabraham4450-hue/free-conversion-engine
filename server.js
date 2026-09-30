const express = require("express");
const multer = require("multer");
const fs = require("fs");
const path = require("path");
const { execFile } = require("child_process");

const app = express();

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

  if (!job || job.status !== "finished" || !job.outputFile) {
    return res.status(404).send("File is not ready.");
  }

  if (!fs.existsSync(job.outputFile)) {
    return res.status(404).send("Output file no longer exists.");
  }

  res.download(
    job.outputFile,
    job.filename,
    () => {
      try {
        fs.unlinkSync(job.outputFile);
      } catch {}
    }
  );
});

app.post("/convert", upload.single("file"), (req, res) => {
  if (!req.file) {
    return res.status(400).json({
      success: false,
      error: "No file uploaded."
    });
  }

  const tool = String(req.body.tool || "");
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
      error: "Unsupported conversion tool."
    });
  }

  const originalName =
    String(req.body.filename || req.file.originalname || "file");

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

  fs.renameSync(req.file.path, inputPath);

  jobs.set(jobId, {
    status: "processing",
    filename:
      `${path.parse(safeName).name}.${outputFormat}`
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

      const files = fs.readdirSync(outputDir);

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

      jobs.set(jobId, {
        status: "finished",
        outputFile,
        filename:
          `${path.parse(safeName).name}.${outputFormat}`
      });
    }
  );
});

function getBaseUrl(req) {
  return `${req.protocol}://${req.get("host")}`;
}

const PORT = process.env.PORT || 10000;

app.listen(PORT, "0.0.0.0", () => {
  console.log(
    `Conversion engine running on port ${PORT}`
  );
});
