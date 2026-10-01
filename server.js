const express = require("express");
const multer = require("multer");
const fs = require("fs");
const path = require("path");
const { execFile } = require("child_process");

const app = express();

/* =========================
   CORS
========================= */

app.use((req, res, next) => {
  res.header("Access-Control-Allow-Origin", "*");
  res.header(
    "Access-Control-Allow-Methods",
    "GET,POST,OPTIONS"
  );
  res.header(
    "Access-Control-Allow-Headers",
    "Content-Type"
  );

  if (req.method === "OPTIONS") {
    return res.sendStatus(204);
  }

  next();
});

/* =========================
   DIRECTORIES
========================= */

const UPLOAD_DIR = "/tmp/uploads";
const OUTPUT_DIR = "/tmp/output";

fs.mkdirSync(UPLOAD_DIR, {
  recursive: true
});

fs.mkdirSync(OUTPUT_DIR, {
  recursive: true
});

/* =========================
   UPLOAD
========================= */

const upload = multer({
  dest: UPLOAD_DIR
});

/* =========================
   JOB STORAGE
========================= */

const jobs = new Map();

/* =========================
   HOME / HEALTH CHECK
========================= */

app.get("/", (req, res) => {
  res.json({
    success: true,
    service: "iLovePDF4 Free Conversion Engine",
    status: "online"
  });
});

/* =========================
   STATUS
========================= */

app.get("/status/:jobId", (req, res) => {
  const jobId = req.params.jobId;

  const job = jobs.get(jobId);

  if (!job) {
    return res.status(404).json({
      status: "error",
      error: "Job not found."
    });
  }

  if (job.status === "finished") {
    return res.json({
      status: "finished",
      url:
        getBaseUrl(req) +
        "/download/" +
        encodeURIComponent(jobId),
      filename: job.filename
    });
  }

  if (job.status === "error") {
    return res.json({
      status: "error",
      error:
        job.error ||
        "Conversion failed."
    });
  }

  return res.json({
    status: "processing"
  });
});

/* =========================
   DOWNLOAD
========================= */

app.get("/download/:jobId", (req, res) => {
  const jobId = req.params.jobId;

  const job = jobs.get(jobId);

  if (
    !job ||
    job.status !== "finished" ||
    !job.outputBuffer
  ) {
    return res.status(404).send(
      "Output file no longer exists."
    );
  }

  res.setHeader(
    "Content-Type",
    job.contentType ||
      "application/octet-stream"
  );

  res.setHeader(
    "Content-Disposition",
    'attachment; filename="' +
      job.filename.replace(/"/g, "") +
      '"'
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

/* =========================
   CONVERSION MAP
========================= */

const conversionMap = {

  /*
   * OFFICE -> PDF
   */

  "word-to-pdf": {
    format: "pdf",
    extension: "pdf",
    contentType: "application/pdf"
  },

  "powerpoint-to-pdf": {
    format: "pdf",
    extension: "pdf",
    contentType: "application/pdf"
  },

  "excel-to-pdf": {
    format: "pdf",
    extension: "pdf",
    contentType: "application/pdf"
  },

  /*
   * PDF -> OFFICE
   */

  "pdf-to-word": {
    format: "docx",
    extension: "docx",
    contentType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  },

  "pdf-to-powerpoint": {
    format: "pptx",
    extension: "pptx",
    contentType:
      "application/vnd.openxmlformats-officedocument.presentationml.presentation"
  },

  "pdf-to-excel": {
    format: "xlsx",
    extension: "xlsx",
    contentType:
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  }
};

/* =========================
   CONVERT
========================= */

app.post(
  "/convert",
  upload.single("file"),
  (req, res) => {

    if (!req.file) {
      return res.status(400).json({
        success: false,
        error: "No file uploaded."
      });
    }

    const tool = String(
      req.body.tool || ""
    )
      .trim()
      .toLowerCase();

    const jobId = String(
      req.body.jobId || ""
    ).trim();

    if (!jobId) {

      safeDelete(req.file.path);

      return res.status(400).json({
        success: false,
        error: "Missing job id."
      });
    }

    const conversion =
      conversionMap[tool];

    /*
     * Only the 6 server-side tools
     * should reach this endpoint.
     */

    if (!conversion) {

      safeDelete(req.file.path);

      return res.status(400).json({
        success: false,
        error:
          "Unsupported conversion tool: " +
          tool
      });
    }

    const originalName = String(
      req.body.filename ||
      req.file.originalname ||
      "file"
    );

    const safeName =
      path.basename(originalName);

    const inputPath = path.join(
      UPLOAD_DIR,
      jobId + "-" + safeName
    );

    const jobOutputDir =
      path.join(
        OUTPUT_DIR,
        jobId
      );

    fs.mkdirSync(
      jobOutputDir,
      {
        recursive: true
      }
    );

    try {

      fs.renameSync(
        req.file.path,
        inputPath
      );

    } catch (error) {

      safeDelete(req.file.path);

      return res.status(500).json({
        success: false,
        error:
          "Unable to prepare uploaded file."
      });
    }

    const outputFilename =
      path.parse(safeName).name +
      "." +
      conversion.extension;

    jobs.set(jobId, {
      status: "processing",
      filename: outputFilename
    });

    /*
     * Respond immediately.
     */

    res.json({
      success: true,
      jobId: jobId,
      status: "processing"
    });

    /*
     * LibreOffice conversion
     */

    const libreOfficeArgs =
  tool === "pdf-to-word"
    ? [
        "--headless",
        "--infilter=writer_pdf_import",
        "--convert-to",
        "docx:MS Word 2007 XML",
        "--outdir",
        jobOutputDir,
        inputPath
      ]
    : [
        "--headless",
        "--convert-to",
        conversion.format,
        "--outdir",
        jobOutputDir,
        inputPath
      ];

execFile(
  "libreoffice",
  libreOfficeArgs,
  {
    timeout: 180000,
    maxBuffer: 10 * 1024 * 1024
  },
  (error, stdout, stderr) => {
        timeout: 180000,
        maxBuffer: 10 * 1024 * 1024
      },
      (error, stdout, stderr) => {

        /*
         * Remove input
         */

        safeDelete(inputPath);

        console.log(
          "Conversion tool:",
          tool
        );

        console.log(
          "LibreOffice stdout:",
          stdout || ""
        );

        console.log(
          "LibreOffice stderr:",
          stderr || ""
        );

        if (error) {

          console.error(
            "LibreOffice conversion failed:",
            error.message
          );

          jobs.set(jobId, {
            status: "error",
            error:
              getLibreOfficeError(
                tool,
                stderr,
                error
              )
          });

          cleanupDirectory(
            jobOutputDir
          );

          return;
        }

        /*
         * Find output
         */

        let files;

        try {

          files =
            fs.readdirSync(
              jobOutputDir
            );

        } catch (readError) {

          jobs.set(jobId, {
            status: "error",
            error:
              "Unable to read conversion output."
          });

          cleanupDirectory(
            jobOutputDir
          );

          return;
        }

        if (!files.length) {

          jobs.set(jobId, {
            status: "error",
            error:
              "Conversion output was not created."
          });

          cleanupDirectory(
            jobOutputDir
          );

          return;
        }

        /*
         * Prefer expected extension.
         */

        let outputName =
          files.find(
            file =>
              path
                .extname(file)
                .toLowerCase() ===
              "." +
                conversion.extension
          );

        /*
         * Fallback to first generated file.
         */

        if (!outputName) {
          outputName = files[0];
        }

        const outputFile =
          path.join(
            jobOutputDir,
            outputName
          );

        try {

          const outputBuffer =
            fs.readFileSync(
              outputFile
            );

          jobs.set(jobId, {
            status: "finished",
            outputBuffer:
              outputBuffer,
            filename:
              outputFilename,
            contentType:
              conversion.contentType
          });

          console.log(
            "Conversion finished:",
            jobId
          );

        } catch (readError) {

          jobs.set(jobId, {
            status: "error",
            error:
              "Unable to read conversion output."
          });
        }

        /*
         * Remove temporary output.
         */

        cleanupDirectory(
          jobOutputDir
        );

        /*
         * Automatically remove job
         * after 10 minutes.
         */

        setTimeout(() => {
          jobs.delete(jobId);
        }, 10 * 60 * 1000);
      }
    );
  }
);

/* =========================
   HELPERS
========================= */

function getBaseUrl(req) {
  return (
    req.protocol +
    "://" +
    req.get("host")
  );
}

function safeDelete(filePath) {
  try {
    if (
      filePath &&
      fs.existsSync(filePath)
    ) {
      fs.unlinkSync(filePath);
    }
  } catch {}
}

function cleanupDirectory(dir) {

  try {

    if (
      fs.existsSync(dir)
    ) {

      const files =
        fs.readdirSync(dir);

      for (const file of files) {

        const fullPath =
          path.join(
            dir,
            file
          );

        try {
          fs.unlinkSync(fullPath);
        } catch {}
      }

      try {
        fs.rmdirSync(dir);
      } catch {}
    }

  } catch {}
}

function getLibreOfficeError(
  tool,
  stderr,
  error
) {

  if (
    stderr &&
    stderr.trim()
  ) {
    return (
      "LibreOffice conversion failed: " +
      stderr.trim()
    );
  }

  return (
    "Conversion failed for " +
    tool +
    ". " +
    (
      error?.message ||
      "Unknown conversion error."
    )
  );
}

/* =========================
   SERVER
========================= */

const PORT =
  process.env.PORT || 10000;

app.listen(
  PORT,
  "0.0.0.0",
  () => {

    console.log(
      "================================="
    );

    console.log(
      "iLovePDF4 Conversion Engine"
    );

    console.log(
      "Running on port " +
        PORT
    );

    console.log(
      "Supported server tools:"
    );

    Object.keys(
      conversionMap
    ).forEach(tool => {
      console.log(
        " - " + tool
      );
    });

    console.log(
      "================================="
    );
  }
);
