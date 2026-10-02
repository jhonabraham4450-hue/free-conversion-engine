const express = require("express");
const multer = require("multer");
const fs = require("fs");
const path = require("path");
const { execFile } = require("child_process");

const {
  Document,
  Packer,
  Paragraph,
  TextRun
} = require("docx");

const app = express();

app.use(express.json({ limit: "10mb" }));

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
const PROFILE_DIR = "/tmp/lo-profiles";

fs.mkdirSync(UPLOAD_DIR, { recursive: true });
fs.mkdirSync(OUTPUT_DIR, { recursive: true });
fs.mkdirSync(PROFILE_DIR, { recursive: true });

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
   PUBLIC BASE URL
========================= */

const PUBLIC_BASE_URL =
  "https://free-conversion-engine.onrender.com";

/* =========================
   HOME / HEALTH
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
      jobId: jobId,
      url:
        PUBLIC_BASE_URL +
        "/download/" +
        encodeURIComponent(jobId),
      filename: job.filename,
      contentType: job.contentType || null
    });
  }

  if (job.status === "error") {
    return res.json({
      status: "error",
      jobId: jobId,
      error: job.error || "Conversion failed."
    });
  }

  return res.json({
    status: "processing",
    jobId: jobId
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

  res.setHeader("Pragma", "no-cache");

  res.send(job.outputBuffer);
});

/* =========================
   CONVERSION MAP
========================= */

const conversionMap = {
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

    const conversion = conversionMap[tool];

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

    const safeName = path.basename(
      originalName
    );

    const inputPath = path.join(
      UPLOAD_DIR,
      jobId + "-" + safeName
    );

    const jobOutputDir = path.join(
      OUTPUT_DIR,
      jobId
    );

    const jobProfileDir = path.join(
      PROFILE_DIR,
      jobId
    );

    fs.mkdirSync(jobOutputDir, {
      recursive: true
    });

    fs.mkdirSync(jobProfileDir, {
      recursive: true
    });

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
      filename: outputFilename,
      contentType: conversion.contentType
    });

    res.json({
      success: true,
      jobId: jobId,
      status: "processing"
    });

    /* ==================================================
       PDF TO WORD
       Use pdftotext + DOCX generation instead of
       LibreOffice PDF import.

       This prevents PDF text from being positioned
       outside the Word page and getting clipped.
    ================================================== */

    if (tool === "pdf-to-word") {
      convertPdfToWord(
        inputPath,
        jobOutputDir,
        jobProfileDir,
        jobId,
        outputFilename
      );

      return;
    }

    /* =========================
       LIBREOFFICE ARGUMENTS
    ========================= */

    let libreOfficeArgs = [
      "--headless",
      "--nologo",
      "--nodefault",
      "--nofirststartwizard",
      "--nolockcheck",

      "-env:UserInstallation=file://" +
        jobProfileDir,

      "--convert-to"
    ];

    /* =========================
       WORD TO PDF
    ========================= */

    if (tool === "word-to-pdf") {
      libreOfficeArgs.push(
        "pdf:writer_pdf_Export"
      );

      libreOfficeArgs.push(
        "--outdir",
        jobOutputDir,
        inputPath
      );
    }

    /* =========================
       POWERPOINT TO PDF
    ========================= */

    else if (
      tool === "powerpoint-to-pdf"
    ) {
      libreOfficeArgs.push(
        "pdf:impress_pdf_Export"
      );

      libreOfficeArgs.push(
        "--outdir",
        jobOutputDir,
        inputPath
      );
    }

    /* =========================
       EXCEL TO PDF
    ========================= */

    else if (
      tool === "excel-to-pdf"
    ) {
      libreOfficeArgs.push(
        "pdf:calc_pdf_Export"
      );

      libreOfficeArgs.push(
        "--outdir",
        jobOutputDir,
        inputPath
      );
    }

    /* =========================
       PDF TO POWERPOINT
    ========================= */

    else if (
      tool === "pdf-to-powerpoint"
    ) {
      libreOfficeArgs.push(
        "pptx:Impress MS PowerPoint 2007 XML"
      );

      libreOfficeArgs.push(
        "--infilter=draw_pdf_import",
        "--outdir",
        jobOutputDir,
        inputPath
      );
    }

    /* =========================
       PDF TO EXCEL
    ========================= */

    else if (
      tool === "pdf-to-excel"
    ) {
      libreOfficeArgs.push(
        "xlsx:Calc MS Excel 2007 XML"
      );

      libreOfficeArgs.push(
        "--infilter=draw_pdf_import",
        "--outdir",
        jobOutputDir,
        inputPath
      );
    }

    console.log(
      "================================="
    );

    console.log(
      "Conversion tool:",
      tool
    );

    console.log(
      "Input:",
      inputPath
    );

    console.log(
      "Output directory:",
      jobOutputDir
    );

    console.log(
      "LibreOffice arguments:",
      libreOfficeArgs
    );

    console.log(
      "================================="
    );

    /* =========================
       RUN LIBREOFFICE
    ========================= */

    execFile(
      "libreoffice",
      libreOfficeArgs,
      {
        timeout: 180000,
        maxBuffer: 20 * 1024 * 1024
      },
      (error, stdout, stderr) => {
        console.log(
          "LibreOffice stdout:",
          stdout || ""
        );

        console.log(
          "LibreOffice stderr:",
          stderr || ""
        );

        safeDelete(inputPath);

        if (error) {
          console.error(
            "LibreOffice conversion failed:",
            error.message
          );

          jobs.set(jobId, {
            status: "error",
            error: getLibreOfficeError(
              tool,
              stderr,
              error
            )
          });

          cleanupDirectory(jobOutputDir);
          cleanupDirectory(jobProfileDir);

          return;
        }

        finishLibreOfficeJob(
          jobId,
          jobOutputDir,
          jobProfileDir,
          conversion,
          outputFilename
        );
      }
    );
  }
);

/* ==================================================
   PDF TO WORD CONVERTER
================================================== */

function convertPdfToWord(
  inputPath,
  jobOutputDir,
  jobProfileDir,
  jobId,
  outputFilename
) {
  const textPath = path.join(
    jobOutputDir,
    "extracted.txt"
  );

  console.log(
    "================================="
  );

  console.log(
    "PDF TO WORD ENGINE: pdftotext + DOCX"
  );

  console.log(
    "Input:",
    inputPath
  );

  console.log(
    "Text output:",
    textPath
  );

  console.log(
    "================================="
  );

  execFile(
    "pdftotext",
    [
      "-layout",
      inputPath,
      textPath
    ],
    {
      timeout: 180000,
      maxBuffer: 20 * 1024 * 1024
    },
    async (error, stdout, stderr) => {
      safeDelete(inputPath);

      if (error) {
        console.error(
          "pdftotext failed:",
          error.message
        );

        console.error(
          "pdftotext stderr:",
          stderr || ""
        );

        jobs.set(jobId, {
          status: "error",
          error:
            "PDF text extraction failed. " +
            (stderr || error.message)
        });

        cleanupDirectory(jobOutputDir);
        cleanupDirectory(jobProfileDir);

        return;
      }

      let text;

      try {
        text = fs.readFileSync(
          textPath,
          "utf8"
        );
      } catch (readError) {
        jobs.set(jobId, {
          status: "error",
          error:
            "Unable to read extracted PDF text."
        });

        cleanupDirectory(jobOutputDir);
        cleanupDirectory(jobProfileDir);

        return;
      }

      try {
        const document =
          buildDocxFromPdfText(text);

        const buffer =
          await Packer.toBuffer(
            document
          );

        jobs.set(jobId, {
          status: "finished",
          outputBuffer: buffer,
          filename: outputFilename,
          contentType:
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        });

        console.log(
          "PDF TO WORD finished:",
          jobId
        );

        console.log(
          "Output size:",
          buffer.length,
          "bytes"
        );

        cleanupDirectory(
          jobOutputDir
        );

        cleanupDirectory(
          jobProfileDir
        );

        setTimeout(() => {
          jobs.delete(jobId);
        }, 10 * 60 * 1000);

      } catch (docxError) {
        console.error(
          "DOCX generation failed:",
          docxError.message
        );

        jobs.set(jobId, {
          status: "error",
          error:
            "Unable to create Word document: " +
            docxError.message
        });

        cleanupDirectory(
          jobOutputDir
        );

        cleanupDirectory(
          jobProfileDir
        );
      }
    }
  );
}

/* ==================================================
   BUILD DOCX FROM EXTRACTED PDF TEXT
================================================== */

function buildDocxFromPdfText(text) {
  const sections = text.split("\f");

  const children = [];

  sections.forEach(
    (pageText, pageIndex) => {
      const lines =
        pageText.replace(
          /\r/g,
          ""
        ).split("\n");

      if (
        pageIndex > 0 &&
        children.length > 0
      ) {
        children.push(
          new Paragraph({
            pageBreakBefore: true,
            children: [
              new TextRun("")
            ]
          })
        );
      }

      let hasContent = false;

      lines.forEach((line) => {
        const cleaned =
          line.replace(/\s+$/g, "");

        if (cleaned.trim().length > 0) {
          hasContent = true;

          children.push(
            new Paragraph({
              spacing: {
                after: 0,
                line: 240
              },
              children: [
                new TextRun({
                  text: cleaned,
                  size: 22
                })
              ]
            })
          );
        } else if (hasContent) {
          children.push(
            new Paragraph({
              spacing: {
                after: 0
              },
              children: [
                new TextRun({
                  text: "",
                  size: 22
                })
              ]
            })
          );
        }
      });
    }
  );

  if (!children.length) {
    children.push(
      new Paragraph({
        children: [
          new TextRun({
            text:
              "No readable text was found in this PDF.",
            size: 22
          })
        ]
      })
    );
  }

  return new Document({
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: 720,
              right: 720,
              bottom: 720,
              left: 720
            }
          }
        },
        children
      }
    ]
  });
}

/* ==================================================
   FINISH LIBREOFFICE JOB
================================================== */

function finishLibreOfficeJob(
  jobId,
  jobOutputDir,
  jobProfileDir,
  conversion,
  outputFilename
) {
  let files;

  try {
    files = fs.readdirSync(
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

    cleanupDirectory(
      jobProfileDir
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

    cleanupDirectory(
      jobProfileDir
    );

    return;
  }

  let outputName = files.find(
    file =>
      path
        .extname(file)
        .toLowerCase() ===
      "." +
        conversion.extension
  );

  if (!outputName) {
    outputName = files[0];
  }

  const outputFile = path.join(
    jobOutputDir,
    outputName
  );

  let outputBuffer;

  try {
    outputBuffer =
      fs.readFileSync(
        outputFile
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

    cleanupDirectory(
      jobProfileDir
    );

    return;
  }

  jobs.set(jobId, {
    status: "finished",
    outputBuffer: outputBuffer,
    filename: outputFilename,
    contentType:
      conversion.contentType
  });

  console.log(
    "Conversion finished:",
    jobId
  );

  console.log(
    "Output file:",
    outputFile
  );

  console.log(
    "Output size:",
    outputBuffer.length,
    "bytes"
  );

  cleanupDirectory(
    jobOutputDir
  );

  cleanupDirectory(
    jobProfileDir
  );

  setTimeout(() => {
    jobs.delete(jobId);
  }, 10 * 60 * 1000);
}

/* =========================
   HELPERS
========================= */

function safeDelete(filePath) {
  try {
    if (
      filePath &&
      fs.existsSync(filePath)
    ) {
      fs.unlinkSync(filePath);
    }
  } catch (error) {}
}

function cleanupDirectory(dir) {
  try {
    if (!fs.existsSync(dir)) {
      return;
    }

    const files =
      fs.readdirSync(dir);

    for (const file of files) {
      const fullPath =
        path.join(dir, file);

      try {
        fs.rmSync(fullPath, {
          recursive: true,
          force: true
        });
      } catch (error) {}
    }

    try {
      fs.rmdirSync(dir);
    } catch (error) {}
  } catch (error) {}
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
      "Running on port " + PORT
    );

    console.log(
      "Public URL:",
      PUBLIC_BASE_URL
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
