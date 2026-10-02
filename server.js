const express = require("express");
const multer = require("multer");
const fs = require("fs");
const path = require("path");
const { execFile } = require("child_process");

const { createCanvas } = require("@napi-rs/canvas");
const pdfjsLib = require("pdfjs-dist/legacy/build/pdf.js");

const {
  Document,
  Packer,
  Paragraph,
  ImageRun
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
   PUBLIC URL
========================= */

const PUBLIC_BASE_URL =
  "https://free-conversion-engine.onrender.com";

/* =========================
   HEALTH
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
        PUBLIC_BASE_URL +
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
   SCANNED PDF -> DOCX
   FULL PAGE IMAGE METHOD
========================= */

async function convertScannedPdfToDocx(
  inputPath,
  outputPath
) {
  const pdfBytes = new Uint8Array(
    fs.readFileSync(inputPath)
  );

  const pdf =
    await pdfjsLib.getDocument({
      data: pdfBytes,
      disableWorker: true
    }).promise;

  const sections = [];

  for (
    let pageNumber = 1;
    pageNumber <= pdf.numPages;
    pageNumber++
  ) {
    const page =
      await pdf.getPage(pageNumber);

    /*
     * PDF uses 72 points per inch.
     * Render at 96 DPI.
     */
    const scale = 96 / 72;

    const viewport =
      page.getViewport({
        scale
      });

    const width =
      Math.ceil(viewport.width);

    const height =
      Math.ceil(viewport.height);

    const canvas =
      createCanvas(
        width,
        height
      );

    const context =
      canvas.getContext("2d");

    await page.render({
      canvasContext: context,
      viewport: viewport
    }).promise;

    const imageBuffer =
      canvas.toBuffer("image/png");

    /*
     * PDF points -> DOCX twips
     * 1 point = 20 twips
     */
    const pageWidthTwips =
      Math.round(
        (viewport.width / scale) * 20
      );

    const pageHeightTwips =
      Math.round(
        (viewport.height / scale) * 20
      );

    sections.push({
      properties: {
        page: {
          size: {
            width: pageWidthTwips,
            height: pageHeightTwips
          },

          margin: {
            top: 0,
            bottom: 0,
            left: 0,
            right: 0,
            header: 0,
            footer: 0,
            gutter: 0
          }
        }
      },

      children: [
        new Paragraph({
          spacing: {
            before: 0,
            after: 0,
            line: 0
          },

          children: [
            new ImageRun({
              type: "png",
              data: imageBuffer,

              transformation: {
                width: width,
                height: height
              }
            })
          ]
        })
      ]
    });
  }

  const doc =
    new Document({
      sections: sections
    });

  const buffer =
    await Packer.toBuffer(doc);

  fs.writeFileSync(
    outputPath,
    buffer
  );
}

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

    const inputPath =
      path.join(
        UPLOAD_DIR,
        jobId + "-" + safeName
      );

    const jobOutputDir =
      path.join(
        OUTPUT_DIR,
        jobId
      );

    const jobProfileDir =
      path.join(
        PROFILE_DIR,
        jobId
      );

    fs.mkdirSync(
      jobOutputDir,
      { recursive: true }
    );

    fs.mkdirSync(
      jobProfileDir,
      { recursive: true }
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

    /* =========================
       FINISH CONTROL
    ========================= */

    let finished = false;

    function finishSuccess(outputFile) {

      if (finished) {
        return;
      }

      finished = true;

      try {

        const outputBuffer =
          fs.readFileSync(
            outputFile
          );

        if (!outputBuffer.length) {
          throw new Error(
            "Output file is empty."
          );
        }

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
          "CONVERSION FINISHED:",
          jobId
        );

        console.log(
          "Output size:",
          outputBuffer.length,
          "bytes"
        );

      } catch (error) {

        jobs.set(jobId, {
          status: "error",
          error:
            "Unable to read conversion output."
        });

        console.error(
          "Output read error:",
          error.message
        );
      }

      safeDelete(inputPath);

      cleanupDirectory(
        jobOutputDir
      );

      cleanupDirectory(
        jobProfileDir
      );
    }

    function finishError(message) {

      if (finished) {
        return;
      }

      finished = true;

      jobs.set(jobId, {
        status: "error",
        error: message
      });

      safeDelete(inputPath);

      cleanupDirectory(
        jobOutputDir
      );

      cleanupDirectory(
        jobProfileDir
      );

      console.error(
        "CONVERSION ERROR:",
        message
      );
    }

    /* =========================
       PDF -> WORD
       FULL PAGE CONVERSION
    ========================= */

    if (tool === "pdf-to-word") {

      const outputFile =
        path.join(
          jobOutputDir,
          outputFilename
        );

      console.log(
        "================================="
      );

      console.log(
        "Starting PDF -> WORD conversion"
      );

      console.log(
        "Job:",
        jobId
      );

      console.log(
        "Input:",
        inputPath
      );

      console.log(
        "Output:",
        outputFile
      );

      console.log(
        "================================="
      );

      convertScannedPdfToDocx(
        inputPath,
        outputFile
      )
        .then(() => {

          if (finished) {
            return;
          }

          finishSuccess(
            outputFile
          );

        })
        .catch((error) => {

          console.error(
            "PDF TO WORD ERROR:",
            error
          );

          finishError(
            error && error.message
              ? error.message
              : "PDF to Word conversion failed."
          );

        });

      /*
       * IMPORTANT:
       * Do not start LibreOffice
       * for PDF -> Word.
       */
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

    /* WORD -> PDF */

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

    /* POWERPOINT -> PDF */

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

    /* EXCEL -> PDF */

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

    /* PDF -> POWERPOINT */

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

    /* PDF -> EXCEL */

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
      "Starting conversion:",
      tool
    );

    console.log(
      "Job:",
      jobId
    );

    console.log(
      "================================="
    );

    /* =========================
       START LIBREOFFICE
    ========================= */

    const child =
      execFile(
        "libreoffice",
        libreOfficeArgs,
        {
          timeout: 35000,
          maxBuffer:
            20 * 1024 * 1024
        },
        (
          error,
          stdout,
          stderr
        ) => {

          console.log(
            "LibreOffice stdout:",
            stdout || ""
          );

          console.log(
            "LibreOffice stderr:",
            stderr || ""
          );

          /*
           * If output already exists,
           * use it immediately.
           */
          if (!finished) {

            const outputFile =
              findOutputFile(
                jobOutputDir,
                conversion.extension
              );

            if (outputFile) {

              finishSuccess(
                outputFile
              );

              return;
            }
          }

          if (finished) {
            return;
          }

          if (error) {

            finishError(
              getLibreOfficeError(
                tool,
                stderr,
                error
              )
            );

            return;
          }

          finishError(
            "Conversion finished without creating an output file."
          );
        }
      );

    /* =========================
       FAST OUTPUT DETECTION
    ========================= */

    let stableSize = -1;
    let stableCount = 0;

    const outputWatcher =
      setInterval(() => {

        if (finished) {

          clearInterval(
            outputWatcher
          );

          return;
        }

        const outputFile =
          findOutputFile(
            jobOutputDir,
            conversion.extension
          );

        if (!outputFile) {
          return;
        }

        let currentSize = 0;

        try {

          currentSize =
            fs.statSync(
              outputFile
            ).size;

        } catch {
          return;
        }

        if (
          currentSize > 0 &&
          currentSize === stableSize
        ) {

          stableCount++;

        } else {

          stableSize =
            currentSize;

          stableCount = 0;
        }

        /*
         * File size stable for
         * two checks = safe to read.
         */
        if (
          stableCount >= 2
        ) {

          clearInterval(
            outputWatcher
          );

          finishSuccess(
            outputFile
          );

          /*
           * LibreOffice may still be
           * alive after output creation.
           */
          try {

            if (
              child &&
              !child.killed
            ) {
              child.kill(
                "SIGKILL"
              );
            }

          } catch {}
        }

      }, 300);

    /* =========================
       HARD TIMEOUT
    ========================= */

    setTimeout(() => {

      if (finished) {
        return;
      }

      clearInterval(
        outputWatcher
      );

      try {

        if (
          child &&
          !child.killed
        ) {
          child.kill(
            "SIGKILL"
          );
        }

      } catch {}

      finishError(
        "LibreOffice conversion timed out."
      );

    }, 35000);

    /* =========================
       AUTO DELETE JOB
    ========================= */

    setTimeout(() => {

      jobs.delete(jobId);

    }, 10 * 60 * 1000);
  }
);

/* =========================
   FIND OUTPUT
========================= */

function findOutputFile(
  directory,
  extension
) {

  try {

    if (
      !fs.existsSync(directory)
    ) {
      return null;
    }

    const files =
      fs.readdirSync(
        directory
      );

    const wanted =
      "." +
      extension.toLowerCase();

    const match =
      files.find(
        file =>
          path
            .extname(file)
            .toLowerCase() ===
          wanted
      );

    if (!match) {
      return null;
    }

    return path.join(
      directory,
      match
    );

  } catch {

    return null;
  }
}

/* =========================
   HELPERS
========================= */

function safeDelete(
  filePath
) {

  try {

    if (
      filePath &&
      fs.existsSync(filePath)
    ) {
      fs.unlinkSync(
        filePath
      );
    }

  } catch {}
}

function cleanupDirectory(
  directory
) {

  try {

    if (
      !fs.existsSync(directory)
    ) {
      return;
    }

    fs.rmSync(
      directory,
      {
        recursive: true,
        force: true
      }
    );

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
      "iLovePDF4 Free Conversion Engine"
    );

    console.log(
      "Running on port " +
        PORT
    );

    console.log(
      "================================="
    );

  }
);
