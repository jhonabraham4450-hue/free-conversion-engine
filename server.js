const express = require("express");
const multer = require("multer");
const fs = require("fs");
const path = require("path");
const { execFile } = require("child_process");
const { promisify } = require("util");

const {
  Document,
  Packer,
  Paragraph,
  TextRun
} = require("docx");

const execFileAsync = promisify(execFile);

const app = express();

const PORT = process.env.PORT || 10000;

const PUBLIC_BASE_URL =
  "https://free-conversion-engine.onrender.com";

const uploadDir = "/tmp/uploads";
const outputDir = "/tmp/output";
const profileDir = "/tmp/lo-profiles";

fs.mkdirSync(uploadDir, { recursive: true });
fs.mkdirSync(outputDir, { recursive: true });
fs.mkdirSync(profileDir, { recursive: true });

const upload = multer({
  dest: uploadDir,
  limits: {
    fileSize: 100 * 1024 * 1024
  }
});

const jobs = new Map();

/* =========================================================
   BASIC HELPERS
========================================================= */

function safeName(name) {
  return String(name || "file")
    .replace(/[^a-zA-Z0-9._-]/g, "_");
}

function removeDir(dir) {
  try {
    fs.rmSync(dir, {
      recursive: true,
      force: true
    });
  } catch (e) {}
}

function removeFile(file) {
  try {
    fs.unlinkSync(file);
  } catch (e) {}
}

function cleanupJob(jobId) {
  const job = jobs.get(jobId);

  if (!job) {
    return;
  }

  if (job.inputPath) {
    removeFile(job.inputPath);
  }

  if (job.outputPath) {
    removeFile(job.outputPath);
  }

  if (job.outputBuffer) {
    job.outputBuffer = null;
  }

  if (job.jobOutputDir) {
    removeDir(job.jobOutputDir);
  }

  if (job.jobProfileDir) {
    removeDir(job.jobProfileDir);
  }

  if (job.ocrDir) {
    removeDir(job.ocrDir);
  }
}

/* =========================================================
   HEALTH
========================================================= */

app.get("/", (req, res) => {
  res.json({
    success: true,
    service: "free-conversion-engine",
    status: "online",
    ocr: true,
    languages: ["ben", "eng"]
  });
});

/* =========================================================
   STATUS
========================================================= */

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
      jobId,
      filename: job.outputFilename,
      url:
        PUBLIC_BASE_URL +
        "/download/" +
        encodeURIComponent(jobId),
      contentType:
        job.contentType ||
        "application/octet-stream"
    });
  }

  if (job.status === "error") {
    return res.json({
      status: "error",
      jobId,
      error:
        job.error ||
        "Conversion failed."
    });
  }

  return res.json({
    status: "processing",
    jobId
  });
});

/* =========================================================
   DOWNLOAD
========================================================= */

app.get("/download/:jobId", (req, res) => {
  const jobId = req.params.jobId;

  const job = jobs.get(jobId);

  if (!job) {
    return res.status(404).json({
      success: false,
      error: "Job not found."
    });
  }

  if (job.status !== "finished") {
    return res.status(409).json({
      success: false,
      error: "Conversion is not finished yet.",
      status: job.status
    });
  }

  if (!job.outputBuffer) {
    return res.status(404).json({
      success: false,
      error: "Output file is no longer available."
    });
  }

  res.setHeader(
    "Content-Type",
    job.contentType ||
    "application/octet-stream"
  );

  res.setHeader(
    "Content-Disposition",
    `attachment; filename="${safeName(job.outputFilename)}"`
  );

  res.setHeader(
    "Content-Length",
    job.outputBuffer.length
  );

  return res.send(job.outputBuffer);
});

/* =========================================================
   CONVERSION MAP
========================================================= */

const conversionMap = {
  "word-to-pdf": {
    inputExt: ".docx",
    outputExt: ".pdf",
    contentType: "application/pdf"
  },

  "powerpoint-to-pdf": {
    inputExt: ".pptx",
    outputExt: ".pdf",
    contentType: "application/pdf"
  },

  "excel-to-pdf": {
    inputExt: ".xlsx",
    outputExt: ".pdf",
    contentType: "application/pdf"
  },

  "pdf-to-word": {
    inputExt: ".pdf",
    outputExt: ".docx",
    contentType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  },

  "pdf-to-powerpoint": {
    inputExt: ".pdf",
    outputExt: ".pptx",
    contentType:
      "application/vnd.openxmlformats-officedocument.presentationml.presentation"
  },

  "pdf-to-excel": {
    inputExt: ".pdf",
    outputExt: ".xlsx",
    contentType:
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  }
};

/* =========================================================
   START CONVERSION
========================================================= */

app.post(
  "/convert",
  upload.single("file"),
  async (req, res) => {

    try {

      if (!req.file) {
        return res.status(400).json({
          success: false,
          error: "No file uploaded."
        });
      }

      const tool =
        String(req.body.tool || "")
          .trim()
          .toLowerCase();

      const clientJobId =
        String(req.body.jobId || "")
          .trim();

      const originalFilename =
        req.body.filename ||
        req.file.originalname ||
        "file";

      if (!conversionMap[tool]) {

        removeFile(req.file.path);

        return res.status(400).json({
          success: false,
          error:
            "Unsupported conversion tool: " +
            tool
        });
      }

      const jobId =
        clientJobId ||
        (
          "job-" +
          Date.now() +
          "-" +
          Math.random()
            .toString(36)
            .slice(2, 10)
        );

      const config =
        conversionMap[tool];

      const extension =
        path.extname(
          originalFilename
        ) ||
        config.inputExt;

      const inputPath =
        path.join(
          uploadDir,
          jobId + extension
        );

      fs.renameSync(
        req.file.path,
        inputPath
      );

      const jobOutputDir =
        path.join(
          outputDir,
          jobId
        );

      const jobProfileDir =
        path.join(
          profileDir,
          jobId
        );

      const ocrDir =
        path.join(
          outputDir,
          jobId + "-ocr"
        );

      fs.mkdirSync(
        jobOutputDir,
        { recursive: true }
      );

      fs.mkdirSync(
        jobProfileDir,
        { recursive: true }
      );

      fs.mkdirSync(
        ocrDir,
        { recursive: true }
      );

      const outputFilename =
        path.basename(
          originalFilename,
          path.extname(originalFilename)
        ) +
        config.outputExt;

      const outputPath =
        path.join(
          jobOutputDir,
          outputFilename
        );

      jobs.set(jobId, {
        status: "processing",
        jobId,
        inputPath,
        outputPath,
        outputFilename,
        contentType:
          config.contentType,
        jobOutputDir,
        jobProfileDir,
        ocrDir,
        outputBuffer: null,
        error: null
      });

      res.json({
        success: true,
        jobId,
        status: "processing"
      });

      /* =====================================================
         PDF -> WORD
         OCR FIRST / TEXT EXTRACTION FALLBACK
      ===================================================== */

      if (tool === "pdf-to-word") {

        convertPdfToWord(
          inputPath,
          outputPath,
          jobId,
          jobOutputDir,
          jobProfileDir,
          ocrDir,
          outputFilename
        ).catch((error) => {

          console.error(
            "PDF to Word error:",
            error
          );

          const job = jobs.get(jobId);

          if (job) {
            job.status = "error";
            job.error =
              error?.message ||
              "PDF to Word conversion failed.";
          }
        });

        return;
      }

      /* =====================================================
         OTHER CONVERSIONS
      ===================================================== */

      convertWithLibreOffice(
        inputPath,
        jobId,
        jobOutputDir,
        jobProfileDir,
        tool,
        outputFilename
      ).catch((error) => {

        console.error(
          "LibreOffice conversion error:",
          error
        );

        const job = jobs.get(jobId);

        if (job) {
          job.status = "error";
          job.error =
            error?.message ||
            "Conversion failed.";
        }
      });

    } catch (error) {

      console.error(
        "Start conversion error:",
        error
      );

      if (req.file?.path) {
        removeFile(req.file.path);
      }

      return res.status(500).json({
        success: false,
        error:
          error?.message ||
          "Unable to start conversion."
      });
    }
  }
);

/* =========================================================
   PDF -> WORD
   ========================================================= */

async function convertPdfToWord(
  inputPath,
  outputPath,
  jobId,
  jobOutputDir,
  jobProfileDir,
  ocrDir,
  outputFilename
) {

  const job = jobs.get(jobId);

  if (!job) {
    throw new Error(
      "Conversion job not found."
    );
  }

  /*
   * STEP 1:
   * Try normal PDF text extraction first.
   */

  const textPath =
    path.join(
      ocrDir,
      "extracted.txt"
    );

  let extractedText = "";

  try {

    await execFileAsync(
      "pdftotext",
      [
        "-layout",
        inputPath,
        textPath
      ],
      {
        timeout: 120000,
        maxBuffer: 10 * 1024 * 1024
      }
    );

    if (fs.existsSync(textPath)) {

      extractedText =
        fs.readFileSync(
          textPath,
          "utf8"
        );
    }

  } catch (error) {

    console.log(
      "Normal text extraction unavailable. Starting OCR."
    );
  }

  /*
   * If normal text extraction worked,
   * use it directly.
   *
   * Otherwise use Bengali + English OCR.
   */

  if (
    extractedText &&
    extractedText.trim().length > 20
  ) {

    console.log(
      "Readable PDF text found. Skipping OCR."
    );

  } else {

    console.log(
      "No readable PDF text found. Starting OCR..."
    );

    extractedText =
      await runPdfOcr(
        inputPath,
        ocrDir,
        jobId
      );
  }

  if (
    !extractedText ||
    !extractedText.trim()
  ) {

    throw new Error(
      "OCR could not find readable text in this PDF."
    );
  }

  console.log(
    "Creating DOCX from extracted text..."
  );

  const document =
    buildDocxFromText(
      extractedText
    );

  const buffer =
    await Packer.toBuffer(
      document
    );

  job.outputBuffer = buffer;

  job.status = "finished";

  job.outputFilename =
    outputFilename;

  job.contentType =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

  console.log(
    "PDF to Word finished:",
    jobId
  );

  setTimeout(() => {

    const currentJob =
      jobs.get(jobId);

    if (currentJob) {

      cleanupJob(jobId);

      jobs.delete(jobId);

    }

  }, 10 * 60 * 1000);
}

/* =========================================================
   OCR
========================================================= */

async function runPdfOcr(
  inputPath,
  ocrDir,
  jobId
) {

  const imagePrefix =
    path.join(
      ocrDir,
      "page"
    );

  console.log(
    "Converting PDF pages to images..."
  );

  await execFileAsync(
    "pdftoppm",
    [
      "-jpeg",
      "-r",
      "200",
      "-jpegopt",
      "quality=90",
      inputPath,
      imagePrefix
    ],
    {
      timeout: 300000,
      maxBuffer: 10 * 1024 * 1024
    }
  );

  const files =
    fs.readdirSync(
      ocrDir
    )
      .filter(
        file =>
          /^page-\d+\.jpg$/i.test(file)
      )
      .sort(
        naturalPageSort
      );

  if (!files.length) {

    throw new Error(
      "Could not create images from PDF pages."
    );
  }

  console.log(
    "OCR pages:",
    files.length
  );

  const pageTexts = [];

  for (
    let i = 0;
    i < files.length;
    i++
  ) {

    const imagePath =
      path.join(
        ocrDir,
        files[i]
      );

    const outputBase =
      path.join(
        ocrDir,
        "ocr-" +
        String(i + 1)
      );

    console.log(
      `OCR page ${i + 1}/${files.length}`
    );

    try {

      await execFileAsync(
        "tesseract",
        [
          imagePath,
          outputBase,
          "-l",
          "ben+eng",
          "--psm",
          "3"
        ],
        {
          timeout: 180000,
          maxBuffer:
            10 * 1024 * 1024
        }
      );

    } catch (error) {

      console.error(
        "OCR failed on page:",
        i + 1,
        error?.message
      );

      /*
       * Try again using automatic page
       * segmentation with English only.
       */

      try {

        await execFileAsync(
          "tesseract",
          [
            imagePath,
            outputBase,
            "-l",
            "eng",
            "--psm",
            "6"
          ],
          {
            timeout: 180000,
            maxBuffer:
              10 * 1024 * 1024
          }
        );

      } catch (secondError) {

        console.error(
          "Second OCR attempt failed:",
          secondError?.message
        );
      }
    }

    const txtFile =
      outputBase +
      ".txt";

    let pageText = "";

    if (fs.existsSync(txtFile)) {

      pageText =
        fs.readFileSync(
          txtFile,
          "utf8"
        );
    }

    pageTexts.push(
      pageText
    );
  }

  return pageTexts.join(
    "\f"
  );
}

/* =========================================================
   NATURAL PAGE SORT
========================================================= */

function naturalPageSort(a, b) {

  const na =
    parseInt(
      a.match(/\d+/)?.[0] || "0",
      10
    );

  const nb =
    parseInt(
      b.match(/\d+/)?.[0] || "0",
      10
    );

  return na - nb;
}

/* =========================================================
   DOCX BUILDER
========================================================= */

function buildDocxFromText(
  text
) {

  const pages =
    String(text)
      .replace(/\r\n/g, "\n")
      .replace(/\r/g, "\n")
      .split("\f");

  const children = [];

  pages.forEach(
    (pageText, pageIndex) => {

      const lines =
        pageText.split("\n");

      if (pageIndex > 0) {

        children.push(
          new Paragraph({
            pageBreakBefore: true,
            children: []
          })
        );
      }

      for (
        const line of lines
      ) {

        /*
         * Keep OCR line structure.
         */

        children.push(
          new Paragraph({
            spacing: {
              before: 0,
              after: 0,
              line: 240
            },
            children: [
              new TextRun({
                text:
                  line || " ",
                font: "Noto Sans Bengali",
                size: 22
              })
            ]
          })
        );
      }
    }
  );

  return new Document({

    sections: [
      {
        properties: {

          page: {
            width: 11906,
            height: 16838,

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

/* =========================================================
   LIBREOFFICE CONVERSION
========================================================= */

async function convertWithLibreOffice(
  inputPath,
  jobId,
  jobOutputDir,
  jobProfileDir,
  tool,
  outputFilename
) {

  const job =
    jobs.get(jobId);

  if (!job) {
    throw new Error(
      "Conversion job not found."
    );
  }

  const args = [

    "--headless",

    "--convert-to",

    getLibreOfficeFormat(tool),

    "--outdir",

    jobOutputDir,

    "-env:UserInstallation=file://" +
      jobProfileDir,

    inputPath
  ];

  console.log(
    "Running LibreOffice:",
    args.join(" ")
  );

  try {

    await execFileAsync(
      "libreoffice",
      args,
      {
        timeout: 300000,
        maxBuffer:
          20 * 1024 * 1024
      }
    );

  } catch (error) {

    console.error(
      "LibreOffice stderr:",
      error?.stderr
    );

    throw new Error(
      error?.message ||
      "LibreOffice conversion failed."
    );
  }

  finishLibreOfficeJob(
    jobId,
    jobOutputDir,
    outputFilename
  );
}

/* =========================================================
   LIBREOFFICE FORMAT
========================================================= */

function getLibreOfficeFormat(
  tool
) {

  switch (tool) {

    case "word-to-pdf":
      return "pdf:writer_pdf_Export";

    case "powerpoint-to-pdf":
      return "pdf:impress_pdf_Export";

    case "excel-to-pdf":
      return "pdf:calc_pdf_Export";

    case "pdf-to-powerpoint":
      return "pptx:Impress MS PowerPoint 2007 XML";

    case "pdf-to-excel":
      return "xlsx:Calc MS Excel 2007 XML";

    default:
      throw new Error(
        "Unsupported LibreOffice conversion."
      );
  }
}

/* =========================================================
   FINISH LIBREOFFICE JOB
========================================================= */

function finishLibreOfficeJob(
  jobId,
  jobOutputDir,
  outputFilename
) {

  const job =
    jobs.get(jobId);

  if (!job) {
    throw new Error(
      "Conversion job not found."
    );
  }

  const outputPath =
    path.join(
      jobOutputDir,
      outputFilename
    );

  if (!fs.existsSync(outputPath)) {

    const files =
      fs.readdirSync(
        jobOutputDir
      );

    console.error(
      "Output files:",
      files
    );

    throw new Error(
      "Converted output file was not created."
    );
  }

  const buffer =
    fs.readFileSync(
      outputPath
    );

  job.outputBuffer =
    buffer;

  job.status =
    "finished";

  console.log(
    "Conversion finished:",
    jobId
  );

  setTimeout(() => {

    const currentJob =
      jobs.get(jobId);

    if (currentJob) {

      cleanupJob(jobId);

      jobs.delete(jobId);
    }

  }, 10 * 60 * 1000);
}

/* =========================================================
   ERROR HANDLER
========================================================= */

app.use(
  (
    error,
    req,
    res,
    next
  ) => {

    console.error(
      "Unhandled error:",
      error
    );

    if (res.headersSent) {
      return next(error);
    }

    return res.status(500).json({
      success: false,
      error:
        error?.message ||
        "Internal server error."
    });
  }
);

/* =========================================================
   START SERVER
========================================================= */

app.listen(
  PORT,
  "0.0.0.0",
  () => {

    console.log(
      "Free conversion engine running on port",
      PORT
    );

    console.log(
      "OCR languages: Bengali + English"
    );

    console.log(
      "Public URL:",
      PUBLIC_BASE_URL
    );
  }
);
