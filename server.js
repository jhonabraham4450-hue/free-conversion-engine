const express = require("express");
const multer = require("multer");
const fs = require("fs");
const path = require("path");
const { execFile } = require("child_process");
const { promisify } = require("util");
const { Document, Packer, Paragraph, TextRun } = require("docx");

const execFileAsync = promisify(execFile);

const app = express();

const PORT = process.env.PORT || 10000;
const HOST = "0.0.0.0";

const PUBLIC_BASE_URL =
  "https://free-conversion-engine.onrender.com";

const UPLOAD_DIR = "/tmp/uploads";
const OUTPUT_DIR = "/tmp/output";
const LO_PROFILE_DIR = "/tmp/lo-profiles";

fs.mkdirSync(UPLOAD_DIR, { recursive: true });
fs.mkdirSync(OUTPUT_DIR, { recursive: true });
fs.mkdirSync(LO_PROFILE_DIR, { recursive: true });

const upload = multer({
  dest: UPLOAD_DIR,
  limits: {
    fileSize: 50 * 1024 * 1024
  }
});

const jobs = new Map();

/* =========================================================
   BASIC HELPERS
========================================================= */

function safeName(name) {
  return String(name || "file")
    .replace(/[^a-zA-Z0-9._-]/g, "_")
    .slice(0, 180);
}

function json(data, status = 200) {
  return new Response(
    JSON.stringify(data),
    {
      status,
      headers: {
        "Content-Type": "application/json",
        "Cache-Control": "no-store"
      }
    }
  );
}

function extensionForTool(tool) {
  const map = {
    "word-to-pdf": "pdf",
    "powerpoint-to-pdf": "pdf",
    "excel-to-pdf": "pdf",
    "pdf-to-word": "docx",
    "pdf-to-powerpoint": "pptx",
    "pdf-to-excel": "xlsx"
  };

  return map[tool] || "bin";
}

function cleanupFile(filePath) {
  try {
    if (filePath && fs.existsSync(filePath)) {
      fs.unlinkSync(filePath);
    }
  } catch (_) {}
}

/* =========================================================
   HEALTH
========================================================= */

app.get("/", (req, res) => {
  res.json({
    success: true,
    service: "iLovePDF4 Free Conversion Engine",
    status: "online"
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
      error: "Job not found.",
      jobId
    });
  }

  return res.json({
    status: job.status,
    jobId,
    filename: job.filename || null,
    contentType: job.contentType || null,
    url:
      job.status === "finished"
        ? `${PUBLIC_BASE_URL}/download/${encodeURIComponent(jobId)}`
        : null,
    error: job.error || null
  });
});

/* =========================================================
   DOWNLOAD
========================================================= */

app.get("/download/:jobId", (req, res) => {
  const jobId = req.params.jobId;
  const job = jobs.get(jobId);

  if (!job) {
    return res.status(404).send("Job not found.");
  }

  if (job.status !== "finished") {
    return res.status(409).send("Conversion is not finished.");
  }

  if (!job.outputPath || !fs.existsSync(job.outputPath)) {
    return res.status(404).send("Output file not found.");
  }

  res.download(
    job.outputPath,
    job.outputName || "converted-file"
  );
});

/* =========================================================
   CONVERT
========================================================= */

app.post("/convert", upload.single("file"), async (req, res) => {
  let inputPath = null;

  try {
    const file = req.file;

    if (!file) {
      return res.status(400).json({
        success: false,
        error: "No file uploaded."
      });
    }

    inputPath = file.path;

    const tool = String(
      req.body.tool || ""
    ).trim().toLowerCase();

    const requestedJobId = String(
      req.body.jobId || ""
    ).trim();

    const jobId =
      requestedJobId ||
      `job-${Date.now()}-${Math.random()
        .toString(36)
        .slice(2, 10)}`;

    const supportedTools = [
      "word-to-pdf",
      "powerpoint-to-pdf",
      "excel-to-pdf",
      "pdf-to-word",
      "pdf-to-powerpoint",
      "pdf-to-excel"
    ];

    if (!supportedTools.includes(tool)) {
      cleanupFile(inputPath);

      return res.status(400).json({
        success: false,
        error: `Unsupported conversion tool: ${tool}`
      });
    }

    const outputExtension =
      extensionForTool(tool);

    const originalName =
      file.originalname || "file";

    const baseName =
      path.basename(
        originalName,
        path.extname(originalName)
      );

    const outputName =
      `${safeName(baseName)}.${outputExtension}`;

    jobs.set(jobId, {
      status: "processing",
      filename: originalName,
      outputName,
      outputPath: null,
      contentType: null,
      error: null
    });

    /*
     * IMPORTANT:
     * Return the response BEFORE the heavy conversion.
     * The actual conversion continues asynchronously.
     */
    res.json({
      success: true,
      jobId,
      status: "processing"
    });

    setImmediate(async () => {
      try {
        await processConversion({
          jobId,
          tool,
          inputPath,
          outputName
        });
      } catch (error) {
        console.error(
          "Conversion error:",
          error
        );

        const job = jobs.get(jobId);

        if (job) {
          job.status = "error";
          job.error =
            error?.message ||
            "Conversion failed.";
        }

        cleanupFile(inputPath);
      }
    });

  } catch (error) {
    console.error(
      "Upload/convert request error:",
      error
    );

    cleanupFile(inputPath);

    if (!res.headersSent) {
      return res.status(500).json({
        success: false,
        error:
          error?.message ||
          "Unable to start conversion."
      });
    }
  }
});

/* =========================================================
   PROCESS CONVERSION
========================================================= */

async function processConversion({
  jobId,
  tool,
  inputPath,
  outputName
}) {
  const job = jobs.get(jobId);

  if (!job) {
    cleanupFile(inputPath);
    return;
  }

  console.log(
    "===================================="
  );

  console.log(
    `Starting conversion: ${jobId}`
  );

  console.log(
    `Tool: ${tool}`
  );

  console.log(
    `Input: ${inputPath}`
  );

  try {
    if (tool === "pdf-to-word") {
      await convertPdfToWord(
        jobId,
        inputPath,
        outputName
      );
    } else {
      await convertWithLibreOffice(
        jobId,
        inputPath,
        tool,
        outputName
      );
    }

    cleanupFile(inputPath);

    console.log(
      `Conversion finished: ${jobId}`
    );

  } catch (error) {
    console.error(
      `Conversion failed: ${jobId}`,
      error
    );

    job.status = "error";
    job.error =
      error?.message ||
      "Conversion failed.";

    cleanupFile(inputPath);
  }
}

/* =========================================================
   PDF -> WORD
   LOW MEMORY OCR VERSION
========================================================= */

async function convertPdfToWord(
  jobId,
  inputPath,
  outputName
) {
  console.log(
    "PDF TO WORD ENGINE: low-memory pdftotext + OCR"
  );

  const job = jobs.get(jobId);

  const workDir =
    path.join(
      OUTPUT_DIR,
      jobId
    );

  fs.mkdirSync(
    workDir,
    { recursive: true }
  );

  const textPath =
    path.join(
      workDir,
      "extracted.txt"
    );

  let text = "";

  /* -------------------------------------------------------
     STEP 1: Normal PDF text extraction
  ------------------------------------------------------- */

  try {
    console.log(
      "Trying pdftotext..."
    );

    const result =
      await execFileAsync(
        "pdftotext",
        [
          "-layout",
          inputPath,
          textPath
        ],
        {
          maxBuffer: 5 * 1024 * 1024
        }
      );

    if (
      fs.existsSync(textPath)
    ) {
      text =
        fs.readFileSync(
          textPath,
          "utf8"
        );
    }

    console.log(
      `Extracted text length: ${text.length}`
    );

  } catch (error) {
    console.log(
      "pdftotext failed, switching to OCR."
    );
  }

  /* -------------------------------------------------------
     STEP 2: OCR if normal extraction found no text
  ------------------------------------------------------- */

  if (!text.trim()) {
    console.log(
      "No readable text found."
    );

    console.log(
      "Starting LOW MEMORY OCR..."
    );

    text =
      await runLowMemoryPdfOcr(
        inputPath,
        workDir
      );
  }

  if (!text.trim()) {
    text =
      "No readable text was found in this PDF.";
  }

  /* -------------------------------------------------------
     STEP 3: Build DOCX
  ------------------------------------------------------- */

  console.log(
    "Creating DOCX..."
  );

  const outputPath =
    path.join(
      workDir,
      outputName
    );

  await buildDocxFromText(
    text,
    outputPath
  );

  if (
    !fs.existsSync(outputPath)
  ) {
    throw new Error(
      "DOCX output was not created."
    );
  }

  const stat =
    fs.statSync(outputPath);

  console.log(
    `PDF TO WORD finished: ${jobId}`
  );

  console.log(
    `Output size: ${stat.size} bytes`
  );

  job.status = "finished";
  job.outputPath = outputPath;
  job.outputName = outputName;
  job.contentType =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
}

/* =========================================================
   LOW MEMORY PDF OCR
========================================================= */

async function runLowMemoryPdfOcr(
  inputPath,
  workDir
) {
  let pageCount = 0;

  try {
    const info =
      await execFileAsync(
        "pdfinfo",
        [inputPath],
        {
          maxBuffer: 1024 * 1024
        }
      );

    const match =
      info.stdout.match(
        /Pages:\s+(\d+)/i
      );

    if (match) {
      pageCount =
        parseInt(match[1], 10);
    }

  } catch (error) {
    console.log(
      "Could not read PDF page count."
    );
  }

  if (!pageCount) {
    throw new Error(
      "Could not determine PDF page count."
    );
  }

  console.log(
    `PDF pages: ${pageCount}`
  );

  const pageTexts = [];

  /*
   * IMPORTANT:
   * Only ONE page image exists in memory/disk at a time.
   */
  for (
    let page = 1;
    page <= pageCount;
    page++
  ) {
    console.log(
      `OCR page ${page}/${pageCount}`
    );

    const pagePrefix =
      path.join(
        workDir,
        `page-${page}`
      );

    const imagePath =
      `${pagePrefix}.jpg`;

    try {
      /* -----------------------------------------------
         Render only the current page.
         120 DPI keeps RAM usage low.
      ------------------------------------------------ */

      await execFileAsync(
        "pdftoppm",
        [
          "-f",
          String(page),
          "-singlefile",
          "-jpeg",
          "-r",
          "120",
          "-jpegopt",
          "quality=75",
          inputPath,
          pagePrefix
        ],
        {
          maxBuffer: 1024 * 1024
        }
      );

      if (
        !fs.existsSync(imagePath)
      ) {
        console.log(
          `Page ${page}: image not created.`
        );
        continue;
      }

      let pageText = "";

      /* -----------------------------------------------
         Bengali + English OCR
      ------------------------------------------------ */

      try {
        const result =
          await execFileAsync(
            "tesseract",
            [
              imagePath,
              "stdout",
              "-l",
              "ben+eng",
              "--psm",
              "3"
            ],
            {
              maxBuffer: 4 * 1024 * 1024
            }
          );

        pageText =
          result.stdout || "";

      } catch (ocrError) {
        console.log(
          `Bengali OCR failed on page ${page}, trying English fallback.`
        );

        try {
          const fallback =
            await execFileAsync(
              "tesseract",
              [
                imagePath,
                "stdout",
                "-l",
                "eng",
                "--psm",
                "6"
              ],
              {
                maxBuffer: 4 * 1024 * 1024
              }
            );

          pageText =
            fallback.stdout || "";

        } catch (fallbackError) {
          console.log(
            `OCR failed on page ${page}.`
          );
        }
      }

      pageTexts.push(
        pageText.trim()
      );

    } finally {
      /*
       * VERY IMPORTANT:
       * Delete page image immediately.
       */
      cleanupFile(imagePath);

      if (
        global.gc
      ) {
        try {
          global.gc();
        } catch (_) {}
      }
    }
  }

  return pageTexts
    .filter(Boolean)
    .join("\n\n\f\n\n");
}

/* =========================================================
   DOCX CREATION
========================================================= */

async function buildDocxFromText(
  text,
  outputPath
) {
  const sections =
    String(text)
      .split("\f");

  const paragraphs = [];

  for (
    let i = 0;
    i < sections.length;
    i++
  ) {
    const section =
      sections[i];

    const lines =
      section.split(/\r?\n/);

    for (const line of lines) {
      paragraphs.push(
        new Paragraph({
          children: [
            new TextRun({
              text: line || " ",
              font: "Noto Sans Bengali",
              size: 22
            })
          ],
          spacing: {
            after: 80
          }
        })
      );
    }

    if (
      i < sections.length - 1
    ) {
      paragraphs.push(
        new Paragraph({
          pageBreakBefore: true
        })
      );
    }
  }

  const document =
    new Document({
      sections: [
        {
          properties: {},
          children: paragraphs
        }
      ]
    });

  const buffer =
    await Packer.toBuffer(
      document
    );

  fs.writeFileSync(
    outputPath,
    buffer
  );
}

/* =========================================================
   LIBREOFFICE CONVERSIONS
========================================================= */

async function convertWithLibreOffice(
  jobId,
  inputPath,
  tool,
  outputName
) {
  const job = jobs.get(jobId);

  const workDir =
    path.join(
      OUTPUT_DIR,
      jobId
    );

  fs.mkdirSync(
    workDir,
    { recursive: true }
  );

  const outputPath =
    path.join(
      workDir,
      outputName
    );

  console.log(
    `LibreOffice conversion: ${tool}`
  );

  const profileDir =
    path.join(
      LO_PROFILE_DIR,
      jobId
    );

  fs.mkdirSync(
    profileDir,
    { recursive: true }
  );

  let args = [
    "--headless",
    "--convert-to"
  ];

  if (
    tool === "word-to-pdf"
  ) {
    args.push(
      "pdf:writer_pdf_Export"
    );

  } else if (
    tool === "powerpoint-to-pdf"
  ) {
    args.push(
      "pdf:impress_pdf_Export"
    );

  } else if (
    tool === "excel-to-pdf"
  ) {
    args.push(
      "pdf:calc_pdf_Export"
    );

  } else if (
    tool === "pdf-to-powerpoint"
  ) {
    args.push(
      "pptx:Impress MS PowerPoint 2007 XML"
    );

  } else if (
    tool === "pdf-to-excel"
  ) {
    args.push(
      "xlsx:Calc MS Excel 2007 XML"
    );

  } else {
    throw new Error(
      `Unsupported LibreOffice tool: ${tool}`
    );
  }

  args.push(
    "--outdir",
    workDir,
    inputPath
  );

  const env =
    {
      ...process.env,
      HOME: profileDir
    };

  await execFileAsync(
    "libreoffice",
    args,
    {
      env,
      maxBuffer: 5 * 1024 * 1024,
      timeout: 180000
    }
  );

  const generatedFiles =
    fs.readdirSync(
      workDir
    );

  let generated =
    generatedFiles.find(
      file =>
        file !== outputName
    );

  if (!generated) {
    throw new Error(
      "LibreOffice did not create an output file."
    );
  }

  const generatedPath =
    path.join(
      workDir,
      generated
    );

  /*
   * Rename LibreOffice output
   * to the exact requested filename.
   */
  if (
    generatedPath !== outputPath
  ) {
    cleanupFile(outputPath);

    fs.renameSync(
      generatedPath,
      outputPath
    );
  }

  if (
    !fs.existsSync(outputPath)
  ) {
    throw new Error(
      "Output file was not created."
    );
  }

  const stat =
    fs.statSync(outputPath);

  console.log(
    `LibreOffice output size: ${stat.size} bytes`
  );

  job.status = "finished";
  job.outputPath = outputPath;
  job.outputName = outputName;

  const contentTypes = {
    "word-to-pdf":
      "application/pdf",

    "powerpoint-to-pdf":
      "application/pdf",

    "excel-to-pdf":
      "application/pdf",

    "pdf-to-powerpoint":
      "application/vnd.openxmlformats-officedocument.presentationml.presentation",

    "pdf-to-excel":
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  };

  job.contentType =
    contentTypes[tool] ||
    "application/octet-stream";
}

/* =========================================================
   ERROR HANDLING
========================================================= */

app.use(
  (error, req, res, next) => {
    console.error(
      "Express error:",
      error
    );

    if (
      error &&
      error.code === "LIMIT_FILE_SIZE"
    ) {
      return res.status(413).json({
        success: false,
        error:
          "File is too large. Maximum size is 50 MB."
      });
    }

    if (!res.headersSent) {
      return res.status(500).json({
        success: false,
        error:
          error?.message ||
          "Internal server error."
      });
    }

    next(error);
  }
);

/* =========================================================
   START SERVER
========================================================= */

app.listen(
  PORT,
  HOST,
  () => {
    console.log(
      "===================================="
    );

    console.log(
      "iLovePDF4 Conversion Engine"
    );

    console.log(
      `Running on port ${PORT}`
    );

    console.log(
      `Public URL: ${PUBLIC_BASE_URL}`
    );

    console.log(
      "Supported server tools:"
    );

    console.log(
      " - word-to-pdf"
    );

    console.log(
      " - powerpoint-to-pdf"
    );

    console.log(
      " - excel-to-pdf"
    );

    console.log(
      " - pdf-to-word"
    );

    console.log(
      " - pdf-to-powerpoint"
    );

    console.log(
      " - pdf-to-excel"
    );

    console.log(
      "===================================="
    );
  }
);
