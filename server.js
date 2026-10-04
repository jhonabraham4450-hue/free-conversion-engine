const express = require("express");
const multer = require("multer");
const cors = require("cors");
const fs = require("fs");
const path = require("path");
const os = require("os");
const { execFile } = require("child_process");
const pptxgen = require("pptxgenjs");

const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  ImageRun
} = require("docx");

const app = express();

app.use(cors());
app.use(express.json({ limit: "50mb" }));
app.use(express.urlencoded({ extended: true, limit: "50mb" }));

// ==================================================
// DIRECTORIES
// ==================================================

const uploadDir = path.join(os.tmpdir(), "uploads");
const outputDir = path.join(os.tmpdir(), "output");

fs.mkdirSync(uploadDir, { recursive: true });
fs.mkdirSync(outputDir, { recursive: true });

// ==================================================
// MULTER
// ==================================================

const upload = multer({
  dest: uploadDir,
  limits: {
    fileSize: 512 * 1024 * 1024
  }
});

// ==================================================
// JOB STORAGE
// ==================================================

const jobs = new Map();

// ==================================================
// ROOT
// ==================================================

app.get("/", (req, res) => {
  res.json({
    success: true,
    service: "iLovePDF4 Free Conversion Engine",
    status: "online"
  });
});

// ==================================================
// STATUS
// ==================================================

app.get("/status/:jobId", (req, res) => {
  const job = jobs.get(req.params.jobId);

  if (!job) {
    return res.status(404).json({
      success: false,
      error: "Job not found."
    });
  }

  res.json({
    success: true,
    jobId: req.params.jobId,
    status: job.status,
    error: job.error || null,
    filename: job.filename || null
  });
});

// ==================================================
// DOWNLOAD
// ==================================================

app.get("/download/:jobId", (req, res) => {
  const job = jobs.get(req.params.jobId);

  if (!job) {
    return res.status(404).json({
      success: false,
      error: "Job not found."
    });
  }

  if (job.status !== "finished") {
    return res.status(400).json({
      success: false,
      error: "Conversion is not finished yet."
    });
  }

  if (!job.outputPath || !fs.existsSync(job.outputPath)) {
    return res.status(404).json({
      success: false,
      error: "Output file not found."
    });
  }

  res.download(
    job.outputPath,
    job.filename || "converted-file",
    err => {
      if (err) {
        console.error("Download error:", err.message);
      }
    }
  );
});

// ==================================================
// CONVERSION MAP
// ==================================================

const conversionMap = {
  "word-to-pdf": {
    input: ["doc", "docx"],
    output: "pdf",
    extension: ".pdf"
  },

  "powerpoint-to-pdf": {
    input: ["ppt", "pptx"],
    output: "pdf",
    extension: ".pdf"
  },

  "excel-to-pdf": {
    input: ["xls", "xlsx"],
    output: "pdf",
    extension: ".pdf"
  },

  "pdf-to-word": {
    input: ["pdf"],
    output: "docx",
    extension: ".docx"
  },

  "pdf-to-powerpoint": {
    input: ["pdf"],
    output: "pptx",
    extension: ".pptx"
  },

  "pdf-to-excel": {
    input: ["pdf"],
    output: "xlsx",
    extension: ".xlsx"
  }
};

// ==================================================
// CLEANUP DIRECTORY
// ==================================================

function cleanupDirectory(dir) {
  try {
    if (!fs.existsSync(dir)) return;

    const files = fs.readdirSync(dir);

    for (const file of files) {
      const filePath = path.join(dir, file);

      try {
        const stat = fs.statSync(filePath);

        if (stat.isDirectory()) {
          cleanupDirectory(filePath);
          fs.rmdirSync(filePath);
        } else {
          fs.unlinkSync(filePath);
        }
      } catch {}
    }

    try {
      fs.rmdirSync(dir);
    } catch {}
  } catch {}
}

// ==================================================
// FIND OUTPUT FILE
// ==================================================

function findOutputFile(dir, extension) {
  try {
    if (!fs.existsSync(dir)) {
      return null;
    }

    const files = fs.readdirSync(dir);

    const wantedExtension = extension.toLowerCase();

    for (const file of files) {
      const fullPath = path.join(dir, file);

      try {
        const stat = fs.statSync(fullPath);

        if (
          stat.isFile() &&
          path.extname(file).toLowerCase() === wantedExtension
        ) {
          return fullPath;
        }
      } catch {}
    }
  } catch {}

  return null;
}

// ==================================================
// PDF -> WORD + OCR
// ==================================================

async function convertScannedPdfToDocx(inputPath, outputPath) {
  const tempDir = fs.mkdtempSync(
    path.join(os.tmpdir(), "pdf-render-")
  );

  try {
    const prefix = path.join(tempDir, "page");

    console.log("=================================");
    console.log("PDF -> WORD + OCR");
    console.log("Input:", inputPath);
    console.log("=================================");

    // ------------------------------------------------
    // RENDER PDF
    // ------------------------------------------------

    await new Promise((resolve, reject) => {
      execFile(
        "pdftoppm",
        [
          "-png",
          "-r",
          "96",
          inputPath,
          prefix
        ],
        {
          timeout: 180000,
          maxBuffer: 50 * 1024 * 1024
        },
        (error, stdout, stderr) => {
          if (error) {
            reject(
              new Error(
                stderr?.trim() ||
                error.message ||
                "PDF rendering failed."
              )
            );
            return;
          }

          resolve();
        }
      );
    });

    const files = fs
      .readdirSync(tempDir)
      .filter(file =>
        /^page-\d+\.png$/i.test(file)
      )
      .sort((a, b) => {
        const na = parseInt(
          a.match(/(\d+)/)[1],
          10
        );

        const nb = parseInt(
          b.match(/(\d+)/)[1],
          10
        );

        return na - nb;
      });

    if (!files.length) {
      throw new Error(
        "No PDF pages were rendered."
      );
    }

    console.log(
      "Rendered pages:",
      files.length
    );

    // ------------------------------------------------
    // PNG SIZE
    // ------------------------------------------------

    function getPngSize(filePath) {
      const buffer = fs.readFileSync(filePath);

      if (
        buffer.length < 24 ||
        buffer.readUInt32BE(0) !== 0x89504e47
      ) {
        throw new Error(
          "Invalid PNG output."
        );
      }

      return {
        width: buffer.readUInt32BE(16),
        height: buffer.readUInt32BE(20)
      };
    }

    // ------------------------------------------------
    // OCR
    // ------------------------------------------------

    async function runOCR(imagePath) {
      return new Promise(resolve => {
        execFile(
          "tesseract",
          [
            imagePath,
            "stdout",
            "-l",
            "ben+eng",
            "--psm",
            "6"
          ],
          {
            timeout: 120000,
            maxBuffer: 20 * 1024 * 1024
          },
          (error, stdout, stderr) => {
            if (error) {
              console.error(
                "OCR failed:",
                stderr || error.message
              );

              resolve("");
              return;
            }

            resolve(
              (stdout || "")
                .replace(/\r/g, "")
                .trim()
            );
          }
        );
      });
    }

    // ------------------------------------------------
    // CREATE SECTIONS
    // ------------------------------------------------

    const sections = [];

    for (let i = 0; i < files.length; i++) {
      const file = files[i];

      const imagePath = path.join(
        tempDir,
        file
      );

      console.log(
        `Processing page ${i + 1}/${files.length}`
      );

      const imageBuffer =
        fs.readFileSync(imagePath);

      const {
        width,
        height
      } = getPngSize(imagePath);

      // ------------------------------------------------
      // OCR
      // ------------------------------------------------

      console.log(
        `Running OCR on page ${i + 1}...`
      );

      const ocrText =
        await runOCR(imagePath);

      console.log(
        `OCR page ${i + 1}:`,
        ocrText
          ? `${ocrText.length} characters`
          : "No text detected"
      );

      // ------------------------------------------------
      // PAGE SIZE
      // ------------------------------------------------

      const pageWidthTwips =
        Math.round(width * 15);

      const pageHeightTwips =
        Math.round(height * 15);

      // ------------------------------------------------
      // IMAGE
      // ------------------------------------------------

      const children = [];

      children.push(
        new Paragraph({
          spacing: {
            before: 0,
            after: 0
          },

          children: [
            new ImageRun({
              type: "png",
              data: imageBuffer,

              transformation: {
                width: Math.max(
                  100,
                  width - 8
                ),

                height: Math.max(
                  100,
                  height - 8
                )
              }
            })
          ]
        })
      );

      // ------------------------------------------------
      // OCR TEXT
      // ------------------------------------------------

      if (ocrText) {
        children.push(
          new Paragraph({
            spacing: {
              before: 100,
              after: 100
            },

            children: [
              new TextRun({
                text: ocrText,
                size: 20
              })
            ]
          })
        );
      }

      // ------------------------------------------------
      // SECTION
      // ------------------------------------------------

      sections.push({
        properties: {
          page: {
            size: {
              width: pageWidthTwips,
              height: pageHeightTwips
            },

            margin: {
              top: 120,
              bottom: 120,
              left: 120,
              right: 120,
              header: 0,
              footer: 0,
              gutter: 0
            }
          }
        },

        children
      });
    }

    // ------------------------------------------------
    // CREATE DOCX
    // ------------------------------------------------

    const doc = new Document({
      sections
    });

    const buffer =
      await Packer.toBuffer(doc);

    if (
      !buffer ||
      buffer.length < 1000
    ) {
      throw new Error(
        "DOCX output is invalid or empty."
      );
    }

    fs.writeFileSync(
      outputPath,
      buffer
    );

    console.log(
      "PDF -> WORD + OCR completed:",
      outputPath
    );

    console.log(
      "DOCX size:",
      buffer.length,
      "bytes"
    );

  } finally {
    cleanupDirectory(tempDir);
  }
}

// ==================================================
// LIBREOFFICE ERROR
// ==================================================

function getLibreOfficeError(
  tool,
  stderr,
  error
) {
  const message =
    stderr?.trim() ||
    error?.message ||
    "LibreOffice conversion failed.";

  return `${tool} conversion failed: ${message}`;
}

// ==================================================
// MAIN CONVERT API
// ==================================================

app.post(
  "/convert",
  upload.single("file"),
  async (req, res) => {

    let inputPath = null;

    try {
      // ------------------------------------------------
      // CHECK FILE
      // ------------------------------------------------

      if (!req.file) {
        return res.status(400).json({
          success: false,
          error: "No file uploaded."
        });
      }

      inputPath = req.file.path;

      // Preserve original file extension for LibreOffice
      const originalExt = path.extname(
        req.file.originalname || ""
      ).toLowerCase();

      if (originalExt) {
        const renamedInputPath =
          inputPath + originalExt;

        fs.renameSync(
          inputPath,
          renamedInputPath
        );

        inputPath = renamedInputPath;
      }

      const tool =
        String(
          req.body.tool ||
          req.query.tool ||
          ""
        )
          .trim()
          .toLowerCase();

      console.log(
        "================================="
      );

      console.log(
        "Starting conversion:",
        tool
      );

      console.log(
        "Original filename:",
        req.file.originalname
      );

      console.log(
        "Input:",
        inputPath
      );

      console.log(
        "================================="
      );

      // ------------------------------------------------
      // CHECK TOOL
      // ------------------------------------------------

      const conversion =
        conversionMap[tool];

      if (!conversion) {
        try {
          fs.unlinkSync(inputPath);
        } catch {}

        return res.status(400).json({
          success: false,
          error:
            "Unsupported conversion tool."
        });
      }

      // ------------------------------------------------
      // JOB
      // ------------------------------------------------

      const jobId =
        `${Date.now()}-${Math.random()
          .toString(36)
          .slice(2, 10)}`;

      const jobDir =
        path.join(
          outputDir,
          jobId
        );

      fs.mkdirSync(
        jobDir,
        { recursive: true }
      );

      const jobOutputDir =
        jobDir;

      // ------------------------------------------------
      // OUTPUT NAME
      // ------------------------------------------------

      const originalName =
        path.parse(
          req.file.originalname ||
          "converted-file"
        ).name;

      const safeName =
        originalName
          .replace(
            /[^a-zA-Z0-9._-]/g,
            "_"
          );

      const outputPath =
        path.join(
          jobOutputDir,
          `${safeName}${conversion.extension}`
        );

      let outputFilename =
        `${safeName}${conversion.extension}`;

      jobs.set(jobId, {
        status: "processing",
        outputPath: null,
        filename: outputFilename,
        error: null
      });

      // ------------------------------------------------
      // RESPONSE FIRST
      // ------------------------------------------------

      res.json({
        success: true,
        jobId,
        status: "processing"
      });

      // ------------------------------------------------
      // PDF -> WORD
      // ------------------------------------------------

      if (tool === "pdf-to-word") {

        try {
          await convertScannedPdfToDocx(
            inputPath,
            outputPath
          );

          if (
            !fs.existsSync(outputPath)
          ) {
            throw new Error(
              "DOCX output file was not created."
            );
          }

          const stat =
            fs.statSync(outputPath);

          if (stat.size < 1000) {
            throw new Error(
              "DOCX output is too small."
            );
          }

          jobs.set(jobId, {
            status: "finished",
            outputPath,
            filename: outputFilename,
            error: null
          });

          console.log(
            "Job finished:",
            jobId
          );

        } catch (error) {

          console.error(
            "PDF -> WORD error:",
            error
          );

          jobs.set(jobId, {
            status: "error",
            outputPath: null,
            filename: outputFilename,
            error:
              error.message ||
              "PDF to Word conversion failed."
          });
        }

        try {
          if (inputPath && fs.existsSync(inputPath)) {
            fs.unlinkSync(inputPath);
          }
        } catch {}

        return;
      }
            // ------------------------------------------------
      // PDF -> POWERPOINT
      // ------------------------------------------------

      if (tool === "pdf-to-powerpoint") {

        try {

          console.log(
            "Starting custom PDF -> PowerPoint conversion..."
          );

          const tempDir = fs.mkdtempSync(
            path.join(
              os.tmpdir(),
              "pdf-ppt-"
            )
          );

          try {

            const prefix =
              path.join(
                tempDir,
                "page"
              );

            // ------------------------------------------
            // RENDER PDF PAGES TO PNG
            // ------------------------------------------

            await new Promise(
              (resolve, reject) => {

                execFile(
                  "pdftoppm",
                  [
                    "-png",
                    "-r",
                    "120",
                    inputPath,
                    prefix
                  ],
                  {
                    timeout: 300000,
                    maxBuffer:
                      100 * 1024 * 1024
                  },
                  (
                    error,
                    stdout,
                    stderr
                  ) => {

                    if (error) {

                      reject(
                        new Error(
                          stderr?.trim() ||
                          error.message ||
                          "PDF rendering failed."
                        )
                      );

                      return;
                    }

                    resolve();
                  }
                );
              }
            );

            // ------------------------------------------
            // FIND PNG PAGES
            // ------------------------------------------

            const pageFiles =
              fs
                .readdirSync(tempDir)
                .filter(
                  file =>
                    /^page-\d+\.png$/i.test(
                      file
                    )
                )
                .sort(
                  (a, b) => {

                    const aNum =
                      parseInt(
                        a.match(
                          /(\d+)/
                        )[1],
                        10
                      );

                    const bNum =
                      parseInt(
                        b.match(
                          /(\d+)/
                        )[1],
                        10
                      );

                    return aNum - bNum;
                  }
                );

            if (
              !pageFiles.length
            ) {

              throw new Error(
                "No PDF pages were rendered."
              );
            }

            console.log(
              "PDF pages:",
              pageFiles.length
            );

            // ------------------------------------------
            // CREATE POWERPOINT
            // ------------------------------------------

            const pptx =
              new pptxgen();

            pptx.layout = "LAYOUT_STANDARD (4:3)";
            pptx.author =
              "iLovePDF4";

            pptx.subject =
              "Converted PDF";

            pptx.title =
              safeName;

            pptx.company =
              "iLovePDF4";

            pptx.lang =
              "en-US";

            // Disable automatic layout
            pptx.theme = {
              headFontFace:
                "Arial",
              bodyFontFace:
                "Arial",
              lang:
                "en-US"
            };

            // ------------------------------------------
            // ADD EACH PDF PAGE AS SLIDE IMAGE
            // ------------------------------------------

            for (
              let i = 0;
              i < pageFiles.length;
              i++
            ) {

              const pageFile =
                pageFiles[i];

              const pagePath =
                path.join(
                  tempDir,
                  pageFile
                );

              console.log(
                `Adding slide ${i + 1}/${pageFiles.length}`
              );

              const slide =
                pptx.addSlide();

              slide.background = {
                color: "FFFFFF"
              };

              slide.addImage({
                path: pagePath,
                x: 0,
                y: 0,
                w: 10,
                h: 5.625
              });
            }

            // ------------------------------------------
            // WRITE PPTX
            // ------------------------------------------

            await pptx.writeFile({
              fileName:
                outputPath
            });

            // ------------------------------------------
            // VALIDATE OUTPUT
            // ------------------------------------------

            if (
              !fs.existsSync(
                outputPath
              )
            ) {

              throw new Error(
                "PowerPoint output file was not created."
              );
            }

            const stat =
              fs.statSync(
                outputPath
              );

            console.log(
              "PPTX size:",
              stat.size,
              "bytes"
            );

            if (
              stat.size < 10000
            ) {

              throw new Error(
                "Generated PowerPoint file is invalid or too small."
              );
            }

            // ------------------------------------------
            // FINISHED
            // ------------------------------------------

            jobs.set(jobId, {
              status: "finished",
              outputPath,
              filename:
                outputFilename,
              error: null
            });

            console.log(
              "PDF -> PowerPoint completed:",
              jobId
            );

          } finally {

            cleanupDirectory(
              tempDir
            );
          }

        } catch (error) {

          console.error(
            "PDF -> PowerPoint error:",
            error
          );

          jobs.set(jobId, {
            status: "error",
            outputPath: null,
            filename:
              outputFilename,
            error:
              error.message ||
              "PDF to PowerPoint conversion failed."
          });
        }

        try {

          if (
            inputPath &&
            fs.existsSync(inputPath)
          ) {

            fs.unlinkSync(
              inputPath
            );
          }

        } catch {}

        return;
      }

      // ------------------------------------------------
      // LIBREOFFICE CONVERSIONS
      // ------------------------------------------------

      const profileDir =
        fs.mkdtempSync(
          path.join(
            os.tmpdir(),
            "lo-profile-"
          )
        );

      const libreOfficeOutputDir =
        jobOutputDir;

      let libreOfficeArgs = [
        "--headless",
        "--nologo",
        "--nodefault",
        "--nofirststartwizard",
        "--nolockcheck",
        "-env:UserInstallation=file://" +
          profileDir,
        "--convert-to"
      ];

      // ------------------------------------------------
      // WORD -> PDF
      // ------------------------------------------------

      if (
        tool === "word-to-pdf"
      ) {

        libreOfficeArgs.push(
          "pdf:writer_pdf_Export"
        );
      }

      // ------------------------------------------------
      // POWERPOINT -> PDF
      // ------------------------------------------------

      else if (
        tool === "powerpoint-to-pdf"
      ) {

        libreOfficeArgs.push(
          "pdf:impress_pdf_Export"
        );
      }

      // ------------------------------------------------
      // EXCEL -> PDF
      // ------------------------------------------------

      else if (
        tool === "excel-to-pdf"
      ) {

        libreOfficeArgs.push(
          "pdf:calc_pdf_Export"
        );
      }

      // ------------------------------------------------
      // PDF -> EXCEL
      // ------------------------------------------------

      else if (
        tool === "pdf-to-excel"
      ) {

        libreOfficeArgs.push(
          "xlsx:Calc MS Excel 2007 XML"
        );

        libreOfficeArgs.push(
          "--infilter=draw_pdf_import"
        );
      }

      // ------------------------------------------------
      // UNSUPPORTED
      // ------------------------------------------------

      else {

        throw new Error(
          "Conversion method not available for this tool."
        );
      }

      libreOfficeArgs.push(
        "--outdir",
        libreOfficeOutputDir,
        inputPath
      );

      console.log(
        "LibreOffice command:"
      );

      console.log(
        "soffice",
        libreOfficeArgs.join(" ")
      );

      // ------------------------------------------------
      // RUN LIBREOFFICE
      // ------------------------------------------------

      await new Promise(
        (resolve, reject) => {

          execFile(
            "libreoffice",
            libreOfficeArgs,
            {
              timeout: 300000,
              maxBuffer:
                100 * 1024 * 1024
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

              if (error) {

                reject(
                  new Error(
                    getLibreOfficeError(
                      tool,
                      stderr,
                      error
                    )
                  )
                );

                return;
              }

              resolve();
            }
          );
        }
      );

      // ------------------------------------------------
      // FIND OUTPUT
      // ------------------------------------------------

      const generatedOutput =
        findOutputFile(
          libreOfficeOutputDir,
          conversion.extension
        );

      if (
        !generatedOutput
      ) {

        throw new Error(
          "LibreOffice did not create the expected output file."
        );
      }

      // ------------------------------------------------
      // VALIDATE OUTPUT
      // ------------------------------------------------

      const outputStat =
        fs.statSync(
          generatedOutput
        );

      console.log(
        "Generated output:",
        generatedOutput
      );

      console.log(
        "Output size:",
        outputStat.size,
        "bytes"
      );

      if (
        outputStat.size < 1000
      ) {

        throw new Error(
          "Generated output file is too small."
        );
      }

      // ------------------------------------------------
      // RENAME TO EXPECTED NAME
      // ------------------------------------------------

      if (
        generatedOutput !==
        outputPath
      ) {

        try {

          if (
            fs.existsSync(
              outputPath
            )
          ) {

            fs.unlinkSync(
              outputPath
            );
          }

          fs.renameSync(
            generatedOutput,
            outputPath
          );

        } catch {

          fs.copyFileSync(
            generatedOutput,
            outputPath
          );
        }
      }

      // ------------------------------------------------
      // FINISHED
      // ------------------------------------------------

      jobs.set(jobId, {
        status: "finished",
        outputPath,
        filename:
          outputFilename,
        error: null
      });

      console.log(
        "Conversion completed:",
        jobId
      );

      // ------------------------------------------------
      // CLEANUP INPUT
      // ------------------------------------------------

      try {

        if (
          inputPath &&
          fs.existsSync(inputPath)
        ) {

          fs.unlinkSync(
            inputPath
          );
        }

      } catch {}

      // ------------------------------------------------
      // CLEANUP LIBREOFFICE PROFILE
      // ------------------------------------------------

      cleanupDirectory(
        profileDir
      );

    } catch (error) {

      console.error(
        "Conversion error:",
        error
      );

      // ------------------------------------------------
      // IF RESPONSE WAS NOT SENT
      // ------------------------------------------------

      if (!res.headersSent) {

        return res.status(500).json({
          success: false,
          error:
            error.message ||
            "Conversion failed."
        });
      }

      // ------------------------------------------------
      // UPDATE JOB ERROR
      // ------------------------------------------------

      const jobId =
        Array.from(
          jobs.entries()
        )
          .reverse()
          .find(
            ([id, job]) =>
              job.status ===
              "processing"
          )?.[0];

      if (jobId) {

        const existing =
          jobs.get(jobId);

        jobs.set(jobId, {
          ...existing,
          status: "error",
          outputPath: null,
          error:
            error.message ||
            "Conversion failed."
        });
      }

      // ------------------------------------------------
      // CLEANUP
      // ------------------------------------------------

      try {

        if (
          inputPath &&
          fs.existsSync(inputPath)
        ) {

          fs.unlinkSync(
            inputPath
          );
        }

      } catch {}
    }
  }
);

// ==================================================
// HEALTH CHECK
// ==================================================

app.get("/health", (req, res) => {

  res.json({
    success: true,
    status: "ok",
    service:
      "free-conversion-engine"
  });

});

// ==================================================
// 404
// ==================================================

app.use(
  (req, res) => {

    res.status(404).json({
      success: false,
      error: "Route not found."
    });

  }
);

// ==================================================
// ERROR HANDLER
// ==================================================

app.use(
  (
    error,
    req,
    res,
    next
  ) => {

    console.error(
      "Unhandled server error:",
      error
    );

    if (
      res.headersSent
    ) {
      return next(error);
    }

    res.status(500).json({
      success: false,
      error:
        error.message ||
        "Internal server error."
    });

  }
);

// ==================================================
// PORT
// ==================================================

const PORT =
  process.env.PORT ||
  10000;

app.listen(
  PORT,
  "0.0.0.0",
  () => {

    console.log(
      "================================="
    );

    console.log(
      "Free Conversion Engine started"
    );

    console.log(
      "Port:",
      PORT
    );

    console.log(
      "================================="
    );

  }
);
// ==================================================
// END OF SERVER
// ==================================================
