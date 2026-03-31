import express from "express";
import cors from "cors";
import multer from "multer";
import fs from "node:fs";
import { generateWordBuffer, validateFormData, generateWordBufferCustom } from "./template";

const upload = multer({ 
  dest: "uploads/",
  limits: { fieldSize: 50 * 1024 * 1024 } // 允许 50MB 的字段大小，以防包含多张高清大图的 base64 字符串
});
const app = express();
const port = Number(process.env.PORT || 3011);

app.use(cors());
app.use(express.json({ limit: "50mb" })); // 提升 JSON 解析限制到 50MB
app.use(express.urlencoded({ limit: "50mb", extended: true }));

app.get("/api/health", (_request, response) => {
  response.json({ ok: true });
});

app.post("/api/export-word", (request, response) => {
  try {
    const formData = validateFormData(request.body);
    const wordBuffer = generateWordBuffer(formData);
    const fileName = `${formData.report_no}.docx`;

    response.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
    response.setHeader("Content-Disposition", `attachment; filename*=UTF-8''${encodeURIComponent(fileName)}`);
    response.send(wordBuffer);
  } catch (error: any) {
    console.error(error);
    if (error.properties && error.properties.errors instanceof Array) {
      const errorMessages = error.properties.errors.map(function (error: any) {
          return error.properties.explanation;
      }).join("\n");
      console.log('errorMessages', errorMessages);
    }
    const message = error instanceof Error ? error.message : "导出失败";
    response.status(400).json({ message });
  }
});

app.post("/api/export-word-custom", upload.single("template"), (request, response) => {
  try {
    if (!request.file) {
      throw new Error("请上传模板文件");
    }
    const rawData = JSON.parse(request.body.data);
    const formData = validateFormData(rawData);
    
    // 使用用户上传的模板文件
    const wordBuffer = generateWordBufferCustom(formData, request.file.path);
    const fileName = `${formData.report_no}_custom.docx`;

    response.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
    response.setHeader("Content-Disposition", `attachment; filename*=UTF-8''${encodeURIComponent(fileName)}`);
    response.send(wordBuffer);
  } catch (error: any) {
    console.error(error);
    const message = error instanceof Error ? error.message : "自定义导出失败";
    response.status(400).json({ message });
  } finally {
    // 导出完成后删除临时文件
    if (request.file && fs.existsSync(request.file.path)) {
      fs.unlinkSync(request.file.path);
    }
  }
});

app.listen(port, () => {
  process.stdout.write(`Backend server running at http://localhost:${port}\n`);
});
