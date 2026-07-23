import express from "express";
import multer from "multer";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { comparePurchaseLedgers } from "./lib/compare.mjs";

const app = express();
const root = path.dirname(fileURLToPath(import.meta.url));
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024, files: 2 },
  fileFilter: (_req, file, done) => done(null, /\.xlsx?$/i.test(file.originalname)),
});

app.use(express.static(path.join(root, "public")));
app.post("/api/compare", upload.fields([{ name: "aFile", maxCount: 1 }, { name: "bFile", maxCount: 1 }]), (req, res, next) => {
  try {
    const aFile = req.files?.aFile?.[0];
    const bFile = req.files?.bFile?.[0];
    if (!aFile || !bFile) return res.status(400).json({ error: "A와 B 엑셀 파일을 모두 선택해 주세요." });
    const { output, stats } = comparePurchaseLedgers(aFile.buffer, bFile.buffer);
    res.set({
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "Content-Disposition": "attachment; filename*=UTF-8''B_%EC%83%89%EC%83%81%ED%91%9C%EC%8B%9C_%EB%B9%84%EA%B5%90%EA%B2%B0%EA%B3%BC.xlsx",
      "X-Compare-Stats": encodeURIComponent(JSON.stringify(stats)),
      "Cache-Control": "no-store",
    });
    return res.send(output);
  } catch (error) {
    return next(error);
  }
});
app.use((error, _req, res, _next) => {
  const message = error instanceof multer.MulterError ? "파일은 각각 10MB 이하의 .xls 또는 .xlsx만 업로드할 수 있습니다." : "엑셀 처리 중 오류가 발생했습니다. 파일 형식을 확인해 주세요.";
  res.status(400).json({ error: message });
});

const port = process.env.PORT || 3000;
app.listen(port, "0.0.0.0", () => console.log(`Excel comparer listening on ${port}`));
