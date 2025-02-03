import { NextApiRequest, NextApiResponse } from "next";
import { IncomingForm, File } from "formidable";
import { readFile } from "fs/promises";
import * as XLSX from "xlsx";
import path from "path";

// Next.js에서 FormData를 처리하도록 설정
export const config = {
  api: {
    bodyParser: false,
  },
};

// Formidable을 Promise로 감싸서 파일 업로드 처리
const parseForm = (req: NextApiRequest): Promise<{ files: formidable.Files }> => {
  return new Promise((resolve, reject) => {
    const form = new IncomingForm({
      multiples: false,
      uploadDir: "/tmp",
      keepExtensions: true,
    });

    form.parse(req, (err, fields, files) => {
      if (err) reject(err);
      else resolve({ files });
    });
  });
};

export default async function handler(req: NextApiRequest, res: NextApiResponse) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method Not Allowed" });
  }

  try {
    // 파일 업로드 처리
    const { files } = await parseForm(req);
    console.log("Uploaded files:", files);

    if (!files.file) {
      return res.status(400).json({ error: "File not found or invalid file" });
    }

    const file = Array.isArray(files.file) ? files.file[0] : (files.file as File);
    if (!file.filepath) {
      return res.status(400).json({ error: "Invalid file path" });
    }

    console.log("Processing file:", file.filepath);

    // 파일 읽기 (암호 해제 포함)
    try {
      const workbook = XLSX.readFile(file.filepath, { password: "1928103362" });

      // 첫 번째 시트 가져오기
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      console.log("Excel data loaded:", jsonData.length);

      // 특정 셀 포맷팅 예제 (A:14 → B:3, C:14 → H:3)
      const formattedData = jsonData.slice(14).map((row: any) => ({
        name: row[2], // A열
        phone: row[1], // B열
        amount: row[2], // C열
      }));

      res.json({ formattedData });
    } catch (error) {
      console.error("Failed to read the encrypted file:", error);
      return res.status(500).json({ error: "Failed to read the encrypted file" });
    }
  } catch (error) {
    console.error("Processing error:", error);
    res.status(500).json({ error: "Internal Server Error" });
  }
}