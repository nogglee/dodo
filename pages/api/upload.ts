import { NextApiRequest, NextApiResponse } from "next";
import formidable from "formidable";
import { readFile } from "fs/promises";
import * as XLSX from "xlsx";

export const config = {
  api: {
    bodyParser: false, // FormData 사용 가능하도록 설정
  },
};

export default async function handler(req: NextApiRequest, res: NextApiResponse) {
  if (req.method !== "POST") return res.status(405).json({ error: "Method Not Allowed" });

  // 파일 업로드 처리
  const form = new formidable.IncomingForm();
  form.uploadDir = "/tmp"; // Mac/Linux에서는 /tmp 폴더 사용 가능
  form.keepExtensions = true;

  form.parse(req, async (err, fields, files) => {
    if (err) return res.status(500).json({ error: "File upload failed" });

    const file = files.file as formidable.File;
    const data = await readFile(file.filepath);
    const workbook = XLSX.read(data, { type: "buffer" });

    // 첫 번째 시트 가져오기
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // 특정 셀 포맷팅 예제 (A:14 → B:3, C:14 → H:3)
    const formattedData = jsonData.slice(14).map((row: any) => ({
      name: row[0], // A열
      phone: row[1], // B열
      amount: row[2], // C열
    }));

    res.json({ formattedData });
  });
}