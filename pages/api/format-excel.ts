import { NextApiRequest, NextApiResponse } from "next";
import formidable from "formidable";
import fs from "fs";
import * as XLSX from "xlsx";
import path from "path";
import { promisify } from "util";
import { exec } from "child_process";

const PYTHON_PATH = "/Users/eunjilee/.venvs/dodo_env/bin/python3";

export const config = {
  api: {
    bodyParser: false,
  },
};

// 셀 크기 자동 조절 함수
const autoFitColumns = (worksheet: XLSX.WorkSheet, data: any[][]) => {
  const columnWidths = data[0].map((_, colIndex) => {
    return Math.max(
      ...data.map((row) => (row[colIndex] ? row[colIndex].toString().length : 10))
    );
  });

  worksheet["!cols"] = columnWidths.map((width) => ({ wch: width + 2 }));
};

// 셀 스타일 설정 함수
const setCellStyles = (worksheet: XLSX.WorkSheet) => {
  const columnsToFormat = [
    'C', 'D', 'F', 'H', 'J', 'L', 'N', 'P', 'R', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 
    , 'AB', 'AC', 'AD', 'AE', 'AF', 'AG'
  ];

  const maxRow = worksheet['!ref'] ? parseInt(worksheet['!ref'].split(':')[1].match(/\d+/)?.[0] || "2") : 2;

  columnsToFormat.forEach((col) => {
    for (let row = 2; row <= maxRow; row++) {
      const cellAddress = `${col}${row}`;
      if (worksheet[cellAddress]) {
        worksheet[cellAddress].z = '#,##0'; // 숫자 형식: 3자리마다 콤마
      }
    }
  });
};

const parseForm = async (req: NextApiRequest): Promise<{ files: formidable.Files }> => {
  return new Promise((resolve, reject) => {
    const form = formidable({ multiples: false, uploadDir: "/tmp", keepExtensions: true });

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
    const { files } = await parseForm(req);
    if (!files.file) {
      return res.status(400).json({ error: "File not found or invalid file" });
    }

    const file = Array.isArray(files.file) ? files.file[0] : (files.file as formidable.File);
    if (!file.filepath) {
      return res.status(400).json({ error: "Invalid file path" });
    }

    console.log("📌 Processing encrypted file:", file.filepath);

    const decryptedFilePath = file.filepath.replace(".xlsx", "_decrypted.xlsx");

    const command = `${PYTHON_PATH} decrypt_excel.py "${file.filepath}" "${decryptedFilePath}"`;
    try {
      const { stdout, stderr } = await promisify(exec)(command);
      console.log("📌 Python Output:", stdout);
      console.error("📌 Python Error:", stderr);

      if (!fs.existsSync(decryptedFilePath)) {
        throw new Error(`❌ Decrypted file not found: ${decryptedFilePath}`);
      }
      
      console.log("✅ Decrypted file created:", decryptedFilePath);
    } catch (error) {
      console.error("❌ Failed to decrypt file:", error);
      return res.status(500).json({ error: "Failed to decrypt file. Check password or encryption type." });
    }

    // 암호 해제된 파일을 읽기
    const workbook = XLSX.readFile(decryptedFilePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    console.log("✅ Excel data loaded:", jsonData.length);

    // 새로운 시트 데이터 생성
    const newSheetData: any[][] = [];

    // 고정된 커스텀 헤더 설정
    newSheetData[0] = [
      "성함", "총 오더수", "기본 금액", "지역단가", "배달거리 할증", "", "픽업지 할증", "", "도착지 할증", "", 
      "기상 할증", "", "기타 프로모션 1", "", "기타 프로모션 2", "", "기타 프로모션 3", "", "기타 프로모션 4", "", 
      "정산금액", "총 지원금", "차감내역", "합계금액", "지사 프로모션", "야간 프로모션", "과세 항목", "비과세 항목", "", "", "", "원천세", "정산예정 금액"
    ];
    
    newSheetData[1] = [
      "", "", "", "", "오더 수", "금액", "오더 수", "금액", "오더 수", "금액", 
      "오더 수", "금액", "오더 수", "금액", "오더 수", "금액", "오더 수", "금액", "오더 수", "금액", 
      "", "", "", "", "", "", "총 전산금액", "고용보험", "산재보험", "시간제보험", "보험료소급", "", ""
    ];

    console.log("✅ 헤더 추가됨:", JSON.stringify(newSheetData[0]), JSON.stringify(newSheetData[1]));

    // 본문 데이터 추가 (14행부터)
    jsonData.slice(13).forEach((row, index) => {
      if (!newSheetData[index + 2]) newSheetData[index + 2] = []; // 2줄이 이미 있으므로 index + 2

      newSheetData[index + 2][0] = row[2];  // C열
      newSheetData[index + 2][1] = row[5];  // F열
      newSheetData[index + 2][2] = row[6];  // G열
      newSheetData[index + 2][3] = row[7];  // H열
      newSheetData[index + 2][4] = row[8];  // I열
      newSheetData[index + 2][5] = row[9];  // J열
      newSheetData[index + 2][6] = row[10]; // K열
      newSheetData[index + 2][7] = row[11]; // L열
      newSheetData[index + 2][8] = row[12]; // M열
      newSheetData[index + 2][9] = row[13]; // N열
      newSheetData[index + 2][10] = row[14];// O열
      newSheetData[index + 2][11] = row[15];// P열
      newSheetData[index + 2][12] = row[16];// Q열
      newSheetData[index + 2][13] = row[17];// R열
      newSheetData[index + 2][14] = row[18];// S열
      newSheetData[index + 2][15] = row[19];// T열
      newSheetData[index + 2][16] = row[20];// U열
      newSheetData[index + 2][17] = row[21];// V열
      newSheetData[index + 2][18] = row[22];// W열
      newSheetData[index + 2][19] = row[23];// X열
      newSheetData[index + 2][20] = row[24];// Y열
      newSheetData[index + 2][21] = row[25]; //Z열
      newSheetData[index + 2][22] = row[26]; //AA열
      newSheetData[index + 2][23] = row[27]; //AB열
      newSheetData[index + 2][27] = row[29]; //AD열 / AB열 고용보험
      newSheetData[index + 2][28] = row[31]; //AD열 / AC열 산재보험
      newSheetData[index + 2][29] = row[32]; //AD열 / AC열 산재보험

      // Y열 지사 프로모션 수식 추가
      newSheetData[index + 2][24] = {
        f: `X${index + 3}+IF(B${index + 3}>=180,(B${index + 3}-180)*1000,0)-X${index + 3}`
      };

      // AA열 과세항목 수식 추가
      newSheetData[index + 2][26] = {
        f: `X${index + 3}+Y${index + 3}+Z${index + 3}`
      };

      // AF열 원천세 수식 추가
      newSheetData[index + 2][31] = {
        f: `((AA${index + 3}+AB${index + 3}+AC${index + 3}+AD${index + 3}+AE${index + 3})*0.967)-(AA${index + 3}+AB${index + 3}+AC${index + 3}+AD${index + 3}+AE${index + 3})`
      };

      // AG열 수식 추가
      newSheetData[index + 2][32] = {
        f: `AA${index + 3}+AB${index + 3}+AC${index + 3}+AD${index + 3}+AE${index + 3}+AF${index + 3}`
      };
    });

    console.log("✅ 변환된 데이터:", JSON.stringify(newSheetData));

    const newWorkbook = XLSX.utils.book_new();
    const newSheet = XLSX.utils.aoa_to_sheet(newSheetData);
    setCellStyles(newSheet); // 스타일 적용
    // 🔥 병합 정보 추가
    newSheet["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 1, c: 0 } }, // 성함 (세로 병합)
      { s: { r: 0, c: 1 }, e: { r: 1, c: 1 } }, // 총 오더수 (세로 병합)
      { s: { r: 0, c: 2 }, e: { r: 1, c: 2 } }, // 기본 금액 (세로 병합)
      { s: { r: 0, c: 3 }, e: { r: 1, c: 3 } }, // 지역단가 (세로 병합)
      { s: { r: 0, c: 4 }, e: { r: 0, c: 5 } }, // 배달거리 할증 (가로 병합)
      { s: { r: 0, c: 6 }, e: { r: 0, c: 7 } }, // 픽업지 할증 (가로 병합)
      { s: { r: 0, c: 8 }, e: { r: 0, c: 9 } }, // 도착지 할증 (가로 병합)
      { s: { r: 0, c: 10 }, e: { r: 0, c: 11 } }, // 기상 할증 (가로 병합)
      { s: { r: 0, c: 12 }, e: { r: 0, c: 13 } }, // 기타 프로모션 1 (가로 병합)
      { s: { r: 0, c: 14 }, e: { r: 0, c: 15 } }, // 기타 프로모션 2 (가로 병합)
      { s: { r: 0, c: 16 }, e: { r: 0, c: 17 } }, // 기타 프로모션 3 (가로 병합)
      { s: { r: 0, c: 18 }, e: { r: 0, c: 19 } }, // 기타 프로모션 4 (가로 병합)
      { s: { r: 0, c: 20 }, e: { r: 1, c: 20 } }, // 정산금액 (세로 병합)
      { s: { r: 0, c: 21 }, e: { r: 1, c: 21 } }, // 총 지원금 (세로 병합)
      { s: { r: 0, c: 22 }, e: { r: 1, c: 22 } }, // 차감내역 (세로 병합)
      { s: { r: 0, c: 23 }, e: { r: 1, c: 23 } }, // 합계금액 (세로 병합)
      { s: { r: 0, c: 24 }, e: { r: 1, c: 24 } }, // 지사 프로모션 (세로 병합)
      { s: { r: 0, c: 25 }, e: { r: 1, c: 25 } }, // 야간 프로모션 (세로 병합)
      { s: { r: 0, c: 26 }, e: { r: 1, c: 26 } }, // 과세 항목 / 총 전산금액 (세로 병합)
      { s: { r: 0, c: 27 }, e: { r: 0, c: 30 } }, // 비과세 항목 (가로 병합)
      { s: { r: 0, c: 31 }, e: { r: 1, c: 31 } }, // 원천세 (세로 병합)
      { s: { r: 0, c: 32 }, e: { r: 1, c: 32 } }, // 정산예정 금액 (세로 병합)
    ];
    // 열 너비 자동 조절 적용
    autoFitColumns(newSheet, newSheetData);

    // 엑셀 파일 저장
    XLSX.utils.book_append_sheet(newWorkbook, newSheet, "Formatted");

    const formattedFilePath = decryptedFilePath.replace(".xlsx", "_formatted.xlsx");
    XLSX.writeFile(newWorkbook, formattedFilePath);

    console.log("✅ Formatted file saved:", formattedFilePath);

    res.setHeader("Content-Disposition", "attachment; filename=formatted.xlsx");
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

    const fileBuffer = fs.readFileSync(formattedFilePath);
    res.status(200).send(fileBuffer);
  } catch (error) {
    console.error("❌ Processing error:", error);
    res.status(500).json({ error: "Failed to process the file" });
  }
};