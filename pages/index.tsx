import { useState } from "react";

export default function Home() {
  const [file, setFile] = useState<File | null>(null);
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);

  const handleUpload = async () => {
    if (!file) return alert("파일을 선택하세요!");

    const formData = new FormData();
    formData.append("file", file);

    const res = await fetch("http://localhost:3000/api/format-excel", {
      method: "POST",
      body: formData,
    });
  

    if (!res.ok) {
      alert("엑셀 변환 실패");
      return;
    }

    const blob = await res.blob();
    const url = window.URL.createObjectURL(blob);
    setDownloadUrl(url);
  };

  return (
    <div>
      <h1>엑셀 업로드</h1>
      <input type="file" accept=".xlsx" onChange={(e) => setFile(e.target.files[0])} />
      <button onClick={handleUpload}>업로드 및 변환</button>

      {downloadUrl && (
        <div>
          <a href={downloadUrl} download="formatted.xlsx">
            변환된 엑셀 다운로드
          </a>
        </div>
      )}
    </div>
  );
}