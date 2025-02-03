import { useState } from "react";

export default function Home() {
  const [file, setFile] = useState<File | null>(null);
  const [data, setData] = useState<any[]>([]);

  const handleUpload = async () => {
    if (!file) return alert("파일을 선택하세요!");

    const formData = new FormData();
    formData.append("file", file);

    const res = await fetch("/api/upload", {
      method: "POST",
      body: formData,
    });

    const result = await res.json();
    setData(result.formattedData);
  };

  return (
    <div className="container">
      <h1>엑셀 업로드</h1>
      <input type="file" accept=".xlsx" onChange={(e) => setFile(e.target.files[0])} />
      <button onClick={handleUpload}>업로드</button>

      <h2>결과</h2>
      <ul>
        {data.map((item, index) => (
          <li key={index}>{item.name} - {item.phone} - {item.amount}</li>
        ))}
      </ul>
    </div>
  );
}