import React, { useState } from "react";

function App() {
  const [message, setMessage] = useState("");
  const [status, setStatus] = useState("idle"); // idle, uploaded, converted, processing

  // Функция загрузки (теперь вызывается автоматически)
  const autoUpload = async (selectedFile) => {
    if (!selectedFile) return;

    setStatus("processing");
    setMessage("Uploading file...");

    const formData = new FormData();
    formData.append("file", selectedFile);

    try {
      const res = await fetch("/api/upload", {
        method: "POST",
        body: formData,
      });

      const data = await res.json();

      if (res.ok) {
        setStatus("uploaded");
        setMessage("File uploaded! Now you can convert it.");
      } else {
        setStatus("idle");
        setMessage(data.error || "Upload failed");
      }
    } catch (error) {
      setStatus("idle");
      setMessage("Server connection error");
    }
  };

  // Обработчик изменения инпута
  const onFileChange = (e) => {
    const selectedFile = e.target.files[0];
    if (selectedFile) {
      autoUpload(selectedFile); // Запускаем загрузку сразу
    }
  };

  const handleConvert = async () => {
    setStatus("processing");
    setMessage("Converting PDF to Excel...");

    try {
      const res = await fetch("/api/convert", {
        method: "POST",
      });
      const data = await res.json();

      if (res.ok) {
        setStatus("converted");
        setMessage("Done! Click below to download.");
      } else {
        setStatus("uploaded");
        setMessage(data.error || "Conversion failed");
      }
    } catch (error) {
      setStatus("uploaded");
      setMessage("Error during conversion");
    }
  };

  const handleDownload = async () => {
    try {
      const res = await fetch("/api/download");
      if (res.ok) {
        const blob = await res.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "result.xlsx";
        document.body.appendChild(a);
        a.click();
        a.remove();
        window.URL.revokeObjectURL(url);
      }
    } catch (error) {
      setMessage("Download error");
    }
  };

  return (
    <div style={styles.container}>
      <div style={styles.card}>
        <h2>PDF to Excel</h2>
        
        {/* Инпут всегда доступен для выбора нового файла */}
        <label style={styles.fileLabel}>
          {status === "processing" ? "Processing..." : "Select PDF file"}
          <input 
            type="file" 
            accept="application/pdf" 
            onChange={onFileChange} 
            style={styles.hiddenInput}
            disabled={status === "processing"}
          />
        </label>

        <div style={styles.buttonGroup}>
          {/* Кнопка Convert появляется сама после авто-загрузки */}
          {status === "uploaded" && (
            <button onClick={handleConvert} style={{...styles.button, backgroundColor: "#28a745"}}>
              Convert
            </button>
          )}

          {/* Кнопка Download появляется после конвертации */}
          {status === "converted" && (
            <button onClick={handleDownload} style={{...styles.button, backgroundColor: "#007bff"}}>
              Download
            </button>
          )}
        </div>

        {message && <p style={styles.message}>{message}</p>}
      </div>
    </div>
  );
}

const styles = {
  container: { display: "flex", justifyContent: "center", alignItems: "center", height: "100vh", backgroundColor: "#f0f2f5", fontFamily: "sans-serif" },
  card: { padding: "40px", backgroundColor: "white", borderRadius: "12px", boxShadow: "0 8px 24px rgba(0,0,0,0.1)", textAlign: "center", width: "320px" },
  fileLabel: {
    display: "block",
    padding: "12px",
    backgroundColor: "#eee",
    borderRadius: "6px",
    cursor: "pointer",
    marginBottom: "20px",
    border: "2px dashed #ccc"
  },
  hiddenInput: { display: "none" },
  buttonGroup: { minHeight: "50px", display: "flex", justifyContent: "center", alignItems: "center" },
  button: { padding: "12px 24px", fontSize: "16px", cursor: "pointer", border: "none", borderRadius: "6px", color: "white", fontWeight: "bold" },
  message: { marginTop: "20px", fontSize: "14px", color: "#555" }
};

export default App;