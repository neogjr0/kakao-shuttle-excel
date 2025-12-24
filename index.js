import express from "express";
import bodyParser from "body-parser";
import fs from "fs";
import path from "path";
import { createRequire } from "module";

const require = createRequire(import.meta.url);
const XLSX = require("xlsx");

const app = express();
const FILE_NAME = "chat_log.xlsx";

app.use((req, res, next) => {
  res.setHeader("ngrok-skip-browser-warning", "true");
  next();
});

app.use(bodyParser.json());

// --- 데이터 추출 함수 ---
const parseMessage = (message) => {
  const school = message.match(/(?:학교|학교명)[:\s]*([^\n*-]+)/)?.[1]?.trim() || "";
  const name = message.match(/(?:이름|학생|학생\s*이름)[:\s]*([^\n*-]+)/)?.[1]?.trim() || "";
  const address = message.match(/(?:주소|장소|탑승|탑승\s*장소)[:\s]*([^\n*-]+)/)?.[1]?.trim() || "";
  const phone = message.match(/(?:연락처|전화|폰)[:\s]*([\d-]{10,14})/)?.[1]?.trim() || 
                message.match(/([\d-]{10,14})/)?.[1]?.trim() || "";

  return { school, name, address, phone };
};

const saveToExcel = (userId, message) => {
  const time = new Date().toLocaleString("ko-KR");
  const parsedData = parseMessage(message);

  const newRow = {
    "시간": time,
    "사용자ID": userId,
    "학교": parsedData.school,
    "학생이름": parsedData.name,
    "탑승장소": parsedData.address,
    "연락처": parsedData.phone,
    "비고(원본메세지)": message
  };

  let workbook;
  let data = [];

  if (fs.existsSync(FILE_NAME)) {
    try {
      workbook = XLSX.readFile(FILE_NAME);
      const sheetName = workbook.SheetNames[0];
      data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
    } catch (err) {
      workbook = XLSX.utils.book_new();
    }
  } else {
    workbook = XLSX.utils.book_new();
  }

  data.push(newRow);
  const newWorksheet = XLSX.utils.json_to_sheet(data);
  
  newWorksheet["!cols"] = [
    { wch: 22 }, { wch: 15 }, { wch: 12 }, { wch: 12 }, { wch: 35 }, { wch: 15 }, { wch: 60 }
  ];

  if (workbook.SheetNames.includes("채팅로그")) {
    workbook.Sheets["채팅로그"] = newWorksheet;
  } else {
    XLSX.utils.book_append_sheet(workbook, newWorksheet, "채팅로그");
  }

  XLSX.writeFile(workbook, FILE_NAME);
  console.log(`📊 엑셀 저장 완료: ${parsedData.name} (${parsedData.school})`);
};

app.get("/download", (req, res) => {
  const filePath = path.resolve(FILE_NAME);
  if (fs.existsSync(filePath)) {
    res.download(filePath, "셔틀고고_정리데이터.xlsx");
  } else {
    res.status(404).send("기록된 데이터가 없습니다.");
  }
});

// --- 핵심 수정: 보이지 않는 공백 응답 ---
app.post("/kakao-webhook", (req, res) => {
  const userId = req.body.userRequest?.user?.id || "비회원";
  const message = req.body.userRequest?.utterance || "";

  if (message && message !== "발화 내용") {
    try {
      saveToExcel(userId, message);
    } catch (error) {
      console.error("❌ 저장 에러:", error);
    }
  }

  // 0너비 공백(\u200B)을 보내 카톡의 기본 응답을 덮어씌웁니다.
  res.status(200).send({
    version: "2.0",
    template: {
      outputs: [
        {
          simpleText: {
            text: "\u200B"
          }
        }
      ]
    }
  });
});

const PORT = 3000;
app.listen(PORT, () => console.log(`🚀 서버 가동 중... 포트: ${PORT}`));