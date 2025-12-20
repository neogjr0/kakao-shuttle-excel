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

// --- 데이터 추출 함수 (정규표현식 보강) ---
const parseMessage = (message) => {
  const school = message.match(/-학교:\s*([^\n*]+)/)?.[1]?.trim() || "";
  const name = message.match(/-학생 이름:\s*([^\n*]+)/)?.[1]?.trim() || "";
  const address = message.match(/-주소 및 탑승 장소:\s*([^\n*]+)/)?.[1]?.trim() || "";
  const phone = message.match(/-연락처:\s*([^\n*]+)/)?.[1]?.trim() || "";

  return { school, name, address, phone };
};

const saveToExcel = (userId, message) => {
  const time = new Date().toLocaleString("ko-KR");
  const parsedData = parseMessage(message);

  // 엑셀 시트의 헤더(열 이름)와 매칭될 데이터 구조
  const newRow = {
    "시간": time,
    "사용자ID": userId,
    "학교": parsedData.school,
    "학생이름": parsedData.name,
    "탑승장소": parsedData.address,
    "연락처": parsedData.phone,
    "비고(원본메세지)": message  // 요청하신 대로 원본 메시지를 비고란에 삽입
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
  
  // 열 너비 자동 설정 (내용이 길어도 잘 보이게)
  newWorksheet["!cols"] = [
    { wch: 22 }, // 시간
    { wch: 15 }, // 사용자ID
    { wch: 12 }, // 학교
    { wch: 12 }, // 학생이름
    { wch: 35 }, // 탑승장소
    { wch: 15 }, // 연락처
    { wch: 60 }  // 비고(원본메세지) - 길게 설정
  ];

  if (workbook.SheetNames.includes("채팅로그")) {
    workbook.Sheets["채팅로그"] = newWorksheet;
  } else {
    XLSX.utils.book_append_sheet(workbook, newWorksheet, "채팅로그");
  }

  XLSX.writeFile(workbook, FILE_NAME);
  console.log(`📊 엑셀 저장 완료: ${parsedData.name} (${parsedData.school})`);
};

// 엑셀 다운로드 경로
app.get("/download", (req, res) => {
  const filePath = path.resolve(FILE_NAME);
  if (fs.existsSync(filePath)) {
    res.download(filePath, "셔틀고고_정리데이터.xlsx");
  } else {
    res.status(404).send("기록된 데이터가 없습니다.");
  }
});

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

  res.status(200).send({
    version: "2.0",
    template: {
      outputs: [{ simpleText: { text: "✅ 신청 정보가 표에 기록되었습니다." } }]
    }
  });
});

const PORT = 3000;
app.listen(PORT, () => console.log(`🚀 서버 가동 중... 포트: ${PORT}`));