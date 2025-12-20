import express from "express";
import bodyParser from "body-parser";
import { GoogleSpreadsheet } from "google-spreadsheet";
import { JWT } from "google-auth-library";

const app = express();
app.use(bodyParser.json());

// --- 구글 시트 설정 (나중에 발급받을 키를 넣을 곳) ---
const SPREADSHEET_ID = '여기에_복사한_시트_ID를_넣으세요';
const GOOGLE_SERVICE_ACCOUNT_EMAIL = '발급받을_서비스_계정_이메일';
const GOOGLE_PRIVATE_KEY = '발급받을_개인_키'.replace(/\\n/g, '\n');

const serviceAccountAuth = new JWT({
  email: GOOGLE_SERVICE_ACCOUNT_EMAIL,
  key: GOOGLE_PRIVATE_KEY,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

const doc = new GoogleSpreadsheet(SPREADSHEET_ID, serviceAccountAuth);

const parseMessage = (message) => {
  const school = message.match(/-학교:\s*([^\n*]+)/)?.[1]?.trim() || "";
  const name = message.match(/-학생 이름:\s*([^\n*]+)/)?.[1]?.trim() || "";
  const address = message.match(/-주소 및 탑승 장소:\s*([^\n*]+)/)?.[1]?.trim() || "";
  const phone = message.match(/-연락처:\s*([^\n*]+)/)?.[1]?.trim() || "";
  return { school, name, address, phone };
};

app.post("/kakao-webhook", async (req, res) => {
  const userId = req.body.userRequest?.user?.id || "비회원";
  const message = req.body.userRequest?.utterance || "";

  if (message && message !== "발화 내용") {
    try {
      await doc.loadInfo();
      const sheet = doc.sheetsByIndex[0];
      const parsed = parseMessage(message);
      
      await sheet.addRow({
        "시간": new Date().toLocaleString("ko-KR"),
        "사용자ID": userId,
        "학교": parsed.school,
        "학생이름": parsed.name,
        "탑승장소": parsed.address,
        "연락처": parsed.phone,
        "비고": message
      });
      console.log("✅ 구글 시트 저장 완료!");
    } catch (e) {
      console.error("❌ 구글 저장 실패:", e);
    }
  }

  res.status(200).send({
    version: "2.0",
    template: { outputs: [{ simpleText: { text: "✅ 구글 시트에 실시간 기록되었습니다." } }] }
  });
});

app.listen(3000, () => console.log("🚀 구글 시트 연동 서버 시작!"));