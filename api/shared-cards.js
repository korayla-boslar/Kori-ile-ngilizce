import { list, put } from "@vercel/blob";

const allowedSections = new Set(["words", "definitions"]);

function sendJson(response, statusCode, payload) {
  response.status(statusCode).json(payload);
}

function validatePayload(payload) {
  if (!payload || !Array.isArray(payload.cards) || !payload.cards.length) {
    return "Kart listesi boş olamaz.";
  }

  const hasInvalidCard = payload.cards.some((card) => {
    return !card || typeof card.front !== "string" || typeof card.back !== "string" || !card.front.trim() || !card.back.trim();
  });

  if (hasInvalidCard) {
    return "Kartlarda ön ve arka yüz zorunludur.";
  }

  return "";
}

export default async function handler(request, response) {
  const section = String(request.query.section || "");

  if (!allowedSections.has(section)) {
    sendJson(response, 400, { error: "Geçersiz bölüm." });
    return;
  }

  if (!process.env.BLOB_READ_WRITE_TOKEN) {
    sendJson(response, 503, { error: "Vercel Blob token tanımlı değil." });
    return;
  }

  const pathname = `shared-cards/${section}.json`;

  if (request.method === "GET") {
    const blobs = await list({ prefix: pathname, limit: 1 });
    const blob = blobs.blobs.find((item) => item.pathname === pathname) ?? blobs.blobs[0];

    if (!blob) {
      sendJson(response, 404, { error: "Paylaşılan kart bulunamadı." });
      return;
    }

    const blobResponse = await fetch(blob.url, { cache: "no-store" });
    const payload = await blobResponse.json();
    sendJson(response, 200, payload);
    return;
  }

  if (request.method === "POST") {
    const adminPin = process.env.SHARED_CARDS_ADMIN_PIN;

    if (adminPin && request.headers["x-admin-pin"] !== adminPin) {
      sendJson(response, 401, { error: "Yayınlama yetkisi yok." });
      return;
    }

    const error = validatePayload(request.body);
    if (error) {
      sendJson(response, 400, { error });
      return;
    }

    const payload = {
      cards: request.body.cards,
      fileName: String(request.body.fileName || "Paylaşılan Excel"),
      savedAt: new Date().toISOString(),
    };

    await put(pathname, JSON.stringify(payload), {
      access: "public",
      allowOverwrite: true,
      contentType: "application/json; charset=utf-8",
    });

    sendJson(response, 200, payload);
    return;
  }

  response.setHeader("Allow", "GET, POST");
  sendJson(response, 405, { error: "Bu yöntem desteklenmiyor." });
}
