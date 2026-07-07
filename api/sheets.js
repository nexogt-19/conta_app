// api/sheets.js — Vercel Serverless Function
// Actúa como puente entre la app y Google Sheets (sin CORS)

const GAS_URL = "https://script.google.com/macros/s/AKfycbyQlGGElFc20CwdqgB4978ZfBF53UbxQTcIl2n3A2Gafe8pYSQ9fJWJETYIdO7FRVIi/exec";

export default async function handler(req, res) {
  // Permitir CORS desde cualquier origen
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  try {
    if (req.method === "GET") {
      const tabla = req.query.tabla || "partidas";
      const r = await fetch(`${GAS_URL}?tabla=${tabla}`);
      const data = await r.json();
      return res.status(200).json(data);
    }

    if (req.method === "POST") {
      const body = typeof req.body === "string" 
        ? req.body 
        : JSON.stringify(req.body);
      const r = await fetch(GAS_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain" },
        body
      });
      const data = await r.json();
      return res.status(200).json(data);
    }
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}
