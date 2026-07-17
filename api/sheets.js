const GAS_URL = "https://script.google.com/macros/s/AKfycbyQlGGElFc20CwdqgB4978ZfBF53UbxQTcIl2n3A2Gafe8pYSQ9fJWJETYIdO7FRVIi/exec";

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  
  if (req.method === 'OPTIONS') { res.status(200).end(); return; }

  try {
    const options = {
      method: req.method,
      headers: { 'Content-Type': 'application/json' },
    };
    if (req.method === 'POST') {
      options.body = JSON.stringify(req.body);
    }
    
    const url = req.method === 'GET' 
      ? `${GAS_URL}?${new URLSearchParams(req.query)}`
      : GAS_URL;
    
    const response = await fetch(url, options);
    const data = await response.json();
    res.status(200).json(data);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
}
