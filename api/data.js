// ─────────────────────────────────────────────────────────────
//  API Route — Atmosfera Data Proxy
//  ทำหน้าที่เป็น middleware ระหว่าง frontend และ Google Apps Script
//  เพื่อซ่อน credentials และเพิ่มความปลอดภัย
//
//  Environment Variables ที่ต้องตั้งค่า:
//    API_TOKEN          — Secret token สำหรับ authenticate กับ Google Script
//    GOOGLE_SCRIPT_URL  — URL ของ Google Apps Script web app
// ─────────────────────────────────────────────────────────────

const FETCH_TIMEOUT_MS = 15_000; // 15 วินาที

export default async function handler(req, res) {
  // ── 1. ตรวจสอบ Environment Variables ──────────────────────
  const token = process.env.API_TOKEN;
  const googleScriptUrl = process.env.GOOGLE_SCRIPT_URL;

  if (!googleScriptUrl) {
    console.error('[Atmosfera] Missing GOOGLE_SCRIPT_URL environment variable');
    return res.status(500).json({
      status: 'error',
      message: 'Server configuration error: missing GOOGLE_SCRIPT_URL',
    });
  }

  if (!token) {
    console.error('[Atmosfera] Missing API_TOKEN environment variable');
    return res.status(500).json({
      status: 'error',
      message: 'Server configuration error: missing API_TOKEN',
    });
  }

  // ── 2. Abort Controller สำหรับ Timeout ────────────────────
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), FETCH_TIMEOUT_MS);

  try {
    // ── 3. เรียก doGet ใน Google Apps Script ─────────────────
    //
    //  Security note:
    //  Token อยู่ใน URL query string เพราะ Apps Script's doGet() อ่านจาก
    //  e.parameter.token เท่านั้น (custom headers ถูก Google strip ออก)
    //
    //  สาเหตุที่ยังปลอดภัย:
    //  - Browser เรียก /api/data (ไม่มี token)
    //  - Server (Vercel/Node) เรียก Google Script (มี token) — server-to-server
    //  - Token ไม่เคยถูกเปิดเผยให้ผู้ใช้เห็น
    //
    //  ⚠️ ควรเก็บ token ใน Script Properties ด้วย (ไม่ hardcode ใน script)
    const url = new URL(googleScriptUrl);
    url.searchParams.set('token', token);

    const response = await fetch(url.toString(), {
      method: 'GET',
      signal: controller.signal,
    });

    // ── 4. ตรวจสอบ HTTP Status ก่อน parse JSON ──────────────
    if (!response.ok) {
      const errorText = await response.text().catch(() => '(no body)');
      console.error(`[Atmosfera] Google Script returned HTTP ${response.status}:`, errorText);
      return res.status(502).json({
        status: 'error',
        message: `Upstream error: HTTP ${response.status}`,
      });
    }

    // ── 5. ตรวจสอบ Content-Type ก่อน parse JSON ────────────
    const contentType = response.headers.get('content-type') ?? '';
    if (!contentType.includes('application/json')) {
      const body = await response.text().catch(() => '');
      console.error('[Atmosfera] Google Script returned non-JSON response:', body.slice(0, 200));
      return res.status(502).json({
        status: 'error',
        message: 'Upstream returned an unexpected response format',
      });
    }

    const data = await response.json();

    // ── 6. ส่งข้อมูลกลับ frontend + Cache ──────────────────
    //  Cache ที่ Vercel CDN 5 นาที → request ซ้ำในช่วงนี้ไม่ต้องเรียก
    //  Google Script อีก ลด latency และ cold start ได้มาก
    res.setHeader('Cache-Control', 's-maxage=300, stale-while-revalidate=600');
    return res.status(200).json(data);

  } catch (error) {
    if (error?.name === 'AbortError') {
      console.error('[Atmosfera] Request to Google Script timed out');
      return res.status(504).json({
        status: 'error',
        message: `Request timed out after ${FETCH_TIMEOUT_MS / 1000}s`,
      });
    }

    console.error('[Atmosfera] Unexpected fetch error:', error);
    return res.status(500).json({
      status: 'error',
      message: error.message ?? 'Unknown error',
    });

  } finally {
    clearTimeout(timeoutId);
  }
}