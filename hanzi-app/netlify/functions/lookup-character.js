// Runs on Netlify's servers, never in the visitor's browser -- this is what
// lets ANTHROPIC_API_KEY stay secret. The frontend calls this function at
// /.netlify/functions/lookup-character instead of hitting Anthropic directly.
//
// SECURITY: every request must include a valid Supabase login token in the
// Authorization header. Without this check, anyone (not just visitors using
// the site's UI) could call this URL directly and rack up API charges on
// your account, since the button being hidden in the app doesn't stop a
// request sent straight to the function.

import { createClient } from "@supabase/supabase-js";

export default async (req) => {
  if (req.method !== "POST") {
    return new Response(JSON.stringify({ error: "Method not allowed" }), { status: 405 });
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return new Response(
      JSON.stringify({ error: "Server is missing ANTHROPIC_API_KEY. Set it in Netlify's environment variables." }),
      { status: 500 }
    );
  }

  // --- Require a real logged-in Supabase user ---
  const supabaseUrl = process.env.VITE_SUPABASE_URL;
  const supabaseAnonKey = process.env.VITE_SUPABASE_ANON_KEY;
  const authHeader = req.headers.get("authorization") || "";
  const token = authHeader.replace(/^Bearer\s+/i, "").trim();

  if (!token) {
    return new Response(
      JSON.stringify({ error: "AUTH_REQUIRED", message: "Please log in to use lookup." }),
      { status: 401 }
    );
  }

  const authClient = createClient(supabaseUrl, supabaseAnonKey);
  const { data: userData, error: authError } = await authClient.auth.getUser(token);
  if (authError || !userData || !userData.user) {
    return new Response(
      JSON.stringify({ error: "AUTH_REQUIRED", message: "Your session is invalid or expired. Please log in again." }),
      { status: 401 }
    );
  }
  // --- end auth check ---

  // --- Check and consume one lookup from this user's quota, atomically ---
  // Uses a client scoped to THIS user's own token (not the anon key alone)
  // so the increment_lookup_count() function's auth.uid() resolves
  // correctly and only ever touches their own row.
  const userClient = createClient(supabaseUrl, supabaseAnonKey, {
    global: { headers: { Authorization: `Bearer ${token}` } },
  });
  const { data: quotaRows, error: quotaError } = await userClient.rpc("increment_lookup_count");
  if (quotaError) {
    console.error("Quota check failed:", quotaError);
    return new Response(JSON.stringify({ error: "Could not verify lookup quota" }), { status: 500 });
  }
  const quota = Array.isArray(quotaRows) ? quotaRows[0] : quotaRows;
  if (!quota || !quota.allowed) {
    const reason = quota && quota.reason ? quota.reason : "LIMIT_REACHED";
    return new Response(
      JSON.stringify({
        error: reason,
        message: reason === "DISABLED" ? "This account has been disabled." : "You've used all your free lookups.",
        lookup_count: quota ? quota.new_count : 0,
        lookup_limit: quota ? quota.limit_value : 0,
      }),
      { status: 403 }
    );
  }
  // --- end quota check ---

  let char;
  try {
    const body = await req.json();
    char = body.char;
  } catch (e) {
    return new Response(JSON.stringify({ error: "Invalid request body" }), { status: 400 });
  }

  if (!char || typeof char !== "string") {
    return new Response(JSON.stringify({ error: "Missing 'char' field" }), { status: 400 });
  }

  try {
    const anthropicRes = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: "claude-sonnet-4-6",
        max_tokens: 600,
        system:
          'You are a Chinese dictionary lookup tool. Given a single Chinese character, respond with ONLY a raw JSON object (no markdown, no code fences, no extra text) in exactly this shape: {"pinyin": "...", "meaning": "...", "meaning_vi": "...", "sino_vietnamese": "...", "components": [{"char": "...", "pinyin": "...", "meaning": "...", "meaning_vi": "...", "sino_vietnamese": "..."}]}. "pinyin" is Hanyu Pinyin with tone marks (e.g. hǎo). "meaning" is a short English gloss, a few words. "meaning_vi" is a short Vietnamese-language TRANSLATION of that meaning (e.g. "tốt" for 好) -- this is DIFFERENT from "sino_vietnamese": meaning_vi is what the character MEANS in Vietnamese (an ordinary Vietnamese word/phrase a reader would use), while sino_vietnamese is how the character is PRONOUNCED using the Sino-Vietnamese reading system (e.g. "hảo" for 好) -- do not confuse these two, they are usually different words entirely. "sino_vietnamese" is the standard Sino-Vietnamese (Hán Việt) reading, lowercase Vietnamese with correct diacritics. "components" is the ordered list of bushou (radical/graphical components, e.g. 女+子 for 好) that combine to form the character, each with its own pinyin/meaning/meaning_vi/sino_vietnamese in the same style; if the character is itself a single indivisible radical, components can be an empty array. If given more than one character or something unrecognized, respond with {"pinyin": "", "meaning": "", "meaning_vi": "", "sino_vietnamese": "", "components": []}.',
        messages: [{ role: "user", content: char }],
      }),
    });

    if (!anthropicRes.ok) {
      const errText = await anthropicRes.text();
      console.error("Anthropic API error:", anthropicRes.status, errText);
      return new Response(JSON.stringify({ error: "Lookup failed upstream" }), { status: 502 });
    }

    const data = await anthropicRes.json();
    data.lookup_count = quota.new_count;
    data.lookup_limit = quota.limit_value;
    return new Response(JSON.stringify(data), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    });
  } catch (err) {
    console.error("Function error:", err);
    return new Response(JSON.stringify({ error: "Lookup failed" }), { status: 500 });
  }
};
