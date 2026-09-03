// Same pattern as lookup-character.js, but for the word as a whole (pinyin,
// meaning, Sino-Vietnamese) rather than a single character's radicals.
//
// SECURITY: requires a valid Supabase login token, same reasoning as
// lookup-character.js -- without this, the URL itself is a free-for-all
// that anyone can hit directly, not just visitors using the site's UI.

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

  let word;
  try {
    const body = await req.json();
    word = body.word;
  } catch (e) {
    return new Response(JSON.stringify({ error: "Invalid request body" }), { status: 400 });
  }

  if (!word || typeof word !== "string") {
    return new Response(JSON.stringify({ error: "Missing 'word' field" }), { status: 400 });
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
        max_tokens: 350,
        system:
          'You are a Chinese dictionary lookup tool. Given a multi-character Chinese word or phrase, respond with ONLY a raw JSON object (no markdown, no code fences, no extra text) in exactly this shape: {"pinyin": "...", "meaning": "...", "meaning_vi": "...", "sino_vietnamese": "..."}. "pinyin" is the Hanyu Pinyin for the whole word with tone marks, one syllable per character separated by a space (e.g. "nǐ hǎo"). "meaning" is a short English gloss for the word as a whole, a few words. "meaning_vi" is a short Vietnamese-language TRANSLATION of that meaning (e.g. "xin chào" for 你好) -- this is DIFFERENT from "sino_vietnamese": meaning_vi is what the word MEANS in Vietnamese (an ordinary Vietnamese phrase a reader would use), while sino_vietnamese is how the word is PRONOUNCED using the Sino-Vietnamese reading system (e.g. "nễ hảo" for 你好) -- do not confuse these, they are usually completely different words. "sino_vietnamese" is the standard Sino-Vietnamese (Hán Việt) reading of the whole word, lowercase Vietnamese with correct diacritics, one word per character separated by a space. If given something unrecognized, respond with {"pinyin": "", "meaning": "", "meaning_vi": "", "sino_vietnamese": ""}.',
        messages: [{ role: "user", content: word }],
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
