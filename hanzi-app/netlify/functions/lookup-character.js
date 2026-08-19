// Runs on Netlify's servers, never in the visitor's browser -- this is what
// lets ANTHROPIC_API_KEY stay secret. The frontend calls this function at
// /.netlify/functions/lookup-character instead of hitting Anthropic directly.

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
        max_tokens: 500,
        system:
          'You are a Chinese dictionary lookup tool. Given a single Chinese character, respond with ONLY a raw JSON object (no markdown, no code fences, no extra text) in exactly this shape: {"pinyin": "...", "meaning": "...", "sino_vietnamese": "...", "components": [{"char": "...", "pinyin": "...", "meaning": "...", "sino_vietnamese": "..."}]}. "pinyin" is Hanyu Pinyin with tone marks (e.g. hǎo). "meaning" is a short English gloss, a few words. "sino_vietnamese" is the standard Sino-Vietnamese (Hán Việt) reading, lowercase Vietnamese with correct diacritics (e.g. hảo). "components" is the ordered list of bushou (radical/graphical components, e.g. 女+子 for 好) that combine to form the character, each with its own pinyin/meaning/sino_vietnamese in the same style; if the character is itself a single indivisible radical, components can be an empty array. If given more than one character or something unrecognized, respond with {"pinyin": "", "meaning": "", "sino_vietnamese": "", "components": []}.',
        messages: [{ role: "user", content: char }],
      }),
    });

    if (!anthropicRes.ok) {
      const errText = await anthropicRes.text();
      console.error("Anthropic API error:", anthropicRes.status, errText);
      return new Response(JSON.stringify({ error: "Lookup failed upstream" }), { status: 502 });
    }

    const data = await anthropicRes.json();
    return new Response(JSON.stringify(data), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    });
  } catch (err) {
    console.error("Function error:", err);
    return new Response(JSON.stringify({ error: "Lookup failed" }), { status: 500 });
  }
};
