// Same pattern as lookup-character.js, but for the word as a whole (pinyin,
// meaning, Sino-Vietnamese) rather than a single character's radicals.

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
        max_tokens: 300,
        system:
          'You are a Chinese dictionary lookup tool. Given a multi-character Chinese word or phrase, respond with ONLY a raw JSON object (no markdown, no code fences, no extra text) in exactly this shape: {"pinyin": "...", "meaning": "...", "sino_vietnamese": "..."}. "pinyin" is the Hanyu Pinyin for the whole word with tone marks, one syllable per character separated by a space (e.g. "nǐ hǎo"). "meaning" is a short English gloss for the word as a whole, a few words. "sino_vietnamese" is the standard Sino-Vietnamese (Hán Việt) reading of the whole word, lowercase Vietnamese with correct diacritics, one word per character separated by a space (e.g. "nễ hảo"). If given something unrecognized, respond with {"pinyin": "", "meaning": "", "sino_vietnamese": ""}.',
        messages: [{ role: "user", content: word }],
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
