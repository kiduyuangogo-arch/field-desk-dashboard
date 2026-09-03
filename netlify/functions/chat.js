// Runs on Netlify's servers, never in the browser, so the API key stays private.
export const handler = async (event) => {
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, body: JSON.stringify({ error: "Method not allowed" }) };
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return {
      statusCode: 500,
      body: JSON.stringify({
        error: "No ANTHROPIC_API_KEY is set for this site yet. Add it under Project configuration > Environment variables in Netlify, then redeploy.",
      }),
    };
  }

  let question, profile;
  try {
    ({ question, profile } = JSON.parse(event.body || "{}"));
  } catch {
    return { statusCode: 400, body: JSON.stringify({ error: "Invalid request body" }) };
  }

  if (!question || !profile) {
    return { statusCode: 400, body: JSON.stringify({ error: "Missing question or data profile" }) };
  }

  try {
    const response = await fetch("https://api.anthropic.com/v1/messages", {
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
          "You are a data analyst helping someone understand a survey dataset. " +
          "You only see a text summary of the data (row/column counts, field types, top category values), never raw rows. " +
          "Answer clearly and concisely, in plain language, referencing the specific fields and numbers given.",
        messages: [
          {
            role: "user",
            content: `Dataset summary:\n${profile}\n\nQuestion: ${question}`,
          },
        ],
      }),
    });

    if (!response.ok) {
      const errText = await response.text();
      return { statusCode: response.status, body: JSON.stringify({ error: `Anthropic API error: ${errText}` }) };
    }

    const data = await response.json();
    const answer = data.content?.find((block) => block.type === "text")?.text || "No response text returned.";

    return {
      statusCode: 200,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ answer }),
    };
  } catch (err) {
    return { statusCode: 500, body: JSON.stringify({ error: err.message }) };
  }
};
