// src/taskpane/firebaseClient.ts
export async function requestTrialToken(email: string) {
  const res = await fetch("/.netlify/functions/requestTrial", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ email }),
  });
  if (!res.ok) throw new Error("requestTrial failed");
  return res.json();
}

/**
 * validateTrialToken(token, consume=true)
 * - if consume=true, function will decrement uses on server.
 * - if consume=false, it only returns validity and current uses.
 */
export async function validateTrialToken(token: string, consume = true) {
  const res = await fetch("/.netlify/functions/validateTrial", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ token, consume }),
  });
  if (!res.ok) throw new Error("validateTrial failed");
  return res.json();
}
