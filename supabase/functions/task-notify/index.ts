// task-notify -- sends the PVH task emails through Resend. See supabase/009_notifications.sql.
//
// Called only by the database (pg_net), never by the app or a phone. It answers
// anyone who does not send the shared secret with 401, so deploy it with
// "Verify JWT" turned OFF (the database sends the secret, not a login token).
//
//   { "kind": "assigned", "task_id": "<uuid>" }   one email, to the new assignee
//   { "kind": "digest", "slot": "morning" | "afternoon" }
//                                                  one email to each officer with
//                                                  a task overdue or due today
//   { "kind": "test", "to": "someone@example.com" } one test email, to check Resend
//
// Secrets (Edge Functions -> Secrets):
//   RESEND_API_KEY   from Resend (Sending access, limited to the domain)
//   NOTIFY_SECRET    the same value as notify_config 'secret' in the database
//   APP_URL          optional, the app's address (default below)
//   FROM_EMAIL       optional, e.g. PVH Tasks <tasks@pvhnotify.fyi>
// SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY are supplied by Supabase itself.
//
// An email carries counts, a due date, a priority and a link. Never a plate, a
// person's name other than who assigned it, or a task title.

const SB = Deno.env.get("SUPABASE_URL");
const KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");
const RESEND = Deno.env.get("RESEND_API_KEY");
const SECRET = Deno.env.get("NOTIFY_SECRET");
const FROM = Deno.env.get("FROM_EMAIL") || "PVH Tasks <tasks@pvhnotify.fyi>";
const APP = Deno.env.get("APP_URL") || "https://asheso80.github.io/pvh_search/";
const TZ = "America/Halifax";

function json(obj, status) {
  return new Response(JSON.stringify(obj), { status: status || 200, headers: { "Content-Type": "application/json" } });
}

async function rest(path) {
  const r = await fetch(SB + "/rest/v1/" + path, { headers: { apikey: KEY, Authorization: "Bearer " + KEY } });
  if (!r.ok) throw new Error("Database " + r.status + ": " + (await r.text()));
  return r.json();
}

// Today's date in Halifax, as YYYY-MM-DD (the en-CA format).
function todayHalifax() {
  return new Intl.DateTimeFormat("en-CA", { timeZone: TZ, year: "numeric", month: "2-digit", day: "2-digit" }).format(new Date());
}

// "2026-11-05" -> "Nov 5". Noon UTC so no timezone can move the day.
function niceDate(ds) {
  return new Date(ds + "T12:00:00Z").toLocaleDateString("en-CA", { month: "short", day: "numeric", timeZone: "UTC" });
}

function esc(s) {
  return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function wait(ms) {
  return new Promise((res) => setTimeout(res, ms));
}

async function send(to, subject, lines) {
  const link = APP + "#/tasks";
  const text = lines.join("\n\n") + "\n\nOpen your tasks: " + link + "\n";
  const html = '<div style="font-family:Arial,Helvetica,sans-serif;font-size:16px;line-height:1.5;color:#1a1a1a">' +
    lines.map((l) => "<p>" + esc(l) + "</p>").join("") +
    '<p><a href="' + esc(link) + '" style="display:inline-block;padding:10px 18px;background:#0b5fd1;color:#fff;' +
    'text-decoration:none;border-radius:8px;font-weight:bold">Open your tasks</a></p>' +
    '<p style="font-size:12px;color:#777">Sent by the PVH app. Details stay in the app.</p></div>';
  const r = await fetch("https://api.resend.com/emails", {
    method: "POST",
    headers: { Authorization: "Bearer " + RESEND, "Content-Type": "application/json" },
    body: JSON.stringify({ from: FROM, to: [to], subject: subject, text: text, html: html }),
  });
  const body = await r.text();
  if (!r.ok) console.error("Resend refused " + to + ": " + r.status + " " + body);
  return r.ok;
}

async function assigned(taskId) {
  if (!/^[0-9a-f-]{36}$/i.test(String(taskId || ""))) return { sent: 0, why: "bad task id" };
  const tasks = await rest("tasks?id=eq." + taskId + "&select=status,due_date,priority,assignee_id,author_name");
  const t = tasks[0];
  if (!t || t.status !== "open") return { sent: 0, why: "not an open task" };
  const people = await rest("profiles?id=eq." + t.assignee_id + "&select=notify_email");
  const to = people[0] && people[0].notify_email;
  if (!to) return { sent: 0, why: "no notify_email for the assignee" };
  const who = t.author_name || "Someone";
  const subject = who + " assigned you a task, due " + niceDate(t.due_date) + (t.priority === "high" ? " (High priority)" : "");
  const ok = await send(to, subject, [
    who + " assigned you a task in the PVH app.",
    "Due " + niceDate(t.due_date) + (t.priority === "high" ? ". Priority: High." : "."),
  ]);
  return { sent: ok ? 1 : 0 };
}

async function digest(slot) {
  const today = todayHalifax();
  const rows = await rest("tasks?status=eq.open&due_date=lte." + today + "&select=assignee_id,due_date");
  const by = {};
  rows.forEach((t) => {
    const c = by[t.assignee_id] || (by[t.assignee_id] = { overdue: 0, today: 0 });
    if (t.due_date < today) c.overdue++; else c.today++;
  });
  const ids = Object.keys(by);
  if (!ids.length) return { sent: 0, why: "nobody has anything due" };
  const people = await rest("profiles?id=in.(" + ids.join(",") + ")&select=id,notify_email");
  let sent = 0, skipped = 0;
  for (const p of people) {
    if (!p.notify_email) { skipped++; continue; }
    const c = by[p.id];
    const parts = [];
    if (c.overdue) parts.push(c.overdue + " overdue");
    if (c.today) parts.push(c.today + " due today");
    const subject = (slot === "afternoon" ? "PVH tasks still open: " : "PVH tasks: ") + parts.join(", ");
    const ok = await send(p.notify_email, subject, [
      slot === "afternoon"
        ? "These are still open before the end of the shift: " + parts.join(", ") + "."
        : "Your tasks today: " + parts.join(", ") + ".",
    ]);
    if (ok) sent++;
    await wait(600); // Resend allows two requests a second
  }
  return { sent: sent, skipped_no_email: skipped };
}

Deno.serve(async (req) => {
  if (req.method !== "POST") return json({ error: "POST only" }, 405);
  if (!SECRET || req.headers.get("x-notify-secret") !== SECRET) return json({ error: "Unauthorized" }, 401);
  if (!SB || !KEY || !RESEND) return json({ error: "Function secrets are not set" }, 500);
  let body;
  try { body = await req.json(); } catch { return json({ error: "Bad JSON" }, 400); }
  try {
    if (body.kind === "assigned") return json(await assigned(body.task_id));
    if (body.kind === "digest") return json(await digest(body.slot === "afternoon" ? "afternoon" : "morning"));
    if (body.kind === "test") {
      const ok = await send(String(body.to || ""), "PVH tasks: test email", ["This is a test from the PVH app. If you can read this, sending works."]);
      return json({ sent: ok ? 1 : 0 });
    }
    return json({ error: "Unknown kind" }, 400);
  } catch (e) {
    console.error(String(e));
    return json({ error: String(e && e.message || e) }, 500);
  }
});
