import { useState, useEffect, useRef } from "react";

const BOARD_ID  = "2120641399";
const API_BASE  = "https://mondaycom-middleware.onrender.com";

const TEMPLATES = {
  "AICPA Town Hall Newsletter": `<!doctype html><html><head><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>{{Subject}}</title><style type="text/css">@import url('https://fonts.googleapis.com/css?family=Roboto:300,400,500,700&display=swap');table{border-collapse:collapse}table td{border-collapse:collapse;font-weight:300!important}table td strong{font-weight:700!important}body{background-color:#fff;margin:0}.appleLinks a{color:inherit!important;text-decoration:underline}a{text-decoration:underline;color:#86387f}</style></head><body style="margin:0px;padding:0px;" yahoo="fix" bgcolor="#ffffff"><table bgcolor="#ffffff" border="0" cellpadding="0" cellspacing="0" style="table-layout:fixed;" width="100%"><tbody><tr><td align="center" bgcolor="#ffffff" style="margin:0 auto!important;min-width:680px;padding:0;width:680px;" valign="top" width="680"><div align="center"><table align="center" border="0" cellpadding="0" cellspacing="0" style="width:680px;max-width:680px;" width="680"><tbody><tr><td align="center" valign="top"><table cellpadding="0" cellspacing="0" style="max-width:680px;" width="680"><tbody><tr><td><table align="center" border="0" cellpadding="0" cellspacing="0" style="max-width:680px" width="680"><tbody><tr><td align="center" style="padding:10px 25px;font-family:Roboto,Arial,Helvetica,sans-serif;font-size:16px;color:#231f20;font-weight:300;text-align:center;" valign="top"><p style="font-family:'Roboto',Arial,Helvetica,sans-serif;color:#000000;font-size:16px;font-weight:300;"><strong style="font-weight:700;"><a href="https://www.aicpa-cima.com/cpe-learning/webcast/aicpa-town-hall-series" style="text-decoration:underline;color:#86387f;" target="_blank">{{PreheaderText}}</a></strong></p></td></tr></tbody></table></td></tr></tbody></table></td></tr><tr><td align="center" bgcolor="#ffffff" valign="top"><table cellpadding="0" cellspacing="0"><tbody><tr><td><table align="center" border="0" cellpadding="0" cellspacing="0" role="presentation"><tbody><tr><td align="center"><a href="https://www.cpa.com/town-hall" target="_blank"><img alt="AICPA Town Hall Series Newsletter" border="0" height="170" src="https://marketing.cpa.com/l/701003/2025-05-14/46x5rg/701003/17472392906lBfnhW5/AICPA_TH_Podcast_Banner.png" style="width:680px;height:170px;" width="680"></a></td></tr></tbody></table></td></tr></tbody></table></td></tr><tr><td align="center" bgcolor="#ffffff" valign="top"><table cellpadding="0" cellspacing="0" style="background-color:#FFFFFF;min-width:100%;border-bottom:solid 1px #ccc;" width="100%"><tbody><tr><td style="padding:0px;"><table align="center" bgcolor="#ffffff" border="0" cellpadding="0" cellspacing="0" role="presentation" style="max-width:680px;" width="680"><tbody><tr><td style="padding:25px 50px 15px 50px;line-height:23px;color:#000;font-family:Roboto,Arial,Helvetica,sans-serif;font-size:16px;font-weight:200;"><p style="font-family:Roboto,Arial,Helvetica,sans-serif;font-size:16px;font-weight:300;">Hi {{Recipient.FirstName}},</p><p style="font-family:Roboto,Arial,Helvetica,sans-serif;font-size:16px;font-weight:300;">{{BodyContent}}</p><p style="font-family:Roboto,Arial,Helvetica,sans-serif;font-size:17px;margin-bottom:0;padding-bottom:0;"><strong style="font-weight:700;">{{SectionHeader}}</strong></p></td></tr></tbody></table></td></tr></tbody></table></td></tr><tr><td align="center" bgcolor="#ffffff" valign="top"><table align="center" bgcolor="#000000" border="0" cellpadding="0" cellspacing="0" role="presentation" style="max-width:680px;background:#000000;" width="680"><tbody><tr><td align="center" valign="middle"><div style="display:inline-block;margin:0;max-width:50%;min-width:240px;vertical-align:middle;"><table border="0" cellpadding="0" cellspacing="0" role="presentation" width="100%"><tbody><tr><td style="font-family:Arial,Verdana;font-size:20px;line-height:24px;color:#72246c;padding:10px 50px;text-align:left;"><a href="https://www.cpa.com" style="color:#ffffff;text-decoration:none;" target="_blank">CPA.com</a></td></tr></tbody></table></div></td></tr></tbody></table><table align="center" bgcolor="#eeeeee" cellpadding="0" cellspacing="0" style="width:680px;padding:0 25px;" width="680"><tbody><tr><td style="text-align:center;vertical-align:top;font-size:0;padding:30px 0px 0px"><div><div style="display:inline-block;vertical-align:top;max-width:330px;"><table align="right" border="0" cellpadding="0" cellspacing="0" height="180" width="330"><tbody><tr><td align="right" style="padding:0px 20px 0px 50px;font-family:Roboto,Arial,Helvetica,sans-serif;font-size:12px;color:#63656b;line-height:100%;font-weight:300;text-align:left;" valign="top"><span>1345 Avenue of the Americas, 27th Floor</span><br><span>New York, NY 10019</span><br><span>888.777.7077</span></td></tr></tbody></table></div></div></td></tr></tbody></table></td></tr></tbody></table></div></td></tr></tbody></table></body></html>`,

  "DOTCPA General": `<!doctype html><html><head><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>{{Subject}}</title><style type="text/css">@import url('https://fonts.googleapis.com/css?family=Roboto:300,400,500,700&display=swap');body{background-color:#ffffff;margin:0}table{border-collapse:collapse}</style></head><body bgcolor="#ffffff" style="margin:0px;padding:0px;" yahoo="fix"><table bgcolor="#ffffff" border="0" cellpadding="0" cellspacing="0" style="table-layout:fixed;" width="100%"><tbody><tr><td align="center" bgcolor="#ffffff" style="margin:0 auto!important;min-width:680px;padding:0;width:680px;" valign="top" width="680"><div align="center"><table align="center" border="0" cellpadding="0" cellspacing="0" style="width:640px;max-width:640px;" width="640"><tbody><tr><td align="center" valign="top"><table align="center" border="0" cellpadding="0" cellspacing="0" style="max-width:680px;border-bottom:1px solid #cccccc" width="100%"><tbody><tr><td align="center"><a href="https://register.domains.cpa/"><img alt=".CPA - A service of AICPA and CPA.com" border="0" height="120" src="https://marketing.cpa.com/l/701003/2020-09-15/n6drm/701003/77696/dotCPA_service_email_header.png" style="width:680px;height:120px;" width="680"></a></td></tr></tbody></table></td></tr><tr><td align="center" bgcolor="#ffffff" valign="top"><table align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="max-width:680px" width="100%"><tbody><tr><td style="padding:25px 50px 0px 50px;font-family:'Roboto',Arial,Verdana;font-weight:300;font-size:17px;line-height:23px;color:#63656b;"><p style="font-family:'Roboto',Arial,sans-serif;font-size:16px;color:#000000;">Hi {{Recipient.FirstName}},</p><p style="font-family:'Roboto',Arial,sans-serif;font-size:16px;color:#000000;">{{BodyContent}}</p><table border="0" cellpadding="0" cellspacing="0" width="100%"><tbody><tr><td align="center"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse:separate!important;" width="300"><tbody><tr><td style="background-color:#72246C;color:#FFFFFF;font-size:15px;padding:10px;border-radius:3px;font-family:Arial,Helvetica,sans-serif;text-align:center;"><a href="https://register.domains.cpa/ga/#start" style="color:#FFFFFF;text-decoration:none;">{{ButtonText}}</a></td></tr></tbody></table></td></tr></tbody></table></td></tr></tbody></table></td></tr><tr><td align="center" bgcolor="#ffffff" valign="top"><table align="center" bgcolor="#000000" border="0" cellpadding="0" cellspacing="0" role="presentation" style="max-width:680px;" width="100%"><tbody><tr><td align="center" style="font-size:0;" valign="middle"><div style="display:inline-block;margin:0;max-width:70%;min-width:430px;vertical-align:middle;"><table border="0" cellpadding="0" cellspacing="0" role="presentation" width="100%"><tbody><tr><td style="font-family:'Roboto',Arial,Verdana;font-size:16px;font-weight:300;color:#72246c;padding:5px 0px 5px 30px;text-align:left;"><a href="http://www.cpa.com/dotcpa" style="color:#ffffff;text-decoration:none;">Domains.CPA</a></td></tr></tbody></table></div></td></tr></tbody></table><table align="center" bgcolor="#eeeeee" cellpadding="0" cellspacing="0" style="max-width:680px;" width="680"><tbody><tr><td style="padding:0px 25px 7px 35px;font-family:'Roboto',Arial;font-size:11px;color:#000;line-height:14px;text-align:left;">{{JobNumber}}</td></tr></tbody></table></td></tr></tbody></table></div></td></tr></tbody></table></body></html>`,

  "CPACOM General": `<!DOCTYPE html><html><head><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>{{Subject}}</title><style type="text/css">@import url('https://fonts.googleapis.com/css?family=Roboto:300,400,500,700&display=swap');table{border-collapse:collapse}table td{border-collapse:collapse;font-weight:300!important}body{background-color:#fff;margin:0}a{text-decoration:underline;color:#86387f}</style></head><body style="margin:0px;padding:0px;" yahoo="fix" bgcolor="#ffffff"><table bgcolor="#ffffff" border="0" cellpadding="0" cellspacing="0" style="table-layout:fixed;" width="100%"><tbody><tr><td align="center" bgcolor="#ffffff" style="margin:0 auto!important;min-width:680px;padding:0;width:680px;" valign="top" width="680"><div align="center"><table align="center" border="0" cellpadding="0" cellspacing="0" style="width:680px;max-width:680px;" width="680"><tbody><tr><td align="center" valign="top"><table cellpadding="0" cellspacing="0" style="width:680px;padding:0;" width="680"><tbody><tr><td align="center" style="padding:0px;border-bottom:solid 1px #CCC;" valign="top"><img alt="CPA.com Header" border="0" height="125" src="https://marketing.cpa.com/l/701003/2022-09-19/3j68h5/701003/1663614409hZUzc2TK/cpacom_header.jpg" style="width:680px;height:125px;" width="680"></td></tr></tbody></table></td></tr><tr><td align="center" bgcolor="#ffffff" valign="top"><table align="center" bgcolor="#ffffff" border="0" cellpadding="0" cellspacing="0" role="presentation" style="max-width:680px;" width="680"><tbody><tr><td style="padding:20px 50px 0px 50px;text-align:left;color:#000;line-height:22px;font-family:Roboto,Arial,Helvetica,sans-serif;font-size:16px;"><p style="font-family:Roboto,Arial,sans-serif;font-size:16px;font-weight:300;">Hi {{Recipient.FirstName}},</p><p style="font-family:Roboto,Arial,sans-serif;font-size:16px;font-weight:300;">{{BodyContent}}</p></td></tr></tbody></table></td></tr><tr><td align="center" bgcolor="#ffffff" valign="top"><table align="center" bgcolor="#0F206C" border="0" cellpadding="0" cellspacing="0" role="presentation" style="max-width:680px;" width="100%"><tbody><tr><td align="center" valign="middle"><div style="display:inline-block;margin:0;max-width:50%;min-width:240px;vertical-align:middle;"><table border="0" cellpadding="0" cellspacing="0" role="presentation" width="100%"><tbody><tr><td style="font-family:Arial,Verdana;font-size:18px;line-height:24px;color:#72246c;padding:10px 50px;text-align:left;"><a href="https://www.cpa.com" style="color:#ffffff;text-decoration:none;" target="_blank">CPA.com/</a></td></tr></tbody></table></div></td></tr></tbody></table><table align="center" bgcolor="#ffffff" cellpadding="0" cellspacing="0" style="max-width:680px;" width="680"><tbody><tr><td align="center" style="padding:0px 25px 7px 35px;font-family:'Roboto',Arial;font-size:11px;color:#000;line-height:14px;text-align:left;" valign="top">{{JobNumber}}</td></tr></tbody></table></td></tr></tbody></table></div></td></tr></tbody></table></body></html>`
};

const TEMPLATE_KEYS = Object.keys(TEMPLATES);

const STATUS_COLORS = {
  "Not started": "#c4c4c4", "In Development": "#0096f5", "Ready": "#00c875",
  "Template Complete": "#9cd326", "Proofing": "#ff9900", "Approved": "#00c875",
  "Scheduled": "#fdab3d", "Sent": "#9cd326", "Stuck": "#e2445c",
  "Overdue": "#e2445c", "Running": "#0096f5", "Canceled": "#797e93",
  "Holiday": "#9cd326", "Awaiting Feedback": "#ffcb00", "Reschedule": "#ff9900",
  "Not started": "#c4c4c4",
};

function Badge({ label, color }) {
  return (
    <span style={{ background: color || "#c4c4c4", color: "#fff", borderRadius: 4, padding: "2px 10px", fontSize: 12, fontWeight: 500, whiteSpace: "nowrap" }}>
      {label}
    </span>
  );
}

async function fetchNewRequestItems() {
  const res = await fetch(`${API_BASE}/api/tickets`);
  if (!res.ok) throw new Error(`Server error: ${res.status}`);
  return await res.json();
}

async function uploadFileToMonday(itemId, fileName, html) {
  const res = await fetch(`${API_BASE}/api/upload`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ itemId, fileName, html }),
  });
  if (!res.ok) throw new Error(`Upload error: ${res.status}`);
  return await res.json();
}

async function generateEmailHTML(ticket, templateName) {
  const templateHtml = TEMPLATES[templateName];
  const res = await fetch(`${API_BASE}/api/generate`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ ticket, templateName, templateHtml }),
  });
  if (!res.ok) throw new Error(`Generate error: ${res.status}`);
  const data = await res.json();
  return data.html ?? "";
}

export default function App() {
  const [tab, setTab] = useState("queue");
  const [tickets, setTickets] = useState([]);
  const [loadingTickets, setLoadingTickets] = useState(false);
  const [ticketsError, setTicketsError] = useState(null);
  const [selectedTicket, setSelectedTicket] = useState(null);
  const [selectedTemplate, setSelectedTemplate] = useState(TEMPLATE_KEYS[0]);
  const [generating, setGenerating] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [result, setResult] = useState(null);
  const [uploadStatus, setUploadStatus] = useState(null);
  const [copyLabel, setCopyLabel] = useState("Copy HTML");
  const [manualForm, setManualForm] = useState({ name: "", subjectLine: "", description: "", jobNumber: "", sendDate: "", template: TEMPLATE_KEYS[0] });

  const loadTickets = async () => {
    setLoadingTickets(true);
    setTicketsError(null);
    try {
      const data = await fetchNewRequestItems();
      setTickets(data.items || []);
      if (!data.items?.length) setTicketsError("No items found in the New Requests group, or the group doesn't exist yet.");
    } catch (e) {
      setTicketsError("Failed to load tickets: " + e.message);
    } finally {
      setLoadingTickets(false);
    }
  };

  useEffect(() => { loadTickets(); }, []);

  // Pre-load the test ticket result
  useEffect(() => {
    const testTicket = {
      id: "11050806957",
      name: "AI in Focus - June newsletter",
      text86: "AI in Focus — June 2026",
      long_text7: "Monthly AI in Focus newsletter covering the latest AI trends, tools, and insights for CPA professionals. This edition focuses on emerging agentic AI solutions, practical applications in tax and audit, and how firms can stay ahead of the curve in 2026.",
      formula: "26JUN11050806957",
      date4: "2026-06-09",
      status_1: "Newsletter",
      dropdown2: "Research & Innovation",
      person: "Lexie Matinog",
      status: "Not started"
    };
    setSelectedTicket(testTicket);
  }, []);

  const handleGenerate = async (ticket, templateName) => {
    setGenerating(true);
    setResult(null);
    setUploadStatus(null);
    try {
      const html = await generateEmailHTML(ticket, templateName);
      if (!html) throw new Error("No HTML generated.");
      setResult({ html, ticket, templateName });
      setTab("output");
    } catch (e) {
      alert("Generation failed: " + e.message);
    } finally {
      setGenerating(false);
    }
  };

  const handleUpload = async () => {
    if (!result) return;
    setUploading(true);
    setUploadStatus(null);
    const fileName = `${result.ticket.name.replace(/\s+/g, "_")}_${result.ticket.formula || "draft"}.html`;
    try {
      const res = await uploadFileToMonday(result.ticket.id, fileName, result.html);
      setUploadStatus(res.success ? "success" : "error");
    } catch (e) {
      setUploadStatus("error");
    } finally {
      setUploading(false);
    }
  };

  const handleDownload = () => {
    if (!result) return;
    const blob = new Blob([result.html], { type: "text/html" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${result.ticket.name.replace(/\s+/g, "_")}_${result.ticket.formula || "draft"}.html`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const handleManualGenerate = async () => {
    const fakeTicket = {
      id: "manual",
      name: manualForm.name,
      text86: manualForm.subjectLine,
      long_text7: manualForm.description,
      formula: manualForm.jobNumber,
      date4: manualForm.sendDate,
      status_1: "",
      dropdown2: "",
    };
    await handleGenerate(fakeTicket, manualForm.template);
  };

  const tabStyle = (t) => ({
    padding: "6px 16px", fontSize: 13, cursor: "pointer", borderRadius: 6,
    border: tab === t ? "0.5px solid var(--color-border-primary)" : "0.5px solid transparent",
    background: tab === t ? "var(--color-background-secondary)" : "transparent",
    fontWeight: tab === t ? 500 : 400, color: "var(--color-text-primary)"
  });

  return (
    <div style={{ padding: "1rem 0", fontFamily: "var(--font-sans)", maxWidth: 780 }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 4 }}>
        <h2 style={{ fontSize: 18, fontWeight: 600, margin: 0, color: "var(--color-text-primary)" }}>
          📧 Email Template Generator
        </h2>
        <span style={{ fontSize: 12, color: "#00c875", fontWeight: 500 }}>● Connected to Monday.com</span>
      </div>
      <p style={{ fontSize: 13, color: "var(--color-text-secondary)", margin: "0 0 18px" }}>
        Email Schedule Board · New Requests queue
      </p>

      <div style={{ display: "flex", gap: 6, marginBottom: 18 }}>
        {[["queue", "🗂 New Requests"], ["manual", "✏️ Manual entry"], ["output", "📄 Output"]].map(([t, label]) => (
          <button key={t} onClick={() => setTab(t)} style={tabStyle(t)}>{label}</button>
        ))}
      </div>

      {/* NEW REQUESTS TAB */}
      {tab === "queue" && (
        <div>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
            <span style={{ fontSize: 13, color: "var(--color-text-secondary)" }}>
              {loadingTickets ? "Loading…" : `${tickets.length} ticket${tickets.length !== 1 ? "s" : ""} found`}
            </span>
            <button onClick={loadTickets} disabled={loadingTickets} style={{ fontSize: 12, padding: "4px 12px" }}>
              {loadingTickets ? "Refreshing…" : "↻ Refresh"}
            </button>
          </div>

          {ticketsError && (
            <div style={{ padding: "12px 16px", background: "var(--color-background-secondary)", borderRadius: 8, fontSize: 13, color: "var(--color-text-secondary)", marginBottom: 12 }}>
              {ticketsError}
            </div>
          )}

          {!loadingTickets && tickets.length === 0 && !ticketsError && (
            <div style={{ textAlign: "center", padding: "2rem 0", color: "var(--color-text-secondary)", fontSize: 13 }}>
              No tickets in New Requests group. Use Manual entry to test generation.
            </div>
          )}

          {tickets.map(ticket => (
            <div key={ticket.id} style={{
              border: "0.5px solid var(--color-border-tertiary)", borderRadius: 10, padding: "14px 16px",
              marginBottom: 10, background: "var(--color-background-primary)"
            }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 8 }}>
                <div>
                  <div style={{ fontWeight: 500, fontSize: 14, color: "var(--color-text-primary)", marginBottom: 4 }}>{ticket.name}</div>
                  <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                    {ticket.status && <Badge label={ticket.status} color={STATUS_COLORS[ticket.status]} />}
                    {ticket.status_1 && <Badge label={ticket.status_1} color="#0096f5" />}
                    {ticket.date4 && <Badge label={`Send: ${ticket.date4}`} color="#797e93" />}
                  </div>
                </div>
                <div style={{ fontSize: 11, color: "var(--color-text-secondary)", textAlign: "right" }}>
                  {ticket.formula && <div style={{ fontFamily: "monospace" }}>{ticket.formula}</div>}
                  {ticket.person && <div>{ticket.person}</div>}
                </div>
              </div>

              {ticket.long_text7 && (
                <p style={{ fontSize: 12, color: "var(--color-text-secondary)", margin: "0 0 10px", lineHeight: 1.5, maxHeight: 40, overflow: "hidden" }}>
                  {ticket.long_text7}
                </p>
              )}

              <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                <select
                  value={selectedTicket?.id === ticket.id ? selectedTemplate : TEMPLATE_KEYS[0]}
                  onChange={e => { setSelectedTicket(ticket); setSelectedTemplate(e.target.value); }}
                  onClick={() => setSelectedTicket(ticket)}
                  style={{ fontSize: 12, padding: "4px 8px", borderRadius: 6 }}
                >
                  {TEMPLATE_KEYS.map(k => <option key={k}>{k}</option>)}
                </select>
                <button
                  onClick={() => { setSelectedTicket(ticket); handleGenerate(ticket, selectedTicket?.id === ticket.id ? selectedTemplate : TEMPLATE_KEYS[0]); }}
                  disabled={generating}
                  style={{ fontSize: 12, padding: "5px 14px", cursor: generating ? "not-allowed" : "pointer", opacity: generating ? 0.6 : 1 }}
                >
                  {generating && selectedTicket?.id === ticket.id ? "Generating…" : "Generate →"}
                </button>
              </div>
            </div>
          ))}
        </div>
      )}

      {/* MANUAL TAB */}
      {tab === "manual" && (
        <div style={{ background: "var(--color-background-primary)", border: "0.5px solid var(--color-border-tertiary)", borderRadius: 10, padding: "1.25rem" }}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 12 }}>
            {[["name", "Item name *", "e.g. Nov Town Hall Newsletter"], ["subjectLine", "Subject line *", "Email subject line"], ["jobNumber", "Job number", "e.g. 26MAY12345678"], ["sendDate", "Send date", ""]].map(([k, label, ph]) => (
              <div key={k}>
                <label style={{ fontSize: 12, color: "var(--color-text-secondary)", display: "block", marginBottom: 4 }}>{label}</label>
                <input type={k === "sendDate" ? "date" : "text"} value={manualForm[k]} onChange={e => setManualForm(f => ({ ...f, [k]: e.target.value }))} placeholder={ph} style={{ width: "100%", boxSizing: "border-box" }} />
              </div>
            ))}
          </div>
          <div style={{ marginBottom: 12 }}>
            <label style={{ fontSize: 12, color: "var(--color-text-secondary)", display: "block", marginBottom: 4 }}>Template *</label>
            <select value={manualForm.template} onChange={e => setManualForm(f => ({ ...f, template: e.target.value }))} style={{ width: "100%", boxSizing: "border-box" }}>
              {TEMPLATE_KEYS.map(k => <option key={k}>{k}</option>)}
            </select>
          </div>
          <div style={{ marginBottom: 14 }}>
            <label style={{ fontSize: 12, color: "var(--color-text-secondary)", display: "block", marginBottom: 4 }}>Description / body content *</label>
            <textarea value={manualForm.description} onChange={e => setManualForm(f => ({ ...f, description: e.target.value }))}
              placeholder="Paste ticket description or body content here…"
              rows={5} style={{ width: "100%", boxSizing: "border-box", fontFamily: "var(--font-sans)", fontSize: 13, padding: 8, borderRadius: 8, border: "0.5px solid var(--color-border-secondary)", background: "var(--color-background-primary)", color: "var(--color-text-primary)", resize: "vertical" }} />
          </div>
          <button onClick={handleManualGenerate} disabled={generating || !manualForm.name || !manualForm.subjectLine || !manualForm.description}
            style={{ padding: "8px 20px", cursor: "pointer" }}>
            {generating ? "Generating…" : "Generate email HTML →"}
          </button>
        </div>
      )}

      {/* OUTPUT TAB */}
      {tab === "output" && (
        <div>
          {!result ? (
            <div style={{ textAlign: "center", padding: "3rem 0", color: "var(--color-text-secondary)", fontSize: 13 }}>
              No output yet — select a ticket or use manual entry to generate.
            </div>
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
              <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                <Badge label="Template Complete" color="#9cd326" />
                <span style={{ fontSize: 13, color: "var(--color-text-secondary)" }}>{result.ticket.name}</span>
                <span style={{ fontSize: 12, color: "var(--color-text-secondary)" }}>· {result.templateName}</span>
              </div>

              <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                <button onClick={handleDownload} style={{ fontSize: 13 }}>⬇ Download .html</button>
                <button onClick={() => { navigator.clipboard.writeText(result.html); setCopyLabel("Copied!"); setTimeout(() => setCopyLabel("Copy HTML"), 2000); }} style={{ fontSize: 13 }}>{copyLabel}</button>
                {result.ticket.id !== "manual" && (
                  <button onClick={handleUpload} disabled={uploading}
                    style={{ fontSize: 13, background: uploading ? undefined : "var(--color-background-success)", cursor: uploading ? "not-allowed" : "pointer", opacity: uploading ? 0.6 : 1 }}>
                    {uploading ? "Uploading…" : "⬆ Upload to Monday ticket"}
                  </button>
                )}
              </div>

              {uploadStatus === "success" && (
                <div style={{ padding: "10px 14px", background: "var(--color-background-success)", borderRadius: 8, fontSize: 13, color: "var(--color-text-success)" }}>
                  ✓ HTML file uploaded successfully to the ticket's Files column!
                </div>
              )}
              {uploadStatus === "error" && (
                <div style={{ padding: "10px 14px", background: "var(--color-background-danger)", borderRadius: 8, fontSize: 13, color: "var(--color-text-danger)" }}>
                  ✗ Upload failed. You can download the file and attach it manually to the ticket.
                </div>
              )}

              <div style={{ border: "0.5px solid var(--color-border-tertiary)", borderRadius: 10, overflow: "hidden" }}>
                <div style={{ background: "var(--color-background-secondary)", padding: "8px 14px", fontSize: 12, color: "var(--color-text-secondary)", borderBottom: "0.5px solid var(--color-border-tertiary)" }}>
                  Preview · {result.ticket.text86 || result.ticket.name}
                </div>
                <iframe srcDoc={result.html} style={{ width: "100%", height: 500, border: "none", background: "#fff" }} title="Email preview" sandbox="allow-same-origin" />
              </div>

              <details>
                <summary style={{ fontSize: 13, color: "var(--color-text-secondary)", cursor: "pointer" }}>View raw HTML</summary>
                <pre style={{ marginTop: 8, padding: 12, background: "var(--color-background-secondary)", borderRadius: 8, fontSize: 11, overflowX: "auto", border: "0.5px solid var(--color-border-tertiary)", color: "var(--color-text-primary)", maxHeight: 300, overflowY: "auto", whiteSpace: "pre-wrap", wordBreak: "break-all" }}>{result.html}</pre>
              </details>
            </div>
          )}
        </div>
      )}
    </div>
  );
}