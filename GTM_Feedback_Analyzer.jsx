import { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

const REVIEWS = [
  {id:1,name:"Mim Nam",flavor:"Cookie Dough",taste:"Good",texture:"Too icy",sweetness:"Just right",buy:"maybe",matters:"Low/no sugar",comment:""},
  {id:2,name:"Alexandar Tabilas",flavor:"Cookie Dough",taste:"Good",texture:"Too icy",sweetness:"Just right",buy:"no",matters:"High protein",comment:""},
  {id:3,name:"Henry",flavor:"Cookie Dough",taste:"Good",texture:"Too icy",sweetness:"Just right",buy:"maybe",matters:"Low/no sugar",comment:"Market it as frozen yogurt"},
  {id:4,name:"Iris Yang",flavor:"Cookie Dough",taste:"",texture:"Too thick",sweetness:"Just right",buy:"maybe",matters:"Low/no sugar",comment:""},
  {id:5,name:"Lindsay Ulrey",flavor:"Cookie Dough",taste:"Good",texture:"Too icy",sweetness:"Just right",buy:"maybe",matters:"Low/no sugar",comment:""},
  {id:6,name:"Kene Olusanya",flavor:"Cookie Dough",taste:"Good",texture:"Just right",sweetness:"Just right",buy:"maybe",matters:"both",comment:""},
  {id:7,name:"Mariana Menezes",flavor:"Cookie Dough",taste:"Excellent",texture:"Just right",sweetness:"Just right",buy:"Yes",matters:"Low in sugar",comment:"Cookie dough is all in one place and very big chunks are there try to spread it and minimise the size."},
  {id:8,name:"Yoda Ermias",flavor:"Cookie Dough",taste:"Good",texture:"Too icy",sweetness:"Just right",buy:"yes",matters:"both",comment:""},
  {id:9,name:"Hazen Pike",flavor:"Cookie Dough",taste:"good",texture:"Just right",sweetness:"Just right",buy:"yes",matters:"low/no sugar",comment:"protein cookie dough"},
  {id:10,name:"Kevin Oye",flavor:"Cookie Dough",taste:"Excellent",texture:"Too icy",sweetness:"Just right",buy:"yes",matters:"both",comment:""},
  {id:11,name:"Jonathan",flavor:"Cookie Dough",taste:"neutral",texture:"too icy",sweetness:"Just right",buy:"maybe",matters:"both",comment:""},
  {id:12,name:"Zoe",flavor:"Cookie Dough",taste:"neutral",texture:"too icy",sweetness:"Just right",buy:"maybe",matters:"both",comment:""},
  {id:13,name:"Aditi",flavor:"Cookie Dough",taste:"Good",texture:"Just right",sweetness:"Just right",buy:"maybe",matters:"low/no sugar",comment:""},
  {id:14,name:"Alyssa",flavor:"Cookie Dough",taste:"Good",texture:"Too icy",sweetness:"Just right",buy:"maybe",matters:"low/no sugar",comment:""},
  {id:15,name:"Field",flavor:"Cookie Dough",taste:"Excellent",texture:"Just right",sweetness:"Just right",buy:"Yes",matters:"Low in sugar",comment:""},
  {id:16,name:"Shyna",flavor:"Cookie Dough",taste:"Excellent",texture:"Just right",sweetness:"Just right",buy:"Yes",matters:"Low in sugar",comment:""},
  {id:17,name:"Ben",flavor:"Cookie Dough",taste:"Excellent",texture:"Just right",sweetness:"Just right",buy:"Yes",matters:"Low in sugar",comment:""},
  {id:18,name:"Smriti",flavor:"Cookie Dough",taste:"Good",texture:"Just right",sweetness:"too sweet",buy:"maybe",matters:"both",comment:""},
  {id:19,name:"Sai",flavor:"Cookie Dough",taste:"Excellent",texture:"Just right",sweetness:"Just right",buy:"Yes",matters:"Low in sugar",comment:""},
  {id:20,name:"Kevin Aka",flavor:"Vanilla",taste:"Neutral",texture:"Just right",sweetness:"Just right",buy:"Maybe",matters:"Low/no sugar",comment:"after taste was bad"},
  {id:21,name:"Davae Lynch",flavor:"Vanilla",taste:"Neutral",texture:"Too thick",sweetness:"Just right",buy:"No",matters:"Low/no sugar",comment:"less yogurt kinda flavor"},
  {id:22,name:"Angela Yan",flavor:"Vanilla",taste:"Good",texture:"Just right",sweetness:"not sweet enough",buy:"Maybe",matters:"Taste",comment:"more flavourfull"},
  {id:23,name:"Rajiv Sonawane",flavor:"Vanilla",taste:"Neutral",texture:"Too thick",sweetness:"Just right",buy:"Maybe",matters:"both",comment:"vanila flavor can be better"},
  {id:24,name:"Adrien Tabor",flavor:"Vanilla",taste:"Neutral",texture:"Just right",sweetness:"too sweet",buy:"no",matters:"both",comment:"Feels Unhealthy to eat icecream"},
  {id:25,name:"Jordon Gassatura",flavor:"Vanilla",taste:"Good",texture:"Too icy",sweetness:"Just right",buy:"Maybe",matters:"both",comment:"It's great air, texture is good, should be slightly sweet."},
  {id:26,name:"Jesse Flores",flavor:"Vanilla",taste:"Good",texture:"Just right",sweetness:"Just right",buy:"maybe",matters:"Low/no sugar",comment:""},
  {id:27,name:"Eyobel Gebre",flavor:"Vanilla",taste:"Excellent",texture:"Just right",sweetness:"Just right",buy:"Yes",matters:"High Protein",comment:""},
  {id:28,name:"Megan Yi",flavor:"Vanilla",taste:"Good",texture:"Too icy",sweetness:"Just right",buy:"Maybe",matters:"Low / no sugar",comment:""},
  {id:29,name:"Mariana Menezes",flavor:"Vanilla",taste:"Excellent",texture:"Just right",sweetness:"Just right",buy:"Yes",matters:"High Protein",comment:""},
  {id:30,name:"Aryaa",flavor:"Vanilla",taste:"Neutral",texture:"Too thick",sweetness:"Just right",buy:"Maybe",matters:"High Protein",comment:"Masking the flavor of cottage cheese more"},
  {id:31,name:"Jonah",flavor:"Vanilla",taste:"",texture:"Too thick",sweetness:"Just right",buy:"no",matters:"High Protein",comment:"negan flavors"},
  {id:32,name:"Vorleak Hak",flavor:"Vanilla",taste:"Excellent",texture:"Just right",sweetness:"not sweet enough",buy:"Maybe",matters:"both",comment:""},
  {id:33,name:"Kevin Aka",flavor:"Vanilla",taste:"Neutral",texture:"Just right",sweetness:"Just right",buy:"Maybe",matters:"Low / no sugar",comment:""},
  {id:34,name:"Dave",flavor:"Vanilla",taste:"Neutral",texture:"too thick",sweetness:"just right",buy:"no",matters:"no sugar",comment:"really liked"},
  {id:35,name:"Ava",flavor:"Vanilla",taste:"good",texture:"too thick",sweetness:"not sweet enough",buy:"maybe",matters:"no sugar",comment:""},
  {id:36,name:"Adrein",flavor:"Vanilla",taste:"Neutral",texture:"Just right",sweetness:"too sweet",buy:"no",matters:"low no sugar",comment:""},
  {id:37,name:"Nam",flavor:"Vanilla",taste:"Excellent",texture:"Just right",sweetness:"Just right",buy:"Yes",matters:"low no sugar",comment:""},
  {id:38,name:"Mjf",flavor:"Vanilla",taste:"good",texture:"Too icy",sweetness:"Just right",buy:"maybe",matters:"High Protein",comment:""},
  {id:39,name:"Kevin Oye",flavor:"Cinnamon",taste:"Excellent",texture:"Just right",sweetness:"Just right",buy:"Yes",matters:"Both",comment:"test the label on me"},
  {id:40,name:"Wasif Khan",flavor:"Cinnamon",taste:"Good",texture:"Just right",sweetness:"Just right",buy:"Yes",matters:"High Protein",comment:"Go for Combos too, maybe protein bar with icecream"},
  {id:41,name:"Sill Rarlee",flavor:"Cinnamon",taste:"Excellent",texture:"Just right",sweetness:"Just right/ too sweet",buy:"no",matters:"both",comment:"Icyness turns me off rest i like the flavour"},
  {id:42,name:"MaryAnn",flavor:"Cinnamon",taste:"Excellent",texture:"Just right",sweetness:"Just right",buy:"Yes",matters:"Both",comment:""},
  {id:43,name:"Balakrishnan Anand",flavor:"Cinnamon",taste:"Good",texture:"Just right",sweetness:"not sweet enough",buy:"Yes",matters:"High Protein",comment:"lots of sugar"},
  {id:44,name:"Aryan Raj",flavor:"Cinnamon",taste:"Good",texture:"Just right",sweetness:"not sweet enough",buy:"Yes",matters:"High Protein",comment:"lots of sugar"},
  {id:45,name:"Maura",flavor:"Cinnamon",taste:"Good",texture:"Just right",sweetness:"Just right",buy:"Yes",matters:"Both",comment:""},
  {id:46,name:"Smriti Lal",flavor:"Cinnamon",taste:"Good",texture:"Just right",sweetness:"Just right",buy:"maybe",matters:"Both",comment:""},
  {id:47,name:"Kapil",flavor:"Cinnamon",taste:"Excellent",texture:"Just right",sweetness:"Just right",buy:"Yes",matters:"Both",comment:"some sprinkle on top"},
  {id:48,name:"Hazen Pike",flavor:"Cinnamon",taste:"good",texture:"too icy",sweetness:"Just right",buy:"maybe",matters:"low/no sugar",comment:""},
  {id:49,name:"Field",flavor:"Cinnamon",taste:"good",texture:"too icy",sweetness:"Just right",buy:"maybe",matters:"low/no sugar",comment:""},
  {id:50,name:"Nam",flavor:"Cinnamon",taste:"good",texture:"too icy",sweetness:"Just right",buy:"maybe",matters:"low no sugar",comment:""},
  {id:51,name:"Dylan",flavor:"Cinnamon",taste:"good",texture:"too icy",sweetness:"Just right",buy:"maybe",matters:"High Protein",comment:""},
  {id:52,name:"Rajat",flavor:"Cinnamon",taste:"good",texture:"too icy",sweetness:"Just right",buy:"maybe",matters:"both",comment:""},
  {id:53,name:"Vinit",flavor:"Cinnamon",taste:"good",texture:"too icy",sweetness:"Just right",buy:"no",matters:"High Protein",comment:""},
  {id:54,name:"Samin",flavor:"Cinnamon",taste:"good",texture:"too icy",sweetness:"not sweet enough",buy:"maybe",matters:"High Protein",comment:""},
  {id:55,name:"Marie",flavor:"Cinnamon",taste:"neutral",texture:"too icy",sweetness:"too sweet",buy:"maybe",matters:"both",comment:""},
  {id:56,name:"Shyna",flavor:"Cinnamon",taste:"good",texture:"too icy",sweetness:"too sweet",buy:"maybe",matters:"low no sugar",comment:""},
  {id:57,name:"Ben",flavor:"Cinnamon",taste:"good",texture:"too icy",sweetness:"just right",buy:"maybe",matters:"low no sugar",comment:""},
  {id:58,name:"Augustine",flavor:"Cinnamon",taste:"neutral",texture:"too icy",sweetness:"-",buy:"maybe",matters:"both",comment:""},
  {id:59,name:"Efe",flavor:"Cinnamon",taste:"good",texture:"too icy",sweetness:"just right",buy:"maybe",matters:"neither",comment:""},
  {id:60,name:"Mariana",flavor:"Cinnamon",taste:"neutral",texture:"too icy",sweetness:"just right",buy:"no",matters:"low no sugar",comment:""},
  {id:61,name:"Sai",flavor:"Cinnamon",taste:"neutral",texture:"too icy",sweetness:"just right",buy:"maybe",matters:"low no sugar",comment:""},
];

const FLAVORS = ["Cookie Dough", "Vanilla", "Cinnamon"];
const norm = (s) => (s || "").trim().toLowerCase();

function computeStats(reviews) {
  const byFlavor = {};
  FLAVORS.forEach(f => {
    const sub = reviews.filter(r => r.flavor === f);
    const n = sub.length;
    const yes = sub.filter(r => norm(r.buy) === "yes").length;
    const maybe = sub.filter(r => norm(r.buy) === "maybe").length;
    const no = sub.filter(r => norm(r.buy) === "no").length;
    const excellent = sub.filter(r => norm(r.taste) === "excellent").length;
    const good = sub.filter(r => norm(r.taste) === "good").length;
    const neutral = sub.filter(r => norm(r.taste) === "neutral").length;
    const icy = sub.filter(r => norm(r.texture).includes("icy")).length;
    const thick = sub.filter(r => norm(r.texture).includes("thick")).length;
    const justRight = sub.filter(r => norm(r.texture) === "just right").length;
    const tooSweet = sub.filter(r => norm(r.sweetness).includes("too sweet")).length;
    const notSweet = sub.filter(r => norm(r.sweetness).includes("not sweet")).length;
    const sugarPriority = sub.filter(r => norm(r.matters).includes("sugar") || norm(r.matters).includes("low")).length;
    const proteinPriority = sub.filter(r => norm(r.matters).includes("protein")).length;
    const bothPriority = sub.filter(r => norm(r.matters) === "both").length;
    byFlavor[f] = { n, yes, maybe, no, excellent, good, neutral, icy, thick, justRight, tooSweet, notSweet, sugarPriority, proteinPriority, bothPriority };
  });
  return byFlavor;
}

const callClaude = async (prompt) => {
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: 1000,
      messages: [{ role: "user", content: prompt }],
    }),
  });
  const data = await res.json();
  return data.content.filter(b => b.type === "text").map(b => b.text).join("");
};

const TABS = [
  { key: "overview", label: "Overview", icon: "ti-chart-dots-3" },
  { key: "ai", label: "AI analysis", icon: "ti-sparkles" },
  { key: "gtm", label: "GTM motion", icon: "ti-rocket" },
  { key: "compare", label: "Before / after", icon: "ti-arrows-exchange" },
];

export default function App() {
  const [tab, setTab] = useState("overview");
  const [flavorFilter, setFlavorFilter] = useState("All");
  const [aiResults, setAiResults] = useState(null);
  const [gtmResults, setGtmResults] = useState(null);
  const [loading, setLoading] = useState("");
  const [progress, setProgress] = useState(0);
  const stats = computeStats(REVIEWS);
  const abortRef = useRef(false);

  const filtered = flavorFilter === "All" ? REVIEWS : REVIEWS.filter(r => r.flavor === flavorFilter);
  const totalYes = REVIEWS.filter(r => norm(r.buy) === "yes").length;
  const totalMaybe = REVIEWS.filter(r => norm(r.buy) === "maybe").length;

  const runAiAnalysis = useCallback(async () => {
    setLoading("ai");
    setProgress(0);
    abortRef.current = false;
    const results = [];
    const batchSize = 10;
    const batches = [];
    for (let i = 0; i < REVIEWS.length; i += batchSize) {
      batches.push(REVIEWS.slice(i, i + batchSize));
    }

    for (let bi = 0; bi < batches.length; bi++) {
      if (abortRef.current) break;
      const batch = batches[bi];
      const reviewsText = batch.map((r, i) => 
        `#${bi * batchSize + i + 1} ${r.name} (${r.flavor}): taste=${r.taste||"N/A"}, texture=${r.texture}, sweetness=${r.sweetness}, buy=${r.buy}, priority=${r.matters}, comment="${r.comment || "none"}"`
      ).join("\n");

      const prompt = `You are a product analyst for a protein ice cream startup. Analyze these ${batch.length} taste test reviews and return ONLY a JSON array (no markdown/backticks). Each object must have:
{"name":"...","sentiment":"Positive|Neutral|Negative","score":<1-10>,"intent":<0-100>,"texture_insight":"<max 10 words>","sweetness_insight":"<max 10 words>","purchase_insight":"<max 10 words>","priority_insight":"<max 10 words>","comment_insight":"<max 12 words>","tags":["tag1","tag2"],"recommendation":"<max 15 words>"}

Reviews:
${reviewsText}`;

      try {
        const raw = await callClaude(prompt);
        const clean = raw.replace(/```json|```/g, "").trim();
        const parsed = JSON.parse(clean);
        results.push(...parsed);
      } catch (e) {
        batch.forEach(r => results.push({
          name: r.name, sentiment: "Neutral", score: 5, intent: 50,
          texture_insight: "Analysis pending", sweetness_insight: "Analysis pending",
          purchase_insight: "Analysis pending", priority_insight: "Analysis pending",
          comment_insight: "Analysis pending", tags: ["pending"], recommendation: "Retry analysis"
        }));
      }
      setProgress(Math.round(((bi + 1) / batches.length) * 100));
    }
    setAiResults(results);
    setLoading("");
  }, []);

  const runGtmAnalysis = useCallback(async () => {
    setLoading("gtm");
    const allData = REVIEWS.map(r =>
      `${r.name}|${r.flavor}|${r.taste}|${r.texture}|${r.sweetness}|${r.buy}|${r.matters}|${r.comment || "none"}`
    ).join("\n");

    const prompt = `You are a GTM strategist for a protein ice cream brand. Here are 61 taste test reviews (name|flavor|taste|texture|sweetness|buy|priority|comment):

${allData}

Synthesize ALL reviews into a GTM motion plan. Return ONLY valid JSON (no markdown, no backticks):
{
  "positioning_statement": "<2 sentences: what this product is, who it's for, why it wins>",
  "icp_profiles": [
    {"segment":"<name>","description":"<1 sentence>","size_pct":<0-100>,"buy_signal":"<key phrase>","objection":"<main concern>","messaging":"<1 line pitch>"}
  ],
  "value_props": ["<prop 1>","<prop 2>","<prop 3>"],
  "objection_handling": [
    {"objection":"<common objection from data>","response":"<GTM response>","evidence":"<data point from reviews>"}
  ],
  "flavor_recommendations": [
    {"flavor":"Cookie Dough","verdict":"<Launch/Iterate/Kill>","reason":"<1 sentence>","priority_fix":"<1 action>"},
    {"flavor":"Vanilla","verdict":"<Launch/Iterate/Kill>","reason":"<1 sentence>","priority_fix":"<1 action>"},
    {"flavor":"Cinnamon","verdict":"<Launch/Iterate/Kill>","reason":"<1 sentence>","priority_fix":"<1 action>"}
  ],
  "launch_messaging": {"headline":"<max 8 words>","subhead":"<max 15 words>","cta":"<max 5 words>"},
  "risk_factors": ["<risk 1>","<risk 2>","<risk 3>"],
  "key_metric_targets": {"conversion_target":"<X%>","nps_target":"<X>","repeat_rate_target":"<X%>"}
}

Base EVERYTHING on the actual review data patterns. Be specific with numbers.`;

    try {
      const raw = await callClaude(prompt);
      const clean = raw.replace(/```json|```/g, "").trim();
      setGtmResults(JSON.parse(clean));
    } catch (e) {
      setGtmResults({ error: "GTM analysis failed. Please try again." });
    }
    setLoading("");
  }, []);

  const exportExcel = () => {
    const wb = XLSX.utils.book_new();
    const rawData = REVIEWS.map((r, i) => ({
      "#": i + 1, Name: r.name, Flavor: r.flavor, "Overall Taste": r.taste || "—",
      Texture: r.texture, Sweetness: r.sweetness, "Would Buy": r.buy,
      "Matters Most": r.matters, Comment: r.comment || "—"
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rawData), "Raw Reviews");

    if (aiResults) {
      const aiData = aiResults.map((a, i) => ({
        "#": i + 1, Name: a.name, Sentiment: a.sentiment, "Score (1-10)": a.score,
        "Purchase Intent %": a.intent, "Texture Insight": a.texture_insight,
        "Sweetness Insight": a.sweetness_insight, "Purchase Insight": a.purchase_insight,
        "Comment Insight": a.comment_insight, Tags: (a.tags || []).join(", "),
        Recommendation: a.recommendation
      }));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(aiData), "AI Analysis");

      const baData = aiResults.map((a, i) => {
        const r = REVIEWS[i] || {};
        return {
          Name: a.name, Flavor: r.flavor,
          "Raw Texture": r.texture, "AI Texture Insight": a.texture_insight,
          "Raw Sweetness": r.sweetness, "AI Sweetness Insight": a.sweetness_insight,
          "Raw Buy": r.buy, "AI Purchase Insight": a.purchase_insight,
          "Raw Comment": r.comment || "—", "AI Comment Insight": a.comment_insight,
        };
      });
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(baData), "Before vs After");
    }

    if (gtmResults && !gtmResults.error) {
      const gtmSheet = [
        { Section: "POSITIONING", Content: gtmResults.positioning_statement },
        ...gtmResults.value_props.map((v, i) => ({ Section: `Value Prop ${i + 1}`, Content: v })),
        ...(gtmResults.icp_profiles || []).map(p => ({
          Section: `ICP: ${p.segment}`, Content: `${p.description} | Buy signal: ${p.buy_signal} | Objection: ${p.objection} | Pitch: ${p.messaging}`
        })),
        ...(gtmResults.objection_handling || []).map(o => ({
          Section: `Objection: ${o.objection}`, Content: `Response: ${o.response} | Evidence: ${o.evidence}`
        })),
        ...(gtmResults.flavor_recommendations || []).map(f => ({
          Section: `${f.flavor}: ${f.verdict}`, Content: `${f.reason} | Fix: ${f.priority_fix}`
        })),
        { Section: "LAUNCH HEADLINE", Content: gtmResults.launch_messaging?.headline },
        { Section: "LAUNCH SUBHEAD", Content: gtmResults.launch_messaging?.subhead },
        ...gtmResults.risk_factors.map((r, i) => ({ Section: `Risk ${i + 1}`, Content: r })),
      ];
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(gtmSheet), "GTM Motion");
    }
    XLSX.writeFile(wb, "GTM_Feedback_Analysis.xlsx");
  };

  const sentColor = (s) => s === "Positive" ? "#34d399" : s === "Negative" ? "#f87171" : "#fbbf24";
  const scoreColor = (v) => v >= 7 ? "#34d399" : v >= 5 ? "#fbbf24" : "#f87171";

  return (
    <div style={{ background: "#0f1117", minHeight: "100vh", color: "#f1f0f9", fontFamily: "system-ui, sans-serif" }}>
      {/* Header */}
      <div style={{ background: "#1a1d27", borderBottom: "1px solid rgba(255,255,255,0.08)", padding: "16px 24px", display: "flex", alignItems: "center", gap: 14 }}>
        <div style={{ width: 36, height: 36, borderRadius: 10, background: "linear-gradient(135deg,#7c6fee,#2dd4bf)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18 }}>🍦</div>
        <div style={{ flex: 1 }}>
          <div style={{ fontSize: 16, fontWeight: 600 }}>Ice cream feedback analyzer</div>
          <div style={{ fontSize: 12, color: "#8b8fa8" }}>AI-powered GTM synthesis from 61 taste test interviews</div>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button onClick={exportExcel} style={{ background: "rgba(52,211,153,0.12)", border: "1px solid rgba(52,211,153,0.3)", borderRadius: 8, padding: "7px 14px", color: "#34d399", fontSize: 12, cursor: "pointer", display: "flex", alignItems: "center", gap: 6 }}>
            <i className="ti ti-table-export" style={{ fontSize: 14 }} /> Export Excel
          </button>
          <span style={{ background: "rgba(124,111,238,0.18)", color: "#a78bfa", fontSize: 11, padding: "6px 12px", borderRadius: 20, border: "1px solid rgba(124,111,238,0.3)" }}>✦ Claude AI</span>
        </div>
      </div>

      {/* Tabs */}
      <div style={{ display: "flex", borderBottom: "1px solid rgba(255,255,255,0.06)", background: "#1a1d27", padding: "0 24px" }}>
        {TABS.map(t => (
          <button key={t.key} onClick={() => setTab(t.key)} style={{
            padding: "11px 18px", fontSize: 13, color: tab === t.key ? "#a78bfa" : "#8b8fa8",
            borderBottom: tab === t.key ? "2px solid #7c6fee" : "2px solid transparent",
            background: "none", border: "none", borderTop: "none", borderLeft: "none", borderRight: "none",
            cursor: "pointer", display: "flex", alignItems: "center", gap: 6, transition: "all 0.15s"
          }}>
            <i className={`ti ${t.icon}`} style={{ fontSize: 15 }} /> {t.label}
          </button>
        ))}
      </div>

      <div style={{ padding: "20px 24px", maxWidth: 1100 }}>
        {/* ===== OVERVIEW ===== */}
        {tab === "overview" && (
          <div>
            {/* KPIs */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(5, 1fr)", gap: 10, marginBottom: 22 }}>
              {[
                { label: "Total reviews", val: "61", color: "#f1f0f9" },
                { label: "Would buy (Yes)", val: `${totalYes} (${Math.round(totalYes/61*100)}%)`, color: "#34d399" },
                { label: "On the fence", val: `${totalMaybe} (${Math.round(totalMaybe/61*100)}%)`, color: "#fbbf24" },
                { label: "Top complaint", val: "Icy texture", color: "#f87171" },
                { label: "Top priority", val: "Low sugar", color: "#2dd4bf" },
              ].map((k, i) => (
                <div key={i} style={{ background: "#1a1d27", borderRadius: 10, padding: "14px 16px", border: "1px solid rgba(255,255,255,0.06)" }}>
                  <div style={{ fontSize: 11, color: "#8b8fa8", marginBottom: 4 }}>{k.label}</div>
                  <div style={{ fontSize: 20, fontWeight: 600, color: k.color }}>{k.val}</div>
                </div>
              ))}
            </div>

            {/* Flavor Cards */}
            <div style={{ fontSize: 12, fontWeight: 600, color: "#8b8fa8", marginBottom: 10, textTransform: "uppercase", letterSpacing: "0.05em" }}>Flavor performance breakdown</div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 12, marginBottom: 22 }}>
              {FLAVORS.map(f => {
                const s = stats[f];
                const badges = { "Cookie Dough": ["Best performer", "#34d399", "rgba(52,211,153,0.12)"], "Vanilla": ["Needs work", "#f87171", "rgba(248,113,113,0.12)"], "Cinnamon": ["Quick win", "#fbbf24", "rgba(251,191,36,0.12)"] };
                const [bl, bc, bbg] = badges[f];
                return (
                  <div key={f} style={{ background: "#1a1d27", borderRadius: 12, padding: 16, border: "1px solid rgba(255,255,255,0.06)" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
                      <span style={{ fontSize: 15, fontWeight: 600 }}>{f}</span>
                      <span style={{ fontSize: 10, padding: "3px 10px", borderRadius: 20, background: bbg, color: bc, border: `1px solid ${bc}33` }}>{bl}</span>
                    </div>
                    {[
                      ["Responses", s.n, "#f1f0f9"],
                      ["Buy: Yes / Maybe / No", `${s.yes} / ${s.maybe} / ${s.no}`, "#a78bfa"],
                      ["Taste: Excellent / Good / Neutral", `${s.excellent} / ${s.good} / ${s.neutral}`, "#2dd4bf"],
                      ["Texture issues", `${s.icy} icy, ${s.thick} thick`, s.icy > s.n * 0.4 ? "#f87171" : "#fbbf24"],
                    ].map(([l, v, c], i) => (
                      <div key={i} style={{ display: "flex", justifyContent: "space-between", marginBottom: 6, fontSize: 12 }}>
                        <span style={{ color: "#8b8fa8" }}>{l}</span>
                        <span style={{ fontWeight: 500, color: c }}>{v}</span>
                      </div>
                    ))}
                    <div style={{ marginTop: 8 }}>
                      <div style={{ fontSize: 11, color: "#8b8fa8", marginBottom: 4 }}>Purchase conversion</div>
                      <div style={{ height: 6, background: "#22263a", borderRadius: 3, overflow: "hidden" }}>
                        <div style={{ height: "100%", width: `${Math.round(s.yes / s.n * 100)}%`, background: "#34d399", borderRadius: 3 }} />
                      </div>
                      <div style={{ fontSize: 11, color: "#34d399", marginTop: 3, textAlign: "right" }}>{Math.round(s.yes / s.n * 100)}% conversion</div>
                    </div>
                  </div>
                );
              })}
            </div>

            {/* Review Table */}
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
              <div style={{ fontSize: 12, fontWeight: 600, color: "#8b8fa8", textTransform: "uppercase", letterSpacing: "0.05em" }}>Raw review data</div>
              <div style={{ display: "flex", gap: 6 }}>
                {["All", ...FLAVORS].map(f => (
                  <button key={f} onClick={() => setFlavorFilter(f)} style={{
                    fontSize: 11, padding: "4px 12px", borderRadius: 16, cursor: "pointer",
                    background: flavorFilter === f ? "rgba(124,111,238,0.2)" : "#22263a",
                    color: flavorFilter === f ? "#a78bfa" : "#8b8fa8",
                    border: flavorFilter === f ? "1px solid rgba(124,111,238,0.4)" : "1px solid rgba(255,255,255,0.06)"
                  }}>{f}</button>
                ))}
              </div>
            </div>
            <div style={{ overflowX: "auto", borderRadius: 10, border: "1px solid rgba(255,255,255,0.06)" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                <thead>
                  <tr style={{ background: "#22263a" }}>
                    {["#", "Name", "Flavor", "Taste", "Texture", "Sweetness", "Buy?", "Priority", "Comment"].map(h => (
                      <th key={h} style={{ padding: "10px 12px", textAlign: "left", color: "#8b8fa8", fontWeight: 600, fontSize: 11, borderBottom: "1px solid rgba(255,255,255,0.06)" }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {filtered.map((r, i) => (
                    <tr key={r.id} style={{ background: i % 2 === 0 ? "#1a1d27" : "#171a24" }}>
                      <td style={{ padding: "8px 12px", color: "#8b8fa8" }}>{r.id}</td>
                      <td style={{ padding: "8px 12px", fontWeight: 500 }}>{r.name}</td>
                      <td style={{ padding: "8px 12px", color: "#a78bfa" }}>{r.flavor}</td>
                      <td style={{ padding: "8px 12px", color: norm(r.taste) === "excellent" ? "#34d399" : norm(r.taste) === "good" ? "#2dd4bf" : "#fbbf24" }}>{r.taste || "—"}</td>
                      <td style={{ padding: "8px 12px", color: norm(r.texture) === "just right" ? "#34d399" : "#f87171" }}>{r.texture}</td>
                      <td style={{ padding: "8px 12px", color: norm(r.sweetness) === "just right" ? "#34d399" : "#fbbf24" }}>{r.sweetness}</td>
                      <td style={{ padding: "8px 12px", color: norm(r.buy) === "yes" ? "#34d399" : norm(r.buy) === "no" ? "#f87171" : "#fbbf24" }}>{r.buy}</td>
                      <td style={{ padding: "8px 12px", color: "#8b8fa8" }}>{r.matters}</td>
                      <td style={{ padding: "8px 12px", color: "#8b8fa8", maxWidth: 220, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r.comment || "—"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* ===== AI ANALYSIS ===== */}
        {tab === "ai" && (
          <div>
            {!aiResults && loading !== "ai" && (
              <div style={{ textAlign: "center", padding: "60px 20px" }}>
                <i className="ti ti-sparkles" style={{ fontSize: 48, color: "#7c6fee", display: "block", marginBottom: 16 }} />
                <div style={{ fontSize: 16, fontWeight: 500, marginBottom: 6 }}>Run AI sentiment analysis</div>
                <div style={{ fontSize: 13, color: "#8b8fa8", marginBottom: 20, maxWidth: 400, margin: "0 auto 20px" }}>
                  Claude will analyze each of the 61 reviews for sentiment, purchase intent, and generate actionable insights with before/after reframing.
                </div>
                <button onClick={runAiAnalysis} style={{
                  background: "#7c6fee", border: "none", borderRadius: 10, padding: "12px 28px",
                  color: "#fff", fontSize: 14, fontWeight: 500, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: 8
                }}>
                  <i className="ti ti-sparkles" /> Analyze all 61 reviews with AI
                </button>
              </div>
            )}

            {loading === "ai" && (
              <div style={{ textAlign: "center", padding: "60px 20px" }}>
                <div style={{ fontSize: 14, color: "#a78bfa", marginBottom: 12 }}>Analyzing reviews with Claude...</div>
                <div style={{ width: 300, height: 6, background: "#22263a", borderRadius: 3, margin: "0 auto", overflow: "hidden" }}>
                  <div style={{ height: "100%", width: `${progress}%`, background: "#7c6fee", borderRadius: 3, transition: "width 0.5s" }} />
                </div>
                <div style={{ fontSize: 12, color: "#8b8fa8", marginTop: 8 }}>{progress}% complete</div>
              </div>
            )}

            {aiResults && (
              <div>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 10, marginBottom: 20 }}>
                  {[
                    { l: "Avg sentiment score", v: (aiResults.reduce((a, r) => a + (r.score || 0), 0) / aiResults.length).toFixed(1), c: "#34d399" },
                    { l: "Positive reviews", v: aiResults.filter(r => r.sentiment === "Positive").length, c: "#34d399" },
                    { l: "Neutral reviews", v: aiResults.filter(r => r.sentiment === "Neutral").length, c: "#fbbf24" },
                    { l: "Negative reviews", v: aiResults.filter(r => r.sentiment === "Negative").length, c: "#f87171" },
                  ].map((k, i) => (
                    <div key={i} style={{ background: "#1a1d27", borderRadius: 10, padding: "14px 16px", border: "1px solid rgba(255,255,255,0.06)" }}>
                      <div style={{ fontSize: 11, color: "#8b8fa8", marginBottom: 4 }}>{k.l}</div>
                      <div style={{ fontSize: 22, fontWeight: 600, color: k.c }}>{k.v}</div>
                    </div>
                  ))}
                </div>

                <div style={{ overflowX: "auto", borderRadius: 10, border: "1px solid rgba(255,255,255,0.06)" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                    <thead>
                      <tr style={{ background: "#22263a" }}>
                        {["#", "Name", "Sentiment", "Score", "Intent", "Texture insight", "Sweetness insight", "Comment insight", "Recommendation"].map(h => (
                          <th key={h} style={{ padding: "10px 10px", textAlign: "left", color: "#8b8fa8", fontWeight: 600, fontSize: 11, borderBottom: "1px solid rgba(255,255,255,0.06)", whiteSpace: "nowrap" }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {aiResults.map((a, i) => (
                        <tr key={i} style={{ background: i % 2 === 0 ? "#1a1d27" : "#171a24" }}>
                          <td style={{ padding: "8px 10px", color: "#8b8fa8" }}>{i + 1}</td>
                          <td style={{ padding: "8px 10px", fontWeight: 500 }}>{a.name}</td>
                          <td style={{ padding: "8px 10px" }}>
                            <span style={{ fontSize: 11, padding: "2px 8px", borderRadius: 10, background: `${sentColor(a.sentiment)}18`, color: sentColor(a.sentiment) }}>{a.sentiment}</span>
                          </td>
                          <td style={{ padding: "8px 10px", fontWeight: 600, color: scoreColor(a.score) }}>{a.score}</td>
                          <td style={{ padding: "8px 10px", color: a.intent >= 60 ? "#34d399" : a.intent >= 40 ? "#fbbf24" : "#f87171" }}>{a.intent}%</td>
                          <td style={{ padding: "8px 10px", color: "#2dd4bf", maxWidth: 160, fontSize: 11 }}>{a.texture_insight}</td>
                          <td style={{ padding: "8px 10px", color: "#2dd4bf", maxWidth: 160, fontSize: 11 }}>{a.sweetness_insight}</td>
                          <td style={{ padding: "8px 10px", color: "#a78bfa", maxWidth: 180, fontSize: 11 }}>{a.comment_insight}</td>
                          <td style={{ padding: "8px 10px", color: "#8b8fa8", maxWidth: 200, fontSize: 11 }}>{a.recommendation}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </div>
        )}

        {/* ===== GTM MOTION ===== */}
        {tab === "gtm" && (
          <div>
            {!gtmResults && loading !== "gtm" && (
              <div style={{ textAlign: "center", padding: "60px 20px" }}>
                <i className="ti ti-rocket" style={{ fontSize: 48, color: "#2dd4bf", display: "block", marginBottom: 16 }} />
                <div style={{ fontSize: 16, fontWeight: 500, marginBottom: 6 }}>Generate GTM motion plan</div>
                <div style={{ fontSize: 13, color: "#8b8fa8", maxWidth: 440, margin: "0 auto 20px" }}>
                  Claude will synthesize all 61 reviews into a complete GTM strategy: positioning, ICP profiles, objection handling, launch messaging, and flavor go/no-go verdicts.
                </div>
                <button onClick={runGtmAnalysis} style={{
                  background: "linear-gradient(135deg,#7c6fee,#2dd4bf)", border: "none", borderRadius: 10, padding: "12px 28px",
                  color: "#fff", fontSize: 14, fontWeight: 500, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: 8
                }}>
                  <i className="ti ti-rocket" /> Generate GTM motion with AI
                </button>
              </div>
            )}

            {loading === "gtm" && (
              <div style={{ textAlign: "center", padding: "60px 20px" }}>
                <div style={{ fontSize: 14, color: "#2dd4bf", marginBottom: 8 }}>Synthesizing GTM strategy from 61 reviews...</div>
                <div style={{ fontSize: 12, color: "#8b8fa8" }}>Building positioning, ICP, objections, and launch plan</div>
              </div>
            )}

            {gtmResults && !gtmResults.error && (
              <div>
                {/* Positioning */}
                <div style={{ background: "linear-gradient(135deg,rgba(124,111,238,0.15),rgba(45,212,191,0.1))", borderRadius: 14, padding: 20, marginBottom: 16, border: "1px solid rgba(124,111,238,0.2)" }}>
                  <div style={{ fontSize: 11, fontWeight: 600, color: "#a78bfa", textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 6 }}>Positioning statement</div>
                  <div style={{ fontSize: 15, lineHeight: 1.7 }}>{gtmResults.positioning_statement}</div>
                </div>

                {/* Launch Messaging */}
                {gtmResults.launch_messaging && (
                  <div style={{ background: "#1a1d27", borderRadius: 12, padding: 20, marginBottom: 16, border: "1px solid rgba(255,255,255,0.06)", textAlign: "center" }}>
                    <div style={{ fontSize: 11, fontWeight: 600, color: "#8b8fa8", textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 10 }}>Launch messaging</div>
                    <div style={{ fontSize: 22, fontWeight: 600, marginBottom: 4 }}>{gtmResults.launch_messaging.headline}</div>
                    <div style={{ fontSize: 14, color: "#8b8fa8", marginBottom: 12 }}>{gtmResults.launch_messaging.subhead}</div>
                    <span style={{ background: "#7c6fee", color: "#fff", padding: "8px 24px", borderRadius: 8, fontSize: 13, fontWeight: 500 }}>{gtmResults.launch_messaging.cta}</span>
                  </div>
                )}

                {/* Value Props */}
                <div style={{ fontSize: 12, fontWeight: 600, color: "#8b8fa8", textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 10 }}>Value propositions</div>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 10, marginBottom: 20 }}>
                  {(gtmResults.value_props || []).map((v, i) => (
                    <div key={i} style={{ background: "#1a1d27", borderRadius: 10, padding: "14px 16px", border: "1px solid rgba(255,255,255,0.06)", borderLeft: "3px solid #7c6fee" }}>
                      <div style={{ fontSize: 13, lineHeight: 1.6 }}>{v}</div>
                    </div>
                  ))}
                </div>

                {/* ICP Profiles */}
                <div style={{ fontSize: 12, fontWeight: 600, color: "#8b8fa8", textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 10 }}>ICP segments</div>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(2, 1fr)", gap: 10, marginBottom: 20 }}>
                  {(gtmResults.icp_profiles || []).map((p, i) => (
                    <div key={i} style={{ background: "#1a1d27", borderRadius: 12, padding: 16, border: "1px solid rgba(255,255,255,0.06)" }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                        <span style={{ fontSize: 14, fontWeight: 600 }}>{p.segment}</span>
                        <span style={{ fontSize: 11, color: "#2dd4bf", background: "rgba(45,212,191,0.12)", padding: "2px 10px", borderRadius: 12 }}>{p.size_pct}% of testers</span>
                      </div>
                      <div style={{ fontSize: 12, color: "#8b8fa8", marginBottom: 10, lineHeight: 1.6 }}>{p.description}</div>
                      {[["Buy signal", p.buy_signal, "#34d399"], ["Objection", p.objection, "#f87171"], ["Pitch", p.messaging, "#a78bfa"]].map(([l, v, c]) => (
                        <div key={l} style={{ display: "flex", gap: 8, marginBottom: 5, fontSize: 12 }}>
                          <span style={{ color: "#5a5e75", minWidth: 72 }}>{l}</span>
                          <span style={{ color: c }}>{v}</span>
                        </div>
                      ))}
                    </div>
                  ))}
                </div>

                {/* Flavor Verdicts */}
                <div style={{ fontSize: 12, fontWeight: 600, color: "#8b8fa8", textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 10 }}>Flavor go / no-go</div>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 10, marginBottom: 20 }}>
                  {(gtmResults.flavor_recommendations || []).map((f, i) => {
                    const vc = f.verdict === "Launch" ? "#34d399" : f.verdict === "Iterate" ? "#fbbf24" : "#f87171";
                    return (
                      <div key={i} style={{ background: "#1a1d27", borderRadius: 12, padding: 16, border: "1px solid rgba(255,255,255,0.06)", borderTop: `3px solid ${vc}` }}>
                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                          <span style={{ fontSize: 14, fontWeight: 600 }}>{f.flavor}</span>
                          <span style={{ fontSize: 12, fontWeight: 600, color: vc, background: `${vc}18`, padding: "3px 12px", borderRadius: 8 }}>{f.verdict}</span>
                        </div>
                        <div style={{ fontSize: 12, color: "#8b8fa8", lineHeight: 1.6, marginBottom: 8 }}>{f.reason}</div>
                        <div style={{ fontSize: 12, color: "#2dd4bf" }}>Fix: {f.priority_fix}</div>
                      </div>
                    );
                  })}
                </div>

                {/* Objection Handling */}
                <div style={{ fontSize: 12, fontWeight: 600, color: "#8b8fa8", textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 10 }}>Objection handling playbook</div>
                {(gtmResults.objection_handling || []).map((o, i) => (
                  <div key={i} style={{ background: "#1a1d27", borderRadius: 10, padding: 14, marginBottom: 8, border: "1px solid rgba(255,255,255,0.06)" }}>
                    <div style={{ fontSize: 13, fontWeight: 600, color: "#f87171", marginBottom: 4 }}>"{o.objection}"</div>
                    <div style={{ fontSize: 12, color: "#34d399", marginBottom: 3 }}>→ {o.response}</div>
                    <div style={{ fontSize: 11, color: "#5a5e75" }}>Evidence: {o.evidence}</div>
                  </div>
                ))}

                {/* Risks + Targets */}
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginTop: 16 }}>
                  <div style={{ background: "#1a1d27", borderRadius: 12, padding: 16, border: "1px solid rgba(255,255,255,0.06)" }}>
                    <div style={{ fontSize: 12, fontWeight: 600, color: "#f87171", textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 10 }}>Risk factors</div>
                    {(gtmResults.risk_factors || []).map((r, i) => (
                      <div key={i} style={{ fontSize: 12, color: "#8b8fa8", marginBottom: 6, paddingLeft: 12, borderLeft: "2px solid #f8717144", lineHeight: 1.6 }}>{r}</div>
                    ))}
                  </div>
                  {gtmResults.key_metric_targets && (
                    <div style={{ background: "#1a1d27", borderRadius: 12, padding: 16, border: "1px solid rgba(255,255,255,0.06)" }}>
                      <div style={{ fontSize: 12, fontWeight: 600, color: "#34d399", textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 10 }}>Metric targets</div>
                      {Object.entries(gtmResults.key_metric_targets).map(([k, v]) => (
                        <div key={k} style={{ display: "flex", justifyContent: "space-between", marginBottom: 8, fontSize: 12 }}>
                          <span style={{ color: "#8b8fa8" }}>{k.replace(/_/g, " ")}</span>
                          <span style={{ fontWeight: 600, color: "#34d399" }}>{v}</span>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </div>
            )}
            {gtmResults?.error && <div style={{ color: "#f87171", textAlign: "center", padding: 40 }}>{gtmResults.error}</div>}
          </div>
        )}

        {/* ===== BEFORE / AFTER ===== */}
        {tab === "compare" && (
          <div>
            {!aiResults ? (
              <div style={{ textAlign: "center", padding: "60px 20px" }}>
                <i className="ti ti-arrows-exchange" style={{ fontSize: 48, color: "#fbbf24", display: "block", marginBottom: 16 }} />
                <div style={{ fontSize: 14, color: "#8b8fa8", marginBottom: 16 }}>Run AI Analysis first to see the before/after comparison</div>
                <button onClick={() => { setTab("ai"); }} style={{
                  background: "#7c6fee", border: "none", borderRadius: 10, padding: "10px 24px",
                  color: "#fff", fontSize: 13, cursor: "pointer"
                }}>Go to AI analysis</button>
              </div>
            ) : (
              <div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 0, marginBottom: 16 }}>
                  <div style={{ background: "rgba(251,191,36,0.08)", borderRadius: "10px 0 0 10px", padding: "10px 16px", textAlign: "center", border: "1px solid rgba(251,191,36,0.2)" }}>
                    <span style={{ fontSize: 12, fontWeight: 600, color: "#fbbf24" }}>BEFORE — Raw customer input</span>
                  </div>
                  <div style={{ background: "rgba(52,211,153,0.08)", borderRadius: "0 10px 10px 0", padding: "10px 16px", textAlign: "center", border: "1px solid rgba(52,211,153,0.2)" }}>
                    <span style={{ fontSize: 12, fontWeight: 600, color: "#34d399" }}>AFTER — AI-synthesized insight</span>
                  </div>
                </div>

                <div style={{ overflowX: "auto", borderRadius: 10, border: "1px solid rgba(255,255,255,0.06)" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                    <thead>
                      <tr style={{ background: "#22263a" }}>
                        <th style={{ padding: "10px", color: "#8b8fa8", fontSize: 11, fontWeight: 600, textAlign: "left", borderBottom: "1px solid rgba(255,255,255,0.06)" }}>#</th>
                        <th style={{ padding: "10px", color: "#8b8fa8", fontSize: 11, fontWeight: 600, textAlign: "left", borderBottom: "1px solid rgba(255,255,255,0.06)" }}>Name</th>
                        <th style={{ padding: "10px", color: "#fbbf24", fontSize: 11, fontWeight: 600, textAlign: "left", borderBottom: "1px solid rgba(255,255,255,0.06)" }}>Raw texture</th>
                        <th style={{ padding: "10px", color: "#34d399", fontSize: 11, fontWeight: 600, textAlign: "left", borderBottom: "1px solid rgba(255,255,255,0.06)" }}>AI texture insight</th>
                        <th style={{ padding: "10px", color: "#fbbf24", fontSize: 11, fontWeight: 600, textAlign: "left", borderBottom: "1px solid rgba(255,255,255,0.06)" }}>Raw sweetness</th>
                        <th style={{ padding: "10px", color: "#34d399", fontSize: 11, fontWeight: 600, textAlign: "left", borderBottom: "1px solid rgba(255,255,255,0.06)" }}>AI sweetness insight</th>
                        <th style={{ padding: "10px", color: "#fbbf24", fontSize: 11, fontWeight: 600, textAlign: "left", borderBottom: "1px solid rgba(255,255,255,0.06)" }}>Raw comment</th>
                        <th style={{ padding: "10px", color: "#34d399", fontSize: 11, fontWeight: 600, textAlign: "left", borderBottom: "1px solid rgba(255,255,255,0.06)" }}>AI comment insight</th>
                      </tr>
                    </thead>
                    <tbody>
                      {aiResults.map((a, i) => {
                        const r = REVIEWS[i] || {};
                        return (
                          <tr key={i} style={{ background: i % 2 === 0 ? "#1a1d27" : "#171a24" }}>
                            <td style={{ padding: "8px 10px", color: "#5a5e75" }}>{i + 1}</td>
                            <td style={{ padding: "8px 10px", fontWeight: 500 }}>{a.name}</td>
                            <td style={{ padding: "8px 10px", color: "#fbbf24" }}>{r.texture || "—"}</td>
                            <td style={{ padding: "8px 10px", color: "#34d399", fontSize: 11 }}>{a.texture_insight}</td>
                            <td style={{ padding: "8px 10px", color: "#fbbf24" }}>{r.sweetness || "—"}</td>
                            <td style={{ padding: "8px 10px", color: "#34d399", fontSize: 11 }}>{a.sweetness_insight}</td>
                            <td style={{ padding: "8px 10px", color: "#fbbf24", maxWidth: 140, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r.comment || "—"}</td>
                            <td style={{ padding: "8px 10px", color: "#34d399", fontSize: 11 }}>{a.comment_insight}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
