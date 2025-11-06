import React, { useMemo, useState } from "react";
import { motion } from "framer-motion";
import { FileText, Upload, Wand2, CheckCircle2, Info, Play, Image as ImageIcon, FileDown } from "lucide-react";
import { Document as DocxDocument, Packer, Paragraph, HeadingLevel, Table, TableRow, TableCell, WidthType, ImageRun } from "docx";

/**
 * Unicode safety: use explicit escapes for macron characters (e.g., \u0101 for ā)
 * and avoid smart quotes/emdashes. This prevents toolchains from throwing
 * "Expecting Unicode escape sequence \\uXXXX".
 */

// ---------------------------------------------------------------------------------
// Hoisted ICMP dictionaries (to avoid re-allocating on every render)
const HCC_ICMPS: { name: string; patterns: RegExp[] }[] = [
  { name: "Peacocke ICMP", patterns: [/\bpeacocke\b|\bpeacocks\b/] },
  { name: "Rotokauri ICMP", patterns: [/\brotokauri\b/] },
  { name: "Te Rapa ICMP", patterns: [/te\s?rapa|\bpukete\b|\bnorthgate\b/] },
  { name: "Rototuna ICMP", patterns: [/\brototuna\b|\bflagstaff\b|\bchartwell\b/] },
  { name: "Ruakura ICMP", patterns: [/\bruakura\b|\bhillcrest\b|\bsilverdale\b|\buniversity\b/] },
  { name: "Waitawhiriwhiri ICMP", patterns: [/waitawhiriwhiri|beerescourt|forest\s*lakes?|frankton|clarkin|bryant/] },
  { name: "Mangakotukutuku ICMP", patterns: [/mangakotukutuku|glenview|melville|fitzroy|bader|kahikatea|ohaupo/] },
];

const WDC_ICMPS: { name: string; patterns: RegExp[] }[] = [
  { name: "Ng\u0101ruaw\u0101hia ICMP", patterns: [/ngaruawahia|ng\u0101ruaw\u0101hia|hopuhopu/] },
  { name: "Huntly ICMP", patterns: [/\bhuntly\b|\brahuipokeka\b/] },
  { name: "Te Kauwhata ICMP", patterns: [/te\s*kauwhata|waerenga|meremere/] },
];

function normalizeForMatch(input: string): string {
  // strip diacritics and lower-case for robust matching
  return (input || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/\p{Diacritic}/gu, "");
}

function inferICMP(location: string, councilName: string): string {
  const loc = normalizeForMatch(location);
  if (councilName === "Hamilton City Council") {
    const hit = HCC_ICMPS.find((area) => area.patterns.some((p) => p.test(loc)));
    return hit ? hit.name : "Hamilton City ICMP (area to confirm)";
  } else {
    const hit = WDC_ICMPS.find((area) => area.patterns.some((p) => p.test(loc)));
    return hit ? hit.name : "Waikato District ICMP (area to confirm)";
  }
}

function inferICMPMatches(location: string, councilName: string): string[] {
  const loc = normalizeForMatch(location);
  const dict = councilName === "Hamilton City Council" ? HCC_ICMPS : WDC_ICMPS;
  const matches = dict.filter((area) => area.patterns.some((p) => p.test(loc))).map((a) => a.name);
  if (matches.length === 0) {
    return [councilName === "Hamilton City Council" ? "Hamilton City ICMP (area to confirm)" : "Waikato District ICMP (area to confirm)"];
  }
  return matches;
}

export default function CIAPrototype() {
  // ---------------------------------------------------------------------------------
  // State
  const [council, setCouncil] = useState("Hamilton City Council");
  const [includeHPMO, setIncludeHPMO] = useState(true);
  const [projectName, setProjectName] = useState("Te Awa Industrial Upgrade - Stage 2");
  const [projectLocation, setProjectLocation] = useState("");
  const [inferredICMP, setInferredICMP] = useState<string>("(not set)");
  const [icmpOptions, setIcmpOptions] = useState<string[]>([]);
  const [showIcmpModal, setShowIcmpModal] = useState(false);
  const [icmpTemp, setIcmpTemp] = useState<string>("(not set)");
  const [status] = useState("Ready to analyse");
  const [testResult, setTestResult] = useState<string | null>(null);
  const [figureGallery, setFigureGallery] = useState(
    // placeholder gallery; real build will parse uploads for figures
    Array.from({ length: 8 }).map((_, i) => ({
      id: `fig-${i + 1}`,
      caption: `Figure ${i + 1}: Placeholder diagram/map`,
      // tiny transparent png
      dataUrl:
        "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAAFElEQVQYV2NkYGBgYGBg+M9ABQAKCAIBz1b5AwAAAABJRU5ErkJggg==",
      selected: i < 3,
    }))
  );

  // ---------------------------------------------------------------------------------
  // Handlers
  function handleAutoDetectICMP() {
    const matches = inferICMPMatches(projectLocation, council);
    if (matches.length <= 1) {
      setInferredICMP(
        matches[0] || (council === "Hamilton City Council" ? "Hamilton City ICMP (area to confirm)" : "Waikato District ICMP (area to confirm)")
      );
      setIcmpOptions([]);
      setShowIcmpModal(false);
    } else {
      setIcmpOptions(matches);
      setIcmpTemp(matches[0] || "(not set)");
      setShowIcmpModal(true);
    }
  }

  // ---------------------------------------------------------------------------------
  // Data model
  type TriggerSpec = {
    metrics: string[];
    baselines: string;
    thresholds: string[];
    actions: string[];
    reporting: string;
  };

  type Effects = {
    cultural: string[];
    social: string[];
    environmental: string[];
    spiritual: string[];
  };

  type Finding = {
    category: string; // wai, whenua, whakapapa, wh\u0101nau, mauri, wairua
    issue: string;
    effects: Effects;
    mitigations: string[];
    recommendations: string[];
    triggers: TriggerSpec;
    policyLinks: string[];
    consentClauses: string[];
  };

  const sampleFindings: Finding[] = [
    {
      category: "wai",
      issue:
        "Potential degradation of mauri and clarity in tributary due to earthworks sediment discharges",
      effects: {
        cultural: [
          "Mauri of the awa diminished; disruption to mahinga kai practices",
          "Reduced ability for wh\u0101nau to gather kai safely",
        ],
        social: ["Community concern over river health; trust impacts"],
        environmental: [
          "Elevated NTU and TSS; smothering of habitat; fish passage stress",
        ],
        spiritual: [
          "Tapu/noa balance affected where discharges occur near w\u0101hi tapu",
        ],
      },
      mitigations: [
        "Install staged sediment retention ponds sized to Hamilton District Plan GD05 guidance; monitor turbidity (NTU) daily.",
        "No instream works during tuna migration periods; implement 25 m riparian buffer in sensitive reaches.",
      ],
      recommendations: [
        "Co-design mahinga kai monitoring with mana whenua; quarterly w\u0101nanga to review data and adaptive actions.",
        "Embed Te Ture Whaimana vision-objectives as assessment criteria in contractor EMS (Environmental Management System).",
      ],
      triggers: {
        metrics: [
          "NTU (turbidity)",
          "TSS (mg/L)",
          "E. coli (cfu/100 mL)",
          "Visual clarity (m)",
          "Mahinga kai presence/abundance (tuna/\u012Bnanga)",
        ],
        baselines:
          "Establish 4-week pre-works baseline for NTU/clarity and cultural health index (CHI) with mana whenua.",
        thresholds: [
          "NTU > baseline + 25% for >24h",
          "Clarity < 1.6 m during fine weather",
          "Any exceedance at mahinga kai sites",
        ],
        actions: [
          "Stop high-risk works; inspect ESCP; deploy additional treatment within 24h",
          "Notify mana whenua and Council within 1 working day",
          "Hold hui within 5 working days to agree corrective actions",
        ],
        reporting:
          "Quarterly report + dashboard; immediate incident reports when thresholds tripped.",
      },
      policyLinks: [
        "Te Ture Whaimana - Vision and Objective 1 (health and wellbeing of the Waikato River)",
        "Tai Tumu, Tai Pari, Tai Ao EMP - Wai: water quality, mahinga kai protection",
        "Hamilton District Plan - 25.14 Infrastructure; erosion/sediment control standards",
      ],
      consentClauses: [
        "Prior to works, the consent holder must submit an Erosion and Sediment Control Plan (ESCP) prepared by a suitably qualified person, demonstrating compliance with GD05 and avoiding instream works during identified migration windows for tuna/\u012Bnanga.",
        "Establish a Mauri Monitoring Programme co-developed with mana whenua that sets baseline and trigger levels (including NTU and clarity), provides for mahinga kai assessments, and requires adaptive responses within 10 working days if triggers are exceeded.",
      ],
    },
    {
      category: "whenua",
      issue:
        "Loss of topsoil and disturbance of known urup\u0101 risk area within 200 m of works",
      effects: {
        cultural: [
          "Risk to w\u0101hi tapu/w\u0101hi t\u016Bpuna; mamae if disturbance occurs",
        ],
        social: ["Project delays and conflict if discovery process unclear"],
        environmental: ["Erosion risk; reduced soil productivity if not salvaged"],
        spiritual: ["Tapu breach potential requiring tikanga responses"],
      },
      mitigations: [
        "Cultural discovery protocol with immediate stop-work and notification process.",
        "Topsoil salvage and reuse plan to support revegetation with taonga species.",
      ],
      recommendations: [
        "Archaeological Authority (HNZPT) pre-works; mana whenua monitors present during initial ground-breaking.",
        "GIS mapping layer for w\u0101hi tapu/w\u0101hi t\u016Bpuna integrated into contractor inductions.",
      ],
      triggers: {
        metrics: ["Protocol drills completed", "Monitor hours on-site", "Incidents recorded"],
        baselines: "Zero harm baseline (no unauthorised ground disturbance).",
        thresholds: ["Any suspected k\u014Diwi or taonga triggers stop-work"],
        actions: [
          "Immediate stop-work; protect area; notify mana whenua, HNZPT, Police (if k\u014Diwi)",
          "Undertake tikanga-led process; update methodology before resuming",
        ],
        reporting:
          "Incident log shared within 24h; monthly summary including training and inductions.",
      },
      policyLinks: [
        "Tai Tumu, Tai Pari, Tai Ao EMP - Whenua: protection of w\u0101hi tapu, soils, and landscapes",
        "Hamilton/Waikato District Plan - Heritage and Archaeology provisions",
      ],
      consentClauses: [
        "Implement a Cultural Discovery Protocol approved by mana whenua prior to commencement; all staff to be inducted and protocol kept onsite at all times.",
        "Require mana whenua cultural monitors to be present during initial stripping; consent holder to fund participation and reporting.",
      ],
    },
    {
      category: "whakapapa",
      issue:
        "Fragmentation of ecological corridors reducing connectivity for taonga species",
      effects: {
        cultural: [
          "Disruption to whakapapa relationships among species and habitats",
        ],
        social: [
          "Loss of local amenity and learning opportunities for rangatahi",
        ],
        environmental: [
          "Barrier to fish passage; edge effects increase predators/weeds",
        ],
        spiritual: ["Diminished wairua of place if connections severed"],
      },
      mitigations: [
        "Design wildlife-friendly culverts and fish passage; stage works to maintain connectivity.",
      ],
      recommendations: [
        "Planting palette guided by whakapapa of place (locally-sourced eco-sourced taonga species); 3-year establishment and pest control.",
      ],
      triggers: {
        metrics: [
          "Fish passage scores (NIWA tool)",
          "Survival of plantings (%)",
          "Predator trap-catch",
        ],
        baselines: "Pre-works fish passage survey and habitat mapping.",
        thresholds: ["Fish passage score < baseline", "Plant survival <85%"],
        actions: [
          "Remediate culverts; replace failed plantings; intensify pest control",
        ],
        reporting: "Six-monthly ecological report + w\u0101nanga walkthrough.",
      },
      policyLinks: [
        "Te Ture Whaimana - enhancement of ecological integrity",
        "Tai Tumu, Tai Pari, Tai Ao EMP - Whakapapa: intergenerational stewardship",
      ],
      consentClauses: [
        "Fish passage to meet NIWA fish passage assessment tool thresholds; as-built certification prior to operation.",
      ],
    },
    {
      category: "wh\u0101nau",
      issue:
        "Construction traffic and noise affecting marae access, tangihanga, and daily wh\u0101nau life",
      effects: {
        cultural: [
          "Disruption to marae protocols and ability to host kaupapa including tangihanga",
        ],
        social: [
          "Increased stress; reduced community cohesion if engagement is weak",
        ],
        environmental: ["Dust and vibration affecting nearby sensitive receivers"],
        spiritual: ["Disturbance to wairua during significant wh\u0101nau events"],
      },
      mitigations: [
        "Traffic Management Plan (TMP) with marae input; avoid peak event times",
        "Construction Noise and Vibration Management Plan; onsite dust suppression",
      ],
      recommendations: [
        "Co-design communications plan with mana whenua; 2-week lookahead notices",
        "Identify protected access windows around known marae events/tangihanga",
      ],
      triggers: {
        metrics: ["LAeq dB", "Number of complaints", "Access block incidents"],
        baselines: "Pre-works ambient noise survey and access mapping with wh\u0101nau.",
        thresholds: [">2 substantiated access incidents/month", "Noise exceeds plan limits"],
        actions: [
          "Adjust work hours/routing; deploy additional acoustic barriers",
          "H\u016Btu w\u0101nanga within 5 working days to agree changes",
        ],
        reporting: "Monthly community report; real-time hotline with log shared to mana whenua.",
      },
      policyLinks: [
        "Tai Tumu, Tai Pari, Tai Ao EMP - Wh\u0101nau and participation",
        "Applicable District Plan - Noise/traffic rules and engagement requirements",
      ],
      consentClauses: [
        "Prepare and implement a TMP and Communications Plan co-designed with mana whenua, including protected access windows for marae and mechanisms for event-time pauses.",
        "Maintain a dedicated contact line and incident log accessible to mana whenua; implement corrective actions within 5 working days.",
      ],
    },
    {
      category: "mauri",
      issue: "Residual effects risk during storm events despite controls",
      effects: {
        cultural: ["Perceived degradation of mauri during heavy rain events"],
        social: ["Community anxiety following spill/overflow rumours"],
        environmental: ["Pulse loads of sediments and contaminants"],
        spiritual: ["Loss of balance (tapu/noa) when incidents occur"],
      },
      mitigations: [
        "Adaptive management triggers linked to rainfall intensity thresholds",
        "Contingency spill kits and overflow prevention measures",
      ],
      recommendations: [
        "Integrate mauri indicators in dashboard with traffic-light triggers",
        "Run post-event w\u0101nanga to agree remediation and learning",
      ],
      triggers: {
        metrics: ["Rainfall (mm/hr)", "NTU spikes", "Incident count"],
        baselines: "Event-based baseline using first-flush data",
        thresholds: [">20 mm/hr with NTU > baseline + 40%", "Any overflow"],
        actions: ["Suspend exposed works; stand-up response team; notify within 24h"],
        reporting: "Event summary within 5 days; quarterly trend analysis with mana whenua.",
      },
      policyLinks: [
        "Te Ture Whaimana - maintaining and enhancing the mauri of the Waikato River",
        "Tai Tumu, Tai Pari, Tai Ao EMP - Mauri",
      ],
      consentClauses: [
        "Adopt an Adaptive Management Plan with rainfall-linked triggers and defined corrective actions; co-develop with mana whenua and submit prior to works.",
      ],
    },
    {
      category: "wairua",
      issue:
        "Loss of sense of place at w\u0101hi t\u016Bpuna vista and culturally sensitive viewshafts",
      effects: {
        cultural: ["Erosion of identity where viewshafts are compromised"],
        social: ["Reduced pride and connection to place"],
        environmental: ["Visual amenity effects; vegetation structure changes"],
        spiritual: ["Disruption to wairua associated with the site"],
      },
      mitigations: [
        "Cultural design review panel with mana whenua; protect key sightlines",
      ],
      recommendations: [
        "Develop a viewshaft protection plan and culturally anchored design palette",
      ],
      triggers: {
        metrics: ["Design gate approvals", "Non-conformance count"],
        baselines: "Pre-works photo-simulations agreed with mana whenua",
        thresholds: ["Any deviation from agreed sightline envelope"],
        actions: ["Iterate design to restore sightlines; additional planting/screening"],
        reporting: "Design review minutes; pre/post photo-comparisons filed with CIA updates.",
      },
      policyLinks: [
        "Tai Tumu, Tai Pari, Tai Ao EMP - Wairua and landscapes",
        "District Plan - Landscape/amenity objectives and policies",
      ],
      consentClauses: [
        "Establish a Cultural Design Review Panel with decision checkpoints at concept, developed, and pre-construction stages; implement agreed viewshaft protection measures.",
      ],
    },
  ];

  // ---------------------------------------------------------------------------------
  // Narrative builders (Standard depth A1)
  function buildStandardManaWhenua(findings: Finding[]): string {
    const sections = [
      "Executive Summary",
      "Background and Whakapapa",
      "Methodology (Kaupapa M\u0101ori, W\u0101nanga)",
      "Categories Assessment",
      "ICMP and Policy Alignment",
      "Cultural Monitoring Programme",
      "Consent Conditions and Next Steps",
    ];
    const intro = `# CIA - Mana Whenua Narrative (Standard)\n\n## Project\n${projectName}\n\n## Whakatau\u0101k\u012B / Context\nKo te mana o te awa me te whenua te t\u016B\u0101papa. This kaupapa recognises our relationship to wai, whenua, and all living systems. We have assessed the technical reports and translated key matters into plain language for wh\u0101nau.`;

    const cats = findings
      .map(
        (f, i) => `\n### ${i + 1}. ${f.category.toUpperCase()}\n**Ng\u0101 take / Issue:** ${f.issue}\n\n**Ng\u0101 whakatika / Mitigations:**\n- ${f.mitigations.join("\n- ")}\n\n**Ng\u0101 t\u016Btohunga / Recommendations:**\n- ${f.recommendations.join("\n- ")}\n\n**Monitoring triggers (plain):** ${f.triggers.metrics.join(", ")}\n`
      )
      .join("\n");

    const align = `\n## Te Ture Whaimana and Tai Tumu, Tai Pari, Tai Ao alignment\nWe checked the mahi against the Vision and Objectives of Te Ture Whaimana, the Waikato-Tainui EMP (Tai Tumu, Tai Pari, Tai Ao), and the relevant District Plan. The project is connected to: **${inferredICMP}**.`;

    const outro = `\n## Tikanga and Participation\n- Mana whenua monitors present at ground-break.\n- W\u0101nanga-a-rohe, quarterly, to review monitoring and adapt.\n- Cultural discovery protocol: stop-work, karakia, k\u014Drero, record.\n\n## Ask to Council / Developer\nAdopt the consent conditions listed and fund the co-governed monitoring programme. Partner early on planting design and mahinga kai.`;

    return [intro, `\n## Sections\n- ${sections.join("\n- ")}`, `\n## Categories - Issues, Mitigations, Recommendations\n${cats}`, align, outro].join("\n\n");
  }

  function buildStandardCouncil(findings: Finding[]): string {
    const scope = `## Assessment Scope\n- Technical reports reviewed: EMPs, CMPs, ESCPs, ecology/archaeology/hydrology.\n- Policy instruments: Te Ture Whaimana; Tai Tumu, Tai Pari, Tai Ao EMP; ${
      council === "Hamilton City Council"
        ? "He Pou Manawa Ora; Hamilton District Plan"
        : "Waikato District Plan"
    }.\n- ICMP area: **${inferredICMP}**.`;

    const matrix = findings
      .map(
        (f, i) => `${i + 1}. ${f.category} | ${f.issue}\n   - Mitigation: ${f.mitigations.join("; ")}\n   - Recommendation: ${f.recommendations.join("; ")}\n   - Policy: ${f.policyLinks.join("; ")}`
      )
      .join("\n\n");

    const conditions = findings.flatMap((f) => f.consentClauses).map((c, i) => `${i + 1}. ${c}`).join("\n");

    return `# CIA - Council/Developer Narrative (Standard)\n\n## Project\n${projectName}\n\n${scope}\n\n## Category Matrix (Issues -> Mitigation -> Recommendation -> Policy Link)\n${matrix}\n\n## Proposed Consent Conditions (extract)\n${conditions}\n\n## Monitoring and Adaptive Management\n- Baseline: clarity/NTU, macroinvertebrates, mahinga kai presence.\n- Triggers: set per-site with mana whenua; actions within 10 working days.\n- Reporting: quarterly hui plus written report for Council and mana whenua.`;
  }

  const manaWhenuaNarrative = useMemo(() => buildStandardManaWhenua(sampleFindings), [projectName, inferredICMP]);
  const councilNarrative = useMemo(() => buildStandardCouncil(sampleFindings), [projectName, council, inferredICMP]);
  const selectedFigures = useMemo(() => figureGallery.filter((g) => g.selected), [figureGallery]);

  // ---------------------------------------------------------------------------------
  // Monitoring helpers
  type MonitoringRow = { phase: string; focus: string; role: string; frequency: string };

  const baseMonitoringPlanRows: MonitoringRow[] = [
    { phase: "Pre-construction", focus: "Baseline mauri / mudfish / habitat surveys", role: "Attend site walkover; confirm w\u0101hi tapu avoidance; record baseline (photos/notes)", frequency: "One-off" },
    { phase: "During earthworks", focus: "Sediment discharge (NTU/TSS), discovery protocol readiness", role: "Onsite checks; hold stop-work if tikanga/cultural risk; verify ESCP field controls", frequency: "Daily / storm-event" },
    { phase: "Ecology", focus: "Mudfish / fish passage; planting survival", role: "Guide methods using tikanga; co-observe with ecologist; confirm safe handling/release", frequency: "Monthly / seasonal" },
    { phase: "Close-out", focus: "Verify consent conditions delivered; cultural outcomes", role: "Final site check; sign-off report to Council and mana whenua", frequency: "One-off" },
  ];

  function deriveMonitoringRows(findings: Finding[], councilName: string, includeHPMOFlag: boolean): MonitoringRow[] {
    const rows = [...baseMonitoringPlanRows];
    const text = findings.map((f) => (f.issue + " " + f.recommendations.join(" ")).toLowerCase()).join(" ");
    const hasMudfish = /mudfish|\u012Bnanga/.test(text);
    if (hasMudfish && !rows.find((r) => r.focus.toLowerCase().includes("mudfish"))) {
      rows.splice(1, 0, {
        phase: "Pre-construction",
        focus: "Targeted mudfish presence/absence at drains/wetlands",
        role: "Assist ecologist; apply tikanga for handling; confirm relocation plan if needed",
        frequency: "One-off",
      });
    }
    if (councilName === "Hamilton City Council") {
      rows.push({
        phase: "During earthworks",
        focus: "He Pou Manawa Ora engagement checkpoint",
        role: "Attend engagement checkpoint; confirm cultural measures are active",
        frequency: includeHPMOFlag ? "At each stage-gate" : "(disabled)",
      });
    } else {
      rows.push({
        phase: "During earthworks",
        focus: "Waikato District Plan noise/access checks near marae",
        role: "Check access windows and noise limits with wh\u0101nau",
        frequency: "Weekly",
      });
    }
    return rows;
  }

  // ---------------------------------------------------------------------------------
  // Exports (C1: separate reports)
  async function exportMonitoringDocx(project: string, rows: MonitoringRow[]) {
    try {
      const doc = new DocxDocument({
        sections: [
          {
            properties: {},
            children: [
              new Paragraph({ text: "Cultural Monitoring Programme", heading: HeadingLevel.TITLE }),
              new Paragraph({ text: `Project: ${project}` }),
              new Paragraph({ text: "Timeline & Tasks", heading: HeadingLevel.HEADING_2 }),
              new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: [
                  new TableRow({
                    children: [
                      new TableCell({ children: [new Paragraph("Phase")] }),
                      new TableCell({ children: [new Paragraph("Monitoring focus")] }),
                      new TableCell({ children: [new Paragraph("Role of Cultural Monitor")] }),
                      new TableCell({ children: [new Paragraph("Frequency")] }),
                    ],
                  }),
                  ...rows.map(
                    (r) =>
                      new TableRow({
                        children: [
                          new TableCell({ children: [new Paragraph(r.phase)] }),
                          new TableCell({ children: [new Paragraph(r.focus)] }),
                          new TableCell({ children: [new Paragraph(r.role)] }),
                          new TableCell({ children: [new Paragraph(r.frequency)] }),
                        ],
                      })
                  ),
                ],
              }),
              new Paragraph({ text: "Job Description", heading: HeadingLevel.HEADING_2 }),
              new Paragraph({ text: "• Represent mana whenua onsite and act as kaitiaki of w\u0101hi tapu, wai and whenua." }),
              new Paragraph({ text: "• Hold stop-work authority when tikanga or cultural risk is observed." }),
              new Paragraph({ text: "• Record observations (photos + narrative) into the CIA dashboard." }),
              new Paragraph({ text: "• Attend toolbox talks and ensure contractors understand cultural protocols." }),
              new Paragraph({ text: "• Escalate incident triggers to Project Manager and Council." }),
            ],
          },
        ],
      });

      const blob = await Packer.toBlob(doc);
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `Cultural_Monitoring_Programme_${project.replace(/\s+/g, "_")}.docx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (e) {
      console.error(e);
      alert("DOCX export failed. Check console for details.");
    }
  }

  async function exportNarrativeDocx(title: string, body: string, figures: { caption: string; dataUrl: string }[]) {
    try {
      const figureBlocks = figures.map((f) => {
        const base64 = f.dataUrl.split(",")[1] || "";
        const buf = Uint8Array.from(atob(base64), (c) => c.charCodeAt(0));
        return [
          new Paragraph({ text: f.caption, heading: HeadingLevel.HEADING_3 }),
          new Paragraph({ children: [new ImageRun({ data: buf, transformation: { width: 360, height: 240 } })] }),
        ];
      });

      const doc = new DocxDocument({
        sections: [
          {
            properties: {},
            children: [
              new Paragraph({ text: title, heading: HeadingLevel.TITLE }),
              ...body.split("\n").map((line) => new Paragraph({ text: line })),
              new Paragraph({ text: "Selected Figures (Inline)", heading: HeadingLevel.HEADING_2 }),
              ...figureBlocks.flat(),
            ],
          },
        ],
      });
      const blob = await Packer.toBlob(doc);
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${title.replace(/\s+/g, "_")}.docx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (e) {
      console.error(e);
      alert("DOCX export failed. Check console for details.");
    }
  }

  // ---------------------------------------------------------------------------------
  // Self-checks (acts like lightweight test cases you can run in UI)
  function runSelfChecks(): string {
    try {
      // required fields
      for (const f of sampleFindings) {
        if (!f.category || !f.issue) throw new Error("Missing category or issue");
        ["effects", "mitigations", "recommendations", "triggers", "policyLinks", "consentClauses"].forEach((k) => {
          // @ts-ignore
          if (f[k] == null) throw new Error(`Missing field: ${k}`);
        });
      }
      // unicode spot-checks
      const mw = buildStandardManaWhenua(sampleFindings);
      if (!/Whakatau\\u0101k\\u012B/.test(mw)) throw new Error("Unicode escape missing");
      // council branch
      const expectedBranch = council === "Hamilton City Council" ? "He Pou Manawa Ora" : "Waikato District Plan";
      const cn = buildStandardCouncil(sampleFindings);
      if (!cn.includes(expectedBranch)) throw new Error("Council branch failed");
      // categories present
      const categories = new Set(sampleFindings.map((f) => f.category));
      const expected = ["wai", "whenua", "whakapapa", "wh\u0101nau", "mauri", "wairua"];
      for (const c of expected) if (!categories.has(c)) throw new Error(`Missing category: ${c}`);
      // figures usable
      const sel = figureGallery.filter((g) => g.selected);
      if (sel.length === 0) throw new Error("No figures selected");
      // ICMP inference for extra areas
      const t1 = inferICMP("Waitawhiriwhiri", "Hamilton City Council");
      if (!/Waitawhiriwhiri ICMP/.test(t1)) throw new Error("ICMP inference failed for Waitawhiriwhiri");
      const t2 = inferICMP("Mangakotukutuku", "Hamilton City Council");
      if (!/Mangakotukutuku ICMP/.test(t2)) throw new Error("ICMP inference failed for Mangakotukutuku");
      return "All self-checks passed";
    } catch (e: any) {
      return `Test failure: ${e.message || String(e)}`;
    }
  }

  // ---------------------------------------------------------------------------------
  // UI
  return (
    <div className="min-h-screen w-full bg-white text-gray-900">
      <div className="mx-auto max-w-6xl px-6 py-10">
        <motion.h1 initial={{ opacity: 0, y: -10 }} animate={{ opacity: 1, y: 0 }} transition={{ duration: 0.5 }} className="text-3xl font-bold tracking-tight">
          Cultural Impact Assessment (CIA) - App Prototype
        </motion.h1>

        <p className="mt-2 text-sm text-gray-600">
          Upload your technical documents (EMPs, CMPs, ecology, hydrology, archaeology). Select council. The AI analyses, maps findings to Categories (wai, whenua, whakapapa, wh\u0101nau, mauri, wairua), checks against Te Ture Whaimana and Tai Tumu, Tai Pari, Tai Ao, and drafts two parallel narratives plus consent conditions.
        </p>

        {/* Controls */}
        <div className="mt-8 grid grid-cols-1 gap-4 md:grid-cols-3">
          {/* Project */}
          <div className="rounded-2xl border p-4 shadow-sm">
            <div className="flex items-center gap-2">
              <FileText className="h-5 w-5" />
              <div className="font-semibold">Project</div>
            </div>
            <input className="mt-3 w-full rounded-xl border px-3 py-2" value={projectName} onChange={(e) => setProjectName(e.target.value)} placeholder="Enter project name" />
            <input className="mt-3 w-full rounded-xl border px-3 py-2" value={projectLocation} onChange={(e) => setProjectLocation(e.target.value)} placeholder="Enter project location (suburb/area or address)" />
            <div className="mt-2 text-xs text-gray-600">
              Inferred ICMP: <span className="font-medium">{inferredICMP}</span>
            </div>
            <button className="mt-2 rounded-xl border px-3 py-1 text-sm" onClick={handleAutoDetectICMP}>
              Auto-detect ICMP from location
            </button>
          </div>

          {/* Council */}
          <div className="rounded-2xl border p-4 shadow-sm">
            <div className="flex items-center gap-2">
              <Info className="h-5 w-5" />
              <div className="font-semibold">Council</div>
            </div>
            <select className="mt-3 w-full rounded-xl border px-3 py-2" value={council} onChange={(e) => { const val = e.target.value; setCouncil(val); setIncludeHPMO(val === "Hamilton City Council"); }}>
              <option>Hamilton City Council</option>
              <option>Waikato District Council</option>
            </select>
            <div className="mt-3 flex items-center justify-between gap-3 rounded-xl bg-gray-50 p-3">
              <span>He Pou Manawa Ora</span>
              <input type="checkbox" checked={includeHPMO} onChange={(e) => setIncludeHPMO(e.target.checked)} disabled={council !== "Hamilton City Council"} />
            </div>
          </div>

          {/* Status */}
          <div className="rounded-2xl border p-4 shadow-sm">
            <div className="flex items-center gap-2">
              <Wand2 className="h-5 w-5" />
              <div className="font-semibold">Status</div>
            </div>
            <div className="mt-3 rounded-xl bg-green-50 p-3 text-sm text-green-800">
              <div className="flex items-center gap-2">
                <CheckCircle2 className="h-4 w-4" /> {status}
              </div>
            </div>
          </div>
        </div>

        {/* ICMP chooser modal */}
        {showIcmpModal && (
          <div className="fixed inset-0 z-50 flex items-center justify-center">
            <div className="absolute inset-0 bg-black/40" onClick={() => { setShowIcmpModal(false); setIcmpOptions([]); }} />
            <motion.div initial={{ opacity: 0, scale: 0.96 }} animate={{ opacity: 1, scale: 1 }} transition={{ duration: 0.2 }} className="relative z-10 w-full max-w-lg rounded-2xl bg-white p-6 shadow-2xl">
              <div className="text-lg font-semibold">Multiple ICMP areas matched</div>
              <p className="mt-1 text-sm text-gray-600">Choose the correct ICMP for <span className="font-medium">{projectLocation || "your location"}</span>.</p>
              <div className="mt-4">
                <label className="text-sm">ICMP area</label>
                <select className="mt-1 w-full rounded-lg border px-3 py-2" value={icmpTemp} onChange={(e) => setIcmpTemp(e.target.value)}>
                  {icmpOptions.map((opt) => (
                    <option key={opt} value={opt}>{opt}</option>
                  ))}
                </select>
              </div>
              <div className="mt-4 flex justify-end gap-2">
                <button className="rounded-lg border px-3 py-1" onClick={() => { setShowIcmpModal(false); setIcmpOptions([]); }}>Cancel</button>
                <button className="rounded-lg border px-3 py-1" onClick={() => { setInferredICMP(icmpTemp); setShowIcmpModal(false); setIcmpOptions([]); }}>Use selected ICMP</button>
              </div>
            </motion.div>
          </div>
        )}

        {/* Upload */}
        <div className="mt-6 rounded-2xl border border-dashed p-6">
          <div className="flex items-center gap-3">
            <div className="mx-auto w-fit rounded-full bg-gray-100 p-3">
              <Upload className="h-6 w-6" />
            </div>
            <div>
              <div className="font-medium">Upload technical documents</div>
              <div className="text-xs text-gray-500">PDF, DOCX, XLSX - EMPs, CMPs, ecology, hydrology, archaeology</div>
            </div>
          </div>
        </div>

        {/* Preview / Left: categories, Right: reports */}
        <div className="mt-10 grid grid-cols-1 gap-6 lg:grid-cols-5">
          <div className="lg:col-span-2 space-y-4">
            <h2 className="text-xl font-semibold">Categories Summary</h2>
            <div className="rounded-2xl border p-4 shadow-sm">
              <table className="w-full text-sm">
                <thead className="text-left text-gray-500">
                  <tr>
                    <th className="py-2">Category</th>
                    <th className="py-2">Key issue</th>
                    <th className="py-2">Policy check</th>
                  </tr>
                </thead>
                <tbody>
                  {sampleFindings.map((f, idx) => (
                    <tr key={idx} className="border-t">
                      <td className="py-2 font-medium">{f.category}</td>
                      <td className="py-2 pr-2">{f.issue}</td>
                      <td className="py-2 text-xs text-gray-600">{f.policyLinks.slice(0, 2).join("; ")}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div className="rounded-2xl border p-4 shadow-sm">
              <h3 className="font-semibold">Consent condition library (auto-suggest)</h3>
              <ul className="mt-2 list-disc pl-5 text-sm">
                {sampleFindings.flatMap((f) => f.consentClauses).slice(0, 3).map((c, i) => (
                  <li key={i} className="mb-1">{c}</li>
                ))}
                <li>Contractor EMS must include Te Ture Whaimana alignment statement and training module co-designed with mana whenua.</li>
              </ul>
            </div>

            {/* Effects tables and triggers per Category */}
            <div className="rounded-2xl border p-4 shadow-sm">
              <h3 className="font-semibold">Effects table (per Category)</h3>
              {sampleFindings.map((f, i) => (
                <div key={i} className="mt-4 rounded-xl border p-3">
                  <div className="font-medium">{f.category.toUpperCase()}</div>
                  <table className="mt-2 w-full text-sm">
                    <thead className="text-left text-gray-500">
                      <tr>
                        <th className="py-1">Cultural</th>
                        <th className="py-1">Social</th>
                        <th className="py-1">Environmental</th>
                        <th className="py-1">Spiritual</th>
                      </tr>
                    </thead>
                    <tbody>
                      <tr className="align-top">
                        <td className="py-2"><ul className="list-disc pl-4">{f.effects.cultural.map((e, idx) => (<li key={idx}>{e}</li>))}</ul></td>
                        <td className="py-2"><ul className="list-disc pl-4">{f.effects.social.map((e, idx) => (<li key={idx}>{e}</li>))}</ul></td>
                        <td className="py-2"><ul className="list-disc pl-4">{f.effects.environmental.map((e, idx) => (<li key={idx}>{e}</li>))}</ul></td>
                        <td className="py-2"><ul className="list-disc pl-4">{f.effects.spiritual.map((e, idx) => (<li key={idx}>{e}</li>))}</ul></td>
                      </tr>
                    </tbody>
                  </table>

                  <div className="mt-3 rounded-lg bg-gray-50 p-3 text-xs">
                    <div className="font-semibold">Monitoring triggers and actions</div>
                    <div><span className="font-medium">Metrics:</span> {f.triggers.metrics.join(", ")}</div>
                    <div><span className="font-medium">Baseline:</span> {f.triggers.baselines}</div>
                    <div><span className="font-medium">Thresholds:</span> {f.triggers.thresholds.join("; ")}</div>
                    <div><span className="font-medium">Actions:</span> {f.triggers.actions.join("; ")}</div>
                    <div><span className="font-medium">Reporting:</span> {f.triggers.reporting}</div>
                  </div>
                </div>
              ))}
            </div>

            {/* Evidence gallery with thumbnails and captions */}
            <div className="rounded-2xl border p-4 shadow-sm">
              <div className="flex items-center gap-2">
                <ImageIcon className="h-5 w-5" />
                <h3 className="font-semibold">Evidence (select figures to embed inline)</h3>
              </div>
              <div className="mt-3 grid grid-cols-2 gap-3 md:grid-cols-4">
                {figureGallery.map((fig) => (
                  <label key={fig.id} className={`cursor-pointer rounded-xl border p-2 text-xs ${fig.selected ? "ring-2 ring-black" : ""}`}>
                    <img src={fig.dataUrl} alt={fig.caption} className="aspect-[4/3] w-full rounded-md border bg-gray-100" />
                    <div className="mt-2 line-clamp-2">{fig.caption}</div>
                    <div className="mt-1 flex items-center gap-2">
                      <input
                        type="checkbox"
                        checked={fig.selected}
                        onChange={(e) => setFigureGallery((prev) => prev.map((g) => (g.id === fig.id ? { ...g, selected: e.target.checked } : g)))}
                      />
                      <span>Select</span>
                    </div>
                  </label>
                ))}
              </div>
              <div className="mt-2 text-xs text-gray-600">Selected: {selectedFigures.length}</div>
            </div>
          </div>

          {/* Right column: report previews and exports */}
          <div className="lg:col-span-3 space-y-6">
            <div className="rounded-2xl border p-4 shadow-sm">
              <div className="flex items-center justify-between">
                <h2 className="text-xl font-semibold">1) CIA - Mana Whenua Narrative (preview)</h2>
                <div className="flex items-center gap-2">
                  <button className="inline-flex items-center gap-2 rounded-xl border px-3 py-1 text-sm" onClick={() => setTestResult(runSelfChecks())} title="Run self-checks">
                    <Play className="h-4 w-4" /> Run checks
                  </button>
                  <button className="inline-flex items-center gap-2 rounded-xl border px-3 py-1 text-sm" onClick={() => exportNarrativeDocx("CIA_Mana_Whenua", manaWhenuaNarrative, selectedFigures)} title="Export as DOCX">
                    <FileDown className="h-4 w-4" /> Export DOCX
                  </button>
                </div>
              </div>
              {testResult && (
                <div className={`mt-3 rounded-lg p-2 text-xs ${testResult.includes("passed") ? "bg-green-50 text-green-800" : "bg-red-50 text-red-800"}`}>{testResult}</div>
              )}
              <pre className="mt-3 max-h-80 overflow-auto whitespace-pre-wrap text-sm">{manaWhenuaNarrative}\n\n## Selected Figures (inline)\n{selectedFigures.map((f) => `- ${f.caption}`).join("\n")}</pre>
            </div>

            <div className="rounded-2xl border p-4 shadow-sm">
              <div className="flex items-center justify-between">
                <h2 className="text-xl font-semibold">2) CIA - Council/Developer Narrative (preview)</h2>
                <button className="inline-flex items-center gap-2 rounded-xl border px-3 py-1 text-sm" onClick={() => exportNarrativeDocx("CIA_Council_Developer", councilNarrative, selectedFigures)} title="Export as DOCX">
                  <FileDown className="h-4 w-4" /> Export DOCX
                </button>
              </div>
              <pre className="mt-3 max-h-80 overflow-auto whitespace-pre-wrap text-sm">{councilNarrative}\n\n## Selected Figures (inline)\n{selectedFigures.map((f) => `- ${f.caption}`).join("\n")}</pre>
            </div>
          </div>
        </div>

        {/* How it works */}
        <div className="mt-12 rounded-2xl border p-6 shadow-sm">
          <h2 className="text-xl font-semibold">How the AI analysis will work (behind the scenes)</h2>
          <ol className="mt-3 list-decimal space-y-2 pl-6 text-sm">
            <li>Ingest and parse your PDFs/DOCs (OCR if needed) to extract structured text, tables, maps, and plan references.</li>
            <li>Identify technical topics (water, land, ecology, archaeology, traffic, noise) and map to Categories (wai, whenua, whakapapa, wh\u0101nau, mauri, wairua).</li>
            <li>Policy crosswalk: compare issues/mitigations to Te Ture Whaimana and Tai Tumu, Tai Pari, Tai Ao; for HCC also compare to He Pou Manawa Ora + Hamilton District Plan; for WDC compare to Waikato District Plan.</li>
            <li>Generate parallel narratives (25+ pages Standard): mana whenua voice (plain language) and council/developer voice (technical), with inline figures you select.</li>
            <li>Auto-suggest consent conditions and produce a Cultural Monitoring Programme with council-specific toggles and per-Category tasks.</li>
          </ol>
        </div>

        {/* Cultural Monitoring Programme (auto-generated) */}
        <div className="mt-12 rounded-2xl border p-6 shadow-sm bg-gray-50">
          <div className="flex items-center justify-between">
            <h2 className="text-xl font-semibold">Cultural Monitoring Programme</h2>
            <button className="rounded-xl border px-3 py-1 text-sm" onClick={() => exportMonitoringDocx(projectName, deriveMonitoringRows(sampleFindings, council, includeHPMO))} title="Export as DOCX">
              Export DOCX
            </button>
          </div>
          <p className="mt-2 text-sm text-gray-700">Automatically derived from identified issues, policies and triggers.</p>

          <table className="mt-4 w-full text-sm border">
            <thead className="bg-gray-100">
              <tr>
                <th className="p-2 border">Phase</th>
                <th className="p-2 border">Monitoring focus</th>
                <th className="p-2 border">Role of Cultural Monitor</th>
                <th className="p-2 border">Frequency</th>
              </tr>
            </thead>
            <tbody>
              {deriveMonitoringRows(sampleFindings, council, includeHPMO).map((r, i) => (
                <tr key={i}>
                  <td className="p-2 border">{r.phase}</td>
                  <td className="p-2 border">{r.focus}</td>
                  <td className="p-2 border">{r.role}</td>
                  <td className="p-2 border">{r.frequency}</td>
                </tr>
              ))}
            </tbody>
          </table>

          <div className="mt-6">
            <h3 className="font-semibold">Cultural Monitor - Job Description (plain language)</h3>
            <ul className="mt-2 text-sm list-disc pl-6">
              <li>Be our eyes and ears on site for mana whenua.</li>
              <li>If you see a cultural risk, you can ask for work to stop.</li>
              <li>Take clear notes and photos. Upload them to the CIA dashboard.</li>
              <li>Help builders understand our tikanga and why it matters.</li>
              <li>Tell the project lead and Council when there is a problem.</li>
            </ul>
          </div>
        </div>

        {/* Per-Category monitoring tasks */}
        <div className="mt-6 rounded-2xl border p-6 shadow-sm">
          <h3 className="text-lg font-semibold">Per-Category Monitoring Tasks</h3>
          {sampleFindings.map((f, i) => (
            <div key={i} className="mt-4 rounded-xl border p-3">
              <div className="font-medium">{f.category.toUpperCase()}</div>
              <ul className="mt-2 text-sm list-disc pl-6">
                {f.category === "wai" && (
                  <>
                    <li>Check water clarity and NTU each day of earthworks.</li>
                    <li>Watch for fish movement (e.g., tuna, \u012Bnanga) at the right seasons.</li>
                    <li>Log any discolouration or sheen with time and weather.</li>
                  </>
                )}
                {f.category === "whenua" && (
                  <>
                    <li>Check for signs of k\u014Diwi/taonga. Follow discovery protocol.</li>
                    <li>Watch stripping works and keep a safe buffer around risk areas.</li>
                  </>
                )}
                {f.category === "whakapapa" && (
                  <>
                    <li>Check fish passage is clear after works and during flows.</li>
                    <li>Walk planting areas; note survival and pest pressure.</li>
                  </>
                )}
                {f.category === "wh\u0101nau" && (
                  <>
                    <li>Check marae access and traffic plans match what was agreed.</li>
                    <li>Log noise issues and call the site contact if access is blocked.</li>
                  </>
                )}
                {f.category === "mauri" && (
                  <>
                    <li>During storms, verify controls are working (photos before/after).</li>
                    <li>Join post-event hui to decide what needs fixing.</li>
                  </>
                )}
                {f.category === "wairua" && (
                  <>
                    <li>Stand at agreed view points and compare to the pre-works photos.</li>
                    <li>Raise concerns early so designs can be adjusted.</li>
                  </>
                )}
              </ul>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}
